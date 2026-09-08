from __future__ import annotations

import unittest
import json
from pathlib import Path
from tempfile import TemporaryDirectory

from eacc_app.browser_automation import EAccountingBrowserService, LoginCredentials


class LoginCredentialSecurityTests(unittest.TestCase):
    def test_password_is_not_shown_by_credential_representation(self) -> None:
        credentials = LoginCredentials(user_id="N1102100", password="secret-password")

        self.assertIn("N1102100", repr(credentials))
        self.assertNotIn("secret-password", repr(credentials))

    def test_temporary_edge_profile_disables_password_saving(self) -> None:
        with TemporaryDirectory() as directory:
            EAccountingBrowserService._prepare_temporary_edge_profile(Path(directory))
            preference_path = Path(directory) / "Default" / "Preferences"
            preferences = json.loads(preference_path.read_text(encoding="utf-8"))

        self.assertFalse(preferences["credentials_enable_service"])
        self.assertFalse(preferences["profile"]["password_manager_enabled"])
        self.assertFalse(preferences["translate"]["enabled"])

    def test_recent_news_init_script_executes_immediately_on_next_document(self) -> None:
        class FakePage:
            script = ""

            def add_init_script(self, script: str) -> None:
                self.script = script

        page = FakePage()

        EAccountingBrowserService._install_recent_news_auto_close_before_login(page)

        self.assertIn("(() => {", page.script)
        self.assertIn("})();", page.script)
        self.assertIn("closeAlarmPopup", page.script)
        self.assertIn("#popbtnArea", page.script)

    def test_other_user_processing_match_requires_the_actual_lock_message(self) -> None:
        self.assertTrue(
            EAccountingBrowserService._is_other_user_processing(
                "이미 다른 사용자가 처리중입니다."
            )
        )
        self.assertFalse(
            EAccountingBrowserService._is_other_user_processing(
                "다른 사용자가 처리 중인 건이 아니므로 결재요청 가능합니다."
            )
        )

    def test_approval_window_dialog_is_accepted_and_recorded(self) -> None:
        class FakeDialog:
            message = "이미 다른 사용자가 처리중입니다."
            accepted = False

            def accept(self) -> None:
                self.accepted = True

        class FakePopup:
            def __init__(self) -> None:
                self.handler = None

            def on(self, event: str, handler) -> None:
                self.assertEqual(event, "dialog")
                self.handler = handler

            def assertEqual(self, left, right) -> None:
                assert left == right

        popup = FakePopup()
        messages: list[str] = []
        EAccountingBrowserService._install_approval_dialog_handler(popup, messages)
        dialog = FakeDialog()
        popup.handler(dialog)

        self.assertTrue(dialog.accepted)
        self.assertEqual(messages, [dialog.message])

    def test_other_user_alert_is_suppressed_before_approval_popup_loads(self) -> None:
        class FakeContext:
            script = ""

            def add_init_script(self, script: str) -> None:
                self.script = script

        context = FakeContext()
        EAccountingBrowserService._install_other_user_lock_alert_suppressor(context)

        self.assertIn("window.alert", context.script)
        self.assertIn("다른사용자", context.script)
        self.assertIn("__eaccOtherUserLockAlertMessage", context.script)

    def test_all_browser_contexts_are_used_for_popup_protection(self) -> None:
        class FakeBrowser:
            contexts = ("first", "second")

        class FakeContext:
            browser = FakeBrowser()

        self.assertEqual(
            EAccountingBrowserService._all_browser_contexts(FakeContext()),
            ("first", "second"),
        )

    def test_other_user_processing_alert_is_classified_as_an_exception(self) -> None:
        self.assertTrue(
            EAccountingBrowserService._is_other_user_processing(
                "이미 다른 사용자가 처리중입니다."
            )
        )
        self.assertFalse(EAccountingBrowserService._is_other_user_processing("결재요청 하시겠습니까?"))

    def test_login_returns_as_soon_as_eacc_icon_is_visible(self) -> None:
        """The e-Acc function is invoked without waiting for icon rendering."""

        class FakeLocator:
            def __init__(self, page, selector: str) -> None:
                self.page = page
                self.selector = selector

            def wait_for(self, **_kwargs) -> None:
                return None

            def count(self) -> int:
                if self.selector == "#user_id":
                    return 1  # Deliberately remains in the old PortalMain DOM.
                if "goLinkEaccounting" in self.selector:
                    return int(self.page.submitted)
                return 1

            def is_checked(self) -> bool:
                return False

            def is_visible(self) -> bool:
                return self.page.submitted

            def fill(self, value: str) -> None:
                self.page.filled.append((self.selector, value))

            def uncheck(self) -> None:
                return None

            def click(self) -> None:
                self.page.submitted = True

        class FakeEaccountingPage:
            def is_closed(self) -> bool:
                return False

        class FakePage:
            submitted = False

            def __init__(self) -> None:
                self.filled: list[tuple[str, str]] = []
                self.waits = 0
                self.context = None

            def locator(self, selector: str) -> FakeLocator:
                return FakeLocator(self, selector)

            def wait_for_timeout(self, _milliseconds: int) -> None:
                self.waits += 1

            def add_init_script(self, _script: str) -> None:
                return None

            def evaluate(self, script: str):
                if "typeof window.goLinkEaccounting" in script:
                    return self.submitted
                if "window.goLinkEaccounting" in script:
                    assert self.context is not None
                    self.context.pages.append(FakeEaccountingPage())
                return None

        class FakeContext:
            browser = None

            def __init__(self, page: FakePage) -> None:
                self.pages = [page]

        page = FakePage()
        context = FakeContext(page)
        page.context = context
        EAccountingBrowserService._login_to_i_net(
            page,
            LoginCredentials(user_id="N1102100", password="secret-password"),
            context,
        )

        self.assertEqual(page.waits, 0)


if __name__ == "__main__":
    unittest.main()
