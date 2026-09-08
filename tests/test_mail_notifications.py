from __future__ import annotations

import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from eacc_app.mail_notifications import UnprocessedCardUse, render_mail_html
from eacc_app.storage import ImportRepository
from eacc_app.unprocessed_card import unprocessed_use_from_grid_row


class MailNotificationTests(unittest.TestCase):
    def test_html_contains_only_the_approved_columns(self) -> None:
        use = UnprocessedCardUse("one", "홍길동", "경영지원팀", "00123456", "2026-09-08 09:30", "테스트가맹점", Decimal("3300"))
        html = render_mail_html("홍길동", (use,), "2026-09-08")
        for label in ("순번", "승인번호", "이용일자", "가맹점", "사용금액"):
            self.assertIn(label, html)
        self.assertNotIn("카드번호", html)
        self.assertNotIn("첨부파일", html)

    def test_grid_parser_preserves_five_mail_fields(self) -> None:
        parsed = unprocessed_use_from_grid_row(
            1,
            ["성명", "부서명", "승인번호", "이용일자", "이용시간", "상호", "사용금액"],
            ["홍길동", "경영지원팀", "00123456", "20260908", "09:30:00", "테스트가맹점", "3,300"],
        )
        self.assertEqual("2026-09-08 09:30:00", parsed.usage_datetime)
        self.assertEqual(Decimal("3300"), parsed.amount)

    def test_mail_log_masks_address(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            repository = ImportRepository(Path(directory) / "test.db")
            item = repository.record_mail_log(
                recipient_name="홍길동", recipient_email="hong.gildong@example.com",
                department="경영지원팀", transaction_ids=("one",), subject="제목",
                status="발송 완료", reason="Outlook 발송 요청 완료",
            )
            self.assertNotIn("hong.gildong", item.recipient_email)
            self.assertEqual("발송 완료", repository.recent_mail_logs()[0].status)


if __name__ == "__main__":
    unittest.main()
