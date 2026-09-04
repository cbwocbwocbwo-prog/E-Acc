from __future__ import annotations

import os
import queue
import shutil
import socket
import subprocess
import threading
import tempfile
import time
import ctypes
import json
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable
from urllib.parse import parse_qsl, urlencode, urljoin, urlparse, urlsplit, urlunsplit
from urllib.request import urlopen

from .models import ReceiptImageResult, UnsubmittedTransaction
from .parser import GridRowFormatError, transaction_from_grid_row


I_NET_URL = "https://i-net.skons.co.kr/"
E_ACCOUNTING_URL = "https://eaccounting.skons.co.kr/"
E_ACCOUNTING_HOST = "eaccounting.skons.co.kr"
# i-NET의 e-Acc 바로가기 실제 클릭 대상. 링크 전체의 애니메이션·안정화 대기 대신
# 사용자에게 보이는 램프 아이콘 영역을 즉시 클릭한다.
EACC_ICON_SELECTOR = (
    "a[href='javascript:goLinkEaccounting()'] "
    "div.mb-1:has(img[src*='/resources/images/front/icons/lamp.png'])"
)
EACC_FUNCTION_READY_SCRIPT = "() => typeof window.goLinkEaccounting === 'function'"


class BrowserAutomationError(RuntimeError):
    """Base exception for the controlled Edge workflow."""


class BrowserLoginRequired(BrowserAutomationError):
    """Raised when the user must complete the temporary manual login step."""


class NoUnsubmittedTransactions(BrowserAutomationError):
    """Raised when the e-Accounting search completed normally with zero rows."""


class ReceiptNotAvailable(BrowserAutomationError):
    """Raised when the selected transaction does not have a usable receipt image."""


class ActualMerchantRegistrationError(BrowserAutomationError):
    """Raised when a resolved PG merchant cannot be registered safely."""


@dataclass(frozen=True, slots=True)
class LoginCredentials:
    """Credentials supplied for one automation request only.

    Instances are intentionally never written to a file, database, browser
    profile preference, log, or error message.  The temporary Edge profile is
    discarded when this application closes; its authenticated cookies are only
    used to reuse the active session during the current run.
    """

    user_id: str
    password: str = field(repr=False)


@dataclass(frozen=True, slots=True)
class _CollectCommand:
    credentials: LoginCredentials | None
    on_success: Callable[[Path], None]
    on_error: Callable[[Exception], None]


@dataclass(frozen=True, slots=True)
class _CurrentTargetCommand:
    credentials: LoginCredentials | None
    on_success: Callable[[UnsubmittedTransaction], None]
    on_error: Callable[[Exception], None]


@dataclass(frozen=True, slots=True)
class _ReceiptCommand:
    transaction: UnsubmittedTransaction
    credentials: LoginCredentials | None
    on_success: Callable[[ReceiptImageResult], None]
    on_error: Callable[[Exception], None]


@dataclass(frozen=True, slots=True)
class _ActualMerchantRegistrationCommand:
    transaction: UnsubmittedTransaction
    business_number: str
    merchant_name: str
    credentials: LoginCredentials | None
    on_success: Callable[[str], None]
    on_error: Callable[[Exception], None]


@dataclass(frozen=True, slots=True)
class _OpenApprovalLineCommand:
    """Open the approval-line popup for exactly one already-validated row.

    This command deliberately ends before the popup's ``결재요청`` action.
    """

    transaction: UnsubmittedTransaction
    credentials: LoginCredentials | None
    on_success: Callable[[str], None]
    on_error: Callable[[Exception], None]


def original_image_url(base_url: str, image_source: str) -> tuple[str, str, str]:
    """Return the authenticated original-image URL, DocIRN and CorpNo."""
    absolute = urljoin(base_url, image_source)
    parts = urlsplit(absolute)
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    # OfficeViewer가 캐시 무효화 값을 두 번째 '?' 뒤에 붙이는 경우가 있다.
    corp_no = query.get("CorpNo", "").split("?", 1)[0]
    doc_irn = query.get("DocIRN", "")
    if not doc_irn or not corp_no or not query.get("UserID"):
        raise ReceiptNotAvailable("영수증 주소에서 DocIRN, UserID 또는 CorpNo를 찾지 못했습니다.")
    query["ImgType"] = "original"
    query["CorpNo"] = corp_no
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query), "")), doc_irn, corp_no


class EAccountingBrowserService:
    """Own Playwright and Edge on one long-lived worker thread."""

    def __init__(self, download_directory: str | Path) -> None:
        self.download_directory = Path(download_directory)
        self.download_directory.mkdir(parents=True, exist_ok=True)
        self._commands: queue.Queue[
            _CollectCommand
            | _CurrentTargetCommand
            | _ReceiptCommand
            | _ActualMerchantRegistrationCommand
            | _OpenApprovalLineCommand
            | None
        ] = queue.Queue()
        self._owned_edge_process: subprocess.Popen | None = None
        self._thread = threading.Thread(
            target=self._worker,
            name="eaccounting-browser-worker",
            daemon=True,
        )
        self._thread.start()

    def collect_unsubmitted(
        self,
        credentials: LoginCredentials | None,
        on_success: Callable[[Path], None],
        on_error: Callable[[Exception], None],
    ) -> None:
        self._commands.put(
            _CollectCommand(
                credentials=credentials,
                on_success=on_success,
                on_error=on_error,
            )
        )

    def read_current_first_target(
        self,
        credentials: LoginCredentials | None,
        on_success: Callable[[UnsubmittedTransaction], None],
        on_error: Callable[[Exception], None],
    ) -> None:
        """Read the current first e-Acc row without relying on Excel row positions."""
        self._commands.put(
            _CurrentTargetCommand(
                credentials=credentials,
                on_success=on_success,
                on_error=on_error,
            )
        )

    def download_receipt(
        self,
        transaction: UnsubmittedTransaction,
        credentials: LoginCredentials | None,
        on_success: Callable[[ReceiptImageResult], None],
        on_error: Callable[[Exception], None],
    ) -> None:
        self._commands.put(
            _ReceiptCommand(transaction, credentials, on_success, on_error)
        )

    def register_actual_merchant(
        self,
        transaction: UnsubmittedTransaction,
        business_number: str,
        merchant_name: str,
        credentials: LoginCredentials | None,
        on_success: Callable[[str], None],
        on_error: Callable[[Exception], None],
    ) -> None:
        """Register one user-confirmed merchant in the e-Accounting popup."""
        self._commands.put(
            _ActualMerchantRegistrationCommand(
                transaction=transaction,
                business_number=business_number,
                merchant_name=merchant_name,
                credentials=credentials,
                on_success=on_success,
                on_error=on_error,
            )
        )

    def open_approval_line(
        self,
        transaction: UnsubmittedTransaction,
        credentials: LoginCredentials | None,
        on_success: Callable[[str], None],
        on_error: Callable[[Exception], None],
    ) -> None:
        """Select one row and open its approval-line popup without submitting."""
        self._commands.put(
            _OpenApprovalLineCommand(
                transaction=transaction,
                credentials=credentials,
                on_success=on_success,
                on_error=on_error,
            )
        )

    def close(self) -> None:
        self._commands.put(None)
        # 종료 후 이전 계정 세션이 새 실행에서 재사용되지 않도록 자체 실행한
        # 자동화용 Edge를 종료한다. 외부 개발용 Edge에는 영향을 주지 않는다.
        self._terminate_owned_edge_process()
        self._thread.join(timeout=5)

    def _terminate_owned_edge_process(self) -> None:
        edge_process = self._owned_edge_process
        if edge_process is None or edge_process.poll() is not None:
            return
        if os.name == "nt":
            # Edge는 브라우저·렌더러·GPU 하위 프로세스를 따로 두므로 부모만
            # 종료하면 임시 프로필의 쿠키 파일이 잠길 수 있다. 이 서비스가 시작한
            # 정확한 PID의 프로세스 트리만 종료한다.
            subprocess.run(
                ("taskkill", "/PID", str(edge_process.pid), "/T", "/F"),
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
            )
        else:
            edge_process.terminate()
        try:
            edge_process.wait(timeout=3)
        except subprocess.TimeoutExpired:
            edge_process.kill()

    def _worker(self) -> None:
        profile_directory = edge_process = None
        debugging_port = None
        command_playwright = None
        browser = None
        try:
            while True:
                command = self._commands.get()
                if command is None:
                    break
                cdp_session = None
                try:
                    if debugging_port is None:
                        from playwright.sync_api import sync_playwright

                        # 기존 사용자의 Edge에 연결하는 기능은 개발 검증용으로만 둔다.
                        # 일반 실행은 항상 새 임시 프로필을 사용해야 계정이 섞이지 않는다.
                        test_port = (
                            os.environ.get("EACC_DEBUG_PORT", "").strip()
                            if os.environ.get("EACC_ENABLE_DEBUG_PORT") == "1"
                            else ""
                        )
                        if test_port:
                            debugging_port = int(test_port)
                            self._wait_for_debugging_endpoint(
                                debugging_port, timeout_seconds=3
                            )
                        else:
                            profile_directory = tempfile.TemporaryDirectory(
                                prefix="eacc_edge_profile_"
                            )
                            self._prepare_temporary_edge_profile(Path(profile_directory.name))
                            debugging_port = self._find_free_port()
                            edge_process = subprocess.Popen(
                                (
                                    str(self._edge_executable()),
                                    f"--remote-debugging-port={debugging_port}",
                                    f"--user-data-dir={profile_directory.name}",
                                    "--no-first-run",
                                    "--no-default-browser-check",
                                    # 사용자가 로그인·목록 상태를 확인할 수 있는 기본
                                    # 자동화 Edge 창은 항상 최대화해 연다. 영수증·결재선
                                    # 팝업의 표시 방식에는 영향을 주지 않는다.
                                    "--start-maximized",
                                    "--disable-features=msEdgeSidebarV2",
                                    "--disable-save-password-bubble",
                                    "--disable-features=Translate,TranslateUI,msEdgeTranslate",
                                    I_NET_URL,
                                ),
                            )
                            self._owned_edge_process = edge_process
                            self._wait_for_debugging_endpoint(debugging_port)

                    # 영수증 다운로드 → PG 조회 → 실구매처 등록은 같은 작업 스레드에서
                    # 연속 실행된다. Sync Playwright의 이벤트 루프를 명령마다 재생성하면
                    # 세 번째 작업부터 "event loop is already running" 충돌이 날 수 있어,
                    # 프로그램 종료 시까지 한 번만 생성·유지한다.
                    if command_playwright is None or browser is None or not browser.is_connected():
                        from playwright.sync_api import sync_playwright

                        if command_playwright is None:
                            command_playwright = sync_playwright().start()
                        browser = command_playwright.chromium.connect_over_cdp(
                            f"http://127.0.0.1:{debugging_port}"
                        )
                    if not browser.contexts:
                        raise BrowserAutomationError(
                            "자동화용 Edge의 브라우저 컨텍스트를 찾지 못했습니다."
                        )
                    context = browser.contexts[0]
                    context.set_default_timeout(15_000)
                    cdp_session = browser.new_browser_cdp_session()
                    cdp_session.send(
                        "Browser.setDownloadBehavior",
                        {
                            "behavior": "allow",
                            "downloadPath": str(self.download_directory),
                            "eventsEnabled": True,
                        },
                    )
                    if isinstance(command, _CollectCommand):
                        result = self._collect(context, command.credentials)
                    elif isinstance(command, _CurrentTargetCommand):
                        result = self._read_current_first_target(context, command.credentials)
                    elif isinstance(command, _ReceiptCommand):
                        result = self._download_receipt(
                            context,
                            command.transaction,
                            command.credentials,
                        )
                    elif isinstance(command, _ActualMerchantRegistrationCommand):
                        result = self._register_actual_merchant(
                            context,
                            command.transaction,
                            command.business_number,
                            command.merchant_name,
                            command.credentials,
                        )
                    else:
                        result = self._open_approval_line(
                            context,
                            command.transaction,
                            command.credentials,
                        )
                except Exception as exc:
                    command.on_error(exc)
                else:
                    command.on_success(result)
                finally:
                    if cdp_session is not None:
                        try:
                            cdp_session.detach()
                        except Exception:
                            pass
                    # 큐에서 꺼낸 로그인 정보의 참조를 즉시 제거한다. 인증 쿠키만
                    # 임시 Edge 세션에 남아 다음 메뉴 작업에서 재로그인을 막는다.
                    command = None
        except Exception as exc:
            if "command" in locals() and command is not None:
                command.on_error(BrowserAutomationError(f"Edge 자동화 초기화 실패: {exc}"))
        finally:
            if command_playwright is not None:
                try:
                    command_playwright.stop()
                except Exception:
                    pass
            self._terminate_owned_edge_process()
            self._owned_edge_process = None
            if profile_directory is not None:
                try:
                    profile_directory.cleanup()
                except PermissionError:
                    # taskkill 직후 Edge 하위 프로세스의 파일 잠금 해제가 조금 늦는
                    # 경우 한 번만 재시도한다. 실패해도 프로필은 무작위 임시 경로다.
                    time.sleep(0.5)
                    profile_directory.cleanup()

    @staticmethod
    def _edge_executable() -> Path:
        candidates = (
            Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
            Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
        )
        for candidate in candidates:
            if candidate.is_file():
                return candidate
        raise BrowserAutomationError("Microsoft Edge 실행 파일을 찾지 못했습니다.")

    @staticmethod
    def _prepare_temporary_edge_profile(profile_directory: Path) -> None:
        """Disable Chromium password saving in this disposable automation profile."""
        default_directory = profile_directory / "Default"
        default_directory.mkdir(parents=True, exist_ok=True)
        # Edge uses Chromium's profile preferences.  This prevents its native
        # '암호를 저장하시겠습니까?' bubble as well as any password persistence.
        (default_directory / "Preferences").write_text(
            json.dumps(
                {
                    "credentials_enable_service": False,
                    "profile": {"password_manager_enabled": False},
                    "translate": {"enabled": False},
                }
            ),
            encoding="utf-8",
        )

    @staticmethod
    def _find_free_port() -> int:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as listener:
            listener.bind(("127.0.0.1", 0))
            return int(listener.getsockname()[1])

    @staticmethod
    def _wait_for_debugging_endpoint(port: int, timeout_seconds: float = 15) -> None:
        deadline = time.monotonic() + timeout_seconds
        endpoint = f"http://127.0.0.1:{port}/json/version"
        while time.monotonic() < deadline:
            try:
                with urlopen(endpoint, timeout=1) as response:
                    if response.status == 200:
                        return
            except OSError:
                time.sleep(0.2)
        raise BrowserAutomationError("자동화용 Edge 디버깅 연결 시간이 초과되었습니다.")

    def _ensure_eaccounting_page(self, context, credentials: LoginCredentials | None):
        """Return an authenticated e-Accounting page, logging in only when needed."""
        page = self._find_eaccounting_page(context)
        if page is not None:
            return page

        page = context.pages[0] if context.pages else context.new_page()
        if not page.url or page.url == "about:blank":
            page.goto(I_NET_URL, wait_until="domcontentloaded")

        # A newly-created Edge profile starts on i-NET.  An existing expired
        # e-Accounting session can also be redirected there, so use the same
        # one-time credentials in both cases.  When no credentials were
        # supplied, do not attempt a login or retain anything for later.
        if page.locator("#user_id").count() or "i-net.skons.co.kr" in page.url:
            if credentials is None or not credentials.user_id.strip() or not credentials.password:
                page.bring_to_front()
                raise BrowserLoginRequired(
                    "i-NET 로그인이 필요합니다. 프로그램 상단에 아이디와 비밀번호를 입력한 뒤 "
                    "원하는 작업 버튼을 다시 눌러 주세요."
                )
            self._login_to_i_net(page, credentials)
            # _login_to_i_net은 PortalMain의 e-Acc 램프 아이콘이 보이는 순간
            # 반환한다. 그 뒤에 팝업/포털 준비상태를 다시 점검하면, 팝업이 없는
            # 경우에도 사용자 화면상 10초 이상 멈춘 것처럼 보일 수 있다. 최근소식
            # 닫기는 로그인 전에 설치한 DOM 감시기가 맡고, 아래에서 램프 아이콘을
            # 즉시 강제 클릭한다.

        # i-NET의 e-Acc 아이콘을 반드시 먼저 누른다. 단순히 e-Accounting 주소로
        # 이동하면 i-NET가 발급하는 SSO 연결 정보가 적용되지 않는 사례가 있다.
        # 아이콘은 별도 Edge 창을 열 수 있으므로 새 페이지를 우선 찾고, Playwright가
        # 새 창을 수집하지 못한 경우에만 클릭 이후의 기존 탭을 보조로 사용한다.
        try:
            eacc_icon = page.locator(EACC_ICON_SELECTOR)
            try:
                function_ready = bool(page.evaluate(EACC_FUNCTION_READY_SCRIPT))
            except Exception:
                function_ready = False
            if not function_ready:
                # 구형 포털 등에서 함수가 준비되지 않은 경우에만 화면 아이콘을
                # 보조 경로로 쓴다. 일반 경로에서는 이미지/CSS 표시를 기다리지 않는다.
                eacc_icon.wait_for(state="visible", timeout=15_000)

            # goLinkEaccounting()은 새 창을 만들고 처음에는 about:blank 상태로
            # 둔 뒤 SSO 주소를 채운다. 기존 코드는 URL이 e-Acc로 바뀔 때까지
            # 포털을 보고 있었으므로 사용자가 15~20초간 전환되지 않는 것으로
            # 보게 됐다. 새 창 생성 이벤트 자체를 기다려 즉시 앞으로 가져온다.
            eaccounting_page = None
            try:
                with context.expect_page(timeout=5_000) as new_page_info:
                    if function_ready:
                        # i-NET의 실제 e-Acc 실행 함수다. 아이콘의 렌더링 완료나
                        # Playwright의 요소 안정화 대기 없이 SSO 새 창을 즉시 연다.
                        page.evaluate("() => window.goLinkEaccounting()")
                    else:
                        # 함수 확인이 불가능한 예외 환경에서만 기존 아이콘 클릭으로
                        # 호환한다. force=True는 포털 카드 애니메이션 대기를 건너뛴다.
                        eacc_icon.click(force=True, no_wait_after=True, timeout=2_000)
                eaccounting_page = new_page_info.value
            except Exception:
                # 일부 Edge 환경에서는 window.open 이벤트가 CDP에 전달되지 않는다.
                # 이 경우에는 이미 열린 페이지를 짧게만 탐색한다. 클릭은 반복하지
                # 않아 중복 e-Acc 창을 만들지 않는다.
                deadline = time.monotonic() + 3
                while time.monotonic() < deadline:
                    eaccounting_page = self._find_eaccounting_page(context)
                    if eaccounting_page is not None:
                        break
                    page.wait_for_timeout(100)

            if eaccounting_page is not None:
                # e-Acc의 프레임과 SSO 화면은 이 뒤에도 계속 로딩되지만, 창 전환은
                # 지체하지 않는다. 프레임 준비 대기는 백그라운드에서 이어진다.
                eaccounting_page.bring_to_front()
                self._wait_for_eaccounting_shell(eaccounting_page)
                return eaccounting_page

            # 일부 Edge 환경에서는 goLinkEaccounting()이 운영체제 수준의 새 창을
            # 열어 CDP 페이지 목록에서 즉시 보이지 않는다. 아이콘 클릭으로 SSO가
            # 실행된 뒤에만 기존 탭을 보조 경로로 이동시킨다.
            page.goto(E_ACCOUNTING_URL, wait_until="domcontentloaded", timeout=20_000)
            page.bring_to_front()
            self._wait_for_eaccounting_shell(page)
            return page
        except Exception as exc:
            try:
                page.goto(I_NET_URL, wait_until="domcontentloaded", timeout=10_000)
            except Exception:
                pass
            page.bring_to_front()
            raise BrowserLoginRequired(
                "i-NET 로그인 세션을 확인하지 못했습니다. 아이디와 비밀번호를 다시 입력해 주세요."
            ) from exc

    @staticmethod
    def _login_to_i_net(page, credentials: LoginCredentials) -> None:
        """Submit the i-NET login form without enabling its ID-save cookie."""
        try:
            page.locator("#user_id").wait_for(state="visible", timeout=15_000)
            # i-NET의 로그인 화면은 기본으로 '아이디 저장'이 선택돼 있다. 해제하지
            # 않으면 웹사이트 쿠키에 아이디가 남으므로 자동화에서는 항상 해제한다.
            save_id = page.locator("#saveMail")
            if save_id.count() and save_id.is_checked():
                save_id.uncheck()
            page.locator("#user_id").fill(credentials.user_id.strip())
            page.locator("#user_pass").fill(credentials.password)
            # 로그인 화면이 SPA 방식으로 PortalMain 내용만 교체하는 경우에도
            # 감시기가 즉시 동작하도록 현재 문서에 먼저 설치한다.
            EAccountingBrowserService._install_recent_news_auto_close(page)
            # PortalMain 문서가 생성되는 순간부터 최근소식 팝업을 감시한다.
            # 전체 DOM 로딩 완료보다 먼저 작동하도록 로그인 제출 전에 등록한다.
            EAccountingBrowserService._install_recent_news_auto_close_before_login(page)
            page.locator("button[type='submit']").click()

            deadline = time.monotonic() + 25
            eacc_link = page.locator(EACC_ICON_SELECTOR)
            while time.monotonic() < deadline:
                # 사용자 눈에 보이는 아이콘의 CSS/이미지 준비보다 실제 e-Acc 실행
                # 함수가 먼저 준비될 수 있다. 함수가 생기는 순간 바로 다음 단계에서
                # 실행하도록 반환한다.
                try:
                    if page.evaluate(EACC_FUNCTION_READY_SCRIPT):
                        return
                    if eacc_link.count() and eacc_link.is_visible():
                        return
                except Exception:
                    # SPA가 PortalMain DOM을 교체하는 중이면 다음 짧은 주기에 재확인한다.
                    pass

                # 아이콘 HTML이 없는 이전 포털 버전의 호환용 보조 조건이다.
                if not page.locator("#user_id").count():
                    return
                page.wait_for_timeout(200)
        except BrowserAutomationError:
            raise
        except Exception as exc:
            raise BrowserLoginRequired(
                "i-NET 로그인 처리 중 오류가 발생했습니다. 아이디와 비밀번호를 확인해 주세요."
            ) from exc
        finally:
            # 실패한 로그인 화면에 비밀번호가 남아 있지 않도록 제거한다. 성공 후에는
            # 해당 요소가 없어 무시된다.
            try:
                page.locator("#user_pass").fill("")
            except Exception:
                pass
        raise BrowserLoginRequired(
            "i-NET 로그인이 완료되지 않았습니다. 아이디·비밀번호 또는 계정 상태를 확인해 주세요."
        )

    @staticmethod
    def _close_i_net_popups(context, main_page) -> None:
        """Close post-login i-NET popups before opening e-Accounting."""
        for candidate in tuple(context.pages):
            if candidate is main_page or candidate.is_closed():
                continue
            hostname = (urlparse(candidate.url).hostname or "").lower().rstrip(".")
            if hostname == "i-net.skons.co.kr":
                try:
                    candidate.close()
                except Exception:
                    pass

    @staticmethod
    def _wait_for_portal_ready(page) -> None:
        """Wait for e-Acc while repeatedly closing the late recent-news layer."""
        link = page.locator(EACC_ICON_SELECTOR)
        deadline = time.monotonic() + 15
        while time.monotonic() < deadline:
            # 최근소식은 e-Acc 아이콘보다 늦게 삽입될 수 있다. 단순 wait_for는
            # 이 동안 JavaScript를 다시 실행하지 않으므로, 짧은 주기로 명시적으로
            # popbtnArea의 닫기 함수를 호출한다.
            EAccountingBrowserService._close_recent_news_layer(page)
            try:
                if link.is_visible():
                    return
            except Exception:
                pass
            page.wait_for_timeout(50)
        raise BrowserAutomationError("i-NET 포털에서 e-Acc 아이콘을 찾지 못했습니다.")

    @staticmethod
    def _has_recent_news_layer(page) -> bool:
        """Return immediately when the optional in-page recent-news layer is absent."""
        try:
            return bool(
                page.locator("#alarmOverlayWrap").is_visible()
            )
        except Exception:
            return False

    @staticmethod
    def _install_recent_news_auto_close(page) -> None:
        """Close the optional recent-news layer immediately, including late DWR inserts."""
        try:
            page.evaluate(
                """
                () => {
                    const closeRecentNews = () => {
                        // i-NET 화면의 실제 닫기 HTML은
                        // onclick="javascript:closeAlarmPopup();" 이다. 외부
                        // 컨테이너 ID 변화와 관계없이 이 함수를 우선 호출한다.
                        const close = document.querySelector(
                            "#popbtnArea img[onclick*='closeAlarmPopup'], img[onclick*='closeAlarmPopup']"
                        );
                        if (!(close instanceof HTMLElement)) return;
                        if (typeof window.closeAlarmPopup === 'function') {
                            try { window.closeAlarmPopup(); return; } catch (_) {}
                        }
                        close.click();
                    };
                    window.__eaccRecentNewsObserver?.disconnect();
                    const observer = new MutationObserver(closeRecentNews);
                    observer.observe(document.documentElement, {
                        childList: true,
                        subtree: true,
                        attributes: true,
                        attributeFilter: ['class', 'style'],
                    });
                    window.__eaccRecentNewsObserver = observer;
                    // PortalMain은 일부 환경에서 팝업의 HTML을 먼저 넣고
                    // closeAlarmPopup 함수를 나중에 준비한다. DOM 변동 한 번만
                    // 감시하면 그 사이를 놓칠 수 있어 20ms 재시도를 짧게 둔다.
                    window.clearInterval(window.__eaccRecentNewsCloserTimer);
                    let retryCount = 0;
                    window.__eaccRecentNewsCloserTimer = window.setInterval(() => {
                        closeRecentNews();
                        retryCount += 1;
                        if (retryCount >= 750) {
                            window.clearInterval(window.__eaccRecentNewsCloserTimer);
                        }
                    }, 20);
                    closeRecentNews();
                }
                """
            )
        except Exception:
            # 포털이 비정상 종료된 경우에는 이후 e-Acc 진입 단계에서 명확한 오류를 낸다.
            return

    @staticmethod
    def _install_recent_news_auto_close_before_login(page) -> None:
        """Install the closer in the next PortalMain document before it starts loading."""
        page.add_init_script(
            """
            (() => {
                const closeRecentNews = () => {
                    const close = document.querySelector(
                        "#popbtnArea img[onclick*='closeAlarmPopup'], img[onclick*='closeAlarmPopup']"
                    );
                    if (!(close instanceof HTMLElement)) return;
                    if (typeof window.closeAlarmPopup === 'function') {
                        try { window.closeAlarmPopup(); return; } catch (_) {}
                    }
                    close.click();
                };
                const observer = new MutationObserver(closeRecentNews);
                observer.observe(document.documentElement || document, {
                    childList: true,
                    subtree: true,
                    attributes: true,
                    attributeFilter: ['class', 'style'],
                });
                window.__eaccRecentNewsObserver = observer;
                let retryCount = 0;
                const timer = window.setInterval(() => {
                    closeRecentNews();
                    retryCount += 1;
                    if (retryCount >= 750) window.clearInterval(timer);
                }, 20);
                closeRecentNews();
            })();
            """
        )

    @staticmethod
    def _close_recent_news_layer(page) -> None:
        """Close i-NET's in-page '최근소식' modal before clicking e-Acc.

        The recent-news layer is not a separate browser window, so closing
        popup pages alone does not remove it.  i-NET revisions have changed
        its element IDs, therefore this targets the visible heading and the
        close control inside the same dialog rather than a fragile fixed ID.
        """
        try:
            page.evaluate(
                """
                () => {
                    const close = document.querySelector(
                        "#popbtnArea img[onclick*='closeAlarmPopup'], img[onclick*='closeAlarmPopup']"
                    );
                    if (!(close instanceof HTMLElement)) return false;
                    if (typeof window.closeAlarmPopup === 'function') {
                        window.closeAlarmPopup();
                        return true;
                    }
                    close.click();
                    return true;
                }
                """
            )
        except Exception:
            # 팝업 구조가 바뀌어도 로그인 전체를 지연·실패시키지 않는다. e-Acc
            # 아이콘은 아래에서 DOM 클릭으로 실행해 화면 레이어의 영향을 받지 않는다.
            return

    def _collect(self, context, credentials: LoginCredentials | None) -> Path:
        page = self._ensure_eaccounting_page(context, credentials)

        page.bring_to_front()
        self._open_card_processing_top_menu(page)
        self._open_unsubmitted_menu(page)
        main_frame = self._wait_for_frame(page, "mainFrame")

        card_select = main_frame.locator("select#CARD_NO")
        card_select.wait_for(state="visible", timeout=15_000)
        card_select.select_option(value="C")
        main_frame.evaluate(
            """
            () => {
                window.__eaccQueryDone = false;
                GridObj.attachEvent('onXLE', () => {
                    window.__eaccQueryDone = true;
                    return true;
                });
                doQuery();
            }
            """
        )
        main_frame.wait_for_function(
            "() => window.__eaccQueryDone === true",
            timeout=20_000,
        )
        main_frame.get_by_text("총 :", exact=False).first.wait_for(state="visible", timeout=20_000)
        row_count = int(main_frame.evaluate("() => GridObj.getRowsNum()") or 0)
        if row_count == 0:
            # 검색 0건은 정상 업무 결과다. 엑셀 저장 이벤트가 발생하지 않으므로
            # 다운로드 대기를 시작하지 않고 UI에 빈 결과로 알린다.
            raise NoUnsubmittedTransactions("미상신내역 검색 결과가 없습니다.")

        files_before = {
            path.name: (path.stat().st_mtime_ns, path.stat().st_size)
            for path in self.download_directory.glob("*.xls")
        }
        with page.expect_download(timeout=20_000) as download_info:
            main_frame.evaluate(
                """
                () => {
                    GridObj.contextID = GridObj.contextID || '0_0';
                    doOnMouseRightButtonClick('ExcelSave');
                }
                """
            )
        download = download_info.value

        downloaded_source = self._wait_for_downloaded_file(
            download.suggested_filename,
            files_before,
        )
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        destination = self.download_directory / f"미상신내역_{timestamp}.xls"
        shutil.copy2(downloaded_source, destination)
        if not destination.is_file() or destination.stat().st_size == 0:
            raise BrowserAutomationError("미상신내역 파일 다운로드 결과를 확인할 수 없습니다.")
        return destination

    def _read_current_first_target(
        self,
        context,
        credentials: LoginCredentials | None,
    ) -> UnsubmittedTransaction:
        """Search by cost center and convert the current first grid row safely."""
        page = self._ensure_eaccounting_page(context, credentials)
        page.bring_to_front()
        self._open_card_processing_top_menu(page)
        self._open_unsubmitted_menu(page)
        main_frame = self._wait_for_frame(page, "mainFrame")
        self._query_unsubmitted(main_frame)
        row_count = int(main_frame.evaluate("() => GridObj.getRowsNum()") or 0)
        if row_count == 0:
            raise NoUnsubmittedTransactions("미상신내역 검색 결과가 없습니다.")
        grid_row = main_frame.evaluate(
            """
            () => {
                const rowId = GridObj.getAllRowIds().split(',').filter(Boolean)[0];
                if (!rowId) return null;
                const count = GridObj.getColumnsNum();
                const headers = [];
                const values = [];
                for (let index = 0; index < count; index += 1) {
                    const technicalId = String(GridObj.getColumnId(index) ?? '');
                    const label = String(GridObj.getColLabel(index) ?? '');
                    headers.push(label.trim() || technicalId);
                    values.push(String(GridObj.cells(rowId, index).getValue() ?? '').trim());
                }
                return {headers, values};
            }
            """
        )
        if not grid_row:
            raise BrowserAutomationError("e-Acc 검색 결과의 첫 행을 읽지 못했습니다.")
        try:
            return transaction_from_grid_row(1, grid_row["headers"], grid_row["values"])
        except GridRowFormatError as exc:
            raise BrowserAutomationError(f"e-Acc 첫 행 형식을 해석하지 못했습니다: {exc}") from exc

    def _download_receipt(
        self,
        context,
        transaction: UnsubmittedTransaction,
        credentials: LoginCredentials | None,
    ) -> ReceiptImageResult:
        if transaction.evidence_status != "#":
            raise ReceiptNotAvailable(
                f"증빙유무가 '{transaction.evidence_status}'이므로 첨부 영수증이 없는 행입니다."
            )

        page = self._ensure_eaccounting_page(context, credentials)
        self._open_card_processing_top_menu(page)
        self._open_unsubmitted_menu(page)
        main_frame = self._wait_for_frame(page, "mainFrame")
        if main_frame.evaluate("() => GridObj.getRowsNum()") == 0:
            self._query_unsubmitted(main_frame)

        matches = main_frame.evaluate(
            r"""
            target => {
                const digits = value => String(value ?? '').replace(/\D/g, '');
                const amount = value => String(value ?? '').replace(/[^0-9-]/g, '');
                return GridObj.getAllRowIds().split(',').filter(Boolean).filter(rowId =>
                    String(GridObj.cells(rowId, GridObj.getColIndexById('CARD_NO')).getValue()) === target.cardNumber &&
                    String(GridObj.cells(rowId, GridObj.getColIndexById('APPR_NO')).getValue()) === target.approvalNumber &&
                    digits(GridObj.cells(rowId, GridObj.getColIndexById('BLDAT')).getValue()) === target.evidenceDate &&
                    amount(GridObj.cells(rowId, GridObj.getColIndexById('USED_AMT')).getValue()) === target.amount
                );
            }
            """,
            {
                "cardNumber": transaction.card_number,
                "approvalNumber": transaction.approval_number,
                "evidenceDate": transaction.evidence_date.replace("-", ""),
                "amount": str(transaction.amount),
            },
        )
        if len(matches) != 1:
            raise BrowserAutomationError(
                f"선택 거래와 일치하는 웹 화면 행이 {len(matches)}건입니다. 영수증을 안전하게 연결할 수 없습니다."
            )

        row_id = matches[0]
        popup = None
        popup_windows_before = self._edge_popup_window_handles()
        try:
            # 이름이 같은 이전 결재선 창은 새 선택 행과 섞이지 않도록 닫는다.
            for candidate in tuple(context.pages):
                if "approval_set_list_popup.jsp" in candidate.url:
                    candidate.close()
            main_frame.evaluate(
                """
                rowId => {
                    const selected = GridObj.getColIndexById('SELECTED');
                    GridObj.forEachRow(id => GridObj.cells(id, selected).setValue(0));
                    GridObj.cells(rowId, selected).setValue(1);
                }
                """,
                row_id,
            )
            with page.expect_popup(timeout=15_000) as popup_info:
                main_frame.evaluate("doApproval()")
            popup = popup_info.value
            # 영수증 팝업은 자동 처리 과정에서만 쓰므로 사용자 작업을 가리지 않게 한다.
            # 로그인과 목록 조회에 사용하는 Edge 본창은 그대로 둔다.
            self._hide_receipt_popup(popup, popup_windows_before)
            popup.wait_for_url("**/approval/approval_set_list_popup.jsp?**", timeout=15_000)
            image_frame = self._wait_for_frame(popup, "imageframe", timeout_ms=20_000)
            thumbs = image_frame.locator("img.thumb_img")
            try:
                # imageframe이 열린 뒤에도 썸네일이 없으면 실제 첨부가 없는 경우다.
                # 사용자에게 긴 대기시간을 만들지 않도록 5초 안에 구분한다.
                thumbs.first.wait_for(state="attached", timeout=5_000)
            except Exception as exc:
                # 이미지가 전혀 없는 경우 Playwright의 대기 시간초과를 그대로 노출하면
                # 행의 영수증 사유에 남길 수 없다. 업무상 '영수증 없음'으로 구분한다.
                raise ReceiptNotAvailable(
                    "결재선 지정 창에 첨부 영수증 이미지가 없습니다."
                ) from exc
            sources = thumbs.evaluate_all("images => images.map(image => image.getAttribute('src'))")
            if not sources:
                raise ReceiptNotAvailable("결재선 지정 창에서 영수증 이미지를 찾지 못했습니다.")

            receipt_directory = self.download_directory.parent / "receipts"
            receipt_directory.mkdir(parents=True, exist_ok=True)
            image_paths: list[Path] = []
            doc_irns: list[str] = []
            corp_no = ""
            for image_index, source in enumerate(sources):
                source_url, doc_irn, current_corp_no = original_image_url(image_frame.url, source)
                response = context.request.get(source_url, timeout=20_000)
                content_type = response.headers.get("content-type", "").lower()
                body = response.body()
                if not response.ok or not content_type.startswith("image/") or len(body) < 100:
                    raise ReceiptNotAvailable(
                        f"원본 영수증 다운로드 실패(HTTP {response.status}, {content_type or '형식 없음'})"
                    )
                extension = ".png" if "png" in content_type else ".jpg"
                destination = receipt_directory / (
                    f"{transaction.transaction_id}_{image_index}_{doc_irn}{extension}"
                )
                destination.write_bytes(body)
                image_paths.append(destination)
                doc_irns.append(doc_irn)
                corp_no = current_corp_no

            return ReceiptImageResult(
                transaction_id=transaction.transaction_id,
                image_paths=tuple(image_paths),
                doc_irns=tuple(doc_irns),
                corp_no=corp_no,
            )
        finally:
            try:
                main_frame.evaluate(
                    "rowId => GridObj.cells(rowId, GridObj.getColIndexById('SELECTED')).setValue(0)",
                    row_id,
                )
            except Exception:
                pass
            if popup is not None and not popup.is_closed():
                popup.close()

    def _register_actual_merchant(
        self,
        context,
        transaction: UnsubmittedTransaction,
        business_number: str,
        merchant_name: str,
        credentials: LoginCredentials | None,
    ) -> str:
        """Open e-Acc's real-merchant popup and submit one confirmed value.

        The caller has already checked the OCR business number, BizNo result
        and the user's explicit confirmation.  This method nevertheless refuses
        to overwrite a row that already has a real merchant name.
        """
        if not business_number.isdigit() or len(business_number) != 10:
            raise ActualMerchantRegistrationError("등록할 사업자번호는 하이픈 없는 10자리 숫자여야 합니다.")
        if not merchant_name.strip():
            raise ActualMerchantRegistrationError("등록할 사업자명이 비어 있습니다.")
        if transaction.actual_merchant_name.strip():
            raise ActualMerchantRegistrationError(
                f"이 행에는 이미 실구매처명 '{transaction.actual_merchant_name}'이 등록되어 있어 덮어쓰지 않습니다."
            )

        # e-Acc는 사용자가 선택한 한 행을 기준으로 처리한다. 프로그램이 그리드의
        # 위치·행 번호를 추측해 팝업을 열지 않고, 사용자가 e-Acc에서 실제 해당 행의
        # 돋보기를 열면 그 이미 열린 팝업에만 값을 자동 입력·등록한다.
        popup = next(
            (
                candidate
                for candidate in reversed(context.pages)
                if not candidate.is_closed()
                and "/account/real_store_mgt.jsp" in candidate.url
            ),
            None,
        )
        if popup is None:
            raise ActualMerchantRegistrationError(
                "e-Acc에서 현재 처리할 한 행의 '선택' 돋보기를 더블클릭해 실구매처 등록 팝업을 먼저 열어 주세요."
            )
        try:
            # The popup has exactly the two business inputs displayed to users:
            # 사업자번호 then 사업자명.  Ignore hidden/submit/reset inputs.
            inputs = popup.locator(
                "input:not([type='hidden']):not([type='button']):not([type='submit']):not([type='reset'])"
            )
            if inputs.count() < 2:
                raise ActualMerchantRegistrationError("실구매처 등록 팝업의 사업자번호·사업자명 입력란을 찾지 못했습니다.")
            inputs.nth(0).fill(business_number)
            inputs.nth(1).fill(merchant_name.strip())

            dialog_messages: list[str] = []

            def accept_dialog(dialog) -> None:
                dialog_messages.append(dialog.message)
                dialog.accept()

            popup.on("dialog", accept_dialog)
            register_button = popup.get_by_text("등록", exact=True)
            if register_button.count() != 1:
                raise ActualMerchantRegistrationError("실구매처 등록 팝업의 '등록' 버튼을 하나로 찾지 못했습니다.")
            register_button.evaluate("element => element.click()")
            popup.wait_for_timeout(600)
            # A registered popup normally closes itself or presents a success
            # alert.  Do not claim success if neither signal is observable.
            if popup.is_closed():
                return "실구매처 등록 창이 정상적으로 닫혔습니다."
            message = " / ".join(dialog_messages)
            if message and not any(word in message for word in ("실패", "오류", "잘못", "입력")):
                return message
            if message:
                raise ActualMerchantRegistrationError(f"실구매처 등록이 완료되지 않았습니다: {message}")
            raise ActualMerchantRegistrationError("실구매처 등록 완료 신호를 확인하지 못했습니다. e-Acc 화면을 확인해 주세요.")
        finally:
            if popup is not None and not popup.is_closed():
                popup.close()

    @staticmethod
    def _find_transaction_row_id(main_frame, transaction: UnsubmittedTransaction) -> str:
        matches = main_frame.evaluate(
            r"""
            target => {
                const digits = value => String(value ?? '').replace(/\D/g, '');
                const amount = value => String(value ?? '').replace(/[^0-9-]/g, '');
                return GridObj.getAllRowIds().split(',').filter(Boolean).filter(rowId =>
                    String(GridObj.cells(rowId, GridObj.getColIndexById('CARD_NO')).getValue()) === target.cardNumber &&
                    String(GridObj.cells(rowId, GridObj.getColIndexById('APPR_NO')).getValue()) === target.approvalNumber &&
                    digits(GridObj.cells(rowId, GridObj.getColIndexById('BLDAT')).getValue()) === target.evidenceDate &&
                    amount(GridObj.cells(rowId, GridObj.getColIndexById('USED_AMT')).getValue()) === target.amount
                );
            }
            """,
            {
                "cardNumber": transaction.card_number,
                "approvalNumber": transaction.approval_number,
                "evidenceDate": transaction.evidence_date.replace("-", ""),
                "amount": str(transaction.amount),
            },
        )
        if len(matches) != 1:
            raise ActualMerchantRegistrationError(
                f"선택 거래와 일치하는 웹 화면 행이 {len(matches)}건입니다. 실구매처를 안전하게 등록할 수 없습니다."
            )
        return str(matches[0])

    @staticmethod
    def _hide_receipt_popup(popup, existing_windows: set[int]) -> None:
        """Hide only the new transient receipt popup, without changing the main Edge window."""
        session = None
        try:
            session = popup.context.new_cdp_session(popup)
            window = session.send("Browser.getWindowForTarget")
            session.send(
                "Browser.setWindowBounds",
                {
                    "windowId": window["windowId"],
                    "bounds": {"windowState": "minimized"},
                },
            )
        except Exception:
            # CDP window-state control is a convenience only. The native hide
            # attempt below handles the normal Windows desktop case.
            pass
        finally:
            if session is not None:
                try:
                    session.detach()
                except Exception:
                    pass

        # CDP only supports minimization, which may still leave a taskbar preview.
        # The newly-created browser window is hidden at the Windows level as well.
        # It remains fully usable through Playwright and is closed in the finally block.
        deadline = time.monotonic() + 1.5
        while time.monotonic() < deadline:
            new_windows = EAccountingBrowserService._edge_popup_window_handles() - existing_windows
            if new_windows:
                try:
                    user32 = ctypes.windll.user32
                    for handle in new_windows:
                        user32.ShowWindow(handle, 0)  # SW_HIDE
                except Exception:
                    pass
                return
            time.sleep(0.05)

    @staticmethod
    def _edge_popup_window_handles() -> set[int]:
        """Return visible Edge popup HWNDs only; the regular Edge window is excluded."""
        if os.name != "nt":
            return set()
        try:
            user32 = ctypes.windll.user32
            handles: set[int] = set()
            callback_type = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

            def visit(handle, _parameter):
                if not user32.IsWindowVisible(handle):
                    return True
                length = user32.GetWindowTextLengthW(handle)
                if length <= 0:
                    return True
                title = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(handle, title, len(title))
                if title.value.startswith("팝업") and "Microsoft Edge" in title.value:
                    handles.add(int(handle))
                return True

            user32.EnumWindows(callback_type(visit), 0)
            return handles
        except Exception:
            return set()

    @staticmethod
    def _query_unsubmitted(main_frame) -> None:
        main_frame.locator("select#CARD_NO").select_option(value="C")
        main_frame.evaluate(
            """
            () => {
                window.__eaccQueryDone = false;
                GridObj.attachEvent('onXLE', () => {
                    window.__eaccQueryDone = true;
                    return true;
                });
                doQuery();
            }
            """
        )
        main_frame.wait_for_function("() => window.__eaccQueryDone === true", timeout=20_000)

    def _wait_for_downloaded_file(
        self,
        suggested_filename: str,
        files_before: dict[str, tuple[int, int]],
        timeout_seconds: float = 20,
    ) -> Path:
        deadline = time.monotonic() + timeout_seconds
        candidate = self.download_directory / suggested_filename
        last_size = -1
        stable_checks = 0
        while time.monotonic() < deadline:
            if candidate.is_file():
                stat = candidate.stat()
                previous = files_before.get(candidate.name)
                changed = previous is None or previous != (stat.st_mtime_ns, stat.st_size)
                if changed and stat.st_size > 0:
                    if stat.st_size == last_size:
                        stable_checks += 1
                    else:
                        stable_checks = 0
                    last_size = stat.st_size
                    if stable_checks >= 1:
                        return candidate
            time.sleep(0.2)
        raise BrowserAutomationError("브라우저가 저장한 미상신내역 파일을 찾지 못했습니다.")

    def _open_card_processing_top_menu(self, page) -> None:
        left_frame = self._wait_for_frame(page, "leftFrame")
        if "MUO20090300001" in left_frame.url:
            return

        top_frame = self._wait_for_frame(page, "topFrame")
        # topFrame은 leftFrame보다 먼저 만들어질 수 있다. 이때 go_home를 즉시
        # 실행하면 parent.leftFrame.displayMenu가 아직 준비되지 않아 TypeError가
        # 발생한다. e-Acc 셸 준비를 확인한 뒤 메뉴 이동을 실행한다.
        self._wait_for_eaccounting_shell(page)
        top_frame.evaluate("go_home('0','MUO20090300001')")
        deadline = time.monotonic() + 15
        while time.monotonic() < deadline:
            left_frame = page.frame(name="leftFrame")
            if left_frame is not None and "MUO20090300001" in left_frame.url:
                return
            page.wait_for_timeout(100)
        raise BrowserAutomationError("상단의 '법인카드 전표처리' 메뉴로 이동하지 못했습니다.")

    @staticmethod
    def _find_eaccounting_page(context):
        browser = context.browser
        contexts = browser.contexts if browser is not None else [context]
        pages = [page for current_context in contexts for page in current_context.pages]
        # 결재선 팝업도 같은 호스트를 사용하므로, 메인 프레임 구성을 가진 창을 최우선한다.
        for page in reversed(pages):
            frame_names = {frame.name for frame in page.frames}
            if {"leftFrame", "mainFrame"}.issubset(frame_names):
                return page
        for page in reversed(pages):
            hostname = (urlparse(page.url).hostname or "").lower().rstrip(".")
            if hostname == E_ACCOUNTING_HOST and "menu_process_mgt_sub.jsp" in page.url:
                return page
        return None

    def _open_unsubmitted_menu(self, page) -> None:
        main_frame = page.frame(name="mainFrame")
        if (
            main_frame is not None
            and "/account/card_use_list.jsp" in main_frame.url
            and "page_type=B" in main_frame.url
        ):
            return

        left_frame = self._wait_for_frame(page, "leftFrame")
        left_frame.locator("body").wait_for(state="visible", timeout=15_000)
        menu = left_frame.locator('a[onclick*="card_use_list.jsp?page_type=B"]')
        if menu.count() != 1:
            menu = left_frame.get_by_text("미상신내역", exact=True)
        if menu.count() != 1:
            raise BrowserAutomationError("왼쪽 메뉴에서 '미상신내역'을 찾지 못했습니다.")
        menu.evaluate("element => eval(element.getAttribute('onclick'))")
        deadline = time.monotonic() + 15
        while time.monotonic() < deadline:
            main_frame = page.frame(name="mainFrame")
            if (
                main_frame is not None
                and "/account/card_use_list.jsp" in main_frame.url
                and "page_type=B" in main_frame.url
            ):
                main_frame.locator("select#CARD_NO").wait_for(
                    state="visible", timeout=10_000
                )
                return
            page.wait_for_timeout(100)
        raise BrowserAutomationError("'미상신내역' 화면 이동이 완료되지 않았습니다.")

    @staticmethod
    def _wait_for_frame(page, name: str, timeout_ms: int = 15_000):
        deadline = datetime.now().timestamp() + (timeout_ms / 1000)
        while datetime.now().timestamp() < deadline:
            frame = page.frame(name=name)
            if frame is not None:
                return frame
            page.wait_for_timeout(100)
        raise BrowserAutomationError(f"e-Accounting의 {name} 화면을 찾지 못했습니다.")

    @staticmethod
    def _wait_for_eaccounting_shell(page, timeout_ms: int = 25_000) -> None:
        """Wait until the e-Acc frames can safely receive menu commands.

        Frame objects are attached before their scripts are available.  The
        portal's ``go_home`` function calls ``parent.leftFrame.displayMenu``;
        therefore waiting for frame *existence* alone causes a race during
        SSO start-up.
        """
        deadline = time.monotonic() + (timeout_ms / 1000)
        while time.monotonic() < deadline:
            top_frame = page.frame(name="topFrame")
            left_frame = page.frame(name="leftFrame")
            main_frame = page.frame(name="mainFrame")
            if top_frame is not None and left_frame is not None and main_frame is not None:
                try:
                    top_ready = bool(top_frame.evaluate("() => typeof window.go_home === 'function'"))
                    left_ready = bool(left_frame.evaluate("() => typeof window.displayMenu === 'function'"))
                    main_ready = main_frame.url not in {"", "about:blank"}
                    if top_ready and left_ready and main_ready:
                        return
                except Exception:
                    # 프레임이 새로 고쳐지는 중일 수 있으므로 다음 짧은 주기에 재확인한다.
                    pass
            page.wait_for_timeout(100)
        raise BrowserAutomationError("e-Accounting 메뉴 화면의 초기화가 완료되지 않았습니다.")
