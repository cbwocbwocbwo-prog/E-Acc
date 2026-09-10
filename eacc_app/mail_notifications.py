from __future__ import annotations

"""Corporate-card unprocessed-mail rendering and Outlook delivery helpers.

The module deliberately has no dependency on tkinter or the e-Acc browser.  It
can therefore be tested with sample rows and records exactly what Outlook
accepted (or rejected) for the audit log.
"""

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from html import escape
import json
import os
import time
from pathlib import Path
from typing import Iterable
from uuid import uuid4


MAIL_SUBJECT = "[법인카드] 미등록 사용내역 등록 요청"
LOCAL_SENDER_SETTINGS_FILENAME = "mail_sender.json"
OL_FOLDER_DRAFTS = 16

# Per-message pacing (seconds) between consecutive Send calls.  Now that
# bulk sends queue into the Outbox instead of calling ``Send`` synchronously
# (see below), Outlook no longer blocks the calling thread on Exchange, so
# a small 0.15s spacing is enough to keep the COM automation responsive
# without artificially slowing down the queue-up phase.
BULK_SEND_PACING_SECONDS = 0.15

# Outlook's ``Outbox`` default-folder constant (olFolderOutbox in the
# Outlook object model).  Moving a saved draft into this folder is how the
# bulk-send path enqueues messages without triggering the synchronous
# Exchange handshake that ``MailItem.Send`` runs.
OL_FOLDER_OUTBOX = 4


@dataclass(frozen=True, slots=True)
class UnprocessedCardUse:
    transaction_id: str
    employee_name: str
    department: str
    approval_number: str
    usage_datetime: str
    merchant: str
    amount: Decimal


@dataclass(frozen=True, slots=True)
class MailRecipient:
    name: str
    email: str
    department: str = ""


@dataclass(frozen=True, slots=True)
class MailDeliveryResult:
    recipient: MailRecipient
    transaction_ids: tuple[str, ...]
    status: str
    reason: str = ""


def render_mail_html(
    recipient_name: str,
    uses: Iterable[UnprocessedCardUse],
    query_date: str | None = None,
) -> str:
    """Render the user-approved 720px HTML template with only safe fields."""
    rows = tuple(uses)
    if not rows:
        raise ValueError("메일 본문에 넣을 미처리 사용내역이 없습니다.")
    query_date = query_date or datetime.now().strftime("%Y-%m-%d")
    body_rows = "".join(
        "<tr style=\"background:%s;\">"
        "<td style=\"%s\">%d</td><td style=\"%s\">%s</td>"
        "<td style=\"%s\">%s</td><td style=\"%s\">%s</td>"
        "<td style=\"%s\">%s원</td></tr>"
        % (
            "#f8fafc" if index % 2 == 0 else "#ffffff",
            _cell("center"), index, _cell("center"), escape(item.approval_number),
            _cell("center"), escape(item.usage_datetime), _cell("left"), escape(item.merchant),
            _cell("right"), f"{item.amount:,.0f}",
        )
        for index, item in enumerate(rows, start=1)
    )
    return f"""<!doctype html>
<html lang="ko"><body style="margin:0;padding:0;background:#ffffff;font-family:'Malgun Gothic',Arial,sans-serif;color:#102b4a;font-size:14px;line-height:1.65;">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:720px;max-width:720px;margin:0 auto;"><tr><td style="padding:20px 30px 0;">
<div style="background:#f6f8fb;border-left:4px solid #c76a18;padding:18px 20px;margin-bottom:24px;">안녕하세요, {escape(recipient_name)}님.<br>법인카드 사용내역 중 O-Park에 등록되지 않은 건이 확인되었습니다.<br>아래 내역을 확인하시어 O-Park에 등록해 주세요.</div>
<div style="font-weight:bold;color:#0f3c67;margin-bottom:7px;">미등록 사용내역</div>
<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:collapse;font-size:13px;"><thead><tr style="background:#163a60;color:#ffffff;"><th style="padding:8px;text-align:center;">순번</th><th style="padding:8px;text-align:center;">승인번호</th><th style="padding:8px;text-align:center;">이용일자</th><th style="padding:8px;text-align:left;">가맹점</th><th style="padding:8px;text-align:right;">사용금액</th></tr></thead><tbody>{body_rows}</tbody></table>
<div style="background:#fff5e5;border:1px solid #ead8bc;border-top:0;padding:10px 12px;">대상 건수 <strong style="color:#d45d00;">{len(rows)}건</strong>&nbsp;·&nbsp;조회기준일 <strong style="color:#d45d00;">{escape(query_date)}</strong></div>
<div style="margin-top:29px;">문의사항은 팀 법인카드 담당자로 연락 바랍니다.<br><br>감사합니다.</div>
<div style="margin-top:30px;padding:10px 12px;text-align:center;background:#f6f8fb;border-top:1px solid #dfe3e8;color:#7c8b9a;font-size:12px;">본 메일은 자동 발송되었습니다.</div>
</td></tr></table></body></html>"""


class OutlookSendContext:
    """A reusable Outlook COM handle plus resolved sender-account binding.

    Bulk-send runs call ``Dispatch("Outlook.Application")`` and re-resolve
    the sender account for every message when they use the simple
    ``send_via_outlook`` entry point.  On profiles with several dozen
    recipients each of those COM round-trips adds up, keeps the UI thread
    marshalled to Outlook, and interacts badly with Exchange throttling,
    which is exactly the "PC frozen for minutes" symptom.

    This context is built ONCE per bulk run and reused for every message,
    so the expensive discovery work happens a single time.
    """

    __slots__ = ("outlook", "send_account", "sender_address", "last_send_at")

    def __init__(self, outlook, send_account, sender_address: str) -> None:
        self.outlook = outlook
        self.send_account = send_account
        # Cache the SMTP address as a plain string so the hot path never has
        # to cross the COM boundary just to compare against a config value.
        self.sender_address = sender_address
        self.last_send_at = 0.0


def open_outlook_send_context() -> OutlookSendContext:
    """Build a bulk-send context: resolve Outlook + the sender account once.

    Raises the same errors ``send_via_outlook`` would raise, so the caller
    can surface configuration problems before the first message is queued.
    """
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError("Outlook 자동화 구성요소(pywin32)를 찾을 수 없습니다.") from exc
    outlook = win32com.client.Dispatch("Outlook.Application")
    sender_address = _local_sender_address()
    send_account = None
    if sender_address:
        send_account = _find_outlook_account(outlook, sender_address)
        if send_account is None:
            available = _list_outlook_smtp_addresses(outlook)
            raise RuntimeError(
                f"로컬 발신 계정({sender_address})을 Outlook에서 찾을 수 없습니다. "
                f"현재 Outlook에 구성된 계정: {available or '(없음)'}. "
                "이 PC의 Outlook 계정 설정 또는 mail_sender.json을 확인해 주세요."
            )
    return OutlookSendContext(outlook, send_account, sender_address)


def send_via_outlook(
    recipient: MailRecipient,
    html_body: str,
    context: OutlookSendContext | None = None,
) -> tuple[str, str]:
    """Ask the locally configured Outlook client to deliver one HTML message.

    A successful return means Outlook accepted ``Send``.  It does not claim
    that the recipient has opened or read the message.

    Sender-account resolution:
      1) If %LOCALAPPDATA%\\EAccAutomation\\mail_sender.json exists AND
         specifies ``outlook_sender_address``, the message is bound to that
         specific Outlook account (this PC override).
      2) Otherwise the profile's default Outlook account is used
         (the correct behavior for deployment).

    When ``context`` is provided the expensive Outlook discovery work is
    skipped and per-message pacing is applied automatically, which is the
    right thing to do inside a bulk-send loop.  Callers that only send
    one-off messages can omit it; a fresh single-use context is created
    transparently in that case.
    """
    if not recipient.email or "@" not in recipient.email:
        raise ValueError("유효한 수신자 이메일 주소가 없습니다.")

    owns_context = context is None
    if context is None:
        context = open_outlook_send_context()
    else:
        # Space consecutive Send() calls out so Exchange doesn't throttle
        # (and Outlook doesn't wedge the UI thread waiting for a response).
        elapsed = time.monotonic() - context.last_send_at
        if 0 < elapsed < BULK_SEND_PACING_SECONDS:
            time.sleep(BULK_SEND_PACING_SECONDS - elapsed)

    outlook = context.outlook
    send_account = context.send_account
    sender_address = context.sender_address
    message_id = uuid4().hex

    if send_account is None:
        # Deployment path: no override -> use the profile's default account.
        mail = outlook.CreateItem(0)  # 0 = olMailItem
    else:
        # Per-PC override path.  Outlook's SendUsingAccount is unreliable on
        # profiles where the *default* delivery account differs from the
        # requested one: Outlook may reset SendUsingAccount at Send time and
        # deliver from the profile default anyway.  The reliable workaround
        # is to create the draft *inside* the requested account's own store
        # (its Drafts folder) - Outlook then binds the item to that store's
        # owner account at Send time and honors it.
        mail = _create_mail_in_account_store(send_account)
    mail.To = recipient.email
    mail.Subject = MAIL_SUBJECT
    # Outlook may preserve UserProperties in the Outbox but not expose them
    # consistently while it is synchronizing.  Keep the same opaque ID in an
    # invisible HTML comment too, so later folder checks have a second,
    # non-visual way to recognize this exact message.
    mail.HTMLBody = f"<!-- EAccAutomationMessageId:{message_id} -->{html_body}"
    # A subject is shared by many recipients, so it cannot identify one
    # message safely.  Store a private Outlook user property for subsequent
    # Outbox/Sent Items verification.
    mail.UserProperties.Add("EAccAutomationMessageId", 1, True).Value = message_id
    if send_account is not None:
        # Belt-and-braces: also assign SendUsingAccount.  With the item
        # already living in the account's own store this rarely gets reset,
        # but setting it makes the intent explicit and covers Outlook
        # builds that use it as an additional hint.
        try:
            mail.SendUsingAccount = send_account
        except Exception:
            # Some Outlook versions raise if the property is unavailable for
            # a store-hosted item; safe to ignore because the store binding
            # already selects the sender.
            pass
    # A single Save materializes the draft in the target store and gives
    # Outlook a stable object to bind SendUsingAccount to.  The previous
    # code saved twice (once before and once after re-asserting the
    # account) - the second Save doubled per-message COM/disk cost with
    # no observed benefit now that the draft is created directly inside
    # the requested account's Drafts folder.
    mail.Save()
    if send_account is not None:
        # Verify the binding actually stuck; fail loudly instead of silently
        # sending from the wrong mailbox.
        bound_smtp = _resolve_bound_sender_smtp(mail)
        if bound_smtp and bound_smtp != sender_address.casefold():
            raise RuntimeError(
                f"Outlook이 발신 계정을 {sender_address}(으)로 설정하지 못했습니다. "
                f"(현재 바인딩: {bound_smtp}). "
                "Outlook에서 해당 계정의 '보낸 사람으로 사용' 권한 또는 "
                "기본 데이터 파일(파일 → 계정 설정 → 데이터 파일)을 확인해 주세요."
            )
    # 예전엔 ``mail.Send()``를 바로 호출했다.  Send()는 동기 호출이라
    # Exchange 서버 왕복이 끝날 때까지 Outlook COM 스레드를 붙잡고 있어서,
    # 70명 발송 시 Outlook과 PC가 몇 분씩 얼어붙는 원인이 됐다.
    #
    # 대신 초안(mail)을 요청 계정의 Outbox 폴더로 이동시켜 큐잉만 한다.
    # 그 다음부턴 Outlook이 자기 페이스로 백그라운드 발송을 하고,
    # 우리 프로그램은 폴링으로 '보낸 편지함' 이동을 감지해 로그를 갱신한다.
    #
    # 참고: MoveTo는 로컬 스토어 조작이라 서버 왕복이 없다.  Save 직후
    #       바로 호출해도 안전하다.
    queued = False
    try:
        if send_account is not None:
            outbox = send_account.DeliveryStore.GetDefaultFolder(OL_FOLDER_OUTBOX)
        else:
            outbox = outlook.GetNamespace("MAPI").GetDefaultFolder(OL_FOLDER_OUTBOX)
        if outbox is not None:
            mail.Move(outbox)
            queued = True
    except Exception:
        # 이 프로필에서 Outbox 이동이 실패하면(예: POP3 즉시발송 계정)
        # 옛날 방식으로 폴백한다.  이 경로는 Outlook을 잠깐 붙잡을 수
        # 있지만 최소한 메일은 나간다.
        queued = False
    if not queued:
        mail.Send()

    if not owns_context:
        context.last_send_at = time.monotonic()
    # Outbox에 큐잉된 상태이므로 '발송 대기'가 정확한 즉시 상태다.
    # 이후 폴링이 '발송 완료'로 승격한다.
    return "발송 대기", message_id


def _create_mail_in_account_store(send_account):
    """Create an ``IPM.Note`` item inside the account's own Drafts folder.

    Falling back to ``Application.CreateItem`` when the store is unreachable
    keeps the automation working on profiles that expose the account without
    a bound delivery store (rare, but observed on some POP/IMAP setups).
    """
    try:
        drafts = send_account.DeliveryStore.GetDefaultFolder(OL_FOLDER_DRAFTS)
    except Exception:
        drafts = None
    if drafts is None:
        # Best-effort fallback: at least the message will be composed; the
        # subsequent SendUsingAccount assignment may still take effect.
        outlook = send_account.Application
        return outlook.CreateItem(0)
    return drafts.Items.Add("IPM.Note")


def _resolve_bound_sender_smtp(mail) -> str:
    """Return the SMTP address Outlook will actually send this item from.

    Checks the delivery store's owner account first (this is what Outlook
    honors at Send time when the item lives in that store) and falls back
    to whatever ``SendUsingAccount`` currently reports.
    """
    # 1) Parent folder -> Store -> owning account.  This reflects where the
    #    item is physically saved, which is what Outlook actually uses.
    try:
        parent_store = mail.Parent.Store
    except Exception:
        parent_store = None
    if parent_store is not None:
        try:
            outlook = mail.Application
            store_id = parent_store.StoreID
            for index in range(1, outlook.Session.Accounts.Count + 1):
                account = outlook.Session.Accounts.Item(index)
                try:
                    if account.DeliveryStore.StoreID == store_id:
                        return str(account.SmtpAddress).casefold()
                except Exception:
                    continue
        except Exception:
            pass
    # 2) Fall back to whatever SendUsingAccount currently says.
    try:
        return str(mail.SendUsingAccount.SmtpAddress).casefold()
    except Exception:
        return ""


def _list_outlook_smtp_addresses(outlook) -> str:
    """Return a comma-separated list of SMTP addresses configured in Outlook."""
    addresses = []
    try:
        for index in range(1, outlook.Session.Accounts.Count + 1):
            try:
                addresses.append(str(outlook.Session.Accounts.Item(index).SmtpAddress))
            except Exception:
                continue
    except Exception:
        return ""
    return ", ".join(addr for addr in addresses if addr)


def outlook_delivery_status(message_id: str) -> str:
    """Check a prior message without sending another one."""
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        return "발송 확인 불가"
    return _outlook_delivery_status(win32com.client.Dispatch("Outlook.Application"), message_id)


def _find_outlook_account(outlook, sender_address: str):
    """Return the configured Outlook account matching an SMTP address."""
    target = sender_address.casefold()
    for index in range(1, outlook.Session.Accounts.Count + 1):
        account = outlook.Session.Accounts.Item(index)
        if str(account.SmtpAddress).casefold() == target:
            return account
    return None


def _local_sender_address() -> str:
    """Read an optional per-PC Outlook sender override outside the app bundle."""
    # Desktop/package hosts can override LOCALAPPDATA.  Prefer the real
    # Windows user profile location so this PC's setting is found regardless
    # of whether the app is run from a terminal, a package host, or an exe.
    paths = [
        Path.home() / "AppData" / "Local" / "EAccAutomation" / LOCAL_SENDER_SETTINGS_FILENAME
    ]
    environment_path = (
        Path(os.environ.get("LOCALAPPDATA", Path.home()))
        / "EAccAutomation"
        / LOCAL_SENDER_SETTINGS_FILENAME
    )
    if environment_path not in paths:
        paths.append(environment_path)
    for path in paths:
        try:
            # ``utf-8-sig`` transparently strips a UTF-8 BOM if one is
            # present (PowerShell's ``Set-Content -Encoding UTF8`` writes a
            # BOM, so plain ``utf-8`` would fail here and the override would
            # be silently ignored -> mail would fall back to the default
            # account, which is exactly the bug we are fixing).
            settings = json.loads(path.read_text(encoding="utf-8-sig"))
        except (OSError, json.JSONDecodeError):
            continue
        sender_address = str(settings.get("outlook_sender_address", "")).strip()
        if sender_address:
            return sender_address
    return ""


def _outlook_delivery_status(outlook, message_id: str) -> str:
    namespace = outlook.GetNamespace("MAPI")
    # Sent Items is checked first: a fast online send may leave Outbox before
    # this method is called.
    for folder_id, status in ((5, "발송 완료"), (4, "발송 대기")):
        for folder in _outlook_folders_for_status(outlook, namespace, folder_id):
            for index in range(1, folder.Items.Count + 1):
                item = folder.Items.Item(index)
                try:
                    prop = item.UserProperties.Find("EAccAutomationMessageId")
                    is_matching_property = prop is not None and str(prop.Value) == message_id
                    is_matching_body = (
                        f"EAccAutomationMessageId:{message_id}" in str(item.HTMLBody)
                    )
                    if is_matching_property or is_matching_body:
                        return status
                except Exception:
                    continue
    return "발송 확인 불가"


def _outlook_folders_for_status(outlook, namespace, folder_id: int):
    """Yield the default plus every configured account's matching folder."""
    folders = []
    try:
        folders.append(namespace.GetDefaultFolder(folder_id))
    except Exception:
        pass
    for index in range(1, outlook.Session.Accounts.Count + 1):
        try:
            folder = outlook.Session.Accounts.Item(index).DeliveryStore.GetDefaultFolder(folder_id)
            if all(folder.EntryID != existing.EntryID for existing in folders):
                folders.append(folder)
        except Exception:
            continue
    return tuple(folders)


def _cell(alignment: str) -> str:
    return f"padding:9px 8px;border-bottom:1px solid #d9dde2;text-align:{alignment};"
