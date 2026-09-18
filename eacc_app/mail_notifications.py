from __future__ import annotations

"""Corporate-card unprocessed-mail rendering and Outlook delivery helpers.

The module deliberately has no dependency on tkinter or the e-Acc browser.  It
can therefore be tested with sample rows and records exactly what Outlook
accepted (or rejected) for the audit log.
"""

from dataclasses import dataclass
from datetime import datetime, timedelta
from decimal import Decimal
from html import escape
import json
import os
import time
from pathlib import Path
from typing import Iterable, Mapping
from uuid import uuid4

from .models import ExceptionMailUse


MAIL_SUBJECT = "[법인카드] 미등록 사용내역 등록 요청"
LOCAL_SENDER_SETTINGS_FILENAME = "mail_sender.json"
OL_FOLDER_DRAFTS = 16

# Outlook must receive a real Send request; merely moving a draft into the
# Outbox does not submit it for delivery on every profile.  Pace those real
# requests conservatively so a large batch cannot monopolise Outlook.
BULK_SEND_PACING_SECONDS = 1.0
BULK_SEND_BATCH_SIZE = 20
BULK_SEND_BATCH_PAUSE_SECONDS = 10.0


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


def render_exception_mail_html(
    cost_center: str,
    uses: Iterable[ExceptionMailUse],
    process_date: str | None = None,
) -> str:
    """Render the team exception-mail template without receipt attachments."""
    rows = tuple(uses)
    if not rows:
        raise ValueError("메일 본문에 넣을 예외처리 내역이 없습니다.")
    process_date = process_date or datetime.now().strftime("%Y-%m-%d")
    body_rows = "".join(
        "<tr style=\"background:%s;\">"
        "<td style=\"%s\">%d</td><td style=\"%s\">%s</td>"
        "<td style=\"%s\">%s</td><td style=\"%s\">%s</td><td style=\"%s\">%s</td>"
        "<td style=\"%s\">%s원</td><td style=\"%s\">%s</td></tr>"
        % (
            "#f8fafc" if index % 2 == 0 else "#ffffff",
            _cell("center"), index, _cell("center"), escape(item.approval_number),
            _cell("center"), escape(item.card_holder), _cell("center"), escape(item.evidence_date), _cell("left"), escape(item.merchant),
            _cell("right"), f"{item.amount:,.0f}", _cell("left"), escape(_short_exception_reason(item.reason)),
        )
        for index, item in enumerate(rows, start=1)
    )
    return f"""<!doctype html>
<html lang="ko"><body style="margin:0;padding:0;background:#ffffff;font-family:'Malgun Gothic',Arial,sans-serif;color:#102b4a;font-size:14px;line-height:1.65;">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:800px;max-width:800px;margin:0 auto;"><tr><td style="padding:20px 30px 0;">
<div style="background:#f6f8fb;border-left:4px solid #c76a18;padding:18px 20px;margin-bottom:24px;">안녕하세요, {escape(cost_center)} 담당자님.<br>법인카드 자동 처리 중 담당자 확인이 필요한 예외처리 건이 검출 되었습니다.<br>아래 내역을 확인하시어 e-Acc에서 필요한 조치를 진행해 주세요.</div>
<div style="font-weight:bold;color:#0f3c67;margin-bottom:7px;">법인카드 예외처리 내역</div>
<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:collapse;table-layout:fixed;font-size:13px;"><colgroup><col style="width:5%"><col style="width:11%"><col style="width:10%"><col style="width:12%"><col style="width:17%"><col style="width:11%"><col style="width:34%"></colgroup><thead><tr style="background:#163a60;color:#ffffff;"><th style="padding:8px;white-space:nowrap;text-align:center;">순번</th><th style="padding:8px;white-space:nowrap;text-align:center;">승인번호</th><th style="padding:8px;white-space:nowrap;text-align:center;">카드소지자</th><th style="padding:8px;white-space:nowrap;text-align:center;">증빙일자</th><th style="padding:8px;white-space:nowrap;text-align:left;">거래처</th><th style="padding:8px;white-space:nowrap;text-align:right;">사용금액</th><th style="padding:8px;white-space:nowrap;text-align:left;">예외 사유</th></tr></thead><tbody>{body_rows}</tbody></table>
<div style="background:#fff5e5;border:1px solid #ead8bc;border-top:0;padding:10px 12px;">대상 건수 <strong style="color:#d45d00;">{len(rows)}건</strong>&nbsp;·&nbsp;처리일 <strong style="color:#d45d00;">{escape(process_date)}</strong></div>
<div style="margin-top:29px;">문의사항은 법인카드 자동처리 담당자에게 연락 바랍니다.<br><br>감사합니다.</div>
<div style="margin-top:30px;padding:10px 12px;text-align:center;background:#f6f8fb;border-top:1px solid #dfe3e8;color:#7c8b9a;font-size:12px;">본 메일은 자동 발송되었습니다.</div>
</td></tr></table></body></html>"""


class OutlookSendContext:
    """A reusable Outlook COM handle plus resolved sender-account binding.

    Without a shared context each message would create a new Outlook COM
    proxy and resolve the sender account again. On profiles with several
    dozen recipients those COM round-trips add up and interact badly with
    Exchange throttling.

    This context is built ONCE per bulk run and reused for every message,
    so the expensive discovery work happens a single time.
    """

    __slots__ = ("outlook", "send_account", "sender_address", "last_send_at", "sent_count")

    def __init__(self, outlook, send_account, sender_address: str) -> None:
        self.outlook = outlook
        self.send_account = send_account
        # Cache the SMTP address as a plain string so the hot path never has
        # to cross the COM boundary just to compare against a config value.
        self.sender_address = sender_address
        self.last_send_at = 0.0
        self.sent_count = 0

    def close(self) -> None:
        """Release this program's COM references without closing Outlook."""
        # Do not call Outlook.Application.Quit(): Outlook belongs to the user.
        # Dropping our references lets pywin32 release the automation handles
        # as soon as the bulk worker finishes.
        self.send_account = None
        self.outlook = None


def open_outlook_send_context() -> OutlookSendContext:
    """Build a bulk-send context: resolve Outlook + the sender account once.

    Raises the same errors ``send_via_outlook`` would raise, so the caller
    can surface configuration problems before the first message is submitted.
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
    subject: str = MAIL_SUBJECT,
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
        # Space actual Send() calls out and give Outlook a short recovery
        # window after each group of 20 requests. This runs on the dedicated
        # mail worker, never on the UI thread.
        elapsed = time.monotonic() - context.last_send_at
        if context.sent_count and context.sent_count % BULK_SEND_BATCH_SIZE == 0:
            time.sleep(max(0.0, BULK_SEND_BATCH_PAUSE_SECONDS - elapsed))
        elif 0 < elapsed < BULK_SEND_PACING_SECONDS:
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
    mail.Subject = subject
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
    # This is the real Outlook submission request. Unlike Move(Outbox), it
    # marks the mail for delivery. Outlook/Exchange may complete delivery
    # later, which is why the Log starts at '발송 요청' and is reconciled
    # separately after the whole batch has been submitted.
    mail.Send()

    if not owns_context:
        context.last_send_at = time.monotonic()
        context.sent_count += 1
    return "발송 요청", message_id


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
    """Compatibility wrapper for checking one program message."""
    return outlook_delivery_statuses((message_id,)).get(message_id, "발송 확인 불가")


def outlook_delivery_statuses(
    message_ids: Iterable[str],
    *,
    since: datetime | None = None,
) -> dict[str, str]:
    """Check one batch in one pass through the sender's recent folders.

    Each message ID is a private token embedded in its mail.  The function
    scans a folder once and resolves every token it finds, rather than
    rescanning the same Outlook folder once per recipient.
    """
    requested_ids = {str(value) for value in message_ids if str(value)}
    result = {message_id: "발송 확인 불가" for message_id in requested_ids}
    if not requested_ids:
        return result
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        return result
    outlook = win32com.client.Dispatch("Outlook.Application")
    try:
        return _outlook_delivery_statuses(
            outlook, requested_ids, _local_sender_address(), since=since
        )
    finally:
        # This only releases the COM proxy created by this check. It never
        # terminates the user's Outlook application.
        del outlook


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


def _outlook_delivery_statuses(
    outlook,
    requested_ids: set[str],
    sender_address: str = "",
    *,
    since: datetime | None = None,
) -> dict[str, str]:
    namespace = outlook.GetNamespace("MAPI")
    result = {message_id: "발송 확인 불가" for message_id in requested_ids}
    unresolved_ids = set(requested_ids)
    # A message can be created a few minutes before its audit row is written.
    # This buffer still confines the scan to the active delivery batch rather
    # than Outlook's full historical folder.
    cutoff = since - timedelta(minutes=5) if since is not None else None
    # Sent Items is checked first: a fast online send may leave Outbox before
    # this method is called.
    for folder_id, status in ((5, "발송 완료"), (4, "발송 대기")):
        for folder in _outlook_folders_for_status(outlook, namespace, folder_id, sender_address):
            items = folder.Items
            try:
                # Newest first means a newly sent item is found without a
                # historical mailbox scan. CreationTime works for both Sent
                # Items and Outbox folders.
                items.Sort("[CreationTime]", True)
            except Exception:
                pass
            try:
                item_count = int(items.Count)
            except Exception:
                continue
            for index in range(1, item_count + 1):
                item = items.Item(index)
                try:
                    if cutoff is not None and _outlook_item_is_older_than(item, cutoff):
                        # Items were sorted newest-first, so older history can
                        # never contain this just-created batch.
                        break
                    prop = item.UserProperties.Find("EAccAutomationMessageId")
                    message_id = str(prop.Value) if prop is not None else ""
                    if message_id not in unresolved_ids:
                        message_id = ""
                    if not message_id:
                        # UserProperties can be hidden while Outlook syncs;
                        # the invisible body marker is the reliable fallback.
                        body = str(item.HTMLBody)
                        for candidate in unresolved_ids:
                            if f"EAccAutomationMessageId:{candidate}" in body:
                                message_id = candidate
                                break
                    if message_id in unresolved_ids:
                        result[message_id] = status
                        unresolved_ids.remove(message_id)
                        if not unresolved_ids:
                            return result
                except Exception:
                    continue
    return result


def _outlook_item_is_older_than(item, cutoff: datetime) -> bool:
    """Best-effort CreationTime cutoff that tolerates Outlook date variants."""
    try:
        created_at = item.CreationTime
    except Exception:
        return False
    if not isinstance(created_at, datetime):
        return False
    if created_at.tzinfo is None and cutoff.tzinfo is not None:
        cutoff = cutoff.replace(tzinfo=None)
    elif created_at.tzinfo is not None and cutoff.tzinfo is None:
        created_at = created_at.replace(tzinfo=None)
    return created_at < cutoff


def _outlook_folders_for_status(outlook, namespace, folder_id: int, sender_address: str = ""):
    """Yield only the folder belonging to the actual sending account.

    A per-PC override has a known sender account, so inspecting unrelated
    accounts is incorrect as well as expensive.  Without an override Outlook's
    default folder is the account used by the deployment path.
    """
    folders = []
    if sender_address:
        account = _find_outlook_account(outlook, sender_address)
        if account is not None:
            try:
                return (account.DeliveryStore.GetDefaultFolder(folder_id),)
            except Exception:
                return ()
    try:
        folders.append(namespace.GetDefaultFolder(folder_id))
    except Exception:
        pass
    return tuple(folders)


def _cell(alignment: str) -> str:
    return f"padding:9px 8px;border-bottom:1px solid #d9dde2;text-align:{alignment};"


def _short_exception_reason(reason: str) -> str:
    """Keep only the exception category, excluding volatile OCR detail."""
    text = " ".join(str(reason or "").split())
    # 음료전용 적요 검증은 음식·주류 등 음료 이외의 품목을 감지했을 때만
    # 예외가 된다. 메일에서는 내부 검증 규칙명이 아니라 담당자가 바로
    # 이해할 수 있는 조치 사유로 표시한다.
    if text.startswith("음료전용 적요:"):
        return "음료 외 내역 존재"
    for separator in (":", "："):
        if separator in text:
            text = text.split(separator, 1)[0]
            break
    return text.rstrip(" /·")[:80]
