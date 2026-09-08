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
from pathlib import Path
from typing import Iterable
from uuid import uuid4


MAIL_SUBJECT = "[법인카드] 미등록 사용내역 등록 요청"
LOCAL_SENDER_SETTINGS_FILENAME = "mail_sender.json"
OL_FOLDER_DRAFTS = 16


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


def send_via_outlook(recipient: MailRecipient, html_body: str) -> tuple[str, str]:
    """Ask the locally configured Outlook client to deliver one HTML message.

    A successful return means Outlook accepted ``Send``.  It does not claim
    that the recipient has opened or read the message.
    """
    if not recipient.email or "@" not in recipient.email:
        raise ValueError("유효한 수신자 이메일 주소가 없습니다.")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError("Outlook 자동화 구성요소(pywin32)를 찾을 수 없습니다.") from exc
    outlook = win32com.client.Dispatch("Outlook.Application")
    message_id = uuid4().hex
    sender_address = _local_sender_address()
    send_account = None
    if sender_address:
        send_account = _find_outlook_account(outlook, sender_address)
        if send_account is None:
            raise RuntimeError(
                f"로컬 발신 계정({sender_address})을 Outlook에서 찾을 수 없습니다. "
                "이 PC의 Outlook 또는 로컬 설정을 확인해 주세요."
            )
    if send_account is None:
        # No per-PC override: Outlook's normal default-account behavior.
        mail = outlook.CreateItem(0)
    else:
        # Creating the item in the selected account's Drafts store is the
        # reliable Outlook/Exchange way to bind the message to that account.
        # SendUsingAccount alone may be ignored for items first created in
        # the profile's default store.
        mail = send_account.DeliveryStore.GetDefaultFolder(OL_FOLDER_DRAFTS).Items.Add(
            "IPM.Note"
        )
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
    mail.Save()
    if send_account is not None:
        # Outlook can reset a sender selected on an unsaved message when the
        # first Save creates its Drafts/Outbox item.  Select and persist the
        # account only after that save, immediately before Send.
        mail.SendUsingAccount = send_account
        mail.Save()
    mail.Send()
    # ``Send`` queues the item in Outlook; it is not evidence that mail has
    # already left the Outbox.  Record the truthful immediate state first.
    # A later folder reconciliation can promote this to ``발송 완료``.
    return "발송 대기", message_id


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
            settings = json.loads(path.read_text(encoding="utf-8"))
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
