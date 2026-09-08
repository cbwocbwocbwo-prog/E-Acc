from __future__ import annotations

"""Parser for the live e-Acc '미처리내역' grid."""

import hashlib
from decimal import Decimal, InvalidOperation

from .mail_notifications import UnprocessedCardUse


_ALIASES = {
    "employee_name": ("성명", "이름", "카드소지자"),
    "department": ("부서명", "부서", "본부"),
    "approval_number": ("승인번호",),
    "usage_date": ("이용일자", "이용일", "증빙일자", "전기일"),
    "usage_time": ("이용시간",),
    "merchant": ("상호", "가맹점", "거래처"),
    "amount": ("사용금액", "이용금액"),
}


class UnprocessedCardFormatError(ValueError):
    pass


def unprocessed_use_from_grid_row(source_row_number: int, headers: list[str], values: list[str]) -> UnprocessedCardUse:
    fields = {header.strip().replace(" ", ""): str(value).strip() for header, value in zip(headers, values)}

    def value(name: str, required: bool = True) -> str:
        result = next((fields.get(alias, "") for alias in _ALIASES[name] if fields.get(alias, "")), "")
        if required and not result:
            raise UnprocessedCardFormatError(f"'{_ALIASES[name][0]}' 열 값을 찾지 못했습니다.")
        return result

    employee_name = value("employee_name")
    approval_number = value("approval_number")
    usage_date = _normalize_date(value("usage_date"))
    usage_time = value("usage_time", required=False)
    merchant = value("merchant")
    raw_amount = value("amount").replace(",", "").replace("원", "")
    try:
        amount = Decimal(raw_amount)
    except InvalidOperation as exc:
        raise UnprocessedCardFormatError("사용금액 형식이 올바르지 않습니다.") from exc
    transaction_id = hashlib.sha256(
        "|".join((employee_name, approval_number, usage_date, usage_time, merchant, str(amount))).encode("utf-8")
    ).hexdigest()[:24]
    return UnprocessedCardUse(
        transaction_id=transaction_id, employee_name=employee_name,
        department=value("department", required=False), approval_number=approval_number,
        usage_datetime=" ".join(part for part in (usage_date, usage_time) if part),
        merchant=merchant, amount=amount,
    )


def _normalize_date(value: str) -> str:
    digits = "".join(char for char in value if char.isdigit())
    return f"{digits[:4]}-{digits[4:6]}-{digits[6:]}" if len(digits) == 8 else value
