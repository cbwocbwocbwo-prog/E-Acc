from __future__ import annotations

from dataclasses import asdict, dataclass
from decimal import Decimal
from pathlib import Path
from typing import Final


EXPECTED_UNSUBMITTED_HEADERS: Final[tuple[str, ...]] = (
    "전표상태",
    "증빙유무",
    "카드번호",
    "카드소지자",
    "승인번호",
    "요일",
    "이용시간",
    "전기일",
    "증빙일자",
    "계정명",
    "적요",
    "거래처",
    "사용금액",
    "세금코드",
    "업종",
    "코스트센터",
    "카드발급일",
    "본부",
    "실구매처명",
    "선택",
    "실구매처코드",
)

# 웹 화면의 검색 아이콘 경로인 '선택'은 업무 데이터에서 제외한다.
DISPLAY_HEADERS: Final[tuple[str, ...]] = tuple(
    header for header in EXPECTED_UNSUBMITTED_HEADERS if header != "선택"
)


@dataclass(frozen=True, slots=True)
class UnsubmittedTransaction:
    source_row_number: int
    voucher_status: str
    evidence_status: str
    card_number: str
    card_holder: str
    approval_number: str
    weekday: str
    usage_time: str
    posting_date: str
    evidence_date: str
    account_name: str
    description: str
    merchant: str
    amount: Decimal
    tax_code: str
    business_type: str
    cost_center: str
    card_issue_date: str
    division: str
    actual_merchant_name: str
    actual_merchant_code: str
    transaction_key: str
    transaction_id: str

    def to_json_dict(self) -> dict[str, object]:
        values = asdict(self)
        values["amount"] = str(self.amount)
        return values

    def display_values(self) -> tuple[str, ...]:
        return (
            self.voucher_status,
            self.evidence_status,
            self.card_number,
            self.card_holder,
            self.approval_number,
            self.weekday,
            self.usage_time,
            self.posting_date,
            self.evidence_date,
            self.account_name,
            self.description,
            self.merchant,
            f"{self.amount:,.0f}",
            self.tax_code,
            self.business_type,
            self.cost_center,
            self.card_issue_date,
            self.division,
            self.actual_merchant_name,
            self.actual_merchant_code,
        )


@dataclass(frozen=True, slots=True)
class ParsedRow:
    source_row_number: int
    transaction: UnsubmittedTransaction | None
    raw_values: tuple[str, ...]
    error_message: str = ""


@dataclass(frozen=True, slots=True)
class ImportDisplayRow:
    source_row_number: int
    status: str
    transaction: UnsubmittedTransaction | None
    raw_values: tuple[str, ...]
    error_message: str = ""


@dataclass(frozen=True, slots=True)
class ImportSummary:
    job_id: int
    source_file: str
    total_count: int
    new_count: int
    duplicate_count: int
    error_count: int
    rows: tuple[ImportDisplayRow, ...]


@dataclass(frozen=True, slots=True)
class ImportHistoryItem:
    job_id: int
    imported_at: str
    source_file: str
    total_count: int
    new_count: int
    duplicate_count: int
    error_count: int


@dataclass(frozen=True, slots=True)
class ReceiptImageResult:
    transaction_id: str
    image_paths: tuple[Path, ...]
    doc_irns: tuple[str, ...]
    corp_no: str


@dataclass(frozen=True, slots=True)
class ReceiptFieldCheck:
    field_name: str
    expected_value: str
    detected_value: str
    is_match: bool
    reason: str


@dataclass(frozen=True, slots=True)
class ReceiptValidationResult:
    transaction_id: str
    status: str
    checks: tuple[ReceiptFieldCheck, ...]
    ocr_text: str
    # OCR 전 단계에서 영수증이 없다는 사실도 행 단위 결과로 남긴다.
    # (예: 증빙유무가 #인데 결재선 창에 실제 이미지가 없는 경우)
    reason: str = ""
    rotation_degrees: int = 0
    orientation_ambiguous: bool = False
    orientation_reason: str = ""


@dataclass(frozen=True, slots=True)
class MerchantLookupResult:
    """The external merchant result for a PG transaction, kept only in memory."""

    transaction_id: str
    status: str
    business_number: str = ""
    merchant_name: str = ""
    reason: str = ""
    registration_status: str = "미등록"
    registration_reason: str = ""


@dataclass(frozen=True, slots=True)
class ProcessingEvent:
    """An immutable audit event for one e-Acc transaction."""

    event_id: int
    event_at: str
    transaction_id: str
    approval_number: str
    evidence_date: str
    amount: str
    merchant: str
    status: str
    reason: str
    account_name: str = ""


@dataclass(frozen=True, slots=True)
class MailLogItem:
    log_id: int
    sent_at: str
    recipient_name: str
    recipient_email: str
    department: str
    transaction_count: int
    subject: str
    status: str
    reason: str
    transaction_ids: tuple[str, ...]
    outlook_message_id: str = ""
