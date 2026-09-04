from __future__ import annotations

import hashlib
import re
from decimal import Decimal, InvalidOperation
from html.parser import HTMLParser
from pathlib import Path
from typing import Sequence

from .models import EXPECTED_UNSUBMITTED_HEADERS, ParsedRow, UnsubmittedTransaction


class WorkbookFormatError(ValueError):
    """Raised when the downloaded file is not the expected E-Acc HTML table."""


class GridRowFormatError(ValueError):
    """Raised when a live e-Acc grid row cannot be converted safely."""


class _FirstTableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._inside_target_table = False
        self._target_finished = False
        self._table_depth = 0
        self._inside_row = False
        self._inside_cell = False
        self._cell_parts: list[str] = []
        self._row: list[str] = []
        self.rows: list[list[str]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag == "table" and not self._target_finished:
            if not self._inside_target_table:
                self._inside_target_table = True
                self._table_depth = 1
            else:
                self._table_depth += 1
            return
        if not self._inside_target_table:
            return
        if tag == "tr" and self._table_depth == 1:
            self._inside_row = True
            self._row = []
        elif tag in {"td", "th"} and self._inside_row:
            self._inside_cell = True
            self._cell_parts = []
        elif tag == "br" and self._inside_cell:
            self._cell_parts.append(" ")

    def handle_data(self, data: str) -> None:
        if self._inside_cell:
            self._cell_parts.append(data)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if not self._inside_target_table:
            return
        if tag in {"td", "th"} and self._inside_cell:
            value = re.sub(r"\s+", " ", "".join(self._cell_parts)).strip()
            self._row.append(value)
            self._inside_cell = False
            self._cell_parts = []
        elif tag == "tr" and self._inside_row:
            if self._row:
                self.rows.append(self._row)
            self._inside_row = False
            self._row = []
        elif tag == "table":
            self._table_depth -= 1
            if self._table_depth == 0:
                self._inside_target_table = False
                self._target_finished = True


def _decode_html(path: Path) -> str:
    raw = path.read_bytes()
    if not raw:
        raise WorkbookFormatError("파일이 비어 있습니다.")
    for encoding in ("utf-8-sig", "cp949"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise WorkbookFormatError("UTF-8 또는 한글 Windows 인코딩으로 파일을 읽을 수 없습니다.")


def _normalize_date(value: str, field_name: str) -> str:
    digits = re.sub(r"\D", "", value)
    if len(digits) != 8:
        raise ValueError(f"{field_name} 형식 오류: {value or '(빈 값)'}")
    year, month, day = int(digits[:4]), int(digits[4:6]), int(digits[6:8])
    try:
        from datetime import date

        normalized = date(year, month, day)
    except ValueError as exc:
        raise ValueError(f"{field_name} 유효하지 않은 날짜: {value}") from exc
    return normalized.isoformat()


def _normalize_amount(value: str) -> Decimal:
    cleaned = re.sub(r"[^0-9.\-]", "", value)
    try:
        amount = Decimal(cleaned)
    except InvalidOperation as exc:
        raise ValueError(f"사용금액 형식 오류: {value or '(빈 값)'}") from exc
    if amount != amount.to_integral_value():
        raise ValueError(f"사용금액은 원 단위 정수여야 합니다: {value}")
    return amount


def _required(value: str, field_name: str) -> str:
    normalized = value.strip()
    if not normalized:
        raise ValueError(f"{field_name} 값이 없습니다.")
    return normalized


def _transaction_from_values(source_row_number: int, values: list[str]) -> UnsubmittedTransaction:
    card_number = _required(values[2], "카드번호")
    approval_number = _required(values[4], "승인번호")
    evidence_date = _normalize_date(values[8], "증빙일자")
    amount = _normalize_amount(values[12])
    posting_date = _normalize_date(values[7], "전기일")
    card_issue_date = _normalize_date(values[16], "카드발급일")

    canonical_key = "|".join((card_number, approval_number, evidence_date, str(amount)))
    digest = hashlib.sha256(canonical_key.encode("utf-8")).hexdigest()

    return UnsubmittedTransaction(
        source_row_number=source_row_number,
        voucher_status=values[0],
        evidence_status=values[1],
        card_number=card_number,
        card_holder=values[3],
        approval_number=approval_number,
        weekday=values[5],
        usage_time=values[6],
        posting_date=posting_date,
        evidence_date=evidence_date,
        account_name=values[9],
        description=values[10],
        merchant=values[11],
        amount=amount,
        tax_code=values[13],
        business_type=values[14],
        cost_center=values[15],
        card_issue_date=card_issue_date,
        division=values[17],
        actual_merchant_name=values[18],
        # values[19]는 검색 아이콘 이미지 경로이므로 저장하지 않는다.
        actual_merchant_code=values[20],
        transaction_key=digest,
        transaction_id=f"UNS-{digest[:16].upper()}",
    )


def transaction_from_grid_row(
    source_row_number: int,
    headers: Sequence[str],
    values: Sequence[str],
) -> UnsubmittedTransaction:
    """Convert one visible DHTMLX row into the established Excel data order."""
    if len(headers) != len(values):
        raise GridRowFormatError(
            f"e-Acc 그리드 열 개수 불일치: 제목 {len(headers)}개 / 값 {len(values)}개"
        )

    aliases = {
        "SLIP_STATUS": "전표상태", "VOUCHER_STATUS": "전표상태",
        "EVIDENCE_STATUS": "증빙유무", "EVIDENCE_YN": "증빙유무",
        "CARD_NO": "카드번호", "CARD_HOLDER": "카드소지자", "APPR_NO": "승인번호",
        "WEEKDAY": "요일", "USE_TIME": "이용시간", "POSTING_DATE": "전기일",
        "POST_DATE": "전기일", "BLDAT": "증빙일자", "ACCOUNT_NAME": "계정명",
        "DESCRIPTION": "적요", "MERCHANT": "거래처", "USED_AMT": "사용금액",
        "TAX_CODE": "세금코드", "BUSINESS_TYPE": "업종", "COST_CENTER": "코스트센터",
        "CARD_ISSUE_DATE": "카드발급일", "DIVISION": "본부",
        "ACTUAL_MERCHANT_NAME": "실구매처명", "SELECTED": "선택",
        "ACTUAL_MERCHANT_CODE": "실구매처코드",
    }
    values_by_header: dict[str, str] = {}
    for header, value in zip(headers, values, strict=True):
        normalized = re.sub(r"\s+", "", str(header)).strip()
        canonical = aliases.get(normalized.upper(), normalized)
        if canonical:
            values_by_header[canonical] = str(value).strip()

    missing = [header for header in EXPECTED_UNSUBMITTED_HEADERS if header not in values_by_header]
    # 선택 아이콘과 실구매처코드는 화면 개편에 따라 노출되지 않을 수 있으며
    # 거래 식별에는 쓰지 않는다.
    for optional_header in ("선택", "실구매처코드"):
        if optional_header in missing:
            values_by_header[optional_header] = ""
            missing.remove(optional_header)
    if missing:
        raise GridRowFormatError("e-Acc 그리드에서 필요한 열을 찾지 못했습니다: " + ", ".join(missing))

    try:
        return _transaction_from_values(
            source_row_number,
            [values_by_header[header] for header in EXPECTED_UNSUBMITTED_HEADERS],
        )
    except ValueError as exc:
        raise GridRowFormatError(str(exc)) from exc


def parse_unsubmitted_xls(file_path: str | Path) -> tuple[ParsedRow, ...]:
    path = Path(file_path)
    if path.suffix.lower() != ".xls":
        raise WorkbookFormatError("미상신내역 다운로드 파일(.xls)을 선택해 주세요.")
    if not path.is_file():
        raise WorkbookFormatError(f"파일을 찾을 수 없습니다: {path}")

    parser = _FirstTableParser()
    parser.feed(_decode_html(path))
    parser.close()

    if not parser.rows:
        raise WorkbookFormatError("HTML 표를 찾을 수 없습니다. e-Accounting에서 받은 파일인지 확인해 주세요.")

    actual_headers = tuple(parser.rows[0])
    if actual_headers != EXPECTED_UNSUBMITTED_HEADERS:
        missing = [h for h in EXPECTED_UNSUBMITTED_HEADERS if h not in actual_headers]
        extra = [h for h in actual_headers if h not in EXPECTED_UNSUBMITTED_HEADERS]
        details: list[str] = []
        if len(actual_headers) != len(EXPECTED_UNSUBMITTED_HEADERS):
            details.append(f"열 개수 {len(actual_headers)}개(예상 21개)")
        if missing:
            details.append("누락: " + ", ".join(missing))
        if extra:
            details.append("예상 밖: " + ", ".join(extra))
        if not details:
            details.append("열 순서가 변경됨")
        raise WorkbookFormatError("미상신내역 열 구조가 다릅니다. " + "; ".join(details))

    parsed: list[ParsedRow] = []
    for source_row_number, values in enumerate(parser.rows[1:], start=2):
        raw_values = tuple(values)
        if len(values) != len(EXPECTED_UNSUBMITTED_HEADERS):
            parsed.append(
                ParsedRow(
                    source_row_number=source_row_number,
                    transaction=None,
                    raw_values=raw_values,
                    error_message=f"열 개수 오류: {len(values)}개(예상 21개)",
                )
            )
            continue
        try:
            transaction = _transaction_from_values(source_row_number, values)
        except ValueError as exc:
            parsed.append(
                ParsedRow(
                    source_row_number=source_row_number,
                    transaction=None,
                    raw_values=raw_values,
                    error_message=str(exc),
                )
            )
        else:
            parsed.append(
                ParsedRow(
                    source_row_number=source_row_number,
                    transaction=transaction,
                    raw_values=raw_values,
                )
            )
    return tuple(parsed)
