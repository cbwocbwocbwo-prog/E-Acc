from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import ROUND_CEILING, Decimal
from typing import Iterable

from .models import UnsubmittedTransaction


SPECIAL_OVERTIME_MEAL_ACCOUNT = "특근자식비"
FIELD_SUPPORT_MEAL_ACCOUNT = "일반복리비-현장지원 (현장대리인 활동지원 식음료대)"

# These are exact account-name matches.  In particular, "일반복리비" must
# not accidentally exclude "일반복리비-현장지원".
EXCLUDED_ACCOUNTS = {
    "회의비": "회의비로 예외처리",
    "부서회의비": "부서회의비로 예외처리",
    "업무회의비": "업무회의비로 예외처리",
    "일반복리비": "일반복리비로 예외처리",
}

VEHICLE_RECEIPT_CATEGORIES = {
    "차량유지비-유류대(법인)": (
        "유류",
        "주유",
        "주유소",
        "휘발유",
        "경유",
        "고급유",
        "디젤",
        "gasoline",
        "diesel",
        "lpg",
    ),
    "차량유지비-주차비": ("주차", "파킹", "parking"),
    "차량유지비-세차비": ("세차", "카워시", "carwash", "wash"),
    "차량유지비-통행료": (
        "통행료",
        "하이패스",
        "한국도로공사",
        "고속도로",
        "highpass",
        "toll",
    ),
}

# Product categories, common Korean labels and representative brand names.
# Matching is deliberately case-insensitive and whitespace-insensitive because
# receipt OCR frequently separates letters or inserts spaces.
ALCOHOL_KEYWORDS = (
    "맥주",
    "소주",
    "와인",
    "위스키",
    "위스키",
    "막걸리",
    "사케",
    "하이볼",
    "칵테일",
    "양주",
    "생맥주",
    "라거",
    "에일",
    "스타우트",
    "ipa",
    "beer",
    "wine",
    "whisky",
    "whiskey",
    "vodka",
    "rum",
    "gin",
    "tequila",
    "brandy",
    "liqueur",
    "cocktail",
    "highball",
    "soju",
    "makgeolli",
    "sake",
    "참이슬",
    "처음처럼",
    "진로",
    "카스",
    "테라",
    "클라우드",
    "필라이트",
    "백세주",
    "복분자",
    "샴페인",
)


@dataclass(frozen=True, slots=True)
class AccountValidationResult:
    """The account-specific rule outcome kept separate from OCR validation."""

    transaction_id: str
    status: str  # 정상 / 정상(i) / 예외 / 검증대기
    reasons: tuple[str, ...] = ()
    required_employee_count: int = 0
    matched_employee_names: tuple[str, ...] = ()
    detected_receipt_categories: tuple[str, ...] = ()

    @property
    def reason_text(self) -> str:
        return " / ".join(self.reasons)


def validate_account_rules(
    transaction: UnsubmittedTransaction,
    employee_names: Iterable[str],
    receipt_text: str | None,
) -> AccountValidationResult:
    """Validate user-defined rules for one account.

    ``receipt_text=None`` means receipt OCR has not run yet; an empty string
    means it has run but yielded no comparable words.  The distinction matters
    because vehicle-account rules treat an unreadable category as normal only
    *after* receipt OCR has completed.
    """
    account_name = _normalize(transaction.account_name)
    description = transaction.description.strip()
    reasons: list[str] = []

    excluded_reason = next(
        (reason for name, reason in EXCLUDED_ACCOUNTS.items() if _normalize(name) == account_name),
        None,
    )
    if excluded_reason:
        return _result(transaction, "예외", (excluded_reason,))

    if account_name == SPECIAL_OVERTIME_MEAL_ACCOUNT and transaction.amount >= Decimal("15000"):
        required = int(
            (transaction.amount / Decimal("15000")).to_integral_value(rounding=ROUND_CEILING)
        )
        matched_names = _matched_employee_names(description, employee_names)
        if len(matched_names) < required:
            reasons.append(
                "특근자식비 사용인원 불충족"
                f" (필요 {required}명 / 확인 {len(matched_names)}명)"
            )
        if reasons:
            return _result(
                transaction,
                "예외",
                tuple(reasons),
                required_employee_count=required,
                matched_employee_names=matched_names,
            )
        names = ", ".join(matched_names)
        return _result(
            transaction,
            "정상",
            (
                "특근자식비 사용인원 확인"
                f" (사용금액 {transaction.amount:,.0f}원 / 필요 {required}명 / "
                f"인정 {len(matched_names)}명: {names})",
            ),
            required_employee_count=required,
            matched_employee_names=matched_names,
        )

    if account_name == _normalize(FIELD_SUPPORT_MEAL_ACCOUNT):
        if transaction.amount > Decimal("200000"):
            reasons.append("현장대리인 사용금액 초과")
        if description:
            reasons.append("불필요한 적요 작성")
        if receipt_text is None:
            pending_reason = "영수증 OCR 검증 대기"
            return _result(
                transaction,
                "예외" if reasons else "검증대기",
                tuple((*reasons, pending_reason)),
            )
        alcohol_matches = _matching_keywords(receipt_text, ALCOHOL_KEYWORDS)
        if alcohol_matches:
            reasons.append("영수증에 주류포함" + f" ({', '.join(alcohol_matches)})")
        if reasons:
            return _result(transaction, "예외", tuple(reasons))
        return _result(
            transaction,
            "정상",
            (
                "현장대리인 활동지원 식음료대 확인"
                f" (사용금액 {transaction.amount:,.0f}원 / 적요 없음 / 주류 키워드 없음)",
            ),
        )

    expected_category = next(
        (name for name in VEHICLE_RECEIPT_CATEGORIES if _normalize(name) == account_name),
        None,
    )
    if expected_category is not None:
        if receipt_text is None:
            return _result(transaction, "검증대기", ("영수증 OCR 검증 대기",))
        detected_categories = _detect_vehicle_categories(receipt_text)
        expected_name = expected_category
        # 전자영수증에는 실제 결제 유형과 무관하게 "주유/자동차주차요금"처럼
        # 복수의 공통 분류어가 함께 찍히는 경우가 있다. 이때 목표 계정의
        # 키워드가 OCR 본문에 있으면 해당 계정 영수증으로 인정하되, 혼합
        # 키워드였다는 사실은 정상(i)로 구분해 결과 화면에 남긴다.
        if expected_name in detected_categories:
            if detected_categories == (expected_name,):
                return _result(
                    transaction,
                    "정상",
                    (
                        f"{expected_name} 영수증 키워드 확인"
                        f" ({', '.join(detected_categories)})",
                    ),
                    detected_receipt_categories=detected_categories,
                )
            return _result(
                transaction,
                "정상(i)",
                (
                    f"{expected_name} OCR 키워드 포함 확인"
                    f" (함께 감지: {', '.join(detected_categories)})",
                ),
                detected_receipt_categories=detected_categories,
            )
        if detected_categories:
            return _result(
                transaction,
                "예외",
                ("계정과 상이한 영수증 첨부",),
                detected_receipt_categories=detected_categories,
            )
        # No recognizable comparison word is intentionally considered normal.
        audit_reason = f"{expected_name}: 비교 키워드 미검출(정상 간주)"
        return _result(
            transaction,
            "정상",
            (audit_reason,),
            detected_receipt_categories=detected_categories,
        )

    return _result(transaction, "정상", ("별도 계정별 검증 규칙 대상 아님",))


def _result(
    transaction: UnsubmittedTransaction,
    status: str,
    reasons: tuple[str, ...] = (),
    *,
    required_employee_count: int = 0,
    matched_employee_names: tuple[str, ...] = (),
    detected_receipt_categories: tuple[str, ...] = (),
) -> AccountValidationResult:
    return AccountValidationResult(
        transaction_id=transaction.transaction_id,
        status=status,
        reasons=reasons,
        required_employee_count=required_employee_count,
        matched_employee_names=matched_employee_names,
        detected_receipt_categories=detected_receipt_categories,
    )


def _matched_employee_names(description: str, employee_names: Iterable[str]) -> tuple[str, ...]:
    # Count a supplied employee name once even if it was typed more than once.
    normalized_description = _normalize(description)
    unique_names = {_normalize(name) for name in employee_names if _normalize(name)}
    return tuple(sorted(name for name in unique_names if name in normalized_description))


def _detect_vehicle_categories(receipt_text: str) -> tuple[str, ...]:
    detected: list[str] = []
    for account_name, keywords in VEHICLE_RECEIPT_CATEGORIES.items():
        if _matching_keywords(receipt_text, keywords):
            detected.append(account_name)
    return tuple(detected)


def _matching_keywords(text: str, keywords: Iterable[str]) -> tuple[str, ...]:
    normalized_text = _normalize(text)
    return tuple(keyword for keyword in keywords if _normalize(keyword) in normalized_text)


def _normalize(value: str) -> str:
    return re.sub(r"\s+", "", value).casefold()
