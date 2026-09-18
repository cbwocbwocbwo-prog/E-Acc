from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import ROUND_CEILING, Decimal
from typing import Iterable

from .models import UnsubmittedTransaction


SPECIAL_OVERTIME_MEAL_ACCOUNT = "특근자식비"
SPECIAL_OVERTIME_MEAL_MAX_AMOUNT = Decimal("90000")
SPECIAL_OVERTIME_MEAL_EMPLOYEE_CHECK_AMOUNT = Decimal("15000")
FIELD_SUPPORT_MEAL_ACCOUNT = "일반복리비-현장지원"
FIELD_SUPPORT_MEAL_DESCRIPTION = "현장대리인 활동지원 식음료대"
FIELD_SUPPORT_BEVERAGE_ONLY_DESCRIPTIONS = (
    "현장업무 수행시 식음료",
    "(전담)SIte기초설계 수행시 식음료",
)

# These are exact account-name matches.  In particular, "일반복리비" must
# not accidentally exclude "일반복리비-현장지원".
EXCLUDED_ACCOUNTS = {
    "회의비": "회의비로 예외처리",
    "부서회의비": "부서회의비로 예외처리",
    "부서회의비(캔미팅)": "부서회의비(캔미팅)로 예외처리",
    "업무회의비": "업무회의비로 예외처리",
    "일반복리비-건강지원": "일반복리비-건강지원로 예외처리",
    "일반복리비-경조/화환": "일반복리비-경조/화환로 예외처리",
    "일반복리비-경조사지원비": "일반복리비-경조사지원비로 예외처리",
    "일반복리비-동호회": "일반복리비-동호회로 예외처리",
    "일반복리비-기타": "일반복리비-기타로 예외처리",
    "일반복리비-급여성복리비": "일반복리비-급여성복리비로 예외처리",
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

# 음료 전용 적요에서는 품목명과 금액·수량이 함께 읽힌 줄만 비교한다.
# 상호명에 '베이커리'가 포함되는 것만으로는 예외처리하지 않기 위함이다.
BEVERAGE_KEYWORDS = (
    "음료", "커피", "아메리카노", "에스프레소", "espresso", "라떼", "latte",
    "카푸치노", "cappuccino", "모카", "mocha", "콜드브루", "coldbrew", "프라페",
    "frappe", "티", "tea", "녹차", "홍차", "밀크티", "주스", "juice", "에이드",
    "ade", "레모네이드", "lemonade", "스무디", "smoothie", "탄산", "콜라", "cola",
    "소다", "soda", "생수", "광천수", "삼다수", "water", "요구르트", "요거트",
)
_BEVERAGE_FUZZY_PATTERNS = (
    # Windows OCR may read "아메리카노" as "아대|리카노" on folded thermal
    # receipts.  This is deliberately narrow and is only used in the
    # product-name crop for the beverage-only descriptions.
    ("아메리카노", re.compile(r"아.{0,2}리카노")),
)
FOOD_KEYWORDS = (
    "빵", "케이크", "케익", "베이글", "쿠키", "마카롱", "머핀", "도넛", "크로플",
    "샌드위치", "토스트", "pastry", "cake", "bread", "식사", "밥", "도시락", "김밥",
    "라면", "국수", "햄버거", "치킨", "피자", "떡볶이", "샐러드", "과자",
)
_ITEM_VALUE_OR_QUANTITY_PATTERN = re.compile(
    r"(?:\d{1,3}(?:[,.]\d{3})+|\d+)\s*(?:원|개|잔|병|캔|팩|ea)\b|"
    r"(?:\d{1,3}(?:[,.]\d{3})+|\d+)\s*(?=$|\s)",
    re.IGNORECASE,
)
_RECEIPT_SUMMARY_KEYWORDS = (
    "합계", "총액", "결제", "승인", "부가세", "공급가", "과세", "면세", "할부",
    "카드번호", "거래일시", "결제일시", "영수증번호", "번호", "금액", "사업자",
    "주소", "전화", "대표", "가맹점", "매장", "상호",
)


@dataclass(frozen=True, slots=True)
class AccountValidationResult:
    """The account-specific rule outcome kept separate from OCR validation."""

    transaction_id: str
    status: str  # 정상 / 정상(i) / 정상(w) / 예외 / 검증대기
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

    if (
        account_name == SPECIAL_OVERTIME_MEAL_ACCOUNT
        and transaction.amount > SPECIAL_OVERTIME_MEAL_MAX_AMOUNT
    ):
        return _result(
            transaction,
            "예외",
            (
                "특근자식비 사용금액 초과"
                f" (사용금액 {transaction.amount:,.0f}원 / 한도 90,000원)",
            ),
        )

    # 15,000원까지는 적요에 직원 이름을 적지 않아도 허용한다.  따라서
    # 인원 검증은 15,000원을 *초과*한 경우부터 시작한다.
    if (
        account_name == SPECIAL_OVERTIME_MEAL_ACCOUNT
        and transaction.amount > SPECIAL_OVERTIME_MEAL_EMPLOYEE_CHECK_AMOUNT
    ):
        required = int(
            (transaction.amount / SPECIAL_OVERTIME_MEAL_EMPLOYEE_CHECK_AMOUNT).to_integral_value(
                rounding=ROUND_CEILING
            )
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
        normalized_description = _normalize(description)
        expected_description = _normalize(FIELD_SUPPORT_MEAL_DESCRIPTION)
        beverage_only_descriptions = {
            _normalize(value) for value in FIELD_SUPPORT_BEVERAGE_ONLY_DESCRIPTIONS
        }
        # 적요는 입력값의 완전 일치 여부를 검증하지 않는다. 정해진 문구가
        # 포함되어 있으면 뒤에 "_음료구매" 등 보조 설명이 붙어도 해당 검증
        # 유형으로 분류한다.
        if expected_description in normalized_description:
            if transaction.amount > Decimal("200000"):
                reasons.append("현장대리인 사용금액 초과")
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
                    f" (사용금액 {transaction.amount:,.0f}원 / 적요 일치 / 주류 키워드 없음)",
                ),
            )

        if any(value in normalized_description for value in beverage_only_descriptions):
            return _validate_beverage_only_field_support_receipt(
                transaction,
                description,
                receipt_text,
            )

        # 세 가지 지정 문구에 해당하지 않는 적요 자체는 결재 제외 사유가 아니다.
        # 적요가 검증 대상을 분류할 뿐, 적요의 정확한 일치를 요구하지 않는다.
        return _result(transaction, "정상", ("현장지원 적요: 별도 세부 검증 대상 아님",))

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


def _validate_beverage_only_field_support_receipt(
    transaction: UnsubmittedTransaction,
    description: str,
    receipt_text: str | None,
) -> AccountValidationResult:
    """Accept only beverage item lines for the two beverage-only descriptions.

    The user explicitly allows an unreadable or unclassified item section.
    That permissive case is recorded as ``정상(w)`` so it remains distinguishable from a
    receipt whose priced/quantified product lines were actually identified as
    beverages.
    """
    reasons: list[str] = []
    if transaction.amount > Decimal("200000"):
        reasons.append("현장대리인 사용금액 초과")
    if receipt_text is None:
        return _result(
            transaction,
            "예외" if reasons else "검증대기",
            tuple((*reasons, "영수증 OCR 검증 대기")),
        )

    # 날짜·승인번호·금액 검증에 이미 사용한 OCR 원문에서만 품목을 판정한다.
    # 음료 품목 판정을 위해 OCR을 별도로 재실행하지 않는다.
    item_lines = _receipt_item_lines(receipt_text)
    if not item_lines:
        if reasons:
            return _result(transaction, "예외", tuple(reasons))
        return _result(
            transaction,
            "정상(w)",
            (
                "음료전용 적요: 품목 OCR 미검출 → 정상(w)",
            ),
        )

    alcohol_matches: list[str] = []
    food_matches: list[str] = []
    beverage_matches: list[str] = []
    beverage_line_count = 0
    has_unclassified_item = False
    for line in item_lines:
        alcohol = _matching_keywords(line, ALCOHOL_KEYWORDS)
        food = _matching_keywords(line, FOOD_KEYWORDS)
        beverage = _matching_beverage_keywords(line)
        if alcohol:
            alcohol_matches.extend(alcohol)
        if food:
            food_matches.extend(food)
        if alcohol or food:
            continue
        if beverage:
            beverage_line_count += 1
            beverage_matches.extend(beverage)
        else:
            has_unclassified_item = True

    if alcohol_matches:
        reasons.append(
            "음료전용 적요: 주류 품목 감지"
            f"({', '.join(dict.fromkeys(alcohol_matches))}) → 예외"
        )
    if food_matches:
        reasons.append(
            "음료전용 적요: 음식 품목 감지"
            f"({', '.join(dict.fromkeys(food_matches))}) → 예외"
        )
    if reasons:
        return _result(transaction, "예외", tuple(reasons))
    if has_unclassified_item:
        return _result(
            transaction,
            "정상(w)",
            (
                "음료전용 적요: 품목 미분류 → 정상(w)",
            ),
        )
    if beverage_line_count:
        return _result(
            transaction,
            "정상",
            (
                "음료전용 적요: 음료 품목 확인"
                f"({', '.join(dict.fromkeys(beverage_matches))}) → 정상",
            ),
        )

    # Defensive fallback: item_lines currently always produce a beverage or
    # unclassified result, but keep the approval-safe state explicit.
    return _result(
        transaction,
        "정상(w)",
        ("음료전용 적요: 품목 OCR 미검출 → 정상(w)",),
    )


def _receipt_item_lines(receipt_text: str) -> tuple[str, ...]:
    """Return readable priced/quantified product lines, excluding receipt totals."""
    lines: list[str] = []
    for raw_line in receipt_text.splitlines():
        line = " ".join(raw_line.split())
        normalized_line = _normalize(line)
        if not line or not _ITEM_VALUE_OR_QUANTITY_PATTERN.search(line):
            continue
        if _is_receipt_summary_line(normalized_line):
            continue
        # A product line must contain a readable label in addition to its
        # number. Pure numeric OCR fragments are not treated as a food item.
        label = re.sub(r"[\d\s,.:/()\-]+", "", line)
        if len(label) < 2:
            continue
        lines.append(line)
    return tuple(dict.fromkeys(lines))


def _is_receipt_summary_line(normalized_line: str) -> bool:
    """Exclude only a pure payment/header line, not a product plus total line."""
    remainder = normalized_line
    for keyword in _RECEIPT_SUMMARY_KEYWORDS:
        remainder = remainder.replace(_normalize(keyword), "")
    remainder = re.sub(r"[\d\s,.:/()\-원개잔병캔팩]+", "", remainder)
    return len(remainder) < 2


def _matching_beverage_keywords(text: str) -> tuple[str, ...]:
    """Return exact and narrowly fuzzy beverage names from a product row."""
    normalized_text = _normalize(text)
    exact = [keyword for keyword in BEVERAGE_KEYWORDS if _normalize(keyword) in normalized_text]
    fuzzy = [name for name, pattern in _BEVERAGE_FUZZY_PATTERNS if pattern.search(normalized_text)]
    return tuple(dict.fromkeys((*exact, *fuzzy)))


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
