from __future__ import annotations

import unittest
from dataclasses import replace
from decimal import Decimal

from eacc_app.account_validation import validate_account_rules
from eacc_app.models import UnsubmittedTransaction


def _transaction() -> UnsubmittedTransaction:
    return UnsubmittedTransaction(
        source_row_number=2,
        voucher_status="임시저장",
        evidence_status="#",
        card_number="5587380002950208",
        card_holder="테스트",
        approval_number="023200102",
        weekday="금요일",
        usage_time="08:38:14",
        posting_date="2026-08-28",
        evidence_date="2026-08-28",
        account_name="테스트계정",
        description="",
        merchant="테스트 거래처",
        amount=Decimal("45000"),
        tax_code="M0",
        business_type="일반",
        cost_center="테스트",
        card_issue_date="2025-07-10",
        division="테스트",
        actual_merchant_name="",
        actual_merchant_code="",
        transaction_key="key",
        transaction_id="UNS-TEST",
    )


class AccountValidationTests(unittest.TestCase):
    def test_overtime_meal_requires_ceiling_of_15000_per_unique_employee(self) -> None:
        transaction = replace(
            _transaction(),
            account_name="특근자식비",
            amount=Decimal("50000"),
            description="김가나, 이다라, 박마바",
        )
        result = validate_account_rules(transaction, ("김가나", "이다라", "박마바", "최사아"), None)

        self.assertEqual("예외", result.status)
        self.assertEqual(4, result.required_employee_count)
        self.assertIn("특근자식비 사용인원 불충족", result.reason_text)

        accepted = validate_account_rules(
            replace(transaction, description="김가나, 이다라, 박마바, 최사아"),
            ("김가나", "이다라", "박마바", "최사아"),
            None,
        )
        self.assertEqual("정상", accepted.status)
        self.assertEqual(4, len(accepted.matched_employee_names))

    def test_vehicle_receipt_is_only_rejected_when_a_different_category_is_detected(self) -> None:
        transaction = replace(_transaction(), account_name="차량유지비-유류대(법인)")

        pending = validate_account_rules(transaction, (), None)
        self.assertEqual("검증대기", pending.status)
        no_comparable_word = validate_account_rules(transaction, (), "카드 결제 완료")
        self.assertEqual("정상", no_comparable_word.status)
        fuel = validate_account_rules(transaction, (), "주유소 휘발유 35,000원")
        self.assertEqual("정상", fuel.status)
        parking = validate_account_rules(transaction, (), "주차요금 5,000원")
        self.assertEqual("예외", parking.status)
        self.assertEqual("계정과 상이한 영수증 첨부", parking.reason_text)

        toll_receipt = validate_account_rules(
            replace(transaction, account_name="차량유지비-통행료"),
            (),
            "한국도로공사 고속도로 이용금액 5,000원",
        )
        self.assertEqual("정상", toll_receipt.status)
        wrong_account = validate_account_rules(
            transaction,
            (),
            "한국도로공사 고속도로 이용금액 5,000원",
        )
        self.assertEqual("예외", wrong_account.status)

    def test_only_exact_field_support_meal_account_has_amount_description_and_alcohol_rules(self) -> None:
        transaction = replace(
            _transaction(),
            account_name="일반복리비-현장지원 (현장대리인 활동지원 식음료대)",
            amount=Decimal("200001"),
            description="현장 식사",
        )
        result = validate_account_rules(transaction, (), "카스 맥주 1병")
        self.assertEqual("예외", result.status)
        self.assertIn("현장대리인 사용금액 초과", result.reason_text)
        self.assertIn("불필요한 적요 작성", result.reason_text)
        self.assertIn("영수증에 주류포함", result.reason_text)

        similar = validate_account_rules(
            replace(transaction, account_name="일반복리비-현장지원"), (), "카스 맥주 1병"
        )
        self.assertEqual("정상", similar.status)

    def test_general_welfare_exact_match_is_excluded_without_excluding_field_support_account(self) -> None:
        excluded = validate_account_rules(replace(_transaction(), account_name="일반복리비"), (), None)
        self.assertEqual("예외", excluded.status)
        self.assertEqual("일반복리비로 예외처리", excluded.reason_text)

        field_support = validate_account_rules(
            replace(_transaction(), account_name="일반복리비-현장지원"), (), None
        )
        self.assertEqual("정상", field_support.status)


if __name__ == "__main__":
    unittest.main()
