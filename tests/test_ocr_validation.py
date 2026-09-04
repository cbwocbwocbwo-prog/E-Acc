from __future__ import annotations

import unittest
from dataclasses import replace
from decimal import Decimal

from eacc_app.models import UnsubmittedTransaction
from eacc_app.ocr_validation import evaluate_ocr_text


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
        account_name="차량유지비",
        description="테스트",
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


class OcrValidationTests(unittest.TestCase):
    def test_all_three_fields_match_and_leading_zero_is_preserved(self) -> None:
        result = evaluate_ocr_text(
            _transaction(),
            "승인NO 023200102 승인일시 2026-08-28 합계금액 45,000원",
        )

        self.assertEqual("정상", result.status)
        self.assertTrue(all(check.is_match for check in result.checks))

    def test_missing_approval_number_is_reported_as_abnormal(self) -> None:
        result = evaluate_ocr_text(
            _transaction(),
            "승인NO 99999999 승인일시 260828083815 합계금액 45,000원",
        )

        self.assertEqual("이상", result.status)
        approval = result.checks[0]
        self.assertFalse(approval.is_match)
        self.assertIn("023200102", approval.reason)
        self.assertTrue(result.checks[1].is_match)
        self.assertTrue(result.checks[2].is_match)

    def test_segmented_approval_and_supply_vat_sum_are_accepted(self) -> None:
        transaction = replace(
            _transaction(),
            approval_number="18286514",
            evidence_date="2026-08-26",
            amount=Decimal("63000"),
        )
        result = evaluate_ocr_text(
            transaction,
            "승인번호 1828 6514 거래일시 2026/08/26 공급가액 57,273 부가세 5,727",
        )

        self.assertEqual("정상", result.status)
        self.assertTrue(all(check.is_match for check in result.checks))
        self.assertEqual("공급가액과 부가세의 합계가 일치", result.checks[2].reason)

    def test_slash_approval_and_korean_or_damaged_dates_are_accepted(self) -> None:
        transaction = replace(
            _transaction(),
            approval_number="37097115",
            evidence_date="2026-08-24",
            amount=Decimal("10000"),
        )
        slashed = evaluate_ocr_text(transaction, "승인번호 3/09/115 이용일 20260824 합계 10,000")
        self.assertEqual("정상", slashed.status)

        transaction = replace(transaction, evidence_date="2026-08-20")
        korean_date = evaluate_ocr_text(
            transaction,
            "승인번호 37097115 승인일 2026년 08월20일 합계 10,000",
        )
        self.assertTrue(korean_date.checks[1].is_match)
        damaged_date = evaluate_ocr_text(
            replace(transaction, evidence_date="2026-08-21"),
            "승인번호 37097115 승인일 2026기821 합계 10,000",
        )
        self.assertTrue(damaged_date.checks[1].is_match)

    def test_date_with_split_zero_from_photo_ocr_is_accepted(self) -> None:
        transaction = replace(_transaction(), evidence_date="2026-08-27")
        result = evaluate_ocr_text(
            transaction,
            "승인번호 023200102 판매일 2,)26 08-27 합계 45,000",
        )
        self.assertTrue(result.checks[1].is_match)

    def test_month_day_only_is_accepted_only_in_payment_context(self) -> None:
        transaction = replace(_transaction(), evidence_date="2026-08-27")
        result = evaluate_ocr_text(
            transaction,
            "승인번호 023200102 POS 03 08 27 결제금액 45,000",
        )
        self.assertTrue(result.checks[1].is_match)
        self.assertIn("월·일 OCR 보정", result.checks[1].reason)

        no_context = evaluate_ocr_text(
            transaction,
            "승인번호 023200102 품목코드 08 27 합계 45,000",
        )
        self.assertFalse(no_context.checks[1].is_match)

    def test_single_digit_approval_noise_is_accepted_only_in_card_context(self) -> None:
        transaction = replace(
            _transaction(),
            approval_number="06106101",
            evidence_date="2026-08-27",
            amount=Decimal("3300"),
        )
        result = evaluate_ocr_text(
            transaction,
            "[061061011] KICC결제 승인일 2026-08-27 합계 3,300",
        )
        self.assertEqual("정상", result.status)
        self.assertIn("한 자리 OCR 보정", result.checks[0].reason)

        no_context = evaluate_ocr_text(
            transaction,
            "참조번호 061061011 판매일 2026-08-27 합계 3,300",
        )
        self.assertFalse(no_context.checks[0].is_match)

    def test_percent_sign_in_kicc_approval_context_is_repaired_conservatively(self) -> None:
        transaction = replace(
            _transaction(),
            approval_number="15404901",
            evidence_date="2026-08-26",
            amount=Decimal("47000"),
        )
        result = evaluate_ocr_text(
            transaction,
            "결제금액 47,000 [15404%11] KICC로제출 POS 2026-08-26",
        )
        self.assertEqual("정상", result.status)
        self.assertTrue(result.checks[0].is_match)

        comma_damaged = evaluate_ocr_text(
            transaction,
            "결제금액 47,000 [5404901] KIC,C로제출 POS 2026-08-26",
        )
        self.assertTrue(comma_damaged.checks[0].is_match)

    def test_gas_station_approval_number_lookalikes_are_repaired_only_after_label(self) -> None:
        transaction = replace(_transaction(), approval_number="22502602")
        result = evaluate_ocr_text(
            transaction,
            "[시 카드]\n인 번호 : 牙25d2名02\n승인 금 : 70,000원",
        )

        self.assertTrue(result.checks[0].is_match)
        self.assertEqual("22502602", result.checks[0].detected_value)


if __name__ == "__main__":
    unittest.main()
