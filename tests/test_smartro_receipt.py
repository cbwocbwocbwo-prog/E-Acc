from __future__ import annotations

import unittest
from decimal import Decimal

from eacc_app.models import UnsubmittedTransaction
from eacc_app.ocr_validation import evaluate_ocr_text
from eacc_app.smartro_receipt import merge_smartro_approval_result


def _transaction() -> UnsubmittedTransaction:
    return UnsubmittedTransaction(
        source_row_number=2,
        voucher_status="임시저장",
        evidence_status="#",
        card_number="4336920000005000",
        card_holder="테스트",
        approval_number="15007711",
        weekday="화요일",
        usage_time="17:48:35",
        posting_date="2026-09-02",
        evidence_date="2026-09-02",
        account_name="차량유지비-유류대(법인)",
        description="",
        merchant="전기차충전(비)_스마트로",
        amount=Decimal("7115"),
        tax_code="M0",
        business_type="전기차충전소",
        cost_center="테스트",
        card_issue_date="2025-07-10",
        division="테스트",
        actual_merchant_name="",
        actual_merchant_code="",
        transaction_key="key",
        transaction_id="UNS-SMARTRO",
    )


class SmartroReceiptTests(unittest.TestCase):
    def test_focused_approval_match_completes_otherwise_normal_result(self) -> None:
        transaction = _transaction()
        base = evaluate_ocr_text(
            transaction,
            "거래일시 2026/09/02 17:48:35 합계 7,115원 사업자번호 105-87-79517",
        )
        self.assertEqual("이상", base.status)
        result = merge_smartro_approval_result(transaction, base, "승인번호 15007711")
        self.assertEqual("정상", result.status)
        self.assertTrue(result.checks[0].is_match)
        self.assertIn("SmartroPAY", result.checks[0].reason)

    def test_unconfirmed_focused_number_never_changes_result(self) -> None:
        transaction = _transaction()
        base = evaluate_ocr_text(transaction, "거래일시 2026/09/02 합계 7,115원")
        result = merge_smartro_approval_result(transaction, base, "승인번호 99999999")
        self.assertEqual("이상", result.status)
        self.assertFalse(result.checks[0].is_match)
