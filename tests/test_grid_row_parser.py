from __future__ import annotations

import unittest

from eacc_app.parser import GridRowFormatError, transaction_from_grid_row


class GridRowParserTests(unittest.TestCase):
    def test_korean_live_grid_headers_create_same_transaction_identity(self) -> None:
        headers = (
            "전표상태", "증빙유무", "카드번호", "카드소지자", "승인번호", "요일", "이용시간",
            "전기일", "증빙일자", "계정명", "적요", "거래처", "사용금액", "세금코드", "업종",
            "코스트센터", "카드발급일", "본부", "실구매처명", "선택", "실구매처코드",
        )
        values = (
            "임시저장", "#", "5587380002701692", "유경상", "15908100", "금요일", "12:42:53",
            "20260828", "20260828", "일반복리비-현장지원", "(전담)SIte기초설계 수행시 식음료_두류중흥",
            "이디야커피대구텍스빌점", "8700", "I2", "커피전문점", "경북치국설계팀", "20250710",
            "치국설계", "", "", "",
        )

        transaction = transaction_from_grid_row(1, headers, values)

        self.assertEqual("15908100", transaction.approval_number)
        self.assertEqual("2026-08-28", transaction.evidence_date)
        self.assertEqual("8700", str(transaction.amount))
        self.assertTrue(transaction.transaction_id.startswith("UNS-"))

    def test_missing_required_live_grid_header_is_rejected(self) -> None:
        with self.assertRaises(GridRowFormatError):
            transaction_from_grid_row(1, ("카드번호",), ("5587380002701692",))


if __name__ == "__main__":
    unittest.main()
