from __future__ import annotations

import unittest

from eacc_app.merchant_lookup import (
    extract_bizno_business_name,
    extract_business_name_from_text,
    extract_business_numbers,
    is_pg_business_type,
    is_valid_business_number,
)


class MerchantLookupTests(unittest.TestCase):
    def test_pg_target_uses_contains_match_for_authentication_variants(self) -> None:
        self.assertTrue(is_pg_business_type("PG일반(인증)"))
        self.assertTrue(is_pg_business_type("PG일반(비인증)"))
        self.assertFalse(is_pg_business_type("일반음식점"))

    def test_checksum_valid_number_is_extracted_without_hyphens(self) -> None:
        result = extract_business_numbers("공급자 사업자번호 : 220-81-55597")

        self.assertEqual(("2208155597",), result.candidates)
        self.assertTrue(is_valid_business_number("2208155597"))

    def test_invalid_and_multiple_business_numbers_need_review(self) -> None:
        invalid = extract_business_numbers("사업자번호 220-81-55598")
        self.assertEqual((), invalid.candidates)

        multiple = extract_business_numbers("사업자 220-81-55597 / 304-21-63066")
        self.assertEqual(("2208155597", "3042163066"), multiple.candidates)
        self.assertIn("여러", multiple.reason)

    def test_bizno_business_name_is_read_from_search_result_link(self) -> None:
        self.assertEqual(
            "(주)케이지이니시스",
            extract_bizno_business_name(
                '<a href="/article/2208155597"><h4>(주)케이지이니시스</h4></a>',
                "2208155597",
            ),
        )


if __name__ == "__main__":
    unittest.main()
