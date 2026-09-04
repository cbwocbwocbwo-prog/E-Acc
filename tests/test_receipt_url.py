from __future__ import annotations

import unittest
from urllib.parse import parse_qs, urlsplit

from eacc_app.browser_automation import ReceiptNotAvailable, original_image_url


class ReceiptUrlTests(unittest.TestCase):
    def test_thumb_url_is_changed_to_original_and_cache_suffix_is_removed(self) -> None:
        url, doc_irn, corp_no = original_image_url(
            "https://officeedms.example:9980/OfficeViewer/slip_actor.jsp",
            "DownloadImage.do?ImgType=thumb&DocIRN=DOC123&Idx=0&degree=0&UserID=1102100&CorpNo=1000?0.123",
        )

        query = parse_qs(urlsplit(url).query)
        self.assertEqual(["original"], query["ImgType"])
        self.assertEqual(["DOC123"], query["DocIRN"])
        self.assertEqual(["1102100"], query["UserID"])
        self.assertEqual(["1000"], query["CorpNo"])
        self.assertEqual("DOC123", doc_irn)
        self.assertEqual("1000", corp_no)

    def test_required_receipt_identifiers_are_checked(self) -> None:
        with self.assertRaises(ReceiptNotAvailable):
            original_image_url(
                "https://officeedms.example/OfficeViewer/slip_actor.jsp",
                "DownloadImage.do?ImgType=thumb&DocIRN=DOC123",
            )


if __name__ == "__main__":
    unittest.main()
