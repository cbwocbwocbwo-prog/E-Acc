from __future__ import annotations

import tempfile
import unittest
import os
from pathlib import Path

from eacc_app.models import EXPECTED_UNSUBMITTED_HEADERS
from eacc_app.parser import WorkbookFormatError, parse_unsubmitted_xls


def _find_sample_file() -> Path:
    original = Path(r"C:\Users\user\Downloads\미상신내역.xls")
    if original.exists():
        return original
    download_dir = Path(os.environ.get("LOCALAPPDATA", "")) / "EAccAutomation" / "downloads"
    candidates = sorted(download_dir.glob("미상신내역_*.xls"), key=lambda path: path.stat().st_mtime, reverse=True)
    return candidates[0] if candidates else original


SAMPLE_FILE = _find_sample_file()


class ParserTests(unittest.TestCase):
    @unittest.skipUnless(SAMPLE_FILE.exists(), "사용자 제공 샘플 파일이 없음")
    def test_real_sample_preserves_identifiers_and_normalizes_values(self) -> None:
        rows = parse_unsubmitted_xls(SAMPLE_FILE)

        self.assertGreater(len(rows), 0)
        self.assertTrue(all(row.transaction is not None for row in rows))
        first = rows[0].transaction
        assert first is not None
        self.assertTrue(first.approval_number.isdigit())
        self.assertTrue(first.card_number.isdigit())
        self.assertRegex(first.evidence_date, r"^\d{4}-\d{2}-\d{2}$")
        self.assertGreaterEqual(first.amount, 0)
        self.assertTrue(first.transaction_id.startswith("UNS-"))
        self.assertNotIn("/images/button", first.to_json_dict().values())

    def test_malformed_business_row_is_returned_as_row_error(self) -> None:
        header = "".join(f"<td>{value}</td>" for value in EXPECTED_UNSUBMITTED_HEADERS)
        malformed = "<td>값 하나뿐</td>"
        content = f'<meta charset="UTF-8"><table><tr>{header}</tr><tr>{malformed}</tr></table>'
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "미상신.xls"
            path.write_text(content, encoding="utf-8")
            rows = parse_unsubmitted_xls(path)

        self.assertEqual(1, len(rows))
        self.assertIsNone(rows[0].transaction)
        self.assertIn("열 개수 오류", rows[0].error_message)

    def test_wrong_header_is_rejected(self) -> None:
        headers = list(EXPECTED_UNSUBMITTED_HEADERS)
        headers[0] = "다른열"
        header = "".join(f"<td>{value}</td>" for value in headers)
        content = f'<meta charset="UTF-8"><table><tr>{header}</tr></table>'
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "미상신.xls"
            path.write_text(content, encoding="utf-8")
            with self.assertRaises(WorkbookFormatError):
                parse_unsubmitted_xls(path)


if __name__ == "__main__":
    unittest.main()
