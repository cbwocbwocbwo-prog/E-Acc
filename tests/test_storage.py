from __future__ import annotations

import tempfile
import unittest
import os
from pathlib import Path

from eacc_app.storage import ImportRepository


def _find_sample_file() -> Path:
    original = Path(r"C:\Users\user\Downloads\미상신내역.xls")
    if original.exists():
        return original
    download_dir = Path(os.environ.get("LOCALAPPDATA", "")) / "EAccAutomation" / "downloads"
    candidates = sorted(download_dir.glob("미상신내역_*.xls"), key=lambda path: path.stat().st_mtime, reverse=True)
    return candidates[0] if candidates else original


SAMPLE_FILE = _find_sample_file()


class StorageTests(unittest.TestCase):
    @unittest.skipUnless(SAMPLE_FILE.exists(), "사용자 제공 샘플 파일이 없음")
    def test_reimport_marks_every_existing_transaction_as_duplicate(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            repository = ImportRepository(Path(temp_dir) / "test.db")
            first = repository.import_file(SAMPLE_FILE)
            second = repository.import_file(SAMPLE_FILE)
            history = repository.recent_history()

        self.assertGreater(first.total_count, 0)
        self.assertEqual(first.total_count, first.new_count)
        self.assertEqual(0, first.duplicate_count)
        self.assertEqual(0, first.error_count)
        self.assertEqual(first.total_count, second.total_count)
        self.assertEqual(0, second.new_count)
        self.assertEqual(second.total_count, second.duplicate_count)
        self.assertEqual(0, second.error_count)
        self.assertEqual(2, len(history))
        self.assertEqual(second.job_id, history[0].job_id)

    @unittest.skipUnless(SAMPLE_FILE.exists(), "사용자 제공 샘플 파일이 없음")
    def test_processing_events_keep_history_and_latest_status(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            repository = ImportRepository(Path(temp_dir) / "test.db")
            transaction = next(
                row.transaction
                for row in repository.import_file(SAMPLE_FILE).rows
                if row.transaction is not None
            )
            repository.record_processing_event(transaction, "영수증 정상", "승인번호·일자·금액 일치")
            repository.record_processing_event(transaction, "결재요청 확인대기")
            repository.record_processing_event(transaction, "처리 완료", "최신 목록에서 행이 사라짐")
            events = repository.recent_processing_events()
            latest = repository.latest_processing_statuses()

        self.assertEqual(3, len(events))
        self.assertEqual("처리 완료", events[0].status)
        self.assertEqual("처리 완료", latest[transaction.transaction_id])

    @unittest.skipUnless(SAMPLE_FILE.exists(), "사용자 제공 샘플 파일이 없음")
    def test_filtered_latest_status_ignores_receipt_validation_events(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            repository = ImportRepository(Path(temp_dir) / "test.db")
            transaction = next(
                row.transaction
                for row in repository.import_file(SAMPLE_FILE).rows
                if row.transaction is not None
            )
            repository.record_processing_event(transaction, "결재요청 전송")
            repository.record_processing_event(transaction, "처리 완료")
            repository.record_processing_event(transaction, "영수증 정상")
            latest = repository.latest_processing_statuses()
            approval_latest = repository.latest_processing_statuses(
                ("결재요청 전송", "처리 완료")
            )

        self.assertEqual("영수증 정상", latest[transaction.transaction_id])
        self.assertEqual("처리 완료", approval_latest[transaction.transaction_id])


if __name__ == "__main__":
    unittest.main()
