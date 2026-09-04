"""Run the current receipt OCR validator against saved receipt originals.

This is a read-only diagnostic tool.  It does not alter receipt files,
the application database, or e-Accounting.
"""

from __future__ import annotations

import json
import sqlite3
import sys
from collections import Counter, defaultdict
from decimal import Decimal
from pathlib import Path

# Running this file places ``tools`` rather than the project root on sys.path.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from eacc_app.models import UnsubmittedTransaction
from eacc_app.ocr_validation import validate_receipt_images


RECEIPT_DIRECTORY = Path(r"C:\Users\user\AppData\Local\EAccAutomation\receipts")
DATABASE_PATH = Path(r"C:\Users\user\AppData\Local\EAccAutomation\eacc.db")


def transactions_by_id() -> dict[str, UnsubmittedTransaction]:
    connection = sqlite3.connect(DATABASE_PATH)
    try:
        rows = connection.execute(
            "SELECT transaction_id, source_data_json FROM transactions"
        )
        transactions: dict[str, UnsubmittedTransaction] = {}
        for transaction_id, source_json in rows:
            payload = json.loads(source_json)
            payload["amount"] = Decimal(str(payload["amount"]))
            transactions[transaction_id] = UnsubmittedTransaction(**payload)
        return transactions
    finally:
        connection.close()


def receipt_groups() -> dict[str, list[Path]]:
    groups: dict[str, list[Path]] = defaultdict(list)
    for path in sorted(RECEIPT_DIRECTORY.iterdir()):
        if path.is_file() and path.suffix.lower() in {".jpg", ".jpeg", ".png"}:
            groups[path.name.split("_", 1)[0]].append(path)
    return groups


def main() -> None:
    transactions = transactions_by_id()
    groups = receipt_groups()
    report: list[dict[str, object]] = []

    items = sorted(groups.items())
    for index, (transaction_id, paths) in enumerate(items, start=1):
        transaction = transactions.get(transaction_id)
        if transaction is None:
            report.append(
                {
                    "transaction_id": transaction_id,
                    "files": [path.name for path in paths],
                    "status": "기준정보 없음",
                }
            )
        else:
            try:
                result = validate_receipt_images(transaction, tuple(paths))
                report.append(
                    {
                        "transaction_id": transaction_id,
                        "files": [path.name for path in paths],
                        "status": result.status,
                        "rotation": result.rotation_degrees,
                        "ambiguous": result.orientation_ambiguous,
                        "checks": [
                            {
                                "field": check.field_name,
                                "expected": check.expected_value,
                                "detected": check.detected_value,
                                "match": check.is_match,
                                "reason": check.reason,
                            }
                            for check in result.checks
                        ],
                    }
                )
            except Exception as exc:  # report every receipt rather than stopping the audit
                report.append(
                    {
                        "transaction_id": transaction_id,
                        "files": [path.name for path in paths],
                        "status": "판독 오류",
                        "error": f"{type(exc).__name__}: {exc}",
                    }
                )
        if index % 5 == 0 or index == len(items):
            print(f"PROGRESS {index}/{len(items)}", flush=True)

    counts = Counter(str(row["status"]) for row in report)
    non_normal = [row for row in report if row["status"] != "정상"]
    print(
        "AUDIT_SUMMARY="
        + json.dumps(
            {
                "image_files": sum(len(paths) for paths in groups.values()),
                "transaction_groups": len(groups),
                "database_transactions": len(transactions),
                "counts": dict(counts),
                "non_normal": non_normal,
            },
            ensure_ascii=False,
        ),
        flush=True,
    )


if __name__ == "__main__":
    main()
