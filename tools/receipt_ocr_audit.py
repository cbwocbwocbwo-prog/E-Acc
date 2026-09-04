from __future__ import annotations

"""Run a read-only OCR audit for every receipt image currently downloaded."""

import argparse
import hashlib
import json
import re
import sqlite3
from collections import defaultdict
from decimal import Decimal
from pathlib import Path

from eacc_app.models import UnsubmittedTransaction
from eacc_app.ocr_validation import validate_receipt_images


# These entries were downloaded after the current import DB was created, so they
# do not yet have a DB record.  Keep the source data only for read-only OCR audit.
MANUAL_TARGETS = (
    (9, "09103201", "4336920004009888", "20260827", "47000"), (10, "03300900", "4336920004693749", "20260827", "4500"),
    (17, "06106101", "4336920002275192", "20260827", "3300"), (23, "07103201", "4336920004009888", "20260827", "1800"),
    (39, "15404901", "5587380002755904", "20260826", "47000"), (41, "01300800", "4336920004693749", "20260826", "4000"),
    (47, "00198410", "4336920002275135", "20260826", "7700"), (49, "06207903", "5587380002941256", "20260825", "135000"),
    (50, "07406603", "4336920004948481", "20260825", "66000"), (51, "03908800", "5587380002701692", "20260825", "7600"),
    (67, "16904701", "4336920004062036", "20260825", "5400"), (68, "07082717", "4336920002275408", "20260825", "3600"),
    (69, "06082717", "4336920002275408", "20260825", "3600"), (79, "22502602", "5587380002565238", "20260824", "70000"),
    (83, "13404103", "5587380002328363", "20260824", "14000"), (87, "28506200", "5587380002933147", "20260821", "9000"),
    (89, "20289513", "4336920002275291", "20260821", "6000"), (90, "19408400", "5587380002024442", "20260821", "91950"),
    (92, "01908400", "5587380002701692", "20260821", "4000"), (97, "23101811", "5587380002039622", "20260820", "8000"),
    (100, "13404301", "5587380002755904", "20260820", "5000"), (104, "09007910", "5587380002461065", "20260820", "10000"),
    (107, "14206401", "4336920004135782", "20260819", "113000"), (112, "16408200", "5587380002024442", "20260819", "15000"),
    (113, "17901901", "5587380002426142", "20260819", "32000"), (129, "27807800", "5587380002701692", "20260814", "6000"),
    (133, "23203411", "5587380002437842", "20260813", "9000"), (137, "09205701", "4336920004135782", "20260812", "4500"),
    (147, "13982907", "4336920002275408", "20260807", "50000"), (149, "12982807", "4336920002275408", "20260806", "3000"),
)


def _manual_transaction(row: int, approval: str, card: str, date: str, amount: str) -> UnsubmittedTransaction:
    value = Decimal(amount)
    key = hashlib.sha256(f"{card}|{approval}|{date}|{value}".encode()).hexdigest()
    day = f"{date[:4]}-{date[4:6]}-{date[6:]}"
    return UnsubmittedTransaction(
        source_row_number=row, voucher_status="", evidence_status="#", card_number=card,
        card_holder="", approval_number=approval, weekday="", usage_time="", posting_date=day,
        evidence_date=day, account_name="", description="", merchant="", amount=value,
        tax_code="", business_type="", cost_center="", card_issue_date="", division="",
        actual_merchant_name="", actual_merchant_code="", transaction_key=key,
        transaction_id="UNS-" + key[:16].upper(),
    )


def _database_transactions(database_path: Path) -> dict[str, UnsubmittedTransaction]:
    connection = sqlite3.connect(database_path)
    try:
        records = connection.execute("SELECT source_data_json FROM transactions").fetchall()
    finally:
        connection.close()
    transactions: dict[str, UnsubmittedTransaction] = {}
    for (source_data,) in records:
        data = json.loads(source_data)
        data["amount"] = Decimal(data["amount"])
        transaction = UnsubmittedTransaction(**data)
        transactions[transaction.transaction_id] = transaction
    return transactions


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--offset", type=int, default=0)
    parser.add_argument("--limit", type=int, default=5)
    parser.add_argument("--ids", help="Comma-separated transaction IDs to audit")
    args = parser.parse_args()

    app_dir = Path.home() / "AppData" / "Local" / "EAccAutomation"
    receipt_dir = app_dir / "receipts"
    groups: dict[str, list[Path]] = defaultdict(list)
    for path in receipt_dir.iterdir():
        match = re.match(r"^(UNS-[0-9A-F]+)_", path.name)
        if path.is_file() and match:
            groups[match.group(1)].append(path)

    transactions = _database_transactions(app_dir / "eacc.db")
    for row in MANUAL_TARGETS:
        transaction = _manual_transaction(*row)
        transactions[transaction.transaction_id] = transaction

    if args.ids:
        selected_ids = {value.strip() for value in args.ids.split(",") if value.strip()}
        items = [(transaction_id, paths) for transaction_id, paths in sorted(groups.items()) if transaction_id in selected_ids]
    else:
        items = sorted(groups.items())[args.offset : args.offset + args.limit]
    for transaction_id, paths in items:
        transaction = transactions.get(transaction_id)
        if transaction is None:
            print(json.dumps({"id": transaction_id, "status": "NO_SOURCE_DATA"}, ensure_ascii=False), flush=True)
            continue
        try:
            result = validate_receipt_images(transaction, tuple(sorted(paths)))
            failed = [check.field_name for check in result.checks if not check.is_match]
            print(json.dumps({
                "id": transaction_id,
                "row": transaction.source_row_number,
                "files": len(paths),
                "expected": {
                    "approval": transaction.approval_number,
                    "evidence_date": transaction.evidence_date,
                    "amount": f"{transaction.amount:,.0f}",
                },
                "status": result.status,
                "checks": [
                    {"field": check.field_name, "match": check.is_match, "detected": check.detected_value}
                    for check in result.checks
                ],
                "failed": failed,
            }, ensure_ascii=True), flush=True)
        except Exception as exc:  # report every file rather than stopping the audit
            print(json.dumps({
                "id": transaction_id, "status": "OCR_ERROR", "error": str(exc)
            }, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
