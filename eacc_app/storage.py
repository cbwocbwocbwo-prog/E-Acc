from __future__ import annotations

import hashlib
import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Iterable, Iterator

from .models import (
    ExceptionMailUse,
    ImportDisplayRow,
    ImportHistoryItem,
    ImportSummary,
    ProcessingEvent,
    MailLogItem,
    UnsubmittedTransaction,
)
from .parser import parse_unsubmitted_xls


SCHEMA = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS import_jobs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    imported_at TEXT NOT NULL,
    source_file TEXT NOT NULL,
    source_sha256 TEXT NOT NULL,
    total_count INTEGER NOT NULL DEFAULT 0,
    new_count INTEGER NOT NULL DEFAULT 0,
    duplicate_count INTEGER NOT NULL DEFAULT 0,
    error_count INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS transactions (
    transaction_key TEXT PRIMARY KEY,
    transaction_id TEXT NOT NULL UNIQUE,
    first_seen_at TEXT NOT NULL,
    source_data_json TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS import_rows (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    job_id INTEGER NOT NULL REFERENCES import_jobs(id),
    source_row_number INTEGER NOT NULL,
    transaction_key TEXT,
    import_status TEXT NOT NULL CHECK(import_status IN ('신규', '중복', '오류')),
    error_message TEXT NOT NULL DEFAULT '',
    raw_data_json TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_import_rows_job ON import_rows(job_id);
CREATE INDEX IF NOT EXISTS idx_import_rows_transaction ON import_rows(transaction_key);

CREATE TABLE IF NOT EXISTS processing_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    event_at TEXT NOT NULL,
    transaction_id TEXT NOT NULL,
    approval_number TEXT NOT NULL,
    evidence_date TEXT NOT NULL,
    amount TEXT NOT NULL,
    merchant TEXT NOT NULL,
    account_name TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL,
    reason TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_processing_events_transaction
ON processing_events(transaction_id, id DESC);

CREATE TABLE IF NOT EXISTS mail_delivery_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sent_at TEXT NOT NULL,
    recipient_name TEXT NOT NULL,
    recipient_email TEXT NOT NULL DEFAULT '',
    department TEXT NOT NULL DEFAULT '',
    transaction_count INTEGER NOT NULL DEFAULT 0,
    subject TEXT NOT NULL,
    status TEXT NOT NULL,
    reason TEXT NOT NULL DEFAULT '',
    transaction_ids_json TEXT NOT NULL DEFAULT '[]'
    ,outlook_message_id TEXT NOT NULL DEFAULT ''
    ,mail_batch_id TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_mail_delivery_logs_status
ON mail_delivery_logs(status, id DESC);
"""


class ImportRepository:
    def __init__(self, database_path: str | Path) -> None:
        self.database_path = Path(database_path)
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.database_path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        return connection

    @contextmanager
    def _connection(self) -> Iterator[sqlite3.Connection]:
        connection = self._connect()
        try:
            with connection:
                yield connection
        finally:
            connection.close()

    def _initialize(self) -> None:
        with self._connection() as connection:
            connection.executescript(SCHEMA)
            # Existing user databases were created before account_name was
            # retained with processing events.  Keep their audit data and add
            # the display-only column in place.
            columns = {
                str(row["name"])
                for row in connection.execute("PRAGMA table_info(processing_events)").fetchall()
            }
            if "account_name" not in columns:
                connection.execute(
                    "ALTER TABLE processing_events ADD COLUMN account_name TEXT NOT NULL DEFAULT ''"
                )
            mail_columns = {
                str(row["name"])
                for row in connection.execute("PRAGMA table_info(mail_delivery_logs)").fetchall()
            }
            if "outlook_message_id" not in mail_columns:
                connection.execute(
                    "ALTER TABLE mail_delivery_logs ADD COLUMN outlook_message_id TEXT NOT NULL DEFAULT ''"
                )
            if "mail_batch_id" not in mail_columns:
                connection.execute(
                    "ALTER TABLE mail_delivery_logs ADD COLUMN mail_batch_id TEXT NOT NULL DEFAULT ''"
                )
            # Earlier builds called Outlook.Send() and immediately wrote
            # '발송 완료'.  That only proves the item entered Outlook, not that
            # it reached Sent Items, so preserve the audit row but correct the
            # misleading final state.
            connection.execute(
                """
                UPDATE mail_delivery_logs
                SET status = '발송 확인 불가',
                    reason = '이전 버전은 Outlook 보낸 편지함 확인 없이 발송 완료로 기록했습니다.'
                WHERE status = '발송 완료' AND reason = 'Outlook 발송 요청 완료'
                """
            )
            # Processing events written by earlier versions did not retain the
            # account name.  It is already present in the transaction snapshot,
            # so restore only this display field without changing any result or
            # audit timestamp.
            rows = connection.execute(
                """
                SELECT event.id, txn.source_data_json
                FROM processing_events AS event
                INNER JOIN transactions AS txn
                    ON txn.transaction_id = event.transaction_id
                WHERE event.account_name = ''
                """
            ).fetchall()
            for row in rows:
                try:
                    account_name = str(json.loads(row["source_data_json"]).get("account_name", ""))
                except (TypeError, ValueError, json.JSONDecodeError):
                    account_name = ""
                if account_name:
                    connection.execute(
                        "UPDATE processing_events SET account_name = ? WHERE id = ?",
                        (account_name, row["id"]),
                    )

    def import_file(self, file_path: str | Path) -> ImportSummary:
        path = Path(file_path).resolve()
        source_sha256 = hashlib.sha256(path.read_bytes()).hexdigest()
        parsed_rows = parse_unsubmitted_xls(path)
        imported_at = datetime.now().astimezone().isoformat(timespec="seconds")
        display_rows: list[ImportDisplayRow] = []
        new_count = duplicate_count = error_count = 0

        with self._connection() as connection:
            cursor = connection.execute(
                """
                INSERT INTO import_jobs(imported_at, source_file, source_sha256)
                VALUES (?, ?, ?)
                """,
                (imported_at, str(path), source_sha256),
            )
            job_id = int(cursor.lastrowid)

            for parsed in parsed_rows:
                transaction = parsed.transaction
                if transaction is None:
                    status = "오류"
                    error_count += 1
                    transaction_key = None
                else:
                    exists = connection.execute(
                        "SELECT 1 FROM transactions WHERE transaction_key = ?",
                        (transaction.transaction_key,),
                    ).fetchone()
                    if exists:
                        status = "중복"
                        duplicate_count += 1
                    else:
                        status = "신규"
                        new_count += 1
                        connection.execute(
                            """
                            INSERT INTO transactions(
                                transaction_key, transaction_id, first_seen_at, source_data_json
                            ) VALUES (?, ?, ?, ?)
                            """,
                            (
                                transaction.transaction_key,
                                transaction.transaction_id,
                                imported_at,
                                json.dumps(transaction.to_json_dict(), ensure_ascii=False),
                            ),
                        )
                    transaction_key = transaction.transaction_key

                connection.execute(
                    """
                    INSERT INTO import_rows(
                        job_id, source_row_number, transaction_key,
                        import_status, error_message, raw_data_json
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        job_id,
                        parsed.source_row_number,
                        transaction_key,
                        status,
                        parsed.error_message,
                        json.dumps(parsed.raw_values, ensure_ascii=False),
                    ),
                )
                display_rows.append(
                    ImportDisplayRow(
                        source_row_number=parsed.source_row_number,
                        status=status,
                        transaction=transaction,
                        raw_values=parsed.raw_values,
                        error_message=parsed.error_message,
                    )
                )

            connection.execute(
                """
                UPDATE import_jobs
                SET total_count = ?, new_count = ?, duplicate_count = ?, error_count = ?
                WHERE id = ?
                """,
                (len(parsed_rows), new_count, duplicate_count, error_count, job_id),
            )

        return ImportSummary(
            job_id=job_id,
            source_file=str(path),
            total_count=len(parsed_rows),
            new_count=new_count,
            duplicate_count=duplicate_count,
            error_count=error_count,
            rows=tuple(display_rows),
        )

    def recent_history(self, limit: int = 100) -> tuple[ImportHistoryItem, ...]:
        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT id, imported_at, source_file, total_count,
                       new_count, duplicate_count, error_count
                FROM import_jobs
                ORDER BY id DESC
                LIMIT ?
                """,
                (limit,),
            ).fetchall()
        return tuple(
            ImportHistoryItem(
                job_id=row["id"],
                imported_at=row["imported_at"],
                source_file=row["source_file"],
                total_count=row["total_count"],
                new_count=row["new_count"],
                duplicate_count=row["duplicate_count"],
                error_count=row["error_count"],
            )
            for row in rows
        )

    def record_processing_event(
        self,
        transaction: UnsubmittedTransaction,
        status: str,
        reason: str = "",
    ) -> ProcessingEvent:
        event_at = datetime.now().astimezone().isoformat(timespec="seconds")
        with self._connection() as connection:
            self._upsert_transaction_snapshot(connection, transaction, event_at)
            cursor = connection.execute(
                """
                INSERT INTO processing_events(
                    event_at, transaction_id, approval_number, evidence_date,
                    amount, merchant, account_name, status, reason
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    event_at,
                    transaction.transaction_id,
                    transaction.approval_number,
                    transaction.evidence_date,
                    f"{transaction.amount:,.0f}",
                    transaction.merchant,
                    transaction.account_name,
                    status,
                    reason,
                ),
            )
            event_id = int(cursor.lastrowid)
        return ProcessingEvent(
            event_id=event_id,
            event_at=event_at,
            transaction_id=transaction.transaction_id,
            approval_number=transaction.approval_number,
            evidence_date=transaction.evidence_date,
            amount=f"{transaction.amount:,.0f}",
            merchant=transaction.merchant,
            status=status,
            reason=reason,
            account_name=transaction.account_name,
        )

    def store_transaction_snapshots(
        self,
        transactions: Iterable[UnsubmittedTransaction],
    ) -> None:
        """Persist rows already read from e-Acc for later result display.

        This only writes the grid data already in memory.  It deliberately
        does not initiate an additional browser request or receipt OCR pass.
        """
        snapshots = tuple(transactions)
        if not snapshots:
            return
        stored_at = datetime.now().astimezone().isoformat(timespec="seconds")
        with self._connection() as connection:
            for transaction in snapshots:
                self._upsert_transaction_snapshot(connection, transaction, stored_at)

    @staticmethod
    def _upsert_transaction_snapshot(connection, transaction: UnsubmittedTransaction, stored_at: str) -> None:
        connection.execute(
            """
            INSERT INTO transactions(transaction_key, transaction_id, first_seen_at, source_data_json)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(transaction_id) DO UPDATE SET source_data_json = excluded.source_data_json
            """,
            (
                transaction.transaction_key,
                transaction.transaction_id,
                stored_at,
                json.dumps(transaction.to_json_dict(), ensure_ascii=False),
            ),
        )

    def recent_processing_events(self, limit: int = 500) -> tuple[ProcessingEvent, ...]:
        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT id, event_at, transaction_id, approval_number, evidence_date,
                       amount, merchant, account_name, status, reason
                FROM processing_events
                ORDER BY id DESC
                LIMIT ?
                """,
                (limit,),
            ).fetchall()
        return tuple(
            ProcessingEvent(
                event_id=row["id"],
                event_at=row["event_at"],
                transaction_id=row["transaction_id"],
                approval_number=row["approval_number"],
                evidence_date=row["evidence_date"],
                amount=row["amount"],
                merchant=row["merchant"],
                status=row["status"],
                reason=row["reason"],
                account_name=row["account_name"],
            )
            for row in rows
        )

    def latest_processing_statuses(
        self,
        statuses: tuple[str, ...] | None = None,
    ) -> dict[str, str]:
        """Return the latest event for each transaction.

        When *statuses* is supplied, only those events participate in choosing
        the latest status.  This keeps operational approval results separate
        from supporting events such as OCR validation.
        """
        filter_clause = ""
        parameters: tuple[str, ...] = ()
        if statuses:
            placeholders = ", ".join("?" for _ in statuses)
            filter_clause = f"WHERE status IN ({placeholders})"
            parameters = statuses
        with self._connection() as connection:
            rows = connection.execute(
                f"""
                SELECT event.transaction_id, event.status
                FROM processing_events AS event
                INNER JOIN (
                    SELECT transaction_id, MAX(id) AS latest_id
                    FROM processing_events
                    {filter_clause}
                    GROUP BY transaction_id
                ) AS latest ON latest.latest_id = event.id
                """,
                parameters,
            ).fetchall()
        return {str(row["transaction_id"]): str(row["status"]) for row in rows}

    def transaction_cost_centers(self, transaction_ids: Iterable[str]) -> dict[str, str]:
        """Return the original cost center for displayed processing rows."""
        identifiers = tuple(dict.fromkeys(str(value) for value in transaction_ids if str(value)))
        if not identifiers:
            return {}
        placeholders = ", ".join("?" for _ in identifiers)
        with self._connection() as connection:
            rows = connection.execute(
                f"SELECT transaction_id, source_data_json FROM transactions WHERE transaction_id IN ({placeholders})",
                identifiers,
            ).fetchall()
        result: dict[str, str] = {}
        for row in rows:
            try:
                result[str(row["transaction_id"])] = str(
                    json.loads(row["source_data_json"]).get("cost_center", "")
                ).strip()
            except (TypeError, ValueError, json.JSONDecodeError):
                continue
        return result

    def today_exception_mail_uses(self, process_date: str | None = None) -> tuple[ExceptionMailUse, ...]:
        """Return today's final exceptions that have a usable receipt.

        The processing-event table deliberately stores a compact audit trail;
        cost center and evidence status are recovered from the original
        transaction snapshot.  Selecting only the final terminal state keeps
        a transaction that was later corrected from being emailed as an
        exception.
        """
        target_date = process_date or datetime.now().astimezone().date().isoformat()
        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT terminal.transaction_id, terminal.approval_number,
                       terminal.evidence_date, terminal.amount, terminal.merchant,
                       terminal.reason, source_transaction.source_data_json,
                       (
                           SELECT receipt.status
                           FROM processing_events AS receipt
                           WHERE receipt.transaction_id = terminal.transaction_id
                             AND receipt.status LIKE '영수증 %'
                           ORDER BY receipt.id DESC LIMIT 1
                       ) AS receipt_status
                FROM processing_events AS terminal
                INNER JOIN (
                    SELECT transaction_id, MAX(id) AS latest_id
                    FROM processing_events
                    WHERE substr(event_at, 1, 10) = ?
                      AND status IN ('처리 완료', '예외처리')
                    GROUP BY transaction_id
                ) AS latest ON latest.latest_id = terminal.id
                INNER JOIN transactions AS source_transaction
                    ON source_transaction.transaction_id = terminal.transaction_id
                WHERE terminal.status = '예외처리'
                ORDER BY terminal.id
                """,
                (target_date,),
            ).fetchall()
        result: list[ExceptionMailUse] = []
        for row in rows:
            try:
                source = json.loads(row["source_data_json"])
            except (TypeError, ValueError, json.JSONDecodeError):
                continue
            # Missing/synchronising receipts must never generate a team mail.
            if str(source.get("evidence_status", "")).strip() != "#":
                continue
            if str(row["receipt_status"] or "").strip() == "영수증 재조회 필요":
                continue
            cost_center = str(source.get("cost_center", "")).strip()
            if not cost_center:
                continue
            try:
                amount = Decimal(str(row["amount"]).replace(",", ""))
            except (ValueError, ArithmeticError):
                continue
            result.append(ExceptionMailUse(
                transaction_id=str(row["transaction_id"]), cost_center=cost_center,
                card_holder=str(source.get("card_holder", "")),
                approval_number=str(row["approval_number"]), evidence_date=str(row["evidence_date"]),
                merchant=str(row["merchant"]), amount=amount, reason=str(row["reason"]),
            ))
        return tuple(result)

    def record_mail_log(
        self,
        *,
        recipient_name: str,
        recipient_email: str,
        department: str,
        transaction_ids: tuple[str, ...],
        subject: str,
        status: str,
        reason: str = "",
        outlook_message_id: str = "",
        mail_batch_id: str = "",
    ) -> MailLogItem:
        sent_at = datetime.now().astimezone().isoformat(timespec="seconds")
        logged_email = _mask_email(recipient_email)
        with self._connection() as connection:
            cursor = connection.execute(
                """
                INSERT INTO mail_delivery_logs(
                    sent_at, recipient_name, recipient_email, department,
                    transaction_count, subject, status, reason, transaction_ids_json, outlook_message_id, mail_batch_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    sent_at, recipient_name, logged_email, department,
                    len(transaction_ids), subject, status, reason,
                    json.dumps(transaction_ids, ensure_ascii=False),
                    outlook_message_id, mail_batch_id,
                ),
            )
            log_id = int(cursor.lastrowid)
        return MailLogItem(
            log_id=log_id, sent_at=sent_at, recipient_name=recipient_name,
            recipient_email=logged_email, department=department,
            transaction_count=len(transaction_ids), subject=subject,
            status=status, reason=reason, transaction_ids=transaction_ids,
            outlook_message_id=outlook_message_id, mail_batch_id=mail_batch_id,
        )

    def recent_mail_logs(self, limit: int = 500) -> tuple[MailLogItem, ...]:
        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT id, sent_at, recipient_name, recipient_email, department,
                       transaction_count, subject, status, reason, transaction_ids_json, outlook_message_id, mail_batch_id
                FROM mail_delivery_logs ORDER BY id DESC LIMIT ?
                """,
                (limit,),
            ).fetchall()
        result: list[MailLogItem] = []
        for row in rows:
            try:
                transaction_ids = tuple(str(value) for value in json.loads(row["transaction_ids_json"]))
            except (TypeError, ValueError, json.JSONDecodeError):
                transaction_ids = ()
            result.append(MailLogItem(
                log_id=row["id"], sent_at=row["sent_at"],
                recipient_name=row["recipient_name"], recipient_email=row["recipient_email"],
                department=row["department"], transaction_count=row["transaction_count"],
                subject=row["subject"], status=row["status"], reason=row["reason"],
                transaction_ids=transaction_ids,
                outlook_message_id=row["outlook_message_id"], mail_batch_id=row["mail_batch_id"],
            ))
        return tuple(result)

    def update_mail_status(self, log_id: int, status: str, reason: str) -> None:
        with self._connection() as connection:
            connection.execute(
                "UPDATE mail_delivery_logs SET status = ?, reason = ? WHERE id = ?",
                (status, reason, log_id),
            )


def _mask_email(value: str) -> str:
    """Keep delivery logs useful without persisting the recipient address."""
    local, separator, domain = value.strip().partition("@")
    if not separator:
        return ""
    visible = local[:1] if local else ""
    return f"{visible}{'*' * max(2, len(local) - 1)}@{domain}"
