from __future__ import annotations

import hashlib
import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Iterator

from .models import (
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
        outlook_message_id: str = "",
    ) -> ProcessingEvent:
        event_at = datetime.now().astimezone().isoformat(timespec="seconds")
        with self._connection() as connection:
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
    ) -> MailLogItem:
        sent_at = datetime.now().astimezone().isoformat(timespec="seconds")
        logged_email = _mask_email(recipient_email)
        with self._connection() as connection:
            cursor = connection.execute(
                """
                INSERT INTO mail_delivery_logs(
                    sent_at, recipient_name, recipient_email, department,
                    transaction_count, subject, status, reason, transaction_ids_json, outlook_message_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    sent_at, recipient_name, logged_email, department,
                    len(transaction_ids), subject, status, reason,
                    json.dumps(transaction_ids, ensure_ascii=False),
                    outlook_message_id,
                ),
            )
            log_id = int(cursor.lastrowid)
        return MailLogItem(
            log_id=log_id, sent_at=sent_at, recipient_name=recipient_name,
            recipient_email=logged_email, department=department,
            transaction_count=len(transaction_ids), subject=subject,
            status=status, reason=reason, transaction_ids=transaction_ids,
            outlook_message_id=outlook_message_id,
        )

    def recent_mail_logs(self, limit: int = 500) -> tuple[MailLogItem, ...]:
        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT id, sent_at, recipient_name, recipient_email, department,
                       transaction_count, subject, status, reason, transaction_ids_json, outlook_message_id
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
                outlook_message_id=row["outlook_message_id"],
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
