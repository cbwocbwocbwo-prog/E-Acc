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
    status TEXT NOT NULL,
    reason TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_processing_events_transaction
ON processing_events(transaction_id, id DESC);
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
            cursor = connection.execute(
                """
                INSERT INTO processing_events(
                    event_at, transaction_id, approval_number, evidence_date,
                    amount, merchant, status, reason
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    event_at,
                    transaction.transaction_id,
                    transaction.approval_number,
                    transaction.evidence_date,
                    f"{transaction.amount:,.0f}",
                    transaction.merchant,
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
        )

    def recent_processing_events(self, limit: int = 500) -> tuple[ProcessingEvent, ...]:
        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT id, event_at, transaction_id, approval_number, evidence_date,
                       amount, merchant, status, reason
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
