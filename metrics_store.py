"""
Metrics Store Module

Persists conversion metrics to a local SQLite database for:
- Historical analysis and trending
- Cross-session comparison
- Bottleneck identification
- Performance regression detection

Phase 4 Week 5: Performance Tuning & Metrics Persistence
"""

import sqlite3
import time
import logging
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Callable, List, Dict, Any, Generator

logger = logging.getLogger(__name__)

# Default database path relative to script directory
DEFAULT_DB_NAME = "rom_converter_metrics.db"

# Schema version for future migrations
SCHEMA_VERSION = 1


@dataclass
class ConversionRecord:
    """A single persisted conversion metric record."""
    file_name: str
    input_format: str
    output_format: str
    input_size_bytes: int
    output_size_bytes: int
    duration_seconds: float
    success: bool
    tool_name: str = ""
    error_message: str = ""
    session_id: str = ""
    timestamp: float = field(default_factory=time.time)
    record_id: Optional[int] = None  # Set after DB insertion

    @property
    def throughput_mbps(self) -> float:
        if self.duration_seconds == 0:
            return 0.0
        return (self.input_size_bytes / 1024 / 1024) / self.duration_seconds

    @property
    def compression_ratio(self) -> float:
        if self.input_size_bytes == 0:
            return 0.0
        return self.output_size_bytes / self.input_size_bytes


@dataclass
class SessionSummary:
    """Aggregate summary for a conversion session."""
    session_id: str
    total_files: int
    successful: int
    failed: int
    total_input_bytes: int
    total_output_bytes: int
    total_duration_seconds: float
    avg_throughput_mbps: float
    start_timestamp: float
    end_timestamp: float

    @property
    def success_rate(self) -> float:
        if self.total_files == 0:
            return 0.0
        return self.successful / self.total_files


class MetricsStore:
    """
    SQLite-backed metrics persistence for ROM conversions.

    Stores per-file conversion records and provides querying,
    aggregation, and trend analysis capabilities.
    """

    def __init__(
        self,
        db_path: Optional[str] = None,
        log_callback: Optional[Callable[[str], None]] = None,
    ):
        """
        Initialize the metrics store.

        Args:
            db_path: Path to SQLite database file. Defaults to
                     rom_converter_metrics.db in the current directory.
            log_callback: Optional logging callback
        """
        self.db_path = Path(db_path) if db_path else Path(DEFAULT_DB_NAME)
        self.log_callback = log_callback
        self._initialized = False
        self._init_db()

    def _log(self, message: str) -> None:
        if self.log_callback:
            self.log_callback(message)
        logger.info(message)

    @contextmanager
    def _connect(self) -> Generator[sqlite3.Connection, None, None]:
        """Context manager providing a database connection."""
        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _init_db(self) -> None:
        """Create tables and indexes if they don't exist."""
        with self._connect() as conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS schema_version (
                    version INTEGER PRIMARY KEY
                );

                CREATE TABLE IF NOT EXISTS conversions (
                    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                    session_id          TEXT    NOT NULL DEFAULT '',
                    file_name           TEXT    NOT NULL,
                    input_format        TEXT    NOT NULL DEFAULT '',
                    output_format       TEXT    NOT NULL DEFAULT '',
                    input_size_bytes    INTEGER NOT NULL DEFAULT 0,
                    output_size_bytes   INTEGER NOT NULL DEFAULT 0,
                    duration_seconds    REAL    NOT NULL DEFAULT 0,
                    success             INTEGER NOT NULL DEFAULT 0,
                    tool_name           TEXT    NOT NULL DEFAULT '',
                    error_message       TEXT    NOT NULL DEFAULT '',
                    timestamp           REAL    NOT NULL
                );

                CREATE INDEX IF NOT EXISTS idx_conversions_session
                    ON conversions(session_id);
                CREATE INDEX IF NOT EXISTS idx_conversions_timestamp
                    ON conversions(timestamp);
                CREATE INDEX IF NOT EXISTS idx_conversions_success
                    ON conversions(success);
                CREATE INDEX IF NOT EXISTS idx_conversions_tool
                    ON conversions(tool_name);
            """)

            # Record schema version if missing
            existing = conn.execute("SELECT version FROM schema_version").fetchone()
            if not existing:
                conn.execute(
                    "INSERT INTO schema_version (version) VALUES (?)",
                    (SCHEMA_VERSION,),
                )

        self._initialized = True

    # ------------------------------------------------------------------
    # Write methods
    # ------------------------------------------------------------------

    def record(self, record: ConversionRecord) -> int:
        """
        Persist a single conversion record.

        Args:
            record: ConversionRecord to save

        Returns:
            Row ID of inserted record
        """
        with self._connect() as conn:
            cursor = conn.execute(
                """
                INSERT INTO conversions
                    (session_id, file_name, input_format, output_format,
                     input_size_bytes, output_size_bytes, duration_seconds,
                     success, tool_name, error_message, timestamp)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    record.session_id,
                    record.file_name,
                    record.input_format,
                    record.output_format,
                    record.input_size_bytes,
                    record.output_size_bytes,
                    record.duration_seconds,
                    int(record.success),
                    record.tool_name,
                    record.error_message,
                    record.timestamp,
                ),
            )
            row_id = cursor.lastrowid
            record.record_id = row_id
            return row_id

    def record_batch(self, records: List[ConversionRecord]) -> int:
        """
        Persist multiple records in a single transaction.

        Args:
            records: List of ConversionRecord to save

        Returns:
            Number of records inserted
        """
        with self._connect() as conn:
            rows = [
                (
                    r.session_id, r.file_name, r.input_format, r.output_format,
                    r.input_size_bytes, r.output_size_bytes, r.duration_seconds,
                    int(r.success), r.tool_name, r.error_message, r.timestamp,
                )
                for r in records
            ]
            conn.executemany(
                """
                INSERT INTO conversions
                    (session_id, file_name, input_format, output_format,
                     input_size_bytes, output_size_bytes, duration_seconds,
                     success, tool_name, error_message, timestamp)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                rows,
            )
        return len(records)

    # ------------------------------------------------------------------
    # Query methods
    # ------------------------------------------------------------------

    def get_recent(self, limit: int = 50) -> List[ConversionRecord]:
        """
        Get the most recent conversion records.

        Args:
            limit: Maximum records to return

        Returns:
            List of ConversionRecord ordered by timestamp desc
        """
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM conversions ORDER BY timestamp DESC LIMIT ?",
                (limit,),
            ).fetchall()
        return [self._row_to_record(r) for r in rows]

    def get_by_session(self, session_id: str) -> List[ConversionRecord]:
        """
        Get all records for a session.

        Args:
            session_id: Session identifier

        Returns:
            List of ConversionRecord for the session
        """
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM conversions WHERE session_id = ? ORDER BY timestamp",
                (session_id,),
            ).fetchall()
        return [self._row_to_record(r) for r in rows]

    def get_failures(self, limit: int = 100) -> List[ConversionRecord]:
        """
        Get recent failed conversions.

        Args:
            limit: Maximum records to return

        Returns:
            List of failed ConversionRecord
        """
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM conversions WHERE success = 0 ORDER BY timestamp DESC LIMIT ?",
                (limit,),
            ).fetchall()
        return [self._row_to_record(r) for r in rows]

    def get_stats(self, since_timestamp: Optional[float] = None) -> Dict[str, Any]:
        """
        Get aggregate statistics.

        Args:
            since_timestamp: Only include records after this Unix timestamp.
                             Defaults to all time.

        Returns:
            Dict of aggregate stats
        """
        where = "WHERE timestamp >= ?" if since_timestamp else ""
        params = (since_timestamp,) if since_timestamp else ()

        with self._connect() as conn:
            row = conn.execute(
                f"""
                SELECT
                    COUNT(*)                        AS total,
                    SUM(success)                    AS successful,
                    COUNT(*) - SUM(success)         AS failed,
                    AVG(duration_seconds)           AS avg_duration,
                    MAX(duration_seconds)           AS max_duration,
                    MIN(CASE WHEN success=1 THEN duration_seconds END) AS min_duration,
                    SUM(input_size_bytes)           AS total_input_bytes,
                    SUM(output_size_bytes)          AS total_output_bytes,
                    AVG(CASE WHEN duration_seconds > 0
                        THEN (input_size_bytes / 1048576.0) / duration_seconds
                        ELSE 0 END)                AS avg_throughput_mbps
                FROM conversions {where}
                """,
                params,
            ).fetchone()

        return dict(row) if row else {}

    def get_throughput_trend(
        self,
        limit: int = 50,
        successful_only: bool = True,
    ) -> List[Dict[str, Any]]:
        """
        Get throughput trend over recent conversions.

        Args:
            limit: Number of data points
            successful_only: Only include successful conversions

        Returns:
            List of dicts with timestamp and throughput_mbps
        """
        where = "WHERE success = 1" if successful_only else ""
        with self._connect() as conn:
            rows = conn.execute(
                f"""
                SELECT timestamp,
                       CASE WHEN duration_seconds > 0
                            THEN (input_size_bytes / 1048576.0) / duration_seconds
                            ELSE 0 END AS throughput_mbps
                FROM conversions {where}
                ORDER BY timestamp DESC
                LIMIT ?
                """,
                (limit,),
            ).fetchall()
        return [dict(r) for r in reversed(rows)]  # Chronological order

    def get_slowest_conversions(self, limit: int = 10) -> List[ConversionRecord]:
        """
        Get the slowest successful conversions.

        Args:
            limit: Maximum records to return

        Returns:
            List of slowest ConversionRecord
        """
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT * FROM conversions
                WHERE success = 1
                ORDER BY duration_seconds DESC
                LIMIT ?
                """,
                (limit,),
            ).fetchall()
        return [self._row_to_record(r) for r in rows]

    def get_session_summary(self, session_id: str) -> Optional[SessionSummary]:
        """
        Get aggregated summary for a session.

        Args:
            session_id: Session identifier

        Returns:
            SessionSummary or None if session not found
        """
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT
                    session_id,
                    COUNT(*)                AS total_files,
                    SUM(success)            AS successful,
                    COUNT(*) - SUM(success) AS failed,
                    SUM(input_size_bytes)   AS total_input_bytes,
                    SUM(output_size_bytes)  AS total_output_bytes,
                    SUM(duration_seconds)   AS total_duration_seconds,
                    AVG(CASE WHEN duration_seconds > 0
                        THEN (input_size_bytes / 1048576.0) / duration_seconds
                        ELSE 0 END)         AS avg_throughput_mbps,
                    MIN(timestamp)          AS start_timestamp,
                    MAX(timestamp)          AS end_timestamp
                FROM conversions
                WHERE session_id = ?
                GROUP BY session_id
                """,
                (session_id,),
            ).fetchone()

        if not row:
            return None

        return SessionSummary(
            session_id=row["session_id"],
            total_files=row["total_files"],
            successful=row["successful"],
            failed=row["failed"],
            total_input_bytes=row["total_input_bytes"] or 0,
            total_output_bytes=row["total_output_bytes"] or 0,
            total_duration_seconds=row["total_duration_seconds"] or 0.0,
            avg_throughput_mbps=row["avg_throughput_mbps"] or 0.0,
            start_timestamp=row["start_timestamp"] or 0.0,
            end_timestamp=row["end_timestamp"] or 0.0,
        )

    def get_tool_stats(self) -> List[Dict[str, Any]]:
        """
        Get per-tool aggregate statistics.

        Returns:
            List of dicts with tool_name, total, successful, avg_throughput_mbps
        """
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT
                    tool_name,
                    COUNT(*)        AS total,
                    SUM(success)    AS successful,
                    AVG(CASE WHEN duration_seconds > 0
                        THEN (input_size_bytes / 1048576.0) / duration_seconds
                        ELSE 0 END) AS avg_throughput_mbps
                FROM conversions
                GROUP BY tool_name
                ORDER BY total DESC
                """,
            ).fetchall()
        return [dict(r) for r in rows]

    def get_all_sessions(self) -> List[SessionSummary]:
        """
        Return a SessionSummary for every distinct session_id in the database,
        ordered by start_timestamp ascending.
        """
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT
                    session_id,
                    COUNT(*)                AS total_files,
                    SUM(success)            AS successful,
                    COUNT(*) - SUM(success) AS failed,
                    SUM(input_size_bytes)   AS total_input_bytes,
                    SUM(output_size_bytes)  AS total_output_bytes,
                    SUM(duration_seconds)   AS total_duration_seconds,
                    AVG(CASE WHEN duration_seconds > 0
                        THEN (input_size_bytes / 1048576.0) / duration_seconds
                        ELSE 0 END)         AS avg_throughput_mbps,
                    MIN(timestamp)          AS start_timestamp,
                    MAX(timestamp)          AS end_timestamp
                FROM conversions
                GROUP BY session_id
                ORDER BY MIN(timestamp)
                """
            ).fetchall()
        return [
            SessionSummary(
                session_id=r["session_id"],
                total_files=r["total_files"],
                successful=r["successful"],
                failed=r["failed"],
                total_input_bytes=r["total_input_bytes"] or 0,
                total_output_bytes=r["total_output_bytes"] or 0,
                total_duration_seconds=r["total_duration_seconds"] or 0.0,
                avg_throughput_mbps=r["avg_throughput_mbps"] or 0.0,
                start_timestamp=r["start_timestamp"] or 0.0,
                end_timestamp=r["end_timestamp"] or 0.0,
            )
            for r in rows
        ]

    def get_session_records(self, session_id: str) -> List[ConversionRecord]:
        """Alias for get_by_session — returns all records for a session."""
        return self.get_by_session(session_id)

    def clear_all(self) -> int:
        """Delete every record from the conversions table. Returns deleted count."""
        with self._connect() as conn:
            cursor = conn.execute("DELETE FROM conversions")
            return cursor.rowcount

    def count(self) -> int:
        """Return total number of stored records."""
        with self._connect() as conn:
            return conn.execute("SELECT COUNT(*) FROM conversions").fetchone()[0]

    def purge_old(self, older_than_days: int = 30) -> int:
        """
        Delete records older than N days.

        Args:
            older_than_days: Records older than this many days will be removed

        Returns:
            Number of deleted records
        """
        cutoff = time.time() - (older_than_days * 86400)
        with self._connect() as conn:
            cursor = conn.execute(
                "DELETE FROM conversions WHERE timestamp < ?", (cutoff,)
            )
            deleted = cursor.rowcount
        if deleted:
            self._log(f"Purged {deleted} metrics records older than {older_than_days} days")
        return deleted

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _row_to_record(row: sqlite3.Row) -> ConversionRecord:
        return ConversionRecord(
            record_id=row["id"],
            session_id=row["session_id"],
            file_name=row["file_name"],
            input_format=row["input_format"],
            output_format=row["output_format"],
            input_size_bytes=row["input_size_bytes"],
            output_size_bytes=row["output_size_bytes"],
            duration_seconds=row["duration_seconds"],
            success=bool(row["success"]),
            tool_name=row["tool_name"],
            error_message=row["error_message"],
            timestamp=row["timestamp"],
        )


def create_metrics_store(
    db_path: Optional[str] = None,
    log_callback: Optional[Callable[[str], None]] = None,
) -> MetricsStore:
    """
    Factory function to create a MetricsStore.

    Args:
        db_path: Path to SQLite database (None = default location)
        log_callback: Optional logging callback

    Returns:
        Initialized MetricsStore
    """
    return MetricsStore(db_path=db_path, log_callback=log_callback)
