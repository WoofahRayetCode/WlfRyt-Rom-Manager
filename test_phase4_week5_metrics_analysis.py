"""
Phase 4 Week 5: Metrics Persistence & Performance Analysis Tests

Covers:
- MetricsStore: SQLite persistence, queries, aggregation
- PerformanceAnalyzer: trend detection, bottleneck finding, outlier detection
"""

import time
import tempfile
from pathlib import Path
from typing import Optional

import pytest

from metrics_store import (
    MetricsStore,
    ConversionRecord,
    SessionSummary,
    create_metrics_store,
)
from performance_analyzer import (
    PerformanceAnalyzer,
    TrendDirection,
    ThroughputTrend,
    Bottleneck,
    OutlierRecord,
    AnalysisReport,
    create_performance_analyzer,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def make_record(
    file_name: str = "game.iso",
    input_size: int = 1024 * 1024 * 100,  # 100 MB
    duration: float = 10.0,
    success: bool = True,
    tool_name: str = "chdman",
    session_id: str = "test-session",
    output_size: Optional[int] = None,
    error: str = "",
) -> ConversionRecord:
    return ConversionRecord(
        file_name=file_name,
        input_format="iso",
        output_format="chd",
        input_size_bytes=input_size,
        output_size_bytes=output_size if output_size is not None else input_size // 2,
        duration_seconds=duration,
        success=success,
        tool_name=tool_name,
        error_message=error,
        session_id=session_id,
        timestamp=time.time(),
    )


@pytest.fixture
def store(tmp_path):
    """In-memory MetricsStore (uses :memory: db)."""
    return create_metrics_store(db_path=str(tmp_path / "test.db"))


@pytest.fixture
def populated_store(store):
    """Store pre-loaded with 10 successful + 2 failed records."""
    for i in range(10):
        store.record(make_record(file_name=f"game_{i}.iso", duration=float(i + 1)))
    store.record(make_record(file_name="bad1.iso", success=False, error="exit code 1"))
    store.record(make_record(file_name="bad2.iso", success=False, error="timed out"))
    return store


# ---------------------------------------------------------------------------
# ConversionRecord dataclass
# ---------------------------------------------------------------------------

class TestConversionRecord:
    def test_throughput_mbps(self):
        r = make_record(input_size=1024 * 1024 * 50, duration=10.0)
        assert r.throughput_mbps == pytest.approx(5.0)

    def test_throughput_zero_duration(self):
        r = make_record(duration=0)
        assert r.throughput_mbps == 0.0

    def test_compression_ratio(self):
        r = make_record(input_size=1000, output_size=500)
        assert r.compression_ratio == pytest.approx(0.5)

    def test_compression_ratio_zero_input(self):
        r = ConversionRecord(
            file_name="x.iso", input_format="iso", output_format="chd",
            input_size_bytes=0, output_size_bytes=0, duration_seconds=1,
            success=True,
        )
        assert r.compression_ratio == 0.0


# ---------------------------------------------------------------------------
# MetricsStore
# ---------------------------------------------------------------------------

class TestMetricsStore:
    def test_create_store(self, tmp_path):
        s = create_metrics_store(db_path=str(tmp_path / "m.db"))
        assert s is not None
        assert s.count() == 0

    def test_record_single(self, store):
        rid = store.record(make_record())
        assert isinstance(rid, int)
        assert rid > 0
        assert store.count() == 1

    def test_record_sets_record_id(self, store):
        r = make_record()
        assert r.record_id is None
        store.record(r)
        assert r.record_id is not None

    def test_record_batch(self, store):
        records = [make_record(file_name=f"f{i}.iso") for i in range(5)]
        inserted = store.record_batch(records)
        assert inserted == 5
        assert store.count() == 5

    def test_get_recent(self, populated_store):
        recent = populated_store.get_recent(limit=5)
        assert len(recent) == 5
        assert all(isinstance(r, ConversionRecord) for r in recent)

    def test_get_recent_ordered_desc(self, populated_store):
        recent = populated_store.get_recent(limit=12)
        for a, b in zip(recent, recent[1:]):
            assert a.timestamp >= b.timestamp

    def test_get_by_session(self, store):
        store.record(make_record(session_id="session-A"))
        store.record(make_record(session_id="session-A"))
        store.record(make_record(session_id="session-B"))
        assert len(store.get_by_session("session-A")) == 2
        assert len(store.get_by_session("session-B")) == 1

    def test_get_failures(self, populated_store):
        failures = populated_store.get_failures()
        assert len(failures) == 2
        assert all(not r.success for r in failures)

    def test_get_stats_counts(self, populated_store):
        stats = populated_store.get_stats()
        assert stats["total"] == 12
        assert stats["successful"] == 10
        assert stats["failed"] == 2

    def test_get_stats_throughput(self, populated_store):
        stats = populated_store.get_stats()
        assert stats["avg_throughput_mbps"] > 0

    def test_get_stats_since_timestamp(self, store):
        past = time.time() - 3600
        r = make_record()
        r.timestamp = time.time()
        store.record(r)
        stats = store.get_stats(since_timestamp=past)
        assert stats["total"] >= 1

    def test_get_throughput_trend(self, populated_store):
        trend_data = populated_store.get_throughput_trend(limit=10)
        assert isinstance(trend_data, list)
        assert len(trend_data) <= 10
        for item in trend_data:
            assert "throughput_mbps" in item
            assert "timestamp" in item

    def test_get_slowest_conversions(self, populated_store):
        slowest = populated_store.get_slowest_conversions(limit=3)
        assert len(slowest) == 3
        for a, b in zip(slowest, slowest[1:]):
            assert a.duration_seconds >= b.duration_seconds

    def test_get_session_summary(self, store):
        for i in range(4):
            store.record(make_record(session_id="sess-X"))
        store.record(make_record(session_id="sess-X", success=False))
        summary = store.get_session_summary("sess-X")
        assert summary is not None
        assert summary.total_files == 5
        assert summary.successful == 4
        assert summary.failed == 1
        assert summary.success_rate == pytest.approx(0.8)

    def test_get_session_summary_missing(self, store):
        assert store.get_session_summary("nonexistent") is None

    def test_get_tool_stats(self, store):
        store.record(make_record(tool_name="chdman"))
        store.record(make_record(tool_name="chdman"))
        store.record(make_record(tool_name="maxcso"))
        tool_stats = store.get_tool_stats()
        names = [t["tool_name"] for t in tool_stats]
        assert "chdman" in names
        assert "maxcso" in names

    def test_purge_old(self, store):
        old = make_record()
        old.timestamp = time.time() - (40 * 86400)  # 40 days ago
        store.record(old)
        store.record(make_record())  # recent
        deleted = store.purge_old(older_than_days=30)
        assert deleted == 1
        assert store.count() == 1

    def test_logging_callback(self, tmp_path):
        msgs = []
        s = create_metrics_store(db_path=str(tmp_path / "log.db"), log_callback=msgs.append)
        # Purge with nothing to delete - no message expected
        s.purge_old()
        # Insert old record then purge - message expected
        old = make_record()
        old.timestamp = time.time() - (40 * 86400)
        s.record(old)
        s.purge_old(older_than_days=30)
        assert any("Purged" in m for m in msgs)


# ---------------------------------------------------------------------------
# PerformanceAnalyzer
# ---------------------------------------------------------------------------

class TestPerformanceAnalyzer:
    def test_create_analyzer(self):
        a = create_performance_analyzer()
        assert a is not None

    def test_create_with_store(self, store):
        a = create_performance_analyzer(store=store)
        assert a.store is store

    def test_logging_callback(self, store):
        msgs = []
        a = create_performance_analyzer(store=store, log_callback=msgs.append)
        a._log("hi")
        assert "hi" in msgs

    def test_trend_insufficient_data(self, store):
        # Only 2 records — not enough
        store.record(make_record())
        store.record(make_record())
        trend = create_performance_analyzer(store=store).analyze_throughput_trend()
        assert trend.direction == TrendDirection.INSUFFICIENT_DATA

    def test_trend_stable(self, store):
        # 10 records with identical throughput → stable
        for _ in range(10):
            store.record(make_record(input_size=100 * 1024 * 1024, duration=10.0))
        trend = create_performance_analyzer(store=store).analyze_throughput_trend()
        assert trend.direction == TrendDirection.STABLE

    def test_trend_improving(self, store):
        # Old records slow, new records fast
        now = time.time()
        for i in range(5):
            r = make_record(input_size=10 * 1024 * 1024, duration=100.0)  # ~0.1 MB/s
            r.timestamp = now - 1000 + i
            store.record(r)
        for i in range(5):
            r = make_record(input_size=100 * 1024 * 1024, duration=1.0)   # ~100 MB/s
            r.timestamp = now + i
            store.record(r)
        trend = create_performance_analyzer(store=store).analyze_throughput_trend()
        assert trend.direction == TrendDirection.IMPROVING
        assert trend.change_percent > 0

    def test_trend_degrading(self, store):
        now = time.time()
        for i in range(5):
            r = make_record(input_size=100 * 1024 * 1024, duration=1.0)   # fast
            r.timestamp = now - 1000 + i
            store.record(r)
        for i in range(5):
            r = make_record(input_size=10 * 1024 * 1024, duration=100.0)  # slow
            r.timestamp = now + i
            store.record(r)
        trend = create_performance_analyzer(store=store).analyze_throughput_trend()
        assert trend.direction == TrendDirection.DEGRADING

    def test_detect_bottlenecks_none(self, store):
        # Good throughput, low failure rate
        for _ in range(20):
            store.record(make_record(input_size=100 * 1024 * 1024, duration=5.0))
        bottlenecks = create_performance_analyzer(store=store).detect_bottlenecks()
        # No high-failure-rate or slow-tool bottlenecks expected
        categories = [b.category for b in bottlenecks]
        assert "high_failure_rate" not in categories

    def test_detect_bottlenecks_high_failure(self, store):
        for _ in range(5):
            store.record(make_record(success=True))
        for _ in range(6):
            store.record(make_record(success=False, error="exit code 1"))
        bottlenecks = create_performance_analyzer(store=store).detect_bottlenecks()
        categories = [b.category for b in bottlenecks]
        assert "high_failure_rate" in categories

    def test_detect_bottlenecks_slow_tool(self, store):
        # One tool (fast) and one tool (very slow)
        for _ in range(10):
            store.record(make_record(tool_name="fast_tool", input_size=100 * 1024 * 1024, duration=1.0))
        for _ in range(5):
            store.record(make_record(tool_name="slow_tool", input_size=1 * 1024 * 1024, duration=100.0))
        bottlenecks = create_performance_analyzer(store=store).detect_bottlenecks()
        categories = [b.category for b in bottlenecks]
        assert "slow_tool" in categories

    def test_detect_outliers_none(self, store):
        # All records same throughput
        for _ in range(10):
            store.record(make_record(input_size=100 * 1024 * 1024, duration=10.0))
        outliers = create_performance_analyzer(store=store).detect_outliers()
        assert outliers == []

    def test_detect_outliers_slow(self, store):
        # 9 fast records + 1 extremely slow
        for _ in range(9):
            store.record(make_record(input_size=100 * 1024 * 1024, duration=1.0))
        store.record(make_record(input_size=1 * 1024 * 1024, duration=1000.0))  # almost 0 MB/s
        outliers = create_performance_analyzer(store=store).detect_outliers()
        assert len(outliers) >= 1
        slow_ones = [o for o in outliers if o.reason == "slow"]
        assert len(slow_ones) >= 1

    def test_recommendations_no_issues(self, store):
        for _ in range(20):
            store.record(make_record(input_size=100 * 1024 * 1024, duration=5.0))
        a = create_performance_analyzer(store=store)
        trend = a.analyze_throughput_trend()
        bottlenecks = a.detect_bottlenecks()
        outliers = a.detect_outliers()
        recs = a.generate_recommendations(trend, bottlenecks, outliers)
        assert any("normally" in r or "No performance" in r for r in recs)

    def test_run_full_analysis(self, populated_store):
        a = create_performance_analyzer(store=populated_store)
        report = a.run_full_analysis()
        assert isinstance(report, AnalysisReport)
        assert isinstance(report.trend, ThroughputTrend)
        assert isinstance(report.bottlenecks, list)
        assert isinstance(report.outliers, list)
        assert isinstance(report.recommendations, list)

    def test_format_report(self, populated_store):
        a = create_performance_analyzer(store=populated_store)
        report = a.run_full_analysis()
        text = report.format_report()
        assert "Performance Analysis Report" in text
        assert "Throughput Trend" in text
        assert "Recommendations" in text

    def test_trend_summary_improving(self):
        trend = ThroughputTrend(
            direction=TrendDirection.IMPROVING,
            recent_avg_mbps=20.0,
            baseline_avg_mbps=10.0,
            change_percent=100.0,
            sample_count=10,
        )
        assert "IMPROVING" in trend.summary()
        assert "↑" in trend.summary()

    def test_trend_summary_insufficient(self):
        trend = ThroughputTrend(
            direction=TrendDirection.INSUFFICIENT_DATA,
            recent_avg_mbps=0, baseline_avg_mbps=0,
            change_percent=0, sample_count=2,
        )
        assert "Insufficient" in trend.summary()


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
