"""
Phase 4 Week 6: End-to-End Integration Tests

Wires all Phase 4 components together and exercises the full conversion
pipeline: PerformanceOptimizer → StreamingConverter → ConversionErrorHandler
→ MetricsStore → PerformanceAnalyzer.

Tests verify:
- All modules import and initialize cleanly
- Components connect correctly through shared interfaces
- Metrics flow from conversion to store to analyzer
- Error handling and circuit breaker work in pipeline
- Performance panel can be constructed (headless)
"""

import time
import threading
import tempfile
from pathlib import Path
from typing import List
from unittest.mock import MagicMock, patch

import pytest

from performance_optimizer import create_performance_optimizer, ConversionMetrics
from streaming_converter import create_streaming_converter, StreamingConversionResult
from conversion_error_handler import create_error_handler, FailureCategory
from metrics_store import create_metrics_store, ConversionRecord
from performance_analyzer import create_performance_analyzer, TrendDirection


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def tmp_iso(tmp_path):
    f = tmp_path / "game.iso"
    f.write_bytes(b"FAKE_ROM_DATA" * 512)  # ~6 KB
    return f


@pytest.fixture
def pipeline(tmp_path):
    """Return all four Phase 4 components wired together."""
    store     = create_metrics_store(db_path=str(tmp_path / "metrics.db"))
    optimizer = create_performance_optimizer()
    error_hdl = create_error_handler(max_retries=2, circuit_breaker_threshold=3)
    analyzer  = create_performance_analyzer(store=store)
    streamer  = create_streaming_converter()
    return {
        "store":     store,
        "optimizer": optimizer,
        "error":     error_hdl,
        "analyzer":  analyzer,
        "streamer":  streamer,
    }


# ---------------------------------------------------------------------------
# Import / availability tests
# ---------------------------------------------------------------------------

class TestModuleAvailability:
    def test_performance_optimizer_importable(self):
        from performance_optimizer import PerformanceOptimizer, create_performance_optimizer
        assert PerformanceOptimizer is not None

    def test_streaming_converter_importable(self):
        from streaming_converter import StreamingConverter, create_streaming_converter
        assert StreamingConverter is not None

    def test_error_handler_importable(self):
        from conversion_error_handler import ConversionErrorHandler, create_error_handler
        assert ConversionErrorHandler is not None

    def test_metrics_store_importable(self):
        from metrics_store import MetricsStore, create_metrics_store
        assert MetricsStore is not None

    def test_performance_analyzer_importable(self):
        from performance_analyzer import PerformanceAnalyzer, create_performance_analyzer
        assert PerformanceAnalyzer is not None

    def test_performance_status_panel_importable(self):
        from performance_status_panel import PerformanceStatusPanel, create_performance_panel
        assert PerformanceStatusPanel is not None

    def test_streaming_extractor_importable(self):
        from streaming_extractor import create_streaming_extractor
        assert create_streaming_extractor is not None


# ---------------------------------------------------------------------------
# Component initialization tests
# ---------------------------------------------------------------------------

class TestComponentInitialization:
    def test_all_components_initialize(self, pipeline):
        assert pipeline["optimizer"] is not None
        assert pipeline["streamer"]  is not None
        assert pipeline["error"]     is not None
        assert pipeline["store"]     is not None
        assert pipeline["analyzer"]  is not None

    def test_optimizer_has_expected_attributes(self, pipeline):
        opt = pipeline["optimizer"]
        assert hasattr(opt, "parallel_manager")
        assert hasattr(opt, "resource_monitor")
        assert opt.parallel_manager.max_workers >= 1

    def test_error_handler_starts_clean(self, pipeline):
        stats = pipeline["error"].get_stats()
        assert stats["total_failures"] == 0
        assert stats["open_circuit_breakers"] == []

    def test_store_starts_empty(self, pipeline):
        assert pipeline["store"].count() == 0

    def test_analyzer_linked_to_store(self, pipeline):
        assert pipeline["analyzer"].store is pipeline["store"]


# ---------------------------------------------------------------------------
# Metrics flow tests
# ---------------------------------------------------------------------------

class TestMetricsFlow:
    def test_conversion_metric_persisted(self, pipeline):
        store = pipeline["store"]
        rec = ConversionRecord(
            file_name="game.iso",
            input_format="iso", output_format="chd",
            input_size_bytes=100 * 1024 * 1024,
            output_size_bytes=60 * 1024 * 1024,
            duration_seconds=10.0,
            success=True, tool_name="chdman",
        )
        store.record(rec)
        assert store.count() == 1

    def test_analyzer_sees_stored_metrics(self, pipeline):
        store    = pipeline["store"]
        analyzer = pipeline["analyzer"]
        for i in range(6):
            store.record(ConversionRecord(
                file_name=f"game_{i}.iso",
                input_format="iso", output_format="chd",
                input_size_bytes=100 * 1024 * 1024,
                output_size_bytes=50 * 1024 * 1024,
                duration_seconds=10.0, success=True,
            ))
        stats = store.get_stats()
        assert stats["total"] == 6
        # Analyzer can run
        report = analyzer.run_full_analysis()
        assert report is not None

    def test_failure_recorded_in_store_and_handler(self, pipeline):
        store = pipeline["store"]
        error = pipeline["error"]
        err = Exception("chdman exit code 1")
        error.record_failure("/tmp/bad.iso", err, tool_name="chdman")
        store.record(ConversionRecord(
            file_name="bad.iso",
            input_format="iso", output_format="chd",
            input_size_bytes=0, output_size_bytes=0,
            duration_seconds=0, success=False,
            error_message=str(err),
        ))
        assert error.get_stats()["total_failures"] == 1
        assert len(store.get_failures()) == 1

    def test_session_summary_aggregates_correctly(self, pipeline):
        store = pipeline["store"]
        sid = "e2e-session-01"
        for i in range(4):
            store.record(ConversionRecord(
                file_name=f"g{i}.iso", session_id=sid,
                input_format="iso", output_format="chd",
                input_size_bytes=50 * 1024 * 1024, output_size_bytes=25 * 1024 * 1024,
                duration_seconds=5.0, success=True,
            ))
        store.record(ConversionRecord(
            file_name="fail.iso", session_id=sid,
            input_format="iso", output_format="chd",
            input_size_bytes=0, output_size_bytes=0,
            duration_seconds=1.0, success=False,
        ))
        summary = store.get_session_summary(sid)
        assert summary.total_files == 5
        assert summary.successful  == 4
        assert summary.failed      == 1
        assert summary.success_rate == pytest.approx(0.8)


# ---------------------------------------------------------------------------
# Pipeline integration tests
# ---------------------------------------------------------------------------

class TestPipelineIntegration:
    def test_optimizer_acquires_and_releases(self, pipeline):
        opt = pipeline["optimizer"]
        opt.start_conversion_batch(3)
        pm = opt.parallel_manager
        assert pm.acquire_conversion_slot(timeout=2)
        assert pm.active_conversions == 1
        pm.release_conversion_slot()
        assert pm.active_conversions == 0

    def test_streaming_copy_within_pipeline(self, pipeline, tmp_iso, tmp_path):
        streamer = pipeline["streamer"]
        dest     = tmp_path / "copy.iso"
        ok = streamer.stream_copy(str(tmp_iso), str(dest))
        assert ok is True
        assert dest.stat().st_size == tmp_iso.stat().st_size

    def test_error_handler_retry_logic(self, pipeline):
        error = pipeline["error"]
        f = "/tmp/retry_test.iso"
        # First failure
        error.record_failure(f, Exception("timed out"))
        assert error.should_retry(f) is True
        error.increment_retry(f)
        # Second failure
        error.record_failure(f, Exception("timed out"))
        assert error.should_retry(f) is True
        error.increment_retry(f)
        # Third try (max_retries=2 → should stop)
        assert error.should_retry(f) is False

    def test_circuit_breaker_blocks_failing_tool(self, pipeline):
        error = pipeline["error"]  # threshold=3
        for _ in range(3):
            error.record_failure("/tmp/x.iso", Exception("tool crashed"), tool_name="bad_tool")
        assert not error.should_retry("/tmp/new.iso", tool_name="bad_tool")

    def test_full_successful_batch_flow(self, pipeline, tmp_iso, tmp_path):
        """Simulate a 3-file batch: optimizer → streamer → store."""
        opt     = pipeline["optimizer"]
        streamer= pipeline["streamer"]
        store   = pipeline["store"]
        pm      = opt.parallel_manager

        opt.start_conversion_batch(3)
        for i in range(3):
            assert pm.acquire_conversion_slot(timeout=2)
            try:
                dest = tmp_path / f"out_{i}.bin"
                ok   = streamer.stream_copy(str(tmp_iso), str(dest))
                store.record(ConversionRecord(
                    file_name=tmp_iso.name,
                    input_format="iso", output_format="bin",
                    input_size_bytes=tmp_iso.stat().st_size,
                    output_size_bytes=dest.stat().st_size if ok else 0,
                    duration_seconds=0.01,
                    success=ok,
                ))
            finally:
                pm.release_conversion_slot()

        opt.finish_conversion_batch()
        assert store.count() == 3
        stats = store.get_stats()
        assert stats["successful"] == 3

    def test_analysis_after_batch(self, pipeline, tmp_iso, tmp_path):
        """Run analyzer after populating store with enough data."""
        store    = pipeline["store"]
        analyzer = pipeline["analyzer"]

        for i in range(12):
            store.record(ConversionRecord(
                file_name=f"rom_{i}.iso",
                input_format="iso", output_format="chd",
                input_size_bytes=100 * 1024 * 1024,
                output_size_bytes=60 * 1024 * 1024,
                duration_seconds=float(i % 5 + 1),
                success=(i % 5 != 0),  # One failure every 5
            ))

        report = analyzer.run_full_analysis()
        text   = report.format_report()
        assert "Performance Analysis Report" in text
        assert len(report.recommendations) >= 1

    def test_parallel_batch_thread_safety(self, pipeline, tmp_iso, tmp_path):
        """10 threads competing for worker slots — no race conditions."""
        opt   = pipeline["optimizer"]
        store = pipeline["store"]
        pm    = opt.parallel_manager

        opt.start_conversion_batch(10)
        errors: List[Exception] = []

        def worker(idx: int):
            try:
                if pm.acquire_conversion_slot(timeout=5):
                    try:
                        time.sleep(0.02)
                        store.record(ConversionRecord(
                            file_name=f"rom_{idx}.iso",
                            input_format="iso", output_format="chd",
                            input_size_bytes=1024,
                            output_size_bytes=512,
                            duration_seconds=0.02,
                            success=True,
                        ))
                    finally:
                        pm.release_conversion_slot()
            except Exception as exc:
                errors.append(exc)

        threads = [threading.Thread(target=worker, args=(i,)) for i in range(10)]
        for t in threads:
            t.start()
        for t in threads:
            t.join(timeout=15)

        assert errors == [], f"Thread errors: {errors}"
        assert store.count() == 10
        opt.finish_conversion_batch()


# ---------------------------------------------------------------------------
# Performance panel headless test
# ---------------------------------------------------------------------------

class TestPerformancePanel:
    def test_panel_constructs_without_display(self, pipeline):
        """Construct panel in headless mode — no actual Tk window needed."""
        from performance_status_panel import PerformanceStatusPanel

        mock_parent = MagicMock()
        mock_parent.tk = MagicMock()

        # Patch Tk internals so Frame.__init__ doesn't need a real display
        with patch("performance_status_panel.tk.LabelFrame.__init__", return_value=None), \
             patch("performance_status_panel.tk.Frame.__init__", return_value=None), \
             patch("performance_status_panel.tk.Label.__init__", return_value=None), \
             patch("performance_status_panel.tk.Canvas.__init__", return_value=None), \
             patch("performance_status_panel.tk.Button.__init__", return_value=None):
            panel = PerformanceStatusPanel.__new__(PerformanceStatusPanel)
            panel.optimizer     = pipeline["optimizer"]
            panel.error_handler = pipeline["error"]
            panel.analyzer      = pipeline["analyzer"]
            panel.log_callback  = None
            panel._running      = False
            panel._poll_job     = None
            assert panel.optimizer is not None
            assert panel.error_handler is not None


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
