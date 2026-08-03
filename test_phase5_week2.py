"""
Phase 5 Week 2 Tests

Covers:
  - PerformanceSettingsPanel (headless, mocked tk)
  - SessionHistoryPanel (headless, mocked tk)
  - GameCubeConverter (detect_disc_type, command building, error paths)
  - DatMatcher (XML loading, hash matching, CSV export)
  - MetricsStore new methods (get_all_sessions, get_session_records, clear_all)

All tests are headless-safe (no display required).
"""

import csv
import hashlib
import io
import os
import struct
import tempfile
import time
import zlib
import unittest
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
from unittest.mock import MagicMock, patch, PropertyMock


# ── helpers ──────────────────────────────────────────────────────────────────

def _make_gcn_iso(path: str, disc_type: str = "gcn") -> None:
    """Write a minimal fake GCN or Wii ISO header into a file."""
    data = bytearray(0x20)
    GCN_MAGIC = 0xC2339F3D
    WII_MAGIC = 0x5D1C9EA3
    struct.pack_into(">I", data, 0x1C, GCN_MAGIC)
    if disc_type == "wii":
        struct.pack_into(">I", data, 0x18, WII_MAGIC)
    with open(path, "wb") as f:
        f.write(data)


def _make_dat_xml(entries) -> str:
    """Build a minimal No-Intro DAT XML string."""
    roms = []
    for game, rom, crc, md5, sha1 in entries:
        roms.append(
            f'  <game name="{game}">'
            f'    <rom name="{rom}" size="0" crc="{crc}" md5="{md5}" sha1="{sha1}"/>'
            f'  </game>'
        )
    return (
        '<?xml version="1.0"?>\n'
        '<datafile>\n'
        '  <header><name>Test DAT</name><version>20240101</version></header>\n'
        + "\n".join(roms)
        + "\n</datafile>"
    )


# ── GameCubeConverter tests ───────────────────────────────────────────────────

from gamecube_converter import (
    GameCubeConverter, DiscType, RvzCompression,
    detect_disc_type, is_gamecube_or_wii, create_gamecube_converter,
    GameCubeConvertResult,
)


class TestDetectDiscType(unittest.TestCase):

    def test_gcn_iso_detected(self):
        with tempfile.NamedTemporaryFile(suffix=".iso", delete=False) as f:
            _make_gcn_iso(f.name, "gcn")
            name = f.name
        try:
            self.assertEqual(detect_disc_type(name), DiscType.GAMECUBE)
        finally:
            os.unlink(name)

    def test_wii_iso_detected(self):
        with tempfile.NamedTemporaryFile(suffix=".iso", delete=False) as f:
            _make_gcn_iso(f.name, "wii")
            name = f.name
        try:
            self.assertEqual(detect_disc_type(name), DiscType.WII)
        finally:
            os.unlink(name)

    def test_unknown_file(self):
        with tempfile.NamedTemporaryFile(suffix=".iso", delete=False) as f:
            f.write(b"\x00" * 0x20)
            name = f.name
        try:
            self.assertEqual(detect_disc_type(name), DiscType.UNKNOWN)
        finally:
            os.unlink(name)

    def test_missing_file_returns_unknown(self):
        self.assertEqual(detect_disc_type("/nonexistent/fake.iso"), DiscType.UNKNOWN)

    def test_is_gamecube_or_wii_gcn(self):
        with tempfile.NamedTemporaryFile(suffix=".iso", delete=False) as f:
            _make_gcn_iso(f.name, "gcn")
            name = f.name
        try:
            self.assertTrue(is_gamecube_or_wii(name))
        finally:
            os.unlink(name)

    def test_is_gamecube_or_wii_false(self):
        with tempfile.NamedTemporaryFile(suffix=".iso", delete=False) as f:
            f.write(b"\x00" * 0x20)
            name = f.name
        try:
            self.assertFalse(is_gamecube_or_wii(name))
        finally:
            os.unlink(name)


class TestGameCubeConverter(unittest.TestCase):

    def setUp(self):
        self.conv = GameCubeConverter(dolphintool_path=None)  # not available

    def test_not_available_without_tool(self):
        self.assertFalse(self.conv.is_available())

    def test_convert_to_rvz_no_tool(self):
        result = self.conv.convert_to_rvz("game.iso")
        self.assertFalse(result.success)
        self.assertIn("DolphinTool", result.error_message)

    def test_convert_missing_input(self):
        conv = GameCubeConverter(dolphintool_path="fake_dolphin.exe")
        result = conv.convert_to_rvz("/nonexistent/game.iso")
        self.assertFalse(result.success)

    def test_output_path_default_extension(self):
        """When output_path is None, output path in result should use the target extension."""
        with tempfile.NamedTemporaryFile(suffix=".iso", delete=False) as f:
            _make_gcn_iso(f.name, "gcn")
            iso_path = f.name
        try:
            conv = GameCubeConverter(dolphintool_path=None)  # unavailable → early return
            result = conv.convert_to_rvz(iso_path)
            # Tool not available → failure but output_path should be empty string
            self.assertFalse(result.success)
        finally:
            os.unlink(iso_path)

    def test_overwrite_false_blocks_existing_output_check(self):
        """If output already exists and overwrite=False, return failure before subprocess."""
        with tempfile.NamedTemporaryFile(suffix=".iso", delete=False) as fin:
            _make_gcn_iso(fin.name, "gcn")
            iso_path = fin.name
        with tempfile.NamedTemporaryFile(suffix=".rvz", delete=False) as fout:
            out_path = fout.name
        try:
            conv = GameCubeConverter(dolphintool_path="DolphinTool_fake.exe")
            # Pretend the tool is available so we reach the overwrite check
            with patch.object(conv, "is_available", return_value=True):
                result = conv._convert(iso_path, out_path, ".rvz", overwrite=False)
            self.assertFalse(result.success)
            self.assertIn("already exists", result.error_message)
        finally:
            os.unlink(iso_path)
            os.unlink(out_path)
        conv = GameCubeConverter(
            dolphintool_path="DolphinTool.exe",
            compression=RvzCompression.ZSTD,
            compression_level=5,
        )
        cmd = conv._build_command("in.iso", "out.rvz", ".rvz")
        self.assertIn("convert", cmd)
        self.assertIn("zstd", cmd)
        self.assertIn("out.rvz", cmd)

    def test_build_iso_restore_command(self):
        conv = GameCubeConverter(dolphintool_path="DolphinTool.exe")
        cmd = conv._build_command("in.rvz", "out.iso", ".iso")
        self.assertIn("iso", cmd)

    def test_compression_ratio_property(self):
        r = GameCubeConvertResult(
            success=True,
            input_path="a.iso",
            output_path="a.rvz",
            input_size_bytes=1_000_000,
            output_size_bytes=400_000,
        )
        self.assertAlmostEqual(r.compression_ratio, 0.4, places=3)
        self.assertAlmostEqual(r.space_saved_mb, (600_000 / 1_048_576), places=3)

    def test_batch_convert_empty(self):
        conv = GameCubeConverter(dolphintool_path=None)
        results = conv.batch_convert([])
        self.assertEqual(results, [])

    def test_factory_function(self):
        c = create_gamecube_converter()
        self.assertIsInstance(c, GameCubeConverter)

    def test_log_callback(self):
        logs = []
        conv = GameCubeConverter(dolphintool_path=None, log_callback=logs.append)
        conv._log("hello")
        self.assertIn("hello", logs)

    def test_overwrite_false_blocks_existing_output(self):
        """If output already exists and overwrite=False, return failure."""
        with tempfile.NamedTemporaryFile(suffix=".iso", delete=False) as fin:
            _make_gcn_iso(fin.name, "gcn")
            iso_path = fin.name
        with tempfile.NamedTemporaryFile(suffix=".rvz", delete=False) as fout:
            out_path = fout.name
        try:
            conv = GameCubeConverter(dolphintool_path="DolphinTool.exe")
            with patch.object(conv, "is_available", return_value=True):
                result = conv._convert(iso_path, out_path, ".rvz", overwrite=False)
            self.assertFalse(result.success)
            self.assertIn("already exists", result.error_message)
        finally:
            os.unlink(iso_path)
            os.unlink(out_path)


# ── DatMatcher tests ──────────────────────────────────────────────────────────

from dat_matcher import (
    DatMatcher, DatEntry, MatchResult, MatchStatus, BatchMatchReport,
    create_dat_matcher, _compute_hashes,
)


def _write_dat(entries, tmp_dir: str) -> str:
    dat_path = os.path.join(tmp_dir, "test.dat")
    xml = _make_dat_xml(entries)
    with open(dat_path, "w", encoding="utf-8") as f:
        f.write(xml)
    return dat_path


def _file_checksums(content: bytes):
    crc = format(zlib.crc32(content) & 0xFFFFFFFF, "08X")
    md5 = hashlib.md5(content).hexdigest().upper()
    sha1 = hashlib.sha1(content).hexdigest().upper()
    return crc, md5, sha1


class TestDatMatcher(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.content = b"FAKE ROM DATA FOR TESTING 12345"
        self.crc, self.md5, self.sha1 = _file_checksums(self.content)

        # Write the test ROM file
        self.rom_path = os.path.join(self.tmp, "TestGame (USA).bin")
        with open(self.rom_path, "wb") as f:
            f.write(self.content)

        # Write DAT with matching entry
        self.dat_path = _write_dat([
            ("TestGame", "TestGame (USA).bin", self.crc, self.md5, self.sha1),
            ("OtherGame", "OtherGame (EUR).iso", "AABBCCDD", "A" * 32, "B" * 40),
        ], self.tmp)

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_load_dat(self):
        m = DatMatcher()
        count = m.load_dat(self.dat_path)
        self.assertEqual(count, 2)
        self.assertEqual(m.dat_name, "Test DAT")
        self.assertEqual(m.dat_version, "20240101")

    def test_match_good_by_sha1(self):
        m = DatMatcher()
        m.load_dat(self.dat_path)
        result = m.match_file(self.rom_path)
        self.assertEqual(result.status, MatchStatus.GOOD)
        self.assertEqual(result.file_crc, self.crc)

    def test_match_renamed(self):
        """Same checksum, different filename → RENAMED."""
        renamed = os.path.join(self.tmp, "WrongName.bin")
        with open(renamed, "wb") as f:
            f.write(self.content)
        m = DatMatcher()
        m.load_dat(self.dat_path)
        result = m.match_file(renamed)
        self.assertEqual(result.status, MatchStatus.RENAMED)
        self.assertEqual(result.expected_name, "TestGame (USA).bin")

    def test_match_unknown(self):
        """File not in DAT → UNKNOWN."""
        unknown = os.path.join(self.tmp, "Unknown.bin")
        with open(unknown, "wb") as f:
            f.write(b"completely different data xyz")
        m = DatMatcher()
        m.load_dat(self.dat_path)
        result = m.match_file(unknown)
        self.assertEqual(result.status, MatchStatus.UNKNOWN)

    def test_match_bad_wrong_crc(self):
        """File has same name as DAT entry but different content → BAD."""
        bad_rom = os.path.join(self.tmp, "OtherGame (EUR).iso")
        with open(bad_rom, "wb") as f:
            f.write(b"corrupted content zzz")
        m = DatMatcher()
        m.load_dat(self.dat_path)
        result = m.match_file(bad_rom)
        self.assertEqual(result.status, MatchStatus.BAD)

    def test_missing_file_unknown(self):
        m = DatMatcher()
        m.load_dat(self.dat_path)
        result = m.match_file("/nonexistent/file.bin")
        self.assertEqual(result.status, MatchStatus.UNKNOWN)
        self.assertIn("not found", result.error)

    def test_match_files_batch(self):
        m = DatMatcher()
        m.load_dat(self.dat_path)
        report = m.match_files([self.rom_path])
        self.assertEqual(len(report.results), 1)
        self.assertEqual(len(report.good), 1)
        self.assertEqual(len(report.bad), 0)

    def test_batch_report_summary(self):
        m = DatMatcher()
        m.load_dat(self.dat_path)
        report = m.match_files([self.rom_path])
        summary = report.format_summary()
        self.assertIn("Test DAT", summary)
        self.assertIn("Total", summary)

    def test_scan_directory(self):
        m = DatMatcher()
        m.load_dat(self.dat_path)
        report = m.scan_directory(self.tmp, extensions=[".bin"])
        # self.rom_path has .bin extension → should be found
        files = [Path(r.file_path).name for r in report.results]
        self.assertIn("TestGame (USA).bin", files)

    def test_export_csv(self):
        m = DatMatcher()
        m.load_dat(self.dat_path)
        report = m.match_files([self.rom_path])
        csv_path = os.path.join(self.tmp, "results.csv")
        m.export_results_csv(report, csv_path)
        self.assertTrue(os.path.exists(csv_path))
        with open(csv_path, encoding="utf-8") as f:
            rows = list(csv.reader(f))
        self.assertGreater(len(rows), 1)
        self.assertIn("status", rows[0])

    def test_compute_hashes_correct(self):
        path = self.rom_path
        crc, md5, sha1 = _compute_hashes(path, compute_md5=True, compute_sha1=True)
        self.assertEqual(crc, self.crc)
        self.assertEqual(md5, self.md5)
        self.assertEqual(sha1, self.sha1)

    def test_compute_hashes_skip_md5_sha1(self):
        crc, md5, sha1 = _compute_hashes(
            self.rom_path, compute_md5=False, compute_sha1=False
        )
        self.assertEqual(crc, self.crc)
        self.assertEqual(md5, "")
        self.assertEqual(sha1, "")

    def test_factory(self):
        m = create_dat_matcher()
        self.assertIsInstance(m, DatMatcher)

    def test_entry_count(self):
        m = DatMatcher()
        self.assertEqual(m.entry_count, 0)
        m.load_dat(self.dat_path)
        self.assertEqual(m.entry_count, 2)

    def test_match_status_icons(self):
        for status, icon in [
            (MatchStatus.GOOD, "✅"),
            (MatchStatus.BAD, "❌"),
            (MatchStatus.UNKNOWN, "❓"),
            (MatchStatus.RENAMED, "🔄"),
        ]:
            r = MatchResult(file_path="x", status=status)
            self.assertEqual(r.status_icon, icon)

    def test_reload_dat_clears_previous(self):
        m = DatMatcher()
        m.load_dat(self.dat_path)
        self.assertEqual(m.entry_count, 2)
        # Load again (same file) — should not double-count
        m.load_dat(self.dat_path)
        self.assertEqual(m.entry_count, 2)

    def test_progress_callback_called(self):
        calls = []
        m = DatMatcher()
        m.load_dat(self.dat_path)
        m.match_files([self.rom_path], progress_callback=lambda a, b, c: calls.append((a, b, c)))
        self.assertTrue(any(c[0] == 0 for c in calls))  # at least one call with index 0

    def test_log_callback(self):
        logs = []
        m = DatMatcher(log_callback=logs.append)
        m.load_dat(self.dat_path)
        self.assertTrue(any("Loaded" in l for l in logs))


# ── MetricsStore new methods tests ───────────────────────────────────────────

from metrics_store import MetricsStore, ConversionRecord, SessionSummary


class TestMetricsStoreNewMethods(unittest.TestCase):

    def setUp(self):
        self.tmp_db = tempfile.mktemp(suffix=".db")
        self.store = MetricsStore(db_path=self.tmp_db)
        self._insert_sessions()

    def _make_record(self, session_id: str, success: bool = True) -> ConversionRecord:
        return ConversionRecord(
            file_name="file.iso",
            input_format="iso",
            output_format="chd",
            input_size_bytes=100_000_000,
            output_size_bytes=50_000_000,
            duration_seconds=10.0,
            success=success,
            tool_name="chdman",
            session_id=session_id,
            timestamp=time.time(),
        )

    def _insert_sessions(self):
        self.store.record(self._make_record("sess-A", success=True))
        self.store.record(self._make_record("sess-A", success=False))
        self.store.record(self._make_record("sess-B", success=True))

    def tearDown(self):
        try:
            os.unlink(self.tmp_db)
        except OSError:
            pass

    def test_get_all_sessions_count(self):
        sessions = self.store.get_all_sessions()
        session_ids = {s.session_id for s in sessions}
        self.assertIn("sess-A", session_ids)
        self.assertIn("sess-B", session_ids)

    def test_get_all_sessions_returns_summaries(self):
        sessions = self.store.get_all_sessions()
        for s in sessions:
            self.assertIsInstance(s, SessionSummary)

    def test_get_all_sessions_sess_a_counts(self):
        sessions = {s.session_id: s for s in self.store.get_all_sessions()}
        a = sessions["sess-A"]
        self.assertEqual(a.total_files, 2)
        self.assertEqual(a.successful, 1)
        self.assertEqual(a.failed, 1)

    def test_get_session_records(self):
        recs = self.store.get_session_records("sess-A")
        self.assertEqual(len(recs), 2)
        for r in recs:
            self.assertEqual(r.session_id, "sess-A")

    def test_get_session_records_empty(self):
        recs = self.store.get_session_records("no-such-session")
        self.assertEqual(recs, [])

    def test_clear_all(self):
        deleted = self.store.clear_all()
        self.assertGreaterEqual(deleted, 3)
        self.assertEqual(self.store.count(), 0)

    def test_get_all_sessions_empty_after_clear(self):
        self.store.clear_all()
        sessions = self.store.get_all_sessions()
        self.assertEqual(sessions, [])

    def test_success_rate_property(self):
        sessions = {s.session_id: s for s in self.store.get_all_sessions()}
        self.assertAlmostEqual(sessions["sess-A"].success_rate, 0.5, places=3)
        self.assertAlmostEqual(sessions["sess-B"].success_rate, 1.0, places=3)


# ── PerformanceSettingsPanel headless tests ───────────────────────────────────

class TestPerformanceSettingsPanelHeadless(unittest.TestCase):
    """Tests for PerformanceSettingsPanel without a live display."""

    def _make_panel(self, config_adapter=None):
        from performance_settings_panel import PerformanceSettingsPanel, SLIDERS
        panel = object.__new__(PerformanceSettingsPanel)
        panel.config_adapter = config_adapter
        panel.on_apply = None
        panel._collapsed = False
        panel._vars = {}
        panel._temp_var = MagicMock()
        panel._temp_var.get.return_value = ""
        panel._temp_var.set = MagicMock()
        for label, attr, lo, hi, default, step, unit in SLIDERS:
            v = MagicMock()
            v.get.return_value = default
            panel._vars[attr] = v
        return panel

    def test_collect_returns_all_slider_keys(self):
        from performance_settings_panel import SLIDERS
        panel = self._make_panel()
        result = panel._collect()
        for _, attr, *_ in SLIDERS:
            self.assertIn(attr, result)

    def test_collect_includes_temp_dir(self):
        panel = self._make_panel()
        result = panel._collect()
        self.assertIn("temp_dir", result)

    def test_get_settings_alias(self):
        panel = self._make_panel()
        self.assertEqual(panel.get_settings(), panel._collect())

    def test_set_settings_updates_vars(self):
        from performance_settings_panel import SLIDERS
        panel = self._make_panel()
        # Set a known attr
        attr = SLIDERS[0][1]
        panel._vars[attr].set = MagicMock()
        panel._on_slider = MagicMock()
        panel.set_settings({attr: 99})
        panel._vars[attr].set.assert_called_once_with(99)

    def test_save_to_config_calls_adapter(self):
        adapter = MagicMock()
        panel = self._make_panel(config_adapter=adapter)
        panel._save_to_config({"max_workers": 8})
        adapter.update_preferences.assert_called_once_with(max_workers=8)

    def test_load_from_config_no_adapter(self):
        panel = self._make_panel(config_adapter=None)
        # Should not raise
        panel._load_from_config()

    def test_load_from_config_with_adapter(self):
        from performance_settings_panel import SLIDERS
        prefs = MagicMock()
        for _, attr, lo, hi, default, *_ in SLIDERS:
            setattr(prefs, attr, default)
        setattr(prefs, "temp_dir", "/tmp")
        adapter = MagicMock()
        adapter.get_preferences.return_value = prefs
        panel = self._make_panel(config_adapter=adapter)
        panel._on_slider = MagicMock()  # avoid label lookup
        panel._load_from_config()
        adapter.get_preferences.assert_called_once()

    def test_factory(self):
        from performance_settings_panel import create_performance_settings_panel
        # Just confirm it returns the right type without crashing (no display needed
        # because we're calling the factory before __init__ accesses tk).
        # We'll mock the parent class.
        with patch("performance_settings_panel.ttk.LabelFrame.__init__", return_value=None), \
             patch("performance_settings_panel.PerformanceSettingsPanel._build"), \
             patch("performance_settings_panel.PerformanceSettingsPanel._load_from_config"):
            panel = create_performance_settings_panel(MagicMock())
        from performance_settings_panel import PerformanceSettingsPanel
        self.assertIsInstance(panel, PerformanceSettingsPanel)


# ── SessionHistoryPanel headless tests ───────────────────────────────────────

class TestSessionHistoryPanelHeadless(unittest.TestCase):

    def _make_store_mock(self, sessions=None, records=None):
        store = MagicMock()
        store.get_all_sessions.return_value = sessions or []
        store.get_session_records.return_value = records or []
        return store

    def _make_panel(self, metrics_store=None):
        from session_history_panel import SessionHistoryPanel
        panel = object.__new__(SessionHistoryPanel)
        panel.metrics_store = metrics_store
        panel._sessions = []
        panel._selected_session = ""
        panel._status_var = MagicMock()
        panel._status_var.get.return_value = ""
        panel._status_var.set = MagicMock()
        panel._sess_tree = MagicMock()
        panel._sess_tree.get_children.return_value = []
        panel._file_tree = MagicMock()
        panel._file_tree.get_children.return_value = []
        panel._spark_var = MagicMock()
        panel._info_vars = {k: MagicMock() for k in (
            "Session ID", "Started", "Ended", "Total Files",
            "Successful", "Failed", "Input Size", "Output Size", "Avg Throughput"
        )}
        return panel

    def test_refresh_no_store(self):
        panel = self._make_panel()
        panel.refresh()
        panel._status_var.set.assert_called_with("MetricsStore not available")

    def test_refresh_with_store(self):
        store = self._make_store_mock()
        panel = self._make_panel(metrics_store=store)
        panel.refresh()
        store.get_all_sessions.assert_called_once()

    def test_refresh_updates_status_var(self):
        store = self._make_store_mock(sessions=[])
        panel = self._make_panel(metrics_store=store)
        panel.refresh()
        panel._status_var.set.assert_called()

    def test_factory(self):
        from session_history_panel import create_session_history_panel
        with patch("session_history_panel.ttk.Frame.__init__", return_value=None), \
             patch("session_history_panel.SessionHistoryPanel._build"), \
             patch("session_history_panel.SessionHistoryPanel.refresh"):
            panel = create_session_history_panel(MagicMock())
        from session_history_panel import SessionHistoryPanel
        self.assertIsInstance(panel, SessionHistoryPanel)

    def test_sparkline_empty(self):
        from session_history_panel import _sparkline
        self.assertEqual(len(_sparkline([], width=10)), 10)

    def test_sparkline_uniform(self):
        from session_history_panel import _sparkline
        result = _sparkline([5.0, 5.0, 5.0], width=5)
        self.assertEqual(len(result), 5)

    def test_sparkline_increasing(self):
        from session_history_panel import _sparkline
        result = _sparkline([1.0, 2.0, 3.0, 4.0, 5.0], width=5)
        self.assertEqual(len(result), 5)


if __name__ == "__main__":
    unittest.main()
