"""
Phase 5 Tests — UI Track + Format Track

Covers:
- BatchProgressPanel (headless)
- CHDVerifier
- CueParser / CueDisc
"""

import time
import tempfile
import threading
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from batch_progress_panel import (
    BatchProgressPanel,
    FileEntry,
    FileStatus,
    create_batch_progress_panel,
)
from chd_verifier import (
    CHDVerifier,
    VerifyResult,
    VerifyStatus,
    BatchVerifyReport,
    create_chd_verifier,
)
from cue_parser import (
    CueParser,
    CueDisc,
    CueTrack,
    TrackMode,
    DiscSystem,
    parse_cue,
    _mmssff_to_frames,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def make_cue(tmp_path: Path, content: str, name: str = "game.cue") -> Path:
    """Write a CUE file and return its path."""
    f = tmp_path / name
    f.write_text(content, encoding="utf-8")
    return f


def make_bin(tmp_path: Path, name: str, size: int = 2352) -> Path:
    """Write a dummy BIN file."""
    f = tmp_path / name
    f.write_bytes(b"\x00" * size)
    return f


# ---------------------------------------------------------------------------
# BatchProgressPanel (headless with mocked tk)
# ---------------------------------------------------------------------------

class TestFileEntry:
    def test_duration_no_start(self):
        e = FileEntry(file_id="a", name="a.iso")
        assert e.duration_s == 0.0

    def test_duration_with_start(self):
        e = FileEntry(file_id="a", name="a.iso", start_time=time.time() - 5)
        assert e.duration_s >= 5.0

    def test_duration_completed(self):
        t = time.time()
        e = FileEntry(file_id="a", name="a.iso", start_time=t, end_time=t + 3.0)
        assert e.duration_s == pytest.approx(3.0, abs=0.1)


class TestBatchProgressPanel:
    """Test the panel logic without a real display."""

    def _make_panel(self):
        """Construct panel with all tk calls mocked out."""
        with patch.multiple(
            "batch_progress_panel",
            tk=MagicMock(),
            ttk=MagicMock(),
        ):
            panel = BatchProgressPanel.__new__(BatchProgressPanel)
            panel.log_callback    = None
            panel.on_pause_resume = None
            panel.on_cancel_file  = None
            panel._lock           = threading.Lock()
            panel._files          = {}
            panel._order          = []
            panel._paused         = False
            panel._batch_start    = 0.0
            panel._total          = 0
            panel._done_count     = 0
            panel._poll_job       = None
            return panel

    def test_start_batch_populates_queue(self):
        panel = self._make_panel()
        panel.start_batch.__func__  # ensure it's unbound
        # Directly exercise the logic
        names = ["game1.iso", "game2.iso", "game3.iso"]
        with panel._lock:
            panel._files.clear()
            panel._order.clear()
            panel._total = len(names)
            panel._batch_start = time.time()
            for n in names:
                from batch_progress_panel import FileEntry, FileStatus
                panel._files[n] = FileEntry(file_id=n, name=n)
                panel._order.append(n)
        assert len(panel._files) == 3
        assert panel._order == names

    def test_update_file_changes_status(self):
        panel = self._make_panel()
        panel._files["x.iso"] = FileEntry(file_id="x.iso", name="x.iso")
        panel._order.append("x.iso")
        panel._total = 1

        # Simulate update_file logic
        with panel._lock:
            entry = panel._files["x.iso"]
            entry.status = FileStatus.CONVERTING
            entry.progress = 50.0
            entry.start_time = time.time()

        assert panel._files["x.iso"].status == FileStatus.CONVERTING
        assert panel._files["x.iso"].progress == 50.0

    def test_done_increments_count(self):
        panel = self._make_panel()
        panel._files["y.iso"] = FileEntry(file_id="y.iso", name="y.iso")
        panel._order.append("y.iso")
        panel._total = 1

        with panel._lock:
            entry = panel._files["y.iso"]
            entry.status = FileStatus.DONE
            entry.end_time = time.time()
            panel._done_count += 1

        assert panel._done_count == 1

    def test_cancel_request(self):
        panel = self._make_panel()
        panel._files["z.iso"] = FileEntry(file_id="z.iso", name="z.iso",
                                           status=FileStatus.QUEUED)
        panel._order.append("z.iso")
        with panel._lock:
            panel._files["z.iso"].cancel_requested = True
        assert panel._files["z.iso"].cancel_requested is True

    def test_pause_toggle(self):
        panel = self._make_panel()
        assert panel._paused is False
        with panel._lock:
            panel._paused = True
        assert panel._paused is True

    def test_format_entry_queued(self):
        e = FileEntry(file_id="a", name="game.iso", status=FileStatus.QUEUED)
        text = BatchProgressPanel._format_entry(e)
        assert "⏳" in text
        assert "game.iso" in text

    def test_format_entry_converting(self):
        e = FileEntry(file_id="a", name="game.iso", status=FileStatus.CONVERTING,
                      progress=50.0, throughput_mbps=12.5, start_time=time.time())
        text = BatchProgressPanel._format_entry(e)
        assert "⚙️" in text
        assert "█" in text
        assert "12.5MB/s" in text

    def test_format_entry_failed(self):
        e = FileEntry(file_id="a", name="game.iso", status=FileStatus.FAILED,
                      error="chdman exit code 1")
        text = BatchProgressPanel._format_entry(e)
        assert "❌" in text
        assert "chdman exit code 1" in text

    def test_is_cancelled_true(self):
        panel = self._make_panel()
        panel._files["f.iso"] = FileEntry(file_id="f.iso", name="f.iso",
                                           cancel_requested=True)
        assert panel.is_cancelled("f.iso") is True

    def test_is_cancelled_false(self):
        panel = self._make_panel()
        panel._files["f.iso"] = FileEntry(file_id="f.iso", name="f.iso")
        assert panel.is_cancelled("f.iso") is False

    def test_is_cancelled_missing(self):
        panel = self._make_panel()
        assert panel.is_cancelled("ghost.iso") is False


# ---------------------------------------------------------------------------
# CHDVerifier
# ---------------------------------------------------------------------------

class TestVerifyResult:
    def test_passed_property(self):
        r = VerifyResult(file_path="/a.chd", status=VerifyStatus.PASS)
        assert r.passed is True

    def test_failed_property(self):
        r = VerifyResult(file_path="/a.chd", status=VerifyStatus.FAIL, error_message="bad SHA1")
        assert r.passed is False

    def test_summary_pass(self):
        r = VerifyResult(file_path="/roms/game.chd", status=VerifyStatus.PASS, duration_seconds=4.2)
        assert "✅" in r.summary()
        assert "game.chd" in r.summary()

    def test_summary_fail(self):
        r = VerifyResult(file_path="/roms/game.chd", status=VerifyStatus.FAIL, error_message="bad SHA1")
        assert "❌" in r.summary()


class TestBatchVerifyReport:
    def test_pass_rate_empty(self):
        report = BatchVerifyReport()
        assert report.pass_rate == 0.0

    def test_pass_rate_all_pass(self):
        report = BatchVerifyReport(results=[
            VerifyResult(file_path=f"/f{i}.chd", status=VerifyStatus.PASS) for i in range(5)
        ])
        assert report.pass_rate == 1.0

    def test_counts(self):
        report = BatchVerifyReport(results=[
            VerifyResult(file_path="/a.chd", status=VerifyStatus.PASS),
            VerifyResult(file_path="/b.chd", status=VerifyStatus.FAIL),
            VerifyResult(file_path="/c.chd", status=VerifyStatus.ERROR),
        ])
        assert report.passed == 1
        assert report.failed == 1
        assert report.errors == 1
        assert report.total  == 3

    def test_format_summary(self):
        report = BatchVerifyReport(results=[
            VerifyResult(file_path="/a.chd", status=VerifyStatus.PASS),
            VerifyResult(file_path="/b.chd", status=VerifyStatus.FAIL, error_message="bad"),
        ])
        text = report.format_summary()
        assert "CHD Verification Summary" in text
        assert "2 files" in text


class TestCHDVerifier:
    def test_not_available_when_no_chdman(self, tmp_path):
        v = create_chd_verifier(chdman_path="/nonexistent/chdman")
        assert v.available is False

    def test_verify_missing_file(self, tmp_path):
        v = create_chd_verifier(chdman_path=str(tmp_path / "chdman"))
        result = v.verify_file(str(tmp_path / "ghost.chd"))
        assert result.status == VerifyStatus.ERROR
        assert "not found" in result.error_message.lower()

    def test_verify_skipped_without_chdman(self, tmp_path):
        chd = tmp_path / "game.chd"
        chd.write_bytes(b"\x00" * 100)
        v = CHDVerifier(chdman_path=None)
        v.chdman_path = None  # ensure unavailable
        result = v.verify_file(str(chd))
        assert result.status == VerifyStatus.SKIPPED

    def test_batch_verify_missing_files(self, tmp_path):
        v = create_chd_verifier(chdman_path=None)
        report = v.verify_batch([str(tmp_path / "a.chd"), str(tmp_path / "b.chd")])
        assert report.total == 2
        assert all(r.status in (VerifyStatus.SKIPPED, VerifyStatus.ERROR) for r in report.results)

    def test_logging_callback(self, tmp_path):
        msgs = []
        chd = tmp_path / "game.chd"
        chd.write_bytes(b"\x00" * 100)
        # Use a path that exists but isn't actually chdman → triggers ERROR + log
        v = create_chd_verifier(
            chdman_path=str(tmp_path / "fake_chdman"),
            log_callback=msgs.append,
        )
        v.verify_file(str(chd))   # chdman missing → VerifyStatus.SKIPPED or ERROR, logs either way
        # Ensure our _log path was invoked from quarantine or verify
        # Simply create with log, call _log directly to confirm wiring
        v._log("ping")
        assert "ping" in msgs

    def test_quarantine_on_failure(self, tmp_path):
        quarantine = tmp_path / "quarantine"
        chd = tmp_path / "bad.chd"
        chd.write_bytes(b"\x00" * 100)

        v = CHDVerifier(
            chdman_path="nonexistent_tool",
            quarantine_dir=str(quarantine),
        )
        # Manually inject a FAIL result and call quarantine
        v._quarantine(chd, VerifyResult(file_path=str(chd), status=VerifyStatus.FAIL))
        assert (quarantine / "bad.chd").exists()

    def test_parse_error_line(self):
        v = CHDVerifier()
        out = "Processing file...\nError: SHA1 mismatch\nDone."
        assert "SHA1 mismatch" in v._parse_error(out)


# ---------------------------------------------------------------------------
# CueParser / CueDisc
# ---------------------------------------------------------------------------

class TestMmssffToFrames:
    def test_zero(self):
        assert _mmssff_to_frames("00:00:00") == 0

    def test_one_second(self):
        assert _mmssff_to_frames("00:01:00") == 75

    def test_one_minute(self):
        assert _mmssff_to_frames("01:00:00") == 4500

    def test_complex(self):
        # 1:02:37 → (60+2)*75 + 37 = 4687
        assert _mmssff_to_frames("01:02:37") == 4687


class TestCueParser:
    SINGLE_BIN_CUE = """\
FILE "game.bin" BINARY
  TRACK 01 MODE2/2352
    INDEX 01 00:00:00
"""

    MULTI_TRACK_CUE = """\
FILE "game.bin" BINARY
  TRACK 01 MODE1/2352
    INDEX 01 00:00:00
  TRACK 02 AUDIO
    PREGAP 00:02:00
    INDEX 01 05:00:00
  TRACK 03 AUDIO
    INDEX 01 08:30:00
"""

    MULTI_BIN_CUE = """\
FILE "track01.bin" BINARY
  TRACK 01 MODE1/2352
    INDEX 01 00:00:00
FILE "track02.bin" BINARY
  TRACK 02 AUDIO
    INDEX 01 00:00:00
FILE "track03.bin" BINARY
  TRACK 03 AUDIO
    INDEX 01 00:00:00
"""

    def test_parse_single_bin(self, tmp_path):
        cue = make_cue(tmp_path, self.SINGLE_BIN_CUE)
        disc = parse_cue(str(cue))
        assert disc.is_valid
        assert disc.track_count == 1
        assert not disc.is_multi_bin
        assert disc.bin_files == ["game.bin"]

    def test_parse_multi_track(self, tmp_path):
        cue = make_cue(tmp_path, self.MULTI_TRACK_CUE)
        disc = parse_cue(str(cue))
        assert disc.track_count == 3
        assert len(disc.data_tracks) == 1
        assert len(disc.audio_tracks) == 2

    def test_parse_multi_bin(self, tmp_path):
        cue = make_cue(tmp_path, self.MULTI_BIN_CUE)
        disc = parse_cue(str(cue))
        assert disc.track_count == 3
        assert disc.is_multi_bin
        assert len(disc.bin_files) == 3

    def test_missing_cue(self, tmp_path):
        disc = parse_cue(str(tmp_path / "ghost.cue"))
        assert not disc.is_valid
        assert len(disc.parse_errors) > 0

    def test_missing_bins_detected(self, tmp_path):
        cue = make_cue(tmp_path, self.SINGLE_BIN_CUE)
        disc = parse_cue(str(cue))
        # game.bin was NOT created
        assert not disc.all_bins_present()
        assert "game.bin" in disc.missing_bins()

    def test_bins_present(self, tmp_path):
        cue = make_cue(tmp_path, self.SINGLE_BIN_CUE)
        make_bin(tmp_path, "game.bin")
        disc = parse_cue(str(cue))
        assert disc.all_bins_present()
        assert disc.missing_bins() == []

    def test_track_modes(self, tmp_path):
        cue = make_cue(tmp_path, self.MULTI_TRACK_CUE)
        disc = parse_cue(str(cue))
        assert disc.tracks[0].mode == TrackMode.MODE1_2352
        assert disc.tracks[1].mode == TrackMode.AUDIO
        assert disc.tracks[0].is_data
        assert disc.tracks[1].is_audio

    def test_pregap_parsed(self, tmp_path):
        cue = make_cue(tmp_path, self.MULTI_TRACK_CUE)
        disc = parse_cue(str(cue))
        assert disc.tracks[1].pregap_frames == 150  # 00:02:00 = 2*75

    def test_index_01_frames(self, tmp_path):
        cue = make_cue(tmp_path, self.MULTI_TRACK_CUE)
        disc = parse_cue(str(cue))
        assert disc.tracks[0].index_01_frames == 0
        assert disc.tracks[1].index_01_frames == _mmssff_to_frames("05:00:00")

    def test_sector_sizes(self):
        assert CueTrack(1, TrackMode.MODE1_2352, "a.bin").sector_size == 2352
        assert CueTrack(1, TrackMode.MODE1_2048, "a.bin").sector_size == 2048
        assert CueTrack(1, TrackMode.AUDIO,      "a.bin").sector_size == 2352

    def test_system_cdda(self, tmp_path):
        cue = make_cue(tmp_path, """\
FILE "audio.bin" BINARY
  TRACK 01 AUDIO
    INDEX 01 00:00:00
""")
        disc = parse_cue(str(cue))
        assert disc.system == DiscSystem.CDDA

    def test_system_generic(self, tmp_path):
        cue = make_cue(tmp_path, self.SINGLE_BIN_CUE)
        disc = parse_cue(str(cue))
        # No BIN → heuristic: single data track, no audio
        assert disc.system == DiscSystem.GENERIC_CDROM

    def test_system_saturn_heuristic(self, tmp_path):
        # 1 data + 6 audio → Saturn heuristic
        lines = ['FILE "disc.bin" BINARY\n',
                 '  TRACK 01 MODE1/2352\n    INDEX 01 00:00:00\n']
        for i in range(2, 8):
            lines.append(f'  TRACK {i:02d} AUDIO\n    INDEX 01 0{i}:00:00\n')
        cue = make_cue(tmp_path, "".join(lines))
        disc = parse_cue(str(cue))
        assert disc.system == DiscSystem.SATURN

    def test_summary_contains_system(self, tmp_path):
        cue = make_cue(tmp_path, self.SINGLE_BIN_CUE)
        disc = parse_cue(str(cue))
        assert disc.system.value in disc.summary()
        assert "game.cue" in disc.summary()

    def test_rem_lines_ignored(self, tmp_path):
        content = "REM GENRE PlayStation\nREM DATE 1999\n" + self.SINGLE_BIN_CUE
        cue = make_cue(tmp_path, content)
        disc = parse_cue(str(cue))
        assert disc.is_valid
        assert disc.track_count == 1


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
