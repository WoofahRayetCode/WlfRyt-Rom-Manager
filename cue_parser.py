"""
CUE Sheet Parser — Phase 5 Format Track

Parses CD-DA / CD-ROM CUE sheets to:
- Enumerate tracks and their BIN files
- Detect disc type (PlayStation, Saturn, GameCube, etc.)
- Validate that all referenced BIN files exist
- Produce a normalised CueDisc object ready for chdman

Handles both single-BIN (all tracks in one .bin) and multi-BIN
(one .bin per track, common in Saturn rips) layouts.
"""

import re
import logging
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import List, Optional, Dict

logger = logging.getLogger(__name__)


class TrackMode(Enum):
    """CUE track data mode."""
    AUDIO       = "AUDIO"
    MODE1_2048  = "MODE1/2048"
    MODE1_2352  = "MODE1/2352"
    MODE2_2048  = "MODE2/2048"
    MODE2_2352  = "MODE2/2352"
    MODE2_2336  = "MODE2/2336"
    UNKNOWN     = "UNKNOWN"


class DiscSystem(Enum):
    """Detected disc system / console."""
    PLAYSTATION1   = "PlayStation 1"
    PLAYSTATION2   = "PlayStation 2"
    SATURN         = "Saturn"
    DREAMCAST      = "Dreamcast"
    PC_ENGINE      = "PC-Engine / TurboGrafx"
    PCFX           = "PC-FX"
    GENERIC_CDROM  = "Generic CD-ROM"
    CDDA           = "CD Audio"
    UNKNOWN        = "Unknown"


@dataclass
class CueTrack:
    """A single track within a CUE sheet."""
    number: int                        # 1-based track number
    mode: TrackMode
    bin_file: str                      # Referenced BIN filename
    pregap_frames: int = 0             # PREGAP in 1/75 sec frames
    index_00_frames: Optional[int] = None  # INDEX 00 MM:SS:FF offset
    index_01_frames: int = 0           # INDEX 01 MM:SS:FF offset (start of data)

    @property
    def is_audio(self) -> bool:
        return self.mode == TrackMode.AUDIO

    @property
    def is_data(self) -> bool:
        return not self.is_audio

    @property
    def sector_size(self) -> int:
        """Raw sector size in bytes."""
        _sizes = {
            TrackMode.MODE1_2048: 2048,
            TrackMode.MODE1_2352: 2352,
            TrackMode.MODE2_2048: 2048,
            TrackMode.MODE2_2352: 2352,
            TrackMode.MODE2_2336: 2336,
            TrackMode.AUDIO:      2352,
            TrackMode.UNKNOWN:    2352,
        }
        return _sizes.get(self.mode, 2352)


@dataclass
class CueDisc:
    """Parsed representation of a CUE sheet."""
    cue_path: str
    tracks: List[CueTrack] = field(default_factory=list)
    bin_files: List[str] = field(default_factory=list)   # unique, in appearance order
    system: DiscSystem = DiscSystem.UNKNOWN
    parse_errors: List[str] = field(default_factory=list)

    @property
    def cue_dir(self) -> Path:
        return Path(self.cue_path).parent

    @property
    def track_count(self) -> int:
        return len(self.tracks)

    @property
    def is_multi_bin(self) -> bool:
        """True if each track references its own BIN file."""
        return len(set(t.bin_file for t in self.tracks)) > 1

    @property
    def is_valid(self) -> bool:
        return len(self.tracks) > 0 and len(self.parse_errors) == 0

    @property
    def data_tracks(self) -> List[CueTrack]:
        return [t for t in self.tracks if t.is_data]

    @property
    def audio_tracks(self) -> List[CueTrack]:
        return [t for t in self.tracks if t.is_audio]

    def missing_bins(self) -> List[str]:
        """List BIN files referenced in the CUE that don't exist on disk."""
        missing = []
        for name in self.bin_files:
            if not (self.cue_dir / name).exists():
                missing.append(name)
        return missing

    def all_bins_present(self) -> bool:
        return len(self.missing_bins()) == 0

    def summary(self) -> str:
        ok = "✅" if self.all_bins_present() else "❌ missing BINs"
        layout = "multi-BIN" if self.is_multi_bin else "single-BIN"
        return (
            f"{Path(self.cue_path).name} | {self.system.value} | "
            f"{self.track_count} tracks ({layout}) | {ok}"
        )


def _mmssff_to_frames(mmssff: str) -> int:
    """Convert MM:SS:FF timecode to total 1/75-sec frames."""
    parts = mmssff.strip().split(":")
    if len(parts) != 3:
        return 0
    mm, ss, ff = int(parts[0]), int(parts[1]), int(parts[2])
    return (mm * 60 + ss) * 75 + ff


def _parse_track_mode(token: str) -> TrackMode:
    token = token.upper().strip()
    try:
        return TrackMode(token)
    except ValueError:
        return TrackMode.UNKNOWN


class CueParser:
    """
    Parses CUE sheets into CueDisc objects.

    Handles:
    - Single-BIN and multi-BIN layouts
    - PREGAP and INDEX 00 / INDEX 01 directives
    - REM comments (used by some tools for metadata)
    - Windows and Unix path separators in FILE lines
    - Quoted and unquoted filenames
    """

    # System fingerprinting: detect by data track mode patterns
    _SATURN_MAGIC   = b"SEGA SEGASATURN"
    _PS1_MAGIC      = b"  Licensed  by"
    _PS2_MAGIC      = b"PlayStation 2"
    _DREAMCAST_MAGIC = b"SEGA SEGAKATANA"
    _PCFX_MAGIC     = b"PC-FX:Hu_CD-ROM"

    def __init__(self, log_callback=None):
        self.log_callback = log_callback

    def _log(self, msg: str) -> None:
        if self.log_callback:
            self.log_callback(msg)
        logger.debug(msg)

    def parse(self, cue_path: str) -> CueDisc:
        """
        Parse a CUE sheet file.

        Args:
            cue_path: Absolute or relative path to .cue file

        Returns:
            CueDisc with parsed data (check is_valid and parse_errors)
        """
        disc = CueDisc(cue_path=str(cue_path))
        cue_file = Path(cue_path)

        if not cue_file.exists():
            disc.parse_errors.append(f"CUE file not found: {cue_path}")
            return disc

        try:
            text = cue_file.read_text(encoding="utf-8", errors="replace")
        except Exception as exc:
            disc.parse_errors.append(f"Cannot read CUE: {exc}")
            return disc

        current_bin: Optional[str] = None
        current_track: Optional[CueTrack] = None

        for lineno, raw_line in enumerate(text.splitlines(), 1):
            line = raw_line.strip()
            if not line or line.upper().startswith("REM"):
                continue

            upper = line.upper()

            # FILE "name.bin" BINARY
            if upper.startswith("FILE "):
                m = re.match(r'FILE\s+"([^"]+)"\s+\S+', line, re.IGNORECASE)
                if not m:
                    # Try unquoted
                    m = re.match(r'FILE\s+(\S+)\s+\S+', line, re.IGNORECASE)
                if m:
                    current_bin = Path(m.group(1).replace("\\", "/")).name
                    if current_bin not in disc.bin_files:
                        disc.bin_files.append(current_bin)
                else:
                    disc.parse_errors.append(f"Line {lineno}: cannot parse FILE: {line!r}")
                continue

            # TRACK 01 MODE1/2352
            if upper.startswith("TRACK "):
                if current_track is not None:
                    disc.tracks.append(current_track)
                m = re.match(r'TRACK\s+(\d+)\s+(\S+)', line, re.IGNORECASE)
                if m:
                    num  = int(m.group(1))
                    mode = _parse_track_mode(m.group(2))
                    current_track = CueTrack(
                        number=num,
                        mode=mode,
                        bin_file=current_bin or "",
                    )
                else:
                    disc.parse_errors.append(f"Line {lineno}: cannot parse TRACK: {line!r}")
                continue

            # PREGAP MM:SS:FF
            if upper.startswith("PREGAP ") and current_track:
                m = re.match(r'PREGAP\s+(\d+:\d+:\d+)', line, re.IGNORECASE)
                if m:
                    current_track.pregap_frames = _mmssff_to_frames(m.group(1))
                continue

            # INDEX 00/01 MM:SS:FF
            if upper.startswith("INDEX ") and current_track:
                m = re.match(r'INDEX\s+(\d+)\s+(\d+:\d+:\d+)', line, re.IGNORECASE)
                if m:
                    idx_num    = int(m.group(1))
                    idx_frames = _mmssff_to_frames(m.group(2))
                    if idx_num == 0:
                        current_track.index_00_frames = idx_frames
                    elif idx_num == 1:
                        current_track.index_01_frames = idx_frames
                continue

        # Don't forget the last track
        if current_track is not None:
            disc.tracks.append(current_track)

        if not disc.tracks:
            disc.parse_errors.append("No TRACK entries found in CUE file")

        # Detect system
        disc.system = self._detect_system(disc)
        return disc

    def _detect_system(self, disc: CueDisc) -> DiscSystem:
        """
        Fingerprint the disc system by reading the first data track.
        Falls back to heuristic track-count/mode analysis.
        """
        # Try byte-level fingerprint from first data track BIN
        for track in disc.data_tracks:
            bin_path = disc.cue_dir / track.bin_file
            if bin_path.exists():
                try:
                    sample = bin_path.read_bytes()[16:32]  # sector 0 ID area
                    # Try 2352-byte sector header offset
                    for offset in (0, 16, 24):
                        chunk = bin_path.read_bytes()[offset: offset + 32]
                        if self._SATURN_MAGIC in chunk:
                            return DiscSystem.SATURN
                        if self._PS1_MAGIC in chunk:
                            return DiscSystem.PLAYSTATION1
                        if self._PS2_MAGIC in chunk:
                            return DiscSystem.PLAYSTATION2
                        if self._DREAMCAST_MAGIC in chunk:
                            return DiscSystem.DREAMCAST
                        if self._PCFX_MAGIC in chunk:
                            return DiscSystem.PCFX
                except Exception:
                    pass

        # Heuristic: pure audio = CDDA
        if all(t.is_audio for t in disc.tracks):
            return DiscSystem.CDDA

        # Mixed: data track 1 + audio = likely Saturn or PlayStation
        if disc.data_tracks and disc.audio_tracks:
            if len(disc.audio_tracks) >= 5:
                return DiscSystem.SATURN   # Saturn commonly has many audio tracks
            return DiscSystem.PLAYSTATION1

        # Single data track, no audio
        if disc.data_tracks and not disc.audio_tracks:
            return DiscSystem.GENERIC_CDROM

        return DiscSystem.UNKNOWN


def parse_cue(cue_path: str, log_callback=None) -> CueDisc:
    """
    Convenience function: parse a CUE sheet and return a CueDisc.

    Args:
        cue_path: Path to .cue file
        log_callback: Optional logging callback

    Returns:
        CueDisc
    """
    return CueParser(log_callback=log_callback).parse(cue_path)
