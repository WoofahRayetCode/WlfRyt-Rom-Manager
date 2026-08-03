"""
DAT File Matcher — No-Intro / Redump ROM Verification

Loads No-Intro/Redump XML DAT files and matches ROM files against them
by CRC32, MD5, and/or SHA1. Reports known-good, known-bad, and unmatched files.

Supported DAT format:
  <datafile>
    <header>…</header>
    <game name="…">
      <rom name="…" size="…" crc="…" md5="…" sha1="…"/>
    </game>
  </datafile>

Phase 5 Week 2 – Format Track
"""

import csv
import hashlib
import logging
import os
import struct
import time
import zlib
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Optional, Callable, List, Dict, Tuple
from xml.etree import ElementTree as ET

logger = logging.getLogger(__name__)

_CHUNK = 1 << 20  # 1 MB read chunks for hashing


# ── enums / dataclasses ──────────────────────────────────────────────────────

class MatchStatus(Enum):
    GOOD      = "good"      # ROM matched exactly in DAT
    BAD       = "bad"       # ROM found in DAT but checksums differ
    UNKNOWN   = "unknown"   # ROM not found in DAT at all
    RENAMED   = "renamed"   # ROM matched by checksum but filename differs


@dataclass
class DatEntry:
    """A single ROM entry from the DAT file."""
    game_name: str
    rom_name: str
    size: int = 0
    crc:  str = ""   # 8-char hex, uppercase
    md5:  str = ""   # 32-char hex, uppercase
    sha1: str = ""   # 40-char hex, uppercase


@dataclass
class MatchResult:
    """Result of matching one ROM file against the DAT."""
    file_path: str
    status: MatchStatus
    dat_entry: Optional[DatEntry] = None
    expected_name: str = ""     # DAT rom_name (if matched/renamed)
    file_crc:  str = ""
    file_md5:  str = ""
    file_sha1: str = ""
    file_size: int = 0
    error: str = ""

    @property
    def status_icon(self) -> str:
        return {
            MatchStatus.GOOD:    "✅",
            MatchStatus.BAD:     "❌",
            MatchStatus.UNKNOWN: "❓",
            MatchStatus.RENAMED: "🔄",
        }[self.status]


@dataclass
class BatchMatchReport:
    """Aggregate results for a batch verification run."""
    results: List[MatchResult] = field(default_factory=list)
    dat_name: str = ""
    dat_version: str = ""

    @property
    def good(self)    -> List[MatchResult]: return [r for r in self.results if r.status == MatchStatus.GOOD]
    @property
    def bad(self)     -> List[MatchResult]: return [r for r in self.results if r.status == MatchStatus.BAD]
    @property
    def unknown(self) -> List[MatchResult]: return [r for r in self.results if r.status == MatchStatus.UNKNOWN]
    @property
    def renamed(self) -> List[MatchResult]: return [r for r in self.results if r.status == MatchStatus.RENAMED]

    def format_summary(self) -> str:
        total = len(self.results)
        lines = [
            f"DAT: {self.dat_name} {self.dat_version}",
            f"Total: {total}",
            f"  ✅ Good:    {len(self.good)}",
            f"  🔄 Renamed: {len(self.renamed)}",
            f"  ❌ Bad:     {len(self.bad)}",
            f"  ❓ Unknown: {len(self.unknown)}",
        ]
        return "\n".join(lines)


# ── hashing utilities ────────────────────────────────────────────────────────

def _compute_hashes(
    path: str,
    compute_md5: bool = True,
    compute_sha1: bool = True,
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> Tuple[str, str, str]:
    """
    Compute CRC32, MD5, SHA1 of a file.
    Returns (crc_hex, md5_hex, sha1_hex) — all uppercase.
    Pass compute_md5=False / compute_sha1=False to skip those for speed.
    """
    crc_val = 0
    md5_h   = hashlib.md5()   if compute_md5  else None
    sha1_h  = hashlib.sha1()  if compute_sha1 else None

    total = os.path.getsize(path)
    done  = 0
    with open(path, "rb") as f:
        while True:
            chunk = f.read(_CHUNK)
            if not chunk:
                break
            crc_val = zlib.crc32(chunk, crc_val)
            if md5_h:
                md5_h.update(chunk)
            if sha1_h:
                sha1_h.update(chunk)
            done += len(chunk)
            if progress_callback:
                progress_callback(done, total)

    crc_hex  = format(crc_val & 0xFFFFFFFF, "08X")
    md5_hex  = md5_h.hexdigest().upper()  if md5_h  else ""
    sha1_hex = sha1_h.hexdigest().upper() if sha1_h else ""
    return crc_hex, md5_hex, sha1_hex


# ── DAT parser / matcher ─────────────────────────────────────────────────────

class DatMatcher:
    """
    No-Intro / Redump DAT matcher.

    Usage:
        matcher = DatMatcher()
        matcher.load_dat("No-Intro - Sony - PlayStation (20240101-000000).dat")
        result = matcher.match_file("Crash Bandicoot (USA).bin")
        print(result.status_icon, result.status.value)
    """

    def __init__(
        self,
        compute_md5: bool = True,
        compute_sha1: bool = True,
        log_callback: Optional[Callable[[str], None]] = None,
    ):
        self.compute_md5 = compute_md5
        self.compute_sha1 = compute_sha1
        self.log_callback = log_callback

        self._entries:     List[DatEntry] = []
        self._by_crc:      Dict[str, DatEntry] = {}   # crc → entry
        self._by_md5:      Dict[str, DatEntry] = {}
        self._by_sha1:     Dict[str, DatEntry] = {}
        self._by_name:     Dict[str, DatEntry] = {}   # rom_name (lower) → entry

        self.dat_name    = ""
        self.dat_version = ""

    # ---------------------------------------------------------------- loading

    def load_dat(self, dat_path: str) -> int:
        """
        Parse a No-Intro/Redump XML DAT file.
        Returns the number of ROM entries loaded.
        """
        tree = ET.parse(dat_path)
        root = tree.getroot()

        header = root.find("header")
        if header is not None:
            self.dat_name    = (header.findtext("name")    or "").strip()
            self.dat_version = (header.findtext("version") or "").strip()

        self._entries.clear()
        self._by_crc.clear()
        self._by_md5.clear()
        self._by_sha1.clear()
        self._by_name.clear()

        for game in root.findall("game"):
            game_name = game.get("name", "")
            for rom in game.findall("rom"):
                entry = DatEntry(
                    game_name=game_name,
                    rom_name=rom.get("name", ""),
                    size=int(rom.get("size", 0)),
                    crc=(rom.get("crc", "") or "").upper(),
                    md5=(rom.get("md5", "") or "").upper(),
                    sha1=(rom.get("sha1", "") or "").upper(),
                )
                self._entries.append(entry)
                if entry.crc:
                    self._by_crc[entry.crc]   = entry
                if entry.md5:
                    self._by_md5[entry.md5]   = entry
                if entry.sha1:
                    self._by_sha1[entry.sha1] = entry
                self._by_name[entry.rom_name.lower()] = entry

        self._log(f"[DAT] Loaded {len(self._entries)} entries from {Path(dat_path).name}")
        return len(self._entries)

    @property
    def entry_count(self) -> int:
        return len(self._entries)

    # --------------------------------------------------------------- matching

    def match_file(
        self,
        file_path: str,
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> MatchResult:
        """
        Compute checksums for one file and match against the loaded DAT.

        Match priority: SHA1 → MD5 → CRC32 → filename-only
        """
        p = Path(file_path)
        if not p.exists():
            return MatchResult(
                file_path=file_path, status=MatchStatus.UNKNOWN,
                error=f"File not found: {file_path}",
            )

        file_size = p.stat().st_size

        try:
            crc, md5, sha1 = _compute_hashes(
                file_path,
                compute_md5=self.compute_md5,
                compute_sha1=self.compute_sha1,
                progress_callback=progress_callback,
            )
        except Exception as exc:
            return MatchResult(
                file_path=file_path, status=MatchStatus.UNKNOWN,
                file_size=file_size, error=str(exc),
            )

        result = MatchResult(
            file_path=file_path, status=MatchStatus.UNKNOWN,
            file_crc=crc, file_md5=md5, file_sha1=sha1,
            file_size=file_size,
        )

        # Try checksum lookups (most precise first)
        entry = None
        if sha1 and sha1 in self._by_sha1:
            entry = self._by_sha1[sha1]
        elif md5 and md5 in self._by_md5:
            entry = self._by_md5[md5]
        elif crc and crc in self._by_crc:
            entry = self._by_crc[crc]

        if entry is not None:
            result.dat_entry     = entry
            result.expected_name = entry.rom_name
            # Good if filename matches, renamed if it doesn't
            result.status = (
                MatchStatus.GOOD
                if p.name.lower() == entry.rom_name.lower()
                else MatchStatus.RENAMED
            )
            return result

        # Filename-only fallback (no checksum match)
        name_entry = self._by_name.get(p.name.lower())
        if name_entry:
            result.dat_entry     = name_entry
            result.expected_name = name_entry.rom_name
            result.status        = MatchStatus.BAD  # name match but CRC differs
        else:
            result.status = MatchStatus.UNKNOWN

        return result

    def match_files(
        self,
        file_paths: List[str],
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
    ) -> BatchMatchReport:
        """
        Match a list of files and return a BatchMatchReport.

        progress_callback: (current, total, filename)
        """
        report = BatchMatchReport(
            dat_name=self.dat_name,
            dat_version=self.dat_version,
        )
        total = len(file_paths)
        for i, path in enumerate(file_paths):
            if progress_callback:
                progress_callback(i, total, Path(path).name)
            result = self.match_file(path)
            report.results.append(result)
            icon = result.status_icon
            self._log(f"[DAT] {icon} {Path(path).name} → {result.status.value}")
        if progress_callback:
            progress_callback(total, total, "")
        return report

    def scan_directory(
        self,
        directory: str,
        extensions: Optional[List[str]] = None,
        recursive: bool = False,
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
    ) -> BatchMatchReport:
        """
        Scan all ROM files in a directory and match against the DAT.

        Args:
            directory: Root directory to scan
            extensions: List of extensions to include (e.g. [".iso", ".bin"]).
                        None = all files.
            recursive: If True, scan subdirectories
        """
        root = Path(directory)
        if recursive:
            all_files = list(root.rglob("*"))
        else:
            all_files = list(root.iterdir())

        if extensions:
            exts = {e.lower() for e in extensions}
            all_files = [f for f in all_files if f.is_file() and f.suffix.lower() in exts]
        else:
            all_files = [f for f in all_files if f.is_file()]

        return self.match_files([str(f) for f in all_files], progress_callback)

    def export_results_csv(self, report: BatchMatchReport, output_path: str) -> None:
        """Write a BatchMatchReport to a CSV file."""
        with open(output_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([
                "file", "status", "crc", "md5", "sha1",
                "expected_name", "game_name", "size_bytes", "error",
            ])
            for r in report.results:
                writer.writerow([
                    r.file_path, r.status.value,
                    r.file_crc, r.file_md5, r.file_sha1,
                    r.expected_name,
                    r.dat_entry.game_name if r.dat_entry else "",
                    r.file_size, r.error,
                ])

    # --------------------------------------------------------------- private

    def _log(self, msg: str) -> None:
        if self.log_callback:
            self.log_callback(msg)
        logger.info(msg)


def create_dat_matcher(
    compute_md5: bool = True,
    compute_sha1: bool = True,
    log_callback: Optional[Callable[[str], None]] = None,
) -> DatMatcher:
    """Factory for DatMatcher."""
    return DatMatcher(
        compute_md5=compute_md5,
        compute_sha1=compute_sha1,
        log_callback=log_callback,
    )
