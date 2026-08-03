"""
GameCube / Wii ISO Converter

Detects GameCube and Wii ISOs by reading disc magic bytes and converts
them to RVZ (or WIA) format via DolphinTool.

Supported workflows:
  GCN/Wii .iso  →  .rvz   (default, best compression)
  GCN/Wii .iso  →  .wia   (legacy, wider compat)
  .rvz / .wia   →  .iso   (restore)

Magic bytes:
  GCN:  offset 0x1C = C2 33 9F 3D
  Wii:  offset 0x18 = 5D 1C 9E A3, offset 0x1C = C2 33 9F 3D

Phase 5 Week 2 – Format Track
"""

import logging
import os
import shutil
import subprocess
import struct
import tempfile
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Optional, Callable, List

logger = logging.getLogger(__name__)

# ── magic constants ──────────────────────────────────────────────────────────

_GCN_MAGIC  = 0xC2339F3D   # offset 0x1C
_WII_MAGIC  = 0x5D1C9EA3   # offset 0x18

_MIN_ISO_SIZE = 1_474_560  # 1.4 MB – sanity check

# ── enums / dataclasses ──────────────────────────────────────────────────────

class DiscType(Enum):
    UNKNOWN   = "unknown"
    GAMECUBE  = "gamecube"
    WII       = "wii"


class RvzCompression(Enum):
    NONE   = "none"
    ZSTD   = "zstd"    # recommended
    BZIP2  = "bzip2"
    LZMA   = "lzma"
    LZMA2  = "lzma2"


@dataclass
class GameCubeConvertResult:
    """Result of a single ISO conversion."""
    success: bool
    input_path: str
    output_path: str
    disc_type: DiscType = DiscType.UNKNOWN
    input_size_bytes: int = 0
    output_size_bytes: int = 0
    duration_seconds: float = 0.0
    error_message: str = ""
    stdout: str = ""
    stderr: str = ""

    @property
    def compression_ratio(self) -> float:
        if self.input_size_bytes == 0:
            return 0.0
        return self.output_size_bytes / self.input_size_bytes

    @property
    def space_saved_mb(self) -> float:
        saved = self.input_size_bytes - self.output_size_bytes
        return saved / 1_048_576


# ── detector ────────────────────────────────────────────────────────────────

def detect_disc_type(iso_path: str) -> DiscType:
    """
    Read magic bytes from an ISO to determine if it is a GameCube or Wii disc.
    Returns DiscType.UNKNOWN if the file cannot be identified.
    """
    try:
        with open(iso_path, "rb") as f:
            # Read enough bytes to cover both magic offsets
            header = f.read(0x20)
        if len(header) < 0x20:
            return DiscType.UNKNOWN
        gcn_magic = struct.unpack_from(">I", header, 0x1C)[0]
        wii_magic = struct.unpack_from(">I", header, 0x18)[0]
        if wii_magic == _WII_MAGIC and gcn_magic == _GCN_MAGIC:
            return DiscType.WII
        if gcn_magic == _GCN_MAGIC:
            return DiscType.GAMECUBE
        return DiscType.UNKNOWN
    except (OSError, struct.error) as exc:
        logger.debug("detect_disc_type error for %s: %s", iso_path, exc)
        return DiscType.UNKNOWN


def is_gamecube_or_wii(iso_path: str) -> bool:
    """Convenience wrapper — True if disc is GCN or Wii."""
    return detect_disc_type(iso_path) in (DiscType.GAMECUBE, DiscType.WII)


# ── converter ────────────────────────────────────────────────────────────────

class GameCubeConverter:
    """
    Wraps DolphinTool to convert GCN/Wii ISOs to RVZ/WIA and back.

    Usage:
        conv = GameCubeConverter(dolphintool_path="DolphinTool.exe")
        result = conv.convert_to_rvz("game.iso", "game.rvz")
    """

    def __init__(
        self,
        dolphintool_path: Optional[str] = None,
        compression: RvzCompression = RvzCompression.ZSTD,
        compression_level: int = 5,
        block_size_kb: int = 128,
        log_callback: Optional[Callable[[str], None]] = None,
        timeout: int = 3600,
    ):
        self.dolphintool_path = dolphintool_path or self._find_dolphintool()
        self.compression = compression
        self.compression_level = max(1, min(22, compression_level))
        self.block_size_kb = block_size_kb
        self.log_callback = log_callback
        self.timeout = timeout

    # ---------------------------------------------------------------- public

    def is_available(self) -> bool:
        """Return True if DolphinTool is found and executable."""
        return self.dolphintool_path is not None and Path(self.dolphintool_path).exists()

    def convert_to_rvz(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        overwrite: bool = False,
    ) -> GameCubeConvertResult:
        """Convert a GCN/Wii ISO to RVZ."""
        return self._convert(input_path, output_path, ".rvz", overwrite)

    def convert_to_wia(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        overwrite: bool = False,
    ) -> GameCubeConvertResult:
        """Convert a GCN/Wii ISO to WIA."""
        return self._convert(input_path, output_path, ".wia", overwrite)

    def restore_to_iso(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        overwrite: bool = False,
    ) -> GameCubeConvertResult:
        """Convert RVZ/WIA back to ISO."""
        return self._convert(input_path, output_path, ".iso", overwrite)

    def detect(self, iso_path: str) -> DiscType:
        """Detect disc type of an ISO file."""
        return detect_disc_type(iso_path)

    def batch_convert(
        self,
        input_paths: List[str],
        output_dir: Optional[str] = None,
        target_ext: str = ".rvz",
        overwrite: bool = False,
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
    ) -> List[GameCubeConvertResult]:
        """
        Convert multiple ISOs.

        Args:
            input_paths: List of ISO file paths
            output_dir: Directory for outputs; defaults to same dir as input
            target_ext: ".rvz", ".wia", or ".iso"
            overwrite: Overwrite existing outputs
            progress_callback: (current, total, filename) callback

        Returns:
            List of GameCubeConvertResult
        """
        results = []
        for i, inp in enumerate(input_paths):
            if progress_callback:
                progress_callback(i, len(input_paths), Path(inp).name)
            out = None
            if output_dir:
                out = str(Path(output_dir) / (Path(inp).stem + target_ext))
            result = self._convert(inp, out, target_ext, overwrite)
            results.append(result)
        if progress_callback:
            progress_callback(len(input_paths), len(input_paths), "")
        return results

    # --------------------------------------------------------------- internal

    def _convert(
        self, input_path: str, output_path: Optional[str],
        target_ext: str, overwrite: bool,
    ) -> GameCubeConvertResult:
        import time

        if not self.is_available():
            return GameCubeConvertResult(
                success=False, input_path=input_path,
                output_path=output_path or "",
                error_message="DolphinTool not found. Install Dolphin Emulator.",
            )

        inp = Path(input_path)
        if not inp.exists():
            return GameCubeConvertResult(
                success=False, input_path=input_path,
                output_path=output_path or "",
                error_message=f"Input file not found: {input_path}",
            )

        # Detect disc type for non-ISO inputs only when converting *to* iso
        disc_type = DiscType.UNKNOWN
        if inp.suffix.lower() == ".iso":
            disc_type = detect_disc_type(input_path)

        out_path = Path(output_path) if output_path else inp.with_suffix(target_ext)

        if out_path.exists() and not overwrite:
            return GameCubeConvertResult(
                success=False, input_path=input_path,
                output_path=str(out_path), disc_type=disc_type,
                error_message=f"Output already exists: {out_path}",
            )

        cmd = self._build_command(str(inp), str(out_path), target_ext)
        self._log(f"[GameCube] {' '.join(cmd)}")

        input_size = inp.stat().st_size
        start = time.monotonic()
        try:
            proc = subprocess.run(
                cmd, capture_output=True, text=True, timeout=self.timeout
            )
            duration = time.monotonic() - start

            success = proc.returncode == 0 and out_path.exists()
            output_size = out_path.stat().st_size if out_path.exists() else 0

            if not success:
                error_msg = proc.stderr.strip() or f"Exit code {proc.returncode}"
            else:
                error_msg = ""
                self._log(
                    f"[GameCube] ✅ {inp.name} → {out_path.name} "
                    f"({input_size / 1e6:.1f} MB → {output_size / 1e6:.1f} MB)"
                )

            return GameCubeConvertResult(
                success=success,
                input_path=input_path,
                output_path=str(out_path),
                disc_type=disc_type,
                input_size_bytes=input_size,
                output_size_bytes=output_size,
                duration_seconds=duration,
                error_message=error_msg,
                stdout=proc.stdout,
                stderr=proc.stderr,
            )
        except subprocess.TimeoutExpired:
            return GameCubeConvertResult(
                success=False, input_path=input_path,
                output_path=str(out_path), disc_type=disc_type,
                input_size_bytes=input_size,
                duration_seconds=self.timeout,
                error_message="DolphinTool timed out",
            )
        except Exception as exc:
            return GameCubeConvertResult(
                success=False, input_path=input_path,
                output_path=str(out_path), disc_type=disc_type,
                input_size_bytes=input_size,
                error_message=str(exc),
            )

    def _build_command(self, inp: str, out: str, target_ext: str) -> list:
        """Build the DolphinTool command line."""
        ext = target_ext.lower()
        if ext in (".rvz", ".wia"):
            return [
                self.dolphintool_path,
                "convert",
                "-f", ext.lstrip("."),
                "-i", inp,
                "-o", out,
                "-c", self.compression.value,
                "-l", str(self.compression_level),
                "-b", str(self.block_size_kb * 1024),
            ]
        else:  # restore to iso
            return [
                self.dolphintool_path,
                "convert",
                "-f", "iso",
                "-i", inp,
                "-o", out,
            ]

    @staticmethod
    def _find_dolphintool() -> Optional[str]:
        """Search PATH for DolphinTool."""
        for name in ("DolphinTool.exe", "DolphinTool", "dolphin-tool"):
            found = shutil.which(name)
            if found:
                return found
        return None

    def _log(self, msg: str) -> None:
        if self.log_callback:
            self.log_callback(msg)
        logger.info(msg)


def create_gamecube_converter(
    dolphintool_path: Optional[str] = None,
    log_callback: Optional[Callable[[str], None]] = None,
) -> GameCubeConverter:
    """Factory for GameCubeConverter."""
    return GameCubeConverter(
        dolphintool_path=dolphintool_path,
        log_callback=log_callback,
    )
