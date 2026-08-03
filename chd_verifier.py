"""
CHD Verifier — Phase 5 Format Track

Runs `chdman verify` after each conversion, parses the output,
and returns a structured result. Supports:
- Single file verification
- Batch verification with progress callbacks
- Auto-quarantine of corrupt outputs
- Integration with ConversionErrorHandler and MetricsStore
"""

import subprocess
import shutil
import time
import logging
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Optional, Callable, List, Dict

logger = logging.getLogger(__name__)


class VerifyStatus(Enum):
    PASS      = "pass"
    FAIL      = "fail"
    SKIPPED   = "skipped"   # chdman not available
    ERROR     = "error"     # could not run command


@dataclass
class VerifyResult:
    """Result of a single CHD verification."""
    file_path: str
    status: VerifyStatus
    duration_seconds: float = 0.0
    error_message: str = ""
    chdman_output: str = ""

    @property
    def passed(self) -> bool:
        return self.status == VerifyStatus.PASS

    def summary(self) -> str:
        icon = {"pass": "✅", "fail": "❌", "skipped": "⚠️", "error": "💥"}[self.status.value]
        name = Path(self.file_path).name
        if self.status == VerifyStatus.PASS:
            return f"{icon} {name} — verified OK ({self.duration_seconds:.1f}s)"
        return f"{icon} {name} — {self.status.value}: {self.error_message[:80]}"


@dataclass
class BatchVerifyReport:
    """Aggregate report for a batch verification run."""
    results: List[VerifyResult] = field(default_factory=list)
    quarantine_dir: Optional[str] = None

    @property
    def total(self) -> int:
        return len(self.results)

    @property
    def passed(self) -> int:
        return sum(1 for r in self.results if r.status == VerifyStatus.PASS)

    @property
    def failed(self) -> int:
        return sum(1 for r in self.results if r.status == VerifyStatus.FAIL)

    @property
    def errors(self) -> int:
        return sum(1 for r in self.results if r.status == VerifyStatus.ERROR)

    @property
    def pass_rate(self) -> float:
        return self.passed / self.total if self.total else 0.0

    def format_summary(self) -> str:
        lines = [
            "═" * 50,
            f"CHD Verification Summary — {self.total} files",
            f"  ✅ Passed   : {self.passed}",
            f"  ❌ Failed   : {self.failed}",
            f"  💥 Errors   : {self.errors}",
            f"  Pass rate  : {self.pass_rate * 100:.1f}%",
        ]
        if self.quarantine_dir and self.failed:
            lines.append(f"  Quarantine : {self.quarantine_dir}")
        failures = [r for r in self.results if r.status != VerifyStatus.PASS]
        if failures:
            lines.append("\nFailed / Errored files:")
            for r in failures[:10]:
                lines.append(f"  • {r.summary()}")
            if len(failures) > 10:
                lines.append(f"  … and {len(failures) - 10} more")
        lines.append("═" * 50)
        return "\n".join(lines)


class CHDVerifier:
    """
    Verifies CHD files using chdman's built-in verify command.

    Runs `chdman verify -i <file>` and parses exit code + stdout.
    chdman exits 0 on success, non-zero on any error.
    """

    def __init__(
        self,
        chdman_path: Optional[str] = None,
        quarantine_dir: Optional[str] = None,
        timeout_seconds: int = 300,
        log_callback: Optional[Callable[[str], None]] = None,
    ):
        """
        Args:
            chdman_path: Path to chdman executable. Auto-detected if None.
            quarantine_dir: Move failed CHDs here. None = don't move.
            timeout_seconds: Per-file timeout.
            log_callback: Optional logging callback.
        """
        self.chdman_path     = chdman_path or self._find_chdman()
        self.quarantine_dir  = Path(quarantine_dir) if quarantine_dir else None
        self.timeout_seconds = timeout_seconds
        self.log_callback    = log_callback

    def _log(self, msg: str) -> None:
        if self.log_callback:
            self.log_callback(msg)
        logger.info(msg)

    @staticmethod
    def _find_chdman() -> Optional[str]:
        """Auto-detect chdman on PATH or common locations."""
        found = shutil.which("chdman")
        if found:
            return found
        candidates = [
            Path("chdman.exe"),
            Path("tools") / "chdman.exe",
            Path("chdman"),
            Path("tools") / "chdman",
        ]
        for c in candidates:
            if c.exists():
                return str(c)
        return None

    @property
    def available(self) -> bool:
        """True if chdman executable is usable."""
        return self.chdman_path is not None and Path(self.chdman_path).exists()

    def verify_file(self, chd_path: str) -> VerifyResult:
        """
        Verify a single CHD file.

        Args:
            chd_path: Path to .chd file

        Returns:
            VerifyResult with pass/fail status
        """
        chd = Path(chd_path)
        if not chd.exists():
            return VerifyResult(
                file_path=chd_path,
                status=VerifyStatus.ERROR,
                error_message=f"File not found: {chd_path}",
            )

        if not self.available:
            return VerifyResult(
                file_path=chd_path,
                status=VerifyStatus.SKIPPED,
                error_message="chdman not available",
            )

        start = time.time()
        try:
            proc = subprocess.run(
                [self.chdman_path, "verify", "-i", str(chd)],
                capture_output=True,
                timeout=self.timeout_seconds,
            )
            duration = time.time() - start
            output = (proc.stdout + proc.stderr).decode("utf-8", errors="replace").strip()

            if proc.returncode == 0:
                self._log(f"✅ Verified: {chd.name} ({duration:.1f}s)")
                return VerifyResult(
                    file_path=chd_path,
                    status=VerifyStatus.PASS,
                    duration_seconds=duration,
                    chdman_output=output,
                )
            else:
                err = self._parse_error(output)
                self._log(f"❌ Verification FAILED: {chd.name} — {err}")
                result = VerifyResult(
                    file_path=chd_path,
                    status=VerifyStatus.FAIL,
                    duration_seconds=duration,
                    error_message=err,
                    chdman_output=output,
                )
                if self.quarantine_dir:
                    self._quarantine(chd, result)
                return result

        except subprocess.TimeoutExpired:
            duration = time.time() - start
            self._log(f"💥 Verify timeout: {chd.name}")
            return VerifyResult(
                file_path=chd_path,
                status=VerifyStatus.ERROR,
                duration_seconds=duration,
                error_message=f"Timed out after {self.timeout_seconds}s",
            )
        except Exception as exc:
            duration = time.time() - start
            self._log(f"💥 Verify error: {chd.name} — {exc}")
            return VerifyResult(
                file_path=chd_path,
                status=VerifyStatus.ERROR,
                duration_seconds=duration,
                error_message=str(exc),
            )

    def verify_batch(
        self,
        chd_paths: List[str],
        progress_callback: Optional[Callable[[int, int, VerifyResult], None]] = None,
        stop_on_first_failure: bool = False,
    ) -> BatchVerifyReport:
        """
        Verify multiple CHD files.

        Args:
            chd_paths: List of .chd file paths
            progress_callback: Called with (completed, total, last_result)
            stop_on_first_failure: Abort after first FAIL/ERROR

        Returns:
            BatchVerifyReport
        """
        report = BatchVerifyReport(
            quarantine_dir=str(self.quarantine_dir) if self.quarantine_dir else None
        )
        total = len(chd_paths)

        for i, path in enumerate(chd_paths):
            result = self.verify_file(path)
            report.results.append(result)

            if progress_callback:
                progress_callback(i + 1, total, result)

            if stop_on_first_failure and result.status in (VerifyStatus.FAIL, VerifyStatus.ERROR):
                self._log(f"Stopping batch after first failure: {Path(path).name}")
                break

        self._log(report.format_summary())
        return report

    def _quarantine(self, chd: Path, result: VerifyResult) -> None:
        """Move a failed CHD to the quarantine directory."""
        if not self.quarantine_dir:
            return
        try:
            self.quarantine_dir.mkdir(parents=True, exist_ok=True)
            dest = self.quarantine_dir / chd.name
            chd.rename(dest)
            self._log(f"🔒 Quarantined: {chd.name} → {self.quarantine_dir}")
        except Exception as exc:
            self._log(f"⚠️  Could not quarantine {chd.name}: {exc}")

    @staticmethod
    def _parse_error(output: str) -> str:
        """Extract meaningful error from chdman output."""
        for line in output.splitlines():
            line = line.strip()
            if any(kw in line.lower() for kw in ["error", "fail", "invalid", "corrupt"]):
                return line[:120]
        return output[:120] if output else "Unknown error"


def create_chd_verifier(
    chdman_path: Optional[str] = None,
    quarantine_dir: Optional[str] = None,
    log_callback: Optional[Callable[[str], None]] = None,
) -> CHDVerifier:
    """Factory function for CHDVerifier."""
    return CHDVerifier(
        chdman_path=chdman_path,
        quarantine_dir=quarantine_dir,
        log_callback=log_callback,
    )
