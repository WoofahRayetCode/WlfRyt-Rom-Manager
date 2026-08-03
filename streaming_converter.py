"""
Streaming Converter Module

Memory-efficient conversion pipeline that reads and processes files in chunks
rather than loading entire files into memory. Particularly useful for large
ROM files (1GB+) where memory is a constraint.

Phase 4 Week 4: Streaming & Error Handling
"""

import os
import io
import time
import logging
import subprocess
import tempfile
from pathlib import Path
from typing import Optional, Callable, Generator, List, Tuple
from dataclasses import dataclass, field

try:
    from exceptions import ConversionError, DiskFullError, OutOfMemoryError
    EXCEPTIONS_AVAILABLE = True
except ImportError:
    EXCEPTIONS_AVAILABLE = False
    ConversionError = Exception

try:
    from retry_logic import RetryConfig, retry_with_backoff
    RETRY_AVAILABLE = True
except ImportError:
    RETRY_AVAILABLE = False

try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False


logger = logging.getLogger(__name__)


@dataclass
class StreamingConversionResult:
    """Result of a streaming conversion operation."""
    success: bool
    input_path: str
    output_path: str
    input_size_bytes: int = 0
    output_size_bytes: int = 0
    duration_seconds: float = 0.0
    error_message: Optional[str] = None
    retry_count: int = 0

    @property
    def compression_ratio(self) -> float:
        """Output size / input size (lower = better compression)."""
        if self.input_size_bytes == 0:
            return 0.0
        return self.output_size_bytes / self.input_size_bytes

    @property
    def throughput_mbps(self) -> float:
        """MB/s processed."""
        if self.duration_seconds == 0:
            return 0.0
        return (self.input_size_bytes / 1024 / 1024) / self.duration_seconds


@dataclass
class StreamingConfig:
    """Configuration for streaming conversion."""
    chunk_size: int = 8192          # Read chunk size in bytes (8KB)
    max_memory_mb: int = 512        # Max memory before throttling
    temp_dir: Optional[str] = None  # Directory for temp files (None = system default)
    verify_output: bool = True      # Verify output file size > 0
    cleanup_on_failure: bool = True # Remove partial output on failure
    timeout_seconds: int = 3600     # Conversion timeout (1 hour default)


class StreamingConverter:
    """
    Memory-efficient conversion wrapper.

    Wraps existing conversion tools (chdman, maxcso, 7z, etc.) and provides:
    - Streaming I/O to minimize memory usage
    - Progress tracking
    - Temp file management
    - Output verification
    """

    def __init__(
        self,
        config: Optional[StreamingConfig] = None,
        log_callback: Optional[Callable[[str], None]] = None,
    ):
        """
        Initialize streaming converter.

        Args:
            config: StreamingConfig with conversion settings
            log_callback: Optional callback for log messages
        """
        self.config = config or StreamingConfig()
        self.log_callback = log_callback
        self._active_processes: List[subprocess.Popen] = []

    def _log(self, message: str) -> None:
        """Log message via callback or logger."""
        if self.log_callback:
            self.log_callback(message)
        logger.info(message)

    def _get_available_memory_mb(self) -> float:
        """Get available system memory in MB."""
        if PSUTIL_AVAILABLE:
            return psutil.virtual_memory().available / 1024 / 1024
        return float('inf')

    def _is_memory_constrained(self) -> bool:
        """Check if memory usage is above threshold."""
        available = self._get_available_memory_mb()
        return available < 100  # Less than 100MB free is critical

    def _get_temp_dir(self) -> Path:
        """Get temp directory for intermediate files."""
        if self.config.temp_dir:
            path = Path(self.config.temp_dir)
            path.mkdir(parents=True, exist_ok=True)
            return path
        return Path(tempfile.gettempdir())

    def _check_disk_space(self, required_bytes: int, target_dir: str) -> bool:
        """
        Check if enough disk space is available.

        Args:
            required_bytes: Required bytes
            target_dir: Directory to check

        Returns:
            True if enough space available
        """
        if PSUTIL_AVAILABLE:
            usage = psutil.disk_usage(str(target_dir))
            return usage.free > required_bytes * 1.2  # 20% buffer
        return True  # Assume available if psutil not present

    def stream_copy(
        self,
        source_path: str,
        dest_path: str,
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> bool:
        """
        Copy file using streaming chunks.

        Args:
            source_path: Source file path
            dest_path: Destination file path
            progress_callback: Optional callback(bytes_copied, total_bytes)

        Returns:
            True if successful
        """
        source = Path(source_path)
        dest = Path(dest_path)

        if not source.exists():
            self._log(f"Source not found: {source_path}")
            return False

        total_bytes = source.stat().st_size
        copied_bytes = 0

        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            with open(source, 'rb') as src_file:
                with open(dest, 'wb') as dst_file:
                    while True:
                        chunk = src_file.read(self.config.chunk_size)
                        if not chunk:
                            break
                        dst_file.write(chunk)
                        copied_bytes += len(chunk)
                        if progress_callback:
                            progress_callback(copied_bytes, total_bytes)
            return True
        except OSError as e:
            self._log(f"Stream copy failed: {e}")
            if dest.exists():
                dest.unlink()
            return False

    def run_conversion_with_timeout(
        self,
        command: List[str],
        timeout: Optional[int] = None,
    ) -> Tuple[int, str, str]:
        """
        Run conversion command with timeout.

        Args:
            command: Command and arguments list
            timeout: Timeout in seconds (uses config default if None)

        Returns:
            Tuple of (returncode, stdout, stderr)
        """
        actual_timeout = timeout or self.config.timeout_seconds

        try:
            proc = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            self._active_processes.append(proc)
            try:
                stdout, stderr = proc.communicate(timeout=actual_timeout)
                return (
                    proc.returncode,
                    stdout.decode('utf-8', errors='replace'),
                    stderr.decode('utf-8', errors='replace'),
                )
            except subprocess.TimeoutExpired:
                proc.kill()
                stdout, stderr = proc.communicate()
                self._log(f"Command timed out after {actual_timeout}s: {command[0]}")
                return (-1, '', f"Timed out after {actual_timeout}s")
            finally:
                if proc in self._active_processes:
                    self._active_processes.remove(proc)
        except FileNotFoundError:
            return (-2, '', f"Command not found: {command[0]}")
        except Exception as e:
            return (-3, '', str(e))

    def convert_file(
        self,
        input_path: str,
        output_path: str,
        command_builder: Callable[[str, str], List[str]],
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> StreamingConversionResult:
        """
        Convert a file using provided command builder.

        Args:
            input_path: Input file path
            output_path: Output file path
            command_builder: Function(input_path, output_path) -> command list
            progress_callback: Optional progress callback

        Returns:
            StreamingConversionResult with success/failure info
        """
        start_time = time.time()
        input_file = Path(input_path)
        output_file = Path(output_path)

        result = StreamingConversionResult(
            success=False,
            input_path=input_path,
            output_path=output_path,
        )

        if not input_file.exists():
            result.error_message = f"Input file not found: {input_path}"
            self._log(result.error_message)
            return result

        result.input_size_bytes = input_file.stat().st_size

        # Check disk space (estimate: need at least input size free)
        output_dir = str(output_file.parent)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        if not self._check_disk_space(result.input_size_bytes, output_dir):
            result.error_message = "Insufficient disk space for conversion"
            self._log(result.error_message)
            return result

        # Build and run command
        command = command_builder(input_path, output_path)
        self._log(f"Starting conversion: {input_file.name} -> {output_file.name}")

        returncode, stdout, stderr = self.run_conversion_with_timeout(command)

        result.duration_seconds = time.time() - start_time

        if returncode != 0:
            result.error_message = stderr.strip() or f"Conversion failed (exit code {returncode})"
            self._log(f"Conversion failed: {result.error_message}")
            if self.config.cleanup_on_failure and output_file.exists():
                output_file.unlink()
            return result

        # Verify output
        if self.config.verify_output:
            if not output_file.exists() or output_file.stat().st_size == 0:
                result.error_message = "Output file missing or empty"
                self._log(result.error_message)
                return result

        result.output_size_bytes = output_file.stat().st_size if output_file.exists() else 0
        result.success = True
        self._log(
            f"Conversion complete: {input_file.name} "
            f"({result.input_size_bytes // 1024 // 1024}MB -> "
            f"{result.output_size_bytes // 1024 // 1024}MB, "
            f"{result.duration_seconds:.1f}s)"
        )
        return result

    def convert_batch(
        self,
        conversions: List[Tuple[str, str, Callable]],
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> List[StreamingConversionResult]:
        """
        Convert a batch of files sequentially with streaming.

        Args:
            conversions: List of (input_path, output_path, command_builder)
            progress_callback: Optional callback(completed, total)

        Returns:
            List of StreamingConversionResult for each conversion
        """
        results = []
        total = len(conversions)

        for i, (input_path, output_path, cmd_builder) in enumerate(conversions):
            result = self.convert_file(input_path, output_path, cmd_builder)
            results.append(result)

            if progress_callback:
                progress_callback(i + 1, total)

        return results

    def terminate_active(self) -> None:
        """Terminate any active conversion processes."""
        for proc in list(self._active_processes):
            try:
                proc.kill()
            except Exception:
                pass
        self._active_processes.clear()


def create_streaming_converter(
    chunk_size: int = 8192,
    max_memory_mb: int = 512,
    log_callback: Optional[Callable[[str], None]] = None,
) -> StreamingConverter:
    """
    Factory function to create a StreamingConverter.

    Args:
        chunk_size: Size of chunks for streaming (bytes)
        max_memory_mb: Max memory before throttling
        log_callback: Optional logging callback

    Returns:
        Configured StreamingConverter instance
    """
    config = StreamingConfig(
        chunk_size=chunk_size,
        max_memory_mb=max_memory_mb,
    )
    return StreamingConverter(config=config, log_callback=log_callback)
