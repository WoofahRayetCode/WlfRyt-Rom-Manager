"""
Conversion Error Handler Module

Robust error recovery for ROM conversion failures including:
- Circuit breaker pattern (stop hammering failing tools)
- Failure tracking and categorization
- Error reporting and summaries
- Integration with retry_logic.py

Phase 4 Week 4: Streaming & Error Handling
"""

import time
import logging
from collections import defaultdict
from dataclasses import dataclass, field
from enum import Enum
from typing import Optional, Callable, Dict, List

try:
    from exceptions import (
        ROMConverterError,
        ConversionError,
        TimeoutError as ConverterTimeoutError,
        OutOfMemoryError,
        DiskFullError,
        is_transient_error,
    )
    EXCEPTIONS_AVAILABLE = True
except ImportError:
    EXCEPTIONS_AVAILABLE = False
    ROMConverterError = Exception
    ConversionError = Exception

try:
    from retry_logic import RetryConfig, retry_with_backoff
    RETRY_AVAILABLE = True
except ImportError:
    RETRY_AVAILABLE = False


logger = logging.getLogger(__name__)


class FailureCategory(Enum):
    """Categories of conversion failures."""
    TRANSIENT = "transient"       # Temporary errors (timeout, OOM) - retry
    TOOL_ERROR = "tool_error"     # Tool returned non-zero exit
    DISK_FULL = "disk_full"       # Not enough disk space
    FILE_CORRUPT = "file_corrupt" # Input file is damaged
    PERMISSION = "permission"     # Access denied
    NOT_FOUND = "not_found"       # Tool or file missing
    UNKNOWN = "unknown"           # Unclassified error


@dataclass
class ConversionFailure:
    """Record of a single conversion failure."""
    file_path: str
    error_message: str
    category: FailureCategory
    timestamp: float = field(default_factory=time.time)
    retry_count: int = 0
    is_retryable: bool = False

    def age_seconds(self) -> float:
        """Seconds since this failure occurred."""
        return time.time() - self.timestamp


@dataclass
class CircuitBreakerState:
    """State for circuit breaker pattern."""
    failure_count: int = 0
    last_failure_time: float = 0.0
    is_open: bool = False           # True = circuit tripped, stop retrying
    reset_timeout: float = 300.0    # Seconds before resetting (5 min default)
    failure_threshold: int = 5      # Failures before opening circuit

    def record_failure(self) -> None:
        """Record a failure and potentially open circuit."""
        self.failure_count += 1
        self.last_failure_time = time.time()
        if self.failure_count >= self.failure_threshold:
            self.is_open = True

    def record_success(self) -> None:
        """Record a success and close circuit."""
        self.failure_count = 0
        self.is_open = False

    def should_attempt(self) -> bool:
        """Check if an attempt should be made (circuit closed or reset timeout)."""
        if not self.is_open:
            return True
        # Auto-reset after timeout
        if time.time() - self.last_failure_time > self.reset_timeout:
            self.is_open = False
            self.failure_count = 0
            return True
        return False

    def time_until_reset(self) -> float:
        """Seconds until circuit resets (-1 if closed)."""
        if not self.is_open:
            return -1.0
        remaining = self.reset_timeout - (time.time() - self.last_failure_time)
        return max(0.0, remaining)


class ConversionErrorHandler:
    """
    Central error handler for ROM conversions.

    Features:
    - Categorizes errors by type
    - Tracks failure counts per file and per tool
    - Circuit breaker to stop retrying consistently failing tools
    - Retry logic with exponential backoff
    - Failure summary and reporting
    """

    def __init__(
        self,
        max_retries: int = 3,
        base_retry_delay: float = 1.0,
        circuit_breaker_threshold: int = 5,
        log_callback: Optional[Callable[[str], None]] = None,
    ):
        """
        Initialize error handler.

        Args:
            max_retries: Max retry attempts per file
            base_retry_delay: Initial delay between retries (seconds)
            circuit_breaker_threshold: Failures before opening circuit breaker
            log_callback: Optional logging callback
        """
        self.max_retries = max_retries
        self.base_retry_delay = base_retry_delay
        self.circuit_breaker_threshold = circuit_breaker_threshold
        self.log_callback = log_callback

        # Failure tracking
        self.failures: List[ConversionFailure] = []
        self.file_retry_counts: Dict[str, int] = defaultdict(int)

        # Circuit breakers per tool name
        self.circuit_breakers: Dict[str, CircuitBreakerState] = {}

        # Retry config (uses retry_logic if available)
        if RETRY_AVAILABLE:
            self.retry_config = RetryConfig(
                max_attempts=max_retries,
                base_delay=base_retry_delay,
                max_delay=30.0,
            )
        else:
            self.retry_config = None

    def _log(self, message: str) -> None:
        """Log via callback or logger."""
        if self.log_callback:
            self.log_callback(message)
        logger.info(message)

    def categorize_error(self, error: Exception) -> FailureCategory:
        """
        Categorize an exception into a FailureCategory.

        Args:
            error: Exception to categorize

        Returns:
            Appropriate FailureCategory
        """
        error_str = str(error).lower()

        # Check custom exception types
        if EXCEPTIONS_AVAILABLE:
            if isinstance(error, (ConverterTimeoutError,)):
                return FailureCategory.TRANSIENT
            if isinstance(error, OutOfMemoryError):
                return FailureCategory.TRANSIENT
            if isinstance(error, DiskFullError):
                return FailureCategory.DISK_FULL

        # String-based classification for generic exceptions
        if any(kw in error_str for kw in ['timeout', 'timed out']):
            return FailureCategory.TRANSIENT
        if any(kw in error_str for kw in ['out of memory', 'memory error', 'oom']):
            return FailureCategory.TRANSIENT
        if any(kw in error_str for kw in ['disk full', 'no space', 'enospc']):
            return FailureCategory.DISK_FULL
        if any(kw in error_str for kw in ['corrupt', 'invalid', 'bad header']):
            return FailureCategory.FILE_CORRUPT
        if any(kw in error_str for kw in ['permission', 'access denied', 'eacces']):
            return FailureCategory.PERMISSION
        if any(kw in error_str for kw in ['not found', 'no such file', 'enoent']):
            return FailureCategory.NOT_FOUND
        if any(kw in error_str for kw in ['exit code', 'returncode', 'non-zero']):
            return FailureCategory.TOOL_ERROR

        return FailureCategory.UNKNOWN

    def is_retryable(self, category: FailureCategory) -> bool:
        """
        Check if a failure category is worth retrying.

        Args:
            category: FailureCategory to check

        Returns:
            True if retry is appropriate
        """
        return category in (FailureCategory.TRANSIENT,)

    def record_failure(
        self,
        file_path: str,
        error: Exception,
        tool_name: Optional[str] = None,
    ) -> ConversionFailure:
        """
        Record a conversion failure.

        Args:
            file_path: Path of file that failed
            error: Exception that occurred
            tool_name: Name of conversion tool (for circuit breaker)

        Returns:
            ConversionFailure record
        """
        category = self.categorize_error(error)
        retryable = self.is_retryable(category)
        retry_count = self.file_retry_counts[file_path]

        failure = ConversionFailure(
            file_path=file_path,
            error_message=str(error),
            category=category,
            retry_count=retry_count,
            is_retryable=retryable,
        )
        self.failures.append(failure)

        # Update circuit breaker for tool
        if tool_name:
            if tool_name not in self.circuit_breakers:
                self.circuit_breakers[tool_name] = CircuitBreakerState(
                    failure_threshold=self.circuit_breaker_threshold,
                )
            self.circuit_breakers[tool_name].record_failure()
            if self.circuit_breakers[tool_name].is_open:
                self._log(
                    f"⚠️  Circuit breaker OPEN for {tool_name} "
                    f"after {self.circuit_breaker_threshold} failures. "
                    f"Pausing for {self.circuit_breakers[tool_name].reset_timeout:.0f}s"
                )

        self._log(
            f"❌ Conversion failure [{category.value}]: {Path(file_path).name} "
            f"- {str(error)[:100]}"
        )
        return failure

    def record_success(self, file_path: str, tool_name: Optional[str] = None) -> None:
        """
        Record a successful conversion, closing circuit breaker if needed.

        Args:
            file_path: Path of file that succeeded
            tool_name: Name of conversion tool
        """
        if file_path in self.file_retry_counts:
            del self.file_retry_counts[file_path]

        if tool_name and tool_name in self.circuit_breakers:
            was_open = self.circuit_breakers[tool_name].is_open
            self.circuit_breakers[tool_name].record_success()
            if was_open:
                self._log(f"✅ Circuit breaker CLOSED for {tool_name} after success")

    def should_retry(self, file_path: str, tool_name: Optional[str] = None) -> bool:
        """
        Check if a file should be retried.

        Args:
            file_path: File path to check
            tool_name: Tool name for circuit breaker check

        Returns:
            True if retry is appropriate
        """
        retry_count = self.file_retry_counts[file_path]
        if retry_count >= self.max_retries:
            return False

        # Check circuit breaker
        if tool_name and tool_name in self.circuit_breakers:
            if not self.circuit_breakers[tool_name].should_attempt():
                remaining = self.circuit_breakers[tool_name].time_until_reset()
                self._log(
                    f"⛔ Circuit breaker preventing retry for {tool_name} "
                    f"(resets in {remaining:.0f}s)"
                )
                return False

        return True

    def increment_retry(self, file_path: str) -> int:
        """
        Increment retry count for a file.

        Args:
            file_path: File path

        Returns:
            New retry count
        """
        self.file_retry_counts[file_path] += 1
        return self.file_retry_counts[file_path]

    def get_retry_delay(self, retry_count: int) -> float:
        """
        Get delay before next retry using exponential backoff.

        Args:
            retry_count: Current retry count (1-based)

        Returns:
            Delay in seconds
        """
        if self.retry_config:
            return self.retry_config.calculate_delay(retry_count)
        # Fallback exponential backoff
        delay = self.base_retry_delay * (2 ** (retry_count - 1))
        return min(delay, 30.0)

    def get_failure_summary(self) -> str:
        """
        Generate a human-readable summary of all failures.

        Returns:
            Multi-line summary string
        """
        if not self.failures:
            return "✅ No conversion failures recorded"

        lines = [f"⚠️  Conversion Failure Summary ({len(self.failures)} total failures)"]
        lines.append("=" * 50)

        # Group by category
        by_category: Dict[FailureCategory, List[ConversionFailure]] = defaultdict(list)
        for f in self.failures:
            by_category[f.category].append(f)

        for category, cat_failures in sorted(by_category.items(), key=lambda x: len(x[1]), reverse=True):
            lines.append(f"\n{category.value.upper()} ({len(cat_failures)} failures):")
            for failure in cat_failures[:5]:  # Show at most 5 per category
                name = Path(failure.file_path).name
                lines.append(f"  • {name}: {failure.error_message[:80]}")
            if len(cat_failures) > 5:
                lines.append(f"  ... and {len(cat_failures) - 5} more")

        # Circuit breaker status
        open_breakers = [
            (name, cb) for name, cb in self.circuit_breakers.items() if cb.is_open
        ]
        if open_breakers:
            lines.append("\n⛔ OPEN CIRCUIT BREAKERS (tools paused):")
            for name, cb in open_breakers:
                lines.append(f"  • {name}: resets in {cb.time_until_reset():.0f}s")

        return "\n".join(lines)

    def get_stats(self) -> Dict:
        """
        Get failure statistics as a dictionary.

        Returns:
            Dict with failure counts, categories, etc.
        """
        by_category = defaultdict(int)
        for f in self.failures:
            by_category[f.category.value] += 1

        return {
            "total_failures": len(self.failures),
            "by_category": dict(by_category),
            "files_with_retries": len(self.file_retry_counts),
            "open_circuit_breakers": [
                name for name, cb in self.circuit_breakers.items() if cb.is_open
            ],
        }

    def clear(self) -> None:
        """Clear all failure records and reset circuit breakers."""
        self.failures.clear()
        self.file_retry_counts.clear()
        for cb in self.circuit_breakers.values():
            cb.record_success()


# Import Path for use in record_failure
from pathlib import Path


def create_error_handler(
    max_retries: int = 3,
    circuit_breaker_threshold: int = 5,
    log_callback: Optional[Callable[[str], None]] = None,
) -> ConversionErrorHandler:
    """
    Factory function to create a ConversionErrorHandler.

    Args:
        max_retries: Maximum retries per file
        circuit_breaker_threshold: Failures before opening circuit
        log_callback: Optional logging callback

    Returns:
        Configured ConversionErrorHandler
    """
    return ConversionErrorHandler(
        max_retries=max_retries,
        circuit_breaker_threshold=circuit_breaker_threshold,
        log_callback=log_callback,
    )
