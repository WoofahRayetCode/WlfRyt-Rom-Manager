"""
Retry Logic Integration for ROM Converter - Makes risky operations resilient to transient failures.

This module provides decorators and wrappers that integrate retry_logic.py into rom_converter.py
methods without requiring rewriting of the entire application.

Architecture:
    Risky operations (downloads, extractions, conversions)
         ↓
    Retry wrapper (retry_operation decorator or context manager)
         ↓
    retry_logic.py (exponential backoff, error categorization)
         ↓
    Success or permanent failure (logged to structured logger)

Usage:
    @retry_operation(max_attempts=3, base_delay=1.0)
    def download_file(url):
        return urllib.request.urlopen(url).read()
    
    OR:
    
    for attempt in with_retry_backoff(config):
        with attempt:
            download_file(url)  # auto-retries on transient errors
"""

import logging
import functools
from typing import Callable, TypeVar, Any, Optional
from pathlib import Path

from retry_logic import (
    RetryConfig,
    RetryAttempt,
    retry_with_backoff,
    is_retryable_error,
)
from logging_setup import get_logger

# Type variable for decorator typing
F = TypeVar('F', bound=Callable[..., Any])

# Module logger
logger = get_logger("rom_converter.retry")


class RetryableOperation:
    """Context manager for retryable operations in rom_converter methods.
    
    Provides a simple interface for wrapping risky operations with automatic
    exponential backoff retry logic.
    
    Example:
        with RetryableOperation("download_tool", max_attempts=3) as op:
            response = urllib.request.urlopen(url)
            op.mark_success()
    """
    
    def __init__(
        self,
        operation_name: str,
        max_attempts: int = 3,
        base_delay: float = 1.0,
        parent_logger: Optional[logging.Logger] = None,
    ):
        """Initialize retryable operation wrapper.
        
        Args:
            operation_name: Name of operation for logging (e.g., "download_chdman")
            max_attempts: Maximum retry attempts (default 3)
            base_delay: Initial delay between retries in seconds (default 1.0)
            parent_logger: Logger to use (default: module logger)
        """
        self.operation_name = operation_name
        self.config = RetryConfig(
            max_attempts=max_attempts,
            base_delay=base_delay,
            max_delay=60.0,
        )
        self.logger = parent_logger or logger
        self.current_attempt = 0
        self.last_error = None
        self._retry_iterator = None
    
    def __enter__(self):
        """Start the retry loop."""
        self._retry_iterator = retry_with_backoff(
            self.config,
            self.logger,
            f"[{self.operation_name}]"
        )
        self._attempt_context = next(self._retry_iterator)
        self.current_attempt += 1
        self._attempt_context.__enter__()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Handle retry logic on context exit."""
        try:
            self._attempt_context.__exit__(exc_type, exc_val, exc_tb)
            # If no exception or retryable, will be handled by retry_logic
            return False
        except StopIteration:
            # All retries exhausted
            if exc_type:
                self.logger.error(
                    f"{self.operation_name} failed after {self.current_attempt} attempts: {exc_val}"
                )
            raise
    
    def mark_success(self):
        """Mark operation as successful (stops retry loop)."""
        self.logger.debug(f"{self.operation_name} succeeded on attempt {self.current_attempt}")


def retry_operation(
    max_attempts: int = 3,
    base_delay: float = 1.0,
    operation_name: Optional[str] = None,
):
    """Decorator for retryable rom_converter methods.
    
    Wraps a method with exponential backoff retry logic. Automatically retries
    on transient errors (TimeoutError, OutOfMemoryError, DiskFullError).
    
    Args:
        max_attempts: Maximum retry attempts (default 3)
        base_delay: Initial delay between retries (default 1.0 seconds)
        operation_name: Name for logging (default: function name)
    
    Returns:
        Decorated function with retry logic
    
    Example:
        @retry_operation(max_attempts=3, base_delay=1.0)
        def download_tool(self, url):
            response = urllib.request.urlopen(url)
            return response.read()
    
    Raises:
        Permanent errors after final attempt (DiskFullError, FileNotFoundError)
        Retryable errors only after exhausting max_attempts
    """
    def decorator(func: F) -> F:
        @functools.wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            name = operation_name or func.__name__
            op_logger = logger
            
            # Extract logger if first arg is self (instance method)
            if args and hasattr(args[0], 'logger'):
                op_logger = args[0].logger
            
            config = RetryConfig(
                max_attempts=max_attempts,
                base_delay=base_delay,
                max_delay=60.0,
            )
            
            for attempt in retry_with_backoff(config, op_logger, f"[{name}]"):
                with attempt:
                    return func(*args, **kwargs)
        
        return wrapper
    return decorator


def make_download_retryable(
    func: Callable,
    max_attempts: int = 5,
    base_delay: float = 2.0,
    name: Optional[str] = None,
    logger: Optional[logging.Logger] = None,
) -> Callable:
    """Wrap a download function with retry logic.
    
    Specifically designed for network operations where transient timeouts
    are common and retrying is beneficial.
    
    Args:
        func: Download function to wrap
        max_attempts: Max attempts for downloads (default 5 - more retries for network)
        base_delay: Initial delay (default 2.0 - longer for network backoff)
        name: Optional operation name override for logging
        logger: Optional logger override
    
    Returns:
        Wrapped function with retry logic
    
    Example:
        download_chdman = make_download_retryable(
            original_download_chdman,
            max_attempts=5,
            base_delay=2.0
        )
    """
    config = RetryConfig(
        max_attempts=max_attempts,
        base_delay=base_delay,
        max_delay=120.0,  # Longer max delay for network operations
    )
    
    @functools.wraps(func)
    def wrapper(*args, **kwargs) -> Any:
        op_logger = logger or globals()["logger"]
        
        # Extract logger if available
        if args and hasattr(args[0], 'logger'):
            op_logger = args[0].logger
        
        func_name = name or func.__name__
        for attempt in retry_with_backoff(config, op_logger, f"[{func_name}]"):
            with attempt:
                return func(*args, **kwargs)
    
    return wrapper


def make_extraction_retryable(
    func: Callable,
    max_attempts: int = 3,
    base_delay: float = 1.0,
    name: Optional[str] = None,
    logger: Optional[logging.Logger] = None,
) -> Callable:
    """Wrap an extraction function with retry logic.
    
    Designed for file I/O operations where transient failures can occur
    (disk access issues, temporary locks).
    
    Args:
        func: Extraction function to wrap
        max_attempts: Max attempts (default 3)
        base_delay: Initial delay in seconds (default 1.0)
        name: Optional operation name override for logging
        logger: Optional logger override
    
    Returns:
        Wrapped function with retry logic
    
    Example:
        extract_archive = make_extraction_retryable(
            original_extract_archive,
            max_attempts=3
        )
    """
    config = RetryConfig(
        max_attempts=max_attempts,
        base_delay=base_delay,
        max_delay=60.0,
    )
    
    @functools.wraps(func)
    def wrapper(*args, **kwargs) -> Any:
        op_logger = logger or globals()["logger"]
        
        if args and hasattr(args[0], 'logger'):
            op_logger = args[0].logger
        
        func_name = name or func.__name__
        for attempt in retry_with_backoff(config, op_logger, f"[{func_name}]"):
            with attempt:
                return func(*args, **kwargs)
    
    return wrapper


def make_conversion_retryable(
    func: Callable,
    max_attempts: int = 2,
    base_delay: float = 5.0,
    name: Optional[str] = None,
    logger: Optional[logging.Logger] = None,
) -> Callable:
    """Wrap a conversion function with retry logic.
    
    Designed for long-running tool calls where out-of-memory or disk-full
    errors can occur mid-conversion.
    
    Args:
        func: Conversion function to wrap
        max_attempts: Max attempts (default 2 - conversions expensive)
        base_delay: Initial delay (default 5.0 - allow time for recovery)
        name: Optional operation name override for logging
        logger: Optional logger override
    
    Returns:
        Wrapped function with retry logic
    
    Example:
        convert_to_chd = make_conversion_retryable(
            original_convert_to_chd,
            max_attempts=2,
            base_delay=5.0
        )
    """
    config = RetryConfig(
        max_attempts=max_attempts,
        base_delay=base_delay,
        max_delay=120.0,
    )
    
    @functools.wraps(func)
    def wrapper(*args, **kwargs) -> Any:
        op_logger = logger or globals()["logger"]
        
        if args and hasattr(args[0], 'logger'):
            op_logger = args[0].logger
        
        func_name = name or func.__name__
        for attempt in retry_with_backoff(config, op_logger, f"[{func_name}]"):
            with attempt:
                return func(*args, **kwargs)
    
    return wrapper


# Example usage documentation
__INTEGRATION_EXAMPLES__ = """
# How to use retry_integration.py in rom_converter.py

## Option 1: Decorator approach (least invasive)

```python
from retry_integration import retry_operation, make_download_retryable

class ROMConverter:
    @retry_operation(max_attempts=5, base_delay=2.0, operation_name="download_tools")
    def download_mame_tools(self):
        # Download logic - will auto-retry on transient errors
        response = urllib.request.urlopen(self.MAME_GITHUB_RELEASES_API)
        return json.loads(response.read())
    
    @retry_operation(max_attempts=3)
    def extract_archive(self, archive_path):
        # Extraction logic - will auto-retry on transient I/O errors
        with zipfile.ZipFile(archive_path) as zf:
            zf.extractall(dest_folder)
```

## Option 2: Wrapper approach (more explicit)

```python
from retry_integration import make_download_retryable, make_extraction_retryable

class ROMConverter:
    def __init__(self, master):
        # ... existing init code ...
        
        # Wrap existing methods
        self._original_download = self.download_mame_tools
        self.download_mame_tools = make_download_retryable(
            self._original_download,
            max_attempts=5,
            base_delay=2.0
        )
        
        self._original_extract = self.extract_archive
        self.extract_archive = make_extraction_retryable(
            self._original_extract,
            max_attempts=3
        )
```

## Option 3: Context manager approach (most flexible)

```python
from retry_integration import RetryableOperation

class ROMConverter:
    def download_mame_tools(self):
        with RetryableOperation("download_mame_tools", max_attempts=5) as op:
            response = urllib.request.urlopen(self.MAME_GITHUB_RELEASES_API)
            data = json.loads(response.read())
            op.mark_success()
            return data
```

## Retry Strategies by Operation Type

### Downloads (Network Operations)
- Max attempts: 5 (network is flaky)
- Base delay: 2.0 seconds (give network time)
- Max delay: 120.0 seconds (long backoff for persistent issues)
- Retryable errors: TimeoutError, ConnectionError, urllib.error.URLError

### Extractions (File I/O)
- Max attempts: 3 (file system less flaky than network)
- Base delay: 1.0 seconds
- Max delay: 60.0 seconds
- Retryable errors: IOError, OSError (some), TimeoutError

### Conversions (Tool Calls)
- Max attempts: 2 (conversions are expensive, don't retry too much)
- Base delay: 5.0 seconds (allow time for cleanup)
- Max delay: 120.0 seconds
- Retryable errors: OutOfMemoryError, DiskFullError, TimeoutError
- Non-retryable: FileNotFoundError, PermissionError
"""
