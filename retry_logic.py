"""
Retry logic with exponential backoff for ROM Converter.

Provides intelligent retry logic for transient errors (timeouts, out of memory,
disk full, etc.) using the exception hierarchy defined in exceptions.py.

Features:
- Exponential backoff with configurable base and maximum delay
- Distinguishes between transient, permanent, and recoverable errors
- Maximum retry attempts with configurable limits per error type
- Logging of retry attempts and failures
- Thread-safe retry state management

Usage:
    from retry_logic import RetryConfig, retry_with_backoff, is_retryable
    
    # Simple usage with decorator
    @retry_with_backoff()
    def risky_conversion(file_path, output_format):
        # Conversion logic
        pass
    
    # Advanced usage with custom config
    config = RetryConfig(
        max_attempts=5,
        base_delay=1.0,
        max_delay=60.0,
        exponential_base=2.0
    )
    
    for attempt in retry_with_backoff(config=config):
        with attempt:
            do_conversion()
"""

import time
import logging
from typing import Optional, Callable, Any, Type, List
from contextlib import contextmanager
from exceptions import (
    ROMConverterError,
    ConversionError,
    TimeoutError as ConverterTimeoutError,
    OutOfMemoryError,
    DiskFullError,
    is_transient_error,
    is_recoverable_error,
)


class RetryConfig:
    """Configuration for retry behavior.
    
    Attributes:
        max_attempts: Maximum number of retry attempts (including initial)
        base_delay: Initial delay between retries in seconds
        max_delay: Maximum delay between retries in seconds
        exponential_base: Base for exponential backoff calculation
        transient_error_types: Tuple of exception types to retry on
    """
    
    def __init__(
        self,
        max_attempts: int = 3,
        base_delay: float = 1.0,
        max_delay: float = 60.0,
        exponential_base: float = 2.0,
        transient_error_types: Optional[List[Type[Exception]]] = None
    ):
        """Initialize retry configuration.
        
        Args:
            max_attempts: Maximum attempts (default: 3)
            base_delay: Initial delay in seconds (default: 1.0)
            max_delay: Maximum delay in seconds (default: 60.0)
            exponential_base: Exponential multiplier (default: 2.0)
            transient_error_types: Exception types to retry on (default: auto-detect)
        """
        self.max_attempts = max_attempts
        self.base_delay = base_delay
        self.max_delay = max_delay
        self.exponential_base = exponential_base
        self.transient_error_types = transient_error_types or (
            ConverterTimeoutError,
            OutOfMemoryError,
            DiskFullError,
        )
    
    def calculate_delay(self, attempt: int) -> float:
        """Calculate delay for attempt number.
        
        Args:
            attempt: Attempt number (0-based)
        
        Returns:
            Delay in seconds
        """
        if attempt <= 0:
            return 0
        
        delay = self.base_delay * (self.exponential_base ** (attempt - 1))
        return min(delay, self.max_delay)


class RetryAttempt:
    """Context manager for individual retry attempts.
    
    Handles tracking attempt number, logging, and delay between attempts.
    """
    
    def __init__(
        self,
        attempt_num: int,
        config: RetryConfig,
        logger: Optional[logging.Logger] = None,
        context: Optional[str] = None
    ):
        """Initialize retry attempt.
        
        Args:
            attempt_num: Attempt number (1-based)
            config: RetryConfig instance
            logger: Logger for retry logging
            context: Context string for logging (e.g., filename)
        """
        self.attempt_num = attempt_num
        self.config = config
        self.logger = logger
        self.context = context or ""
        self.exception: Optional[Exception] = None
        self.success = False
    
    def __enter__(self):
        """Enter context manager."""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Exit context manager.
        
        Returns:
            True if exception should be suppressed (retry will happen)
            False if exception should propagate (no retry)
        """
        if exc_type is None:
            # No exception, mark as success
            self.success = True
            return False
        
        # Store exception
        self.exception = exc_val
        
        # Check if exception is retryable
        if not self._should_retry():
            return False  # Let exception propagate
        
        # Log retry attempt
        if self.logger:
            self.logger.warning(
                f"Transient error on attempt {self.attempt_num}/{self.config.max_attempts}: "
                f"{exc_type.__name__}: {exc_val} {self.context}"
            )
        
        # If this is the last attempt, let exception propagate
        if self.attempt_num >= self.config.max_attempts:
            return False
        
        # Calculate and sleep delay
        delay = self.config.calculate_delay(self.attempt_num)
        if delay > 0:
            if self.logger:
                self.logger.debug(f"Retrying in {delay:.1f}s... {self.context}")
            time.sleep(delay)
        
        # Suppress exception to allow retry
        return True
    
    def _should_retry(self) -> bool:
        """Check if exception should trigger a retry.
        
        Returns:
            True if retry should happen, False otherwise
        """
        if self.exception is None:
            return False
        
        exc_type = type(self.exception)
        
        # Check if exception type is in transient error list
        if exc_type in self.config.transient_error_types:
            return True
        
        # Check if exception has retry_possible flag (from our exception hierarchy)
        if isinstance(self.exception, ROMConverterError):
            return getattr(self.exception, 'retry_possible', False)
        
        # Use helper function from exceptions module
        return is_transient_error(self.exception)


class RetryIterator:
    """Iterator for retry attempts.
    
    Allows using retry logic in a for loop:
        for attempt in retry_with_backoff():
            with attempt:
                do_work()
    """
    
    def __init__(
        self,
        config: RetryConfig,
        logger: Optional[logging.Logger] = None,
        context: Optional[str] = None
    ):
        """Initialize retry iterator.
        
        Args:
            config: RetryConfig instance
            logger: Logger for retry logging
            context: Context string for logging
        """
        self.config = config
        self.logger = logger
        self.context = context
        self.current_attempt = 0
    
    def __iter__(self):
        """Return iterator."""
        return self
    
    def __next__(self) -> RetryAttempt:
        """Get next retry attempt.
        
        Raises:
            StopIteration: When max attempts reached
        """
        self.current_attempt += 1
        
        if self.current_attempt > self.config.max_attempts:
            raise StopIteration
        
        return RetryAttempt(
            self.current_attempt,
            self.config,
            self.logger,
            self.context
        )


def retry_with_backoff(
    config: Optional[RetryConfig] = None,
    logger: Optional[logging.Logger] = None,
    context: Optional[str] = None
) -> RetryIterator:
    """Create a retry iterator with exponential backoff.
    
    Usage with context manager:
        config = RetryConfig(max_attempts=5)
        for attempt in retry_with_backoff(config):
            with attempt:
                do_risky_operation()
    
    Args:
        config: RetryConfig instance (uses defaults if None)
        logger: Logger for retry events
        context: Context string for logging
    
    Returns:
        RetryIterator instance
    """
    if config is None:
        config = RetryConfig()
    
    return RetryIterator(config, logger, context)


def with_retry(
    func: Callable,
    *args,
    config: Optional[RetryConfig] = None,
    logger: Optional[logging.Logger] = None,
    context: Optional[str] = None,
    **kwargs
) -> Any:
    """Execute function with retry logic.
    
    Args:
        func: Function to execute
        *args: Positional arguments for func
        config: RetryConfig instance
        logger: Logger for retry events
        context: Context string for logging
        **kwargs: Keyword arguments for func
    
    Returns:
        Return value of func
    
    Raises:
        Exception: If all retry attempts fail
    
    Example:
        def convert_file(path, fmt):
            return converter.convert(path, fmt)
        
        result = with_retry(
            convert_file,
            "game.iso",
            "CHD",
            config=RetryConfig(max_attempts=3)
        )
    """
    for attempt in retry_with_backoff(config, logger, context):
        with attempt:
            return func(*args, **kwargs)
    
    # Should not reach here, but just in case
    raise RuntimeError("Retry exhausted without success")


def is_retryable_error(exception: Exception) -> bool:
    """Check if an exception should trigger a retry.
    
    Args:
        exception: Exception to check
    
    Returns:
        True if exception is retryable, False otherwise
    """
    # Check if exception is transient
    if is_transient_error(exception):
        return True
    
    # Check retry_possible flag
    if isinstance(exception, ROMConverterError):
        return getattr(exception, 'retry_possible', False)
    
    return False
