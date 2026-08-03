"""
Unit tests for retry_logic module.

Tests exponential backoff retry mechanism, error categorization, and retry logic.
"""

import pytest
import time
import logging
from retry_logic import (
    RetryConfig,
    RetryAttempt,
    RetryIterator,
    retry_with_backoff,
    with_retry,
    is_retryable_error,
)
from exceptions import (
    TimeoutError as ConverterTimeoutError,
    OutOfMemoryError,
    DiskFullError,
    ConversionError,
)


@pytest.mark.unit
class TestRetryConfig:
    """Test RetryConfig configuration."""
    
    def test_default_config(self):
        """Test RetryConfig with default values."""
        config = RetryConfig()
        
        assert config.max_attempts == 3
        assert config.base_delay == 1.0
        assert config.max_delay == 60.0
        assert config.exponential_base == 2.0
    
    def test_custom_config(self):
        """Test RetryConfig with custom values."""
        config = RetryConfig(
            max_attempts=5,
            base_delay=0.5,
            max_delay=30.0,
            exponential_base=1.5
        )
        
        assert config.max_attempts == 5
        assert config.base_delay == 0.5
        assert config.max_delay == 30.0
        assert config.exponential_base == 1.5
    
    def test_calculate_delay_first_attempt(self):
        """Test delay calculation for first attempt."""
        config = RetryConfig(base_delay=1.0)
        
        # First attempt should have no delay
        assert config.calculate_delay(0) == 0
    
    def test_calculate_delay_exponential(self):
        """Test exponential backoff calculation."""
        config = RetryConfig(base_delay=1.0, exponential_base=2.0)
        
        # Delays should double each attempt
        assert config.calculate_delay(1) == 1.0  # 1 * 2^0
        assert config.calculate_delay(2) == 2.0  # 1 * 2^1
        assert config.calculate_delay(3) == 4.0  # 1 * 2^2
        assert config.calculate_delay(4) == 8.0  # 1 * 2^3
    
    def test_calculate_delay_max_limit(self):
        """Test that delay respects max_delay."""
        config = RetryConfig(base_delay=10.0, max_delay=20.0, exponential_base=2.0)
        
        # Delay should not exceed max_delay
        assert config.calculate_delay(3) == 20.0  # Would be 80, but capped at 20


@pytest.mark.unit
class TestRetryAttempt:
    """Test RetryAttempt context manager."""
    
    def test_successful_attempt(self):
        """Test successful attempt doesn't trigger retry."""
        config = RetryConfig()
        attempt = RetryAttempt(1, config)
        
        with attempt:
            pass  # No exception
        
        assert attempt.success is True
        assert attempt.exception is None
    
    def test_transient_error_retry(self):
        """Test transient error triggers retry."""
        config = RetryConfig()
        attempt = RetryAttempt(1, config)
        
        exception_raised = False
        try:
            with attempt:
                raise ConverterTimeoutError("Timeout", timeout_seconds=30)
        except ConverterTimeoutError:
            exception_raised = True
        
        # Exception should be suppressed (retry will happen)
        assert not exception_raised
        assert attempt.exception is not None
        assert isinstance(attempt.exception, ConverterTimeoutError)
    
    def test_permanent_error_no_retry(self):
        """Test permanent error doesn't trigger retry."""
        config = RetryConfig()
        attempt = RetryAttempt(1, config)
        
        exception_raised = False
        try:
            with attempt:
                raise ValueError("Permanent error")
        except ValueError:
            exception_raised = True
        
        # Exception should propagate (no retry)
        assert exception_raised
    
    def test_max_attempts_reached(self):
        """Test that max attempts stops retry."""
        config = RetryConfig(max_attempts=3)
        attempt = RetryAttempt(3, config)  # Last attempt
        
        exception_raised = False
        try:
            with attempt:
                raise OutOfMemoryError(required_mb=2000, available_mb=1000)
        except OutOfMemoryError:
            exception_raised = True
        
        # Exception should propagate on last attempt
        assert exception_raised
    
    def test_context_string_logging(self, caplog):
        """Test context string in logging."""
        config = RetryConfig()
        logger = logging.getLogger("test_retry")
        attempt = RetryAttempt(1, config, logger, "[game.iso]")
        
        with caplog.at_level(logging.WARNING):
            try:
                with attempt:
                    raise DiskFullError(required_gb=50.0, available_gb=10.0)
            except DiskFullError:
                pass
        
        assert "[game.iso]" in caplog.text


@pytest.mark.unit
class TestRetryIterator:
    """Test RetryIterator for-loop usage."""
    
    def test_iterator_stops_at_max_attempts(self):
        """Test iterator stops at max attempts."""
        config = RetryConfig(max_attempts=3)
        iterator = retry_with_backoff(config)
        
        attempts = list(iterator)
        assert len(attempts) == 3
    
    def test_iterator_returns_retry_attempts(self):
        """Test iterator returns RetryAttempt objects."""
        config = RetryConfig(max_attempts=2)
        iterator = retry_with_backoff(config)
        
        for attempt in iterator:
            assert isinstance(attempt, RetryAttempt)
            break  # Just check first one


@pytest.mark.unit
class TestRetryWithBackoff:
    """Test retry_with_backoff function."""
    
    def test_successful_on_first_attempt(self):
        """Test function succeeds on first attempt."""
        attempt_count = 0
        
        for attempt in retry_with_backoff():
            with attempt:
                attempt_count += 1
                pass  # Success
            if attempt.success:
                break  # Exit on success
        
        assert attempt_count == 1
    
    def test_retry_on_transient_error(self):
        """Test retry on transient error."""
        attempt_count = 0
        
        for attempt in retry_with_backoff():
            with attempt:
                attempt_count += 1
                if attempt_count < 2:
                    raise ConverterTimeoutError("Timeout", timeout_seconds=30)
            if attempt.success:
                break  # Exit on success
        
        assert attempt_count == 2
    
    def test_stop_on_permanent_error(self):
        """Test stop on permanent error."""
        attempt_count = 0
        
        exception_raised = False
        try:
            for attempt in retry_with_backoff():
                with attempt:
                    attempt_count += 1
                    raise ValueError("Permanent")
        except ValueError:
            exception_raised = True
        
        assert exception_raised
        assert attempt_count == 1
    
    def test_max_attempts_respected(self):
        """Test max attempts limit is respected."""
        config = RetryConfig(max_attempts=3)
        attempt_count = 0
        
        exception_raised = False
        try:
            for attempt in retry_with_backoff(config):
                with attempt:
                    attempt_count += 1
                    raise OutOfMemoryError(required_mb=1000, available_mb=500)
        except OutOfMemoryError:
            exception_raised = True
        
        assert exception_raised
        assert attempt_count == 3


@pytest.mark.unit
class TestWithRetry:
    """Test with_retry function wrapper."""
    
    def test_with_retry_success(self):
        """Test with_retry with successful function."""
        def add(a, b):
            return a + b
        
        result = with_retry(add, 2, 3)
        assert result == 5
    
    def test_with_retry_with_kwargs(self):
        """Test with_retry with keyword arguments."""
        def greet(name, greeting="Hello"):
            return f"{greeting}, {name}"
        
        result = with_retry(greet, "World", greeting="Hi")
        assert result == "Hi, World"
    
    def test_with_retry_succeeds_on_retry(self):
        """Test with_retry retries transient errors."""
        call_count = [0]
        
        def risky_operation():
            call_count[0] += 1
            if call_count[0] < 2:
                raise ConverterTimeoutError("Timeout", timeout_seconds=30)
            return "success"
        
        result = with_retry(risky_operation)
        assert result == "success"
        assert call_count[0] == 2
    
    def test_with_retry_fails_on_permanent_error(self):
        """Test with_retry fails on permanent error."""
        def broken():
            raise ValueError("Broken")
        
        with pytest.raises(ValueError):
            with_retry(broken)


@pytest.mark.unit
class TestIsRetryableError:
    """Test is_retryable_error function."""
    
    def test_timeout_is_retryable(self):
        """Test timeout error is retryable."""
        error = ConverterTimeoutError("Timeout", timeout_seconds=30)
        assert is_retryable_error(error) is True
    
    def test_oom_is_retryable(self):
        """Test OOM error is retryable."""
        error = OutOfMemoryError(required_mb=1000, available_mb=500)
        assert is_retryable_error(error) is True
    
    def test_disk_full_is_retryable(self):
        """Test disk full error is retryable."""
        error = DiskFullError(required_gb=50.0, available_gb=10.0)
        assert is_retryable_error(error) is True
    
    def test_value_error_not_retryable(self):
        """Test ValueError is not retryable."""
        error = ValueError("Bad value")
        assert is_retryable_error(error) is False
    
    def test_conversion_error_retryable_flag(self):
        """Test ConversionError with retry_possible flag."""
        error = ConversionError("Failed", returncode=1, retry_possible=True)
        assert is_retryable_error(error) is True


@pytest.mark.unit
@pytest.mark.slow
def test_exponential_backoff_timing():
    """Test that exponential backoff timing works correctly."""
    config = RetryConfig(base_delay=0.01, exponential_base=2.0, max_delay=1.0)
    
    start_time = time.time()
    attempt_count = 0
    
    for attempt in retry_with_backoff(config):
        with attempt:
            attempt_count += 1
            if attempt_count < 3:
                raise ConverterTimeoutError("Timeout", timeout_seconds=30)
    
    elapsed = time.time() - start_time
    
    # Should have taken at least 0.01 + 0.02 = 0.03 seconds
    assert elapsed >= 0.03
    assert attempt_count == 3
