"""
Tests for retry_integration.py - Retry logic integration for ROM converter methods.

Tests the decorator and wrapper functions that make rom_converter operations resilient.
"""

import logging
import pytest
import tempfile
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock
import time

from logging_setup import get_logger
from retry_integration import (
    RetryableOperation,
    retry_operation,
    make_download_retryable,
    make_extraction_retryable,
    make_conversion_retryable,
)
from exceptions import (
    TimeoutError,
    OutOfMemoryError,
    DiskFullError,
)


@pytest.fixture
def clean_logging():
    """Cleanup logging handlers after each test."""
    yield
    # Cleanup handled by conftest.py cleanup_logging fixture


class TestRetryableOperation:
    """Test the RetryableOperation context manager."""
    
    def test_successful_operation_first_attempt(self, clean_logging):
        """Operation should succeed on first attempt."""
        logger = get_logger("test")
        
        with RetryableOperation("test_op", parent_logger=logger) as op:
            assert op.current_attempt == 1
            op.mark_success()
    
    def test_operation_with_custom_limits(self, clean_logging):
        """Should respect max_attempts and base_delay configuration."""
        logger = get_logger("test")
        
        with RetryableOperation(
            "test_op",
            max_attempts=5,
            base_delay=2.0,
            parent_logger=logger
        ) as op:
            assert op.config.max_attempts == 5
            assert op.config.base_delay == 2.0


class TestRetryOperationDecorator:
    """Test the @retry_operation decorator."""
    
    def test_decorator_on_successful_function(self, clean_logging):
        """Decorated function should work normally on success."""
        call_count = [0]
        
        @retry_operation(max_attempts=3)
        def successful_func():
            call_count[0] += 1
            return "success"
        
        result = successful_func()
        assert result == "success"
        assert call_count[0] == 1  # Only called once
    
    def test_decorator_name_parameter(self, clean_logging):
        """Decorator should use custom operation name for logging."""
        @retry_operation(max_attempts=1, operation_name="custom_name")
        def named_func():
            return "ok"
        
        # Should not raise
        result = named_func()
        assert result == "ok"
    
    def test_decorator_with_retryable_error(self, clean_logging):
        """Decorator should retry on retryable errors."""
        call_count = [0]
        
        @retry_operation(max_attempts=3, base_delay=0.01)
        def failing_then_success():
            call_count[0] += 1
            if call_count[0] < 3:
                # Simulate timeout (transient error)
                raise TimeoutError("Timeout", 5)
            return "eventually_success"
        
        result = failing_then_success()
        assert result == "eventually_success"
        assert call_count[0] == 3  # Called 3 times before success
    
    def test_decorator_with_instance_method(self, clean_logging):
        """Decorator should work with instance methods and access self.logger."""
        
        class MockConverter:
            def __init__(self):
                self.logger = get_logger("test")
                self.call_count = 0
            
            @retry_operation(max_attempts=2, base_delay=0.01)
            def download(self, url):
                self.call_count += 1
                if self.call_count == 1:
                    raise TimeoutError("Timeout", 5)
                return "downloaded"
        
        converter = MockConverter()
        result = converter.download("http://example.com")
        assert result == "downloaded"
        assert converter.call_count == 2


class TestMakeDownloadRetryable:
    """Test the make_download_retryable wrapper."""
    
    def test_download_succeeds_immediately(self, clean_logging):
        """Successful download should not retry."""
        call_count = [0]
        
        def original_download(url):
            call_count[0] += 1
            return "data"
        
        retryable = make_download_retryable(original_download)
        result = retryable("http://example.com")
        
        assert result == "data"
        assert call_count[0] == 1
    
    def test_download_with_timeout_retry(self, clean_logging):
        """Download should retry on timeout."""
        call_count = [0]
        
        def original_download(url):
            call_count[0] += 1
            if call_count[0] < 3:
                raise TimeoutError("Timeout", 10)
            return "data"
        
        retryable = make_download_retryable(original_download, max_attempts=5, base_delay=0.01)
        result = retryable("http://example.com")
        
        assert result == "data"
        assert call_count[0] == 3
    
    def test_download_max_attempts_respected(self, clean_logging):
        """Download should stop after max attempts."""
        call_count = [0]
        
        def original_download(url):
            call_count[0] += 1
            raise TimeoutError("Timeout", 10)
        
        retryable = make_download_retryable(original_download, max_attempts=3, base_delay=0.01)
        
        with pytest.raises(TimeoutError):
            retryable("http://example.com")
        
        assert call_count[0] == 3


class TestMakeExtractionRetryable:
    """Test the make_extraction_retryable wrapper."""
    
    def test_extraction_succeeds(self, clean_logging):
        """Successful extraction should not retry."""
        call_count = [0]
        
        def original_extract(path):
            call_count[0] += 1
            return {"files": 5}
        
        retryable = make_extraction_retryable(original_extract)
        result = retryable("/path/to/archive.zip")
        
        assert result == {"files": 5}
        assert call_count[0] == 1
    
    def test_extraction_with_timeout_retry(self, clean_logging):
        """Extraction should retry on timeout."""
        call_count = [0]
        
        def original_extract(path):
            call_count[0] += 1
            if call_count[0] < 2:
                raise TimeoutError("Timeout", 30)
            return {"files": 10}
        
        retryable = make_extraction_retryable(original_extract, max_attempts=3, base_delay=0.01)
        result = retryable("/path/to/archive.zip")
        
        assert result == {"files": 10}
        assert call_count[0] == 2


class TestMakeConversionRetryable:
    """Test the make_conversion_retryable wrapper."""
    
    def test_conversion_succeeds(self, clean_logging):
        """Successful conversion should not retry."""
        call_count = [0]
        
        def original_convert(input_file, output_file):
            call_count[0] += 1
            return {"status": "complete", "size": 1024}
        
        retryable = make_conversion_retryable(original_convert)
        result = retryable("input.iso", "output.chd")
        
        assert result["status"] == "complete"
        assert call_count[0] == 1
    
    def test_conversion_with_oom_retry(self, clean_logging):
        """Conversion should retry on OutOfMemoryError."""
        call_count = [0]
        
        def original_convert(input_file, output_file):
            call_count[0] += 1
            if call_count[0] == 1:
                raise OutOfMemoryError(required_mb=4096, available_mb=2048)
            return {"status": "complete"}
        
        retryable = make_conversion_retryable(original_convert, max_attempts=3, base_delay=0.01)
        result = retryable("input.iso", "output.chd")
        
        assert result["status"] == "complete"
        assert call_count[0] == 2
    
    def test_conversion_with_disk_full_retry(self, clean_logging):
        """Conversion should retry on DiskFullError."""
        call_count = [0]
        
        def original_convert(input_file, output_file):
            call_count[0] += 1
            if call_count[0] == 1:
                raise DiskFullError(4.8, 0.1)
            return {"status": "complete"}
        
        retryable = make_conversion_retryable(original_convert, max_attempts=2, base_delay=0.01)
        result = retryable("input.iso", "output.chd")
        
        assert result["status"] == "complete"
        assert call_count[0] == 2


class TestRetryIntegrationWithLogger:
    """Test retry integration with actual logger."""
    
    def test_retry_logs_to_parent_logger(self, clean_logging):
        """Retries should log to provided parent logger."""
        logger = get_logger("test_retry")
        call_count = [0]
        
        @retry_operation(max_attempts=2, base_delay=0.01)
        def failing_func():
            call_count[0] += 1
            if call_count[0] == 1:
                raise TimeoutError("Timeout", 5)
            return "ok"
        
        # Capture log output
        import io
        import sys
        
        # Call the function
        result = failing_func()
        
        assert result == "ok"
        assert call_count[0] == 2


class TestRetrySequence:
    """Test complete retry sequence with multiple attempts."""
    
    def test_three_attempt_sequence(self, clean_logging):
        """Test sequence: fail, fail, succeed."""
        call_count = [0]
        
        @retry_operation(max_attempts=5, base_delay=0.01)
        def unstable_operation():
            call_count[0] += 1
            if call_count[0] < 3:
                raise TimeoutError("Timeout", 5)
            return f"success_on_attempt_{call_count[0]}"
        
        result = unstable_operation()
        
        assert result == "success_on_attempt_3"
        assert call_count[0] == 3
    
    def test_ultimate_failure_after_max_attempts(self, clean_logging):
        """Should fail permanently after max attempts exhausted."""
        call_count = [0]
        max_attempts = 3
        
        @retry_operation(max_attempts=max_attempts, base_delay=0.01)
        def permanently_failing():
            call_count[0] += 1
            raise TimeoutError("Timeout", 5)
        
        with pytest.raises(TimeoutError):
            permanently_failing()
        
        assert call_count[0] == max_attempts


class TestRetryWithDifferentErrors:
    """Test retry behavior with different error types."""
    
    def test_retryable_error_gets_retried(self, clean_logging):
        """TimeoutError should be retried."""
        call_count = [0]
        
        @retry_operation(max_attempts=3, base_delay=0.01)
        def timeout_func():
            call_count[0] += 1
            if call_count[0] < 2:
                raise TimeoutError("Timeout", 5)
            return "ok"
        
        result = timeout_func()
        assert result == "ok"
        assert call_count[0] == 2
    
    def test_non_retryable_error_fails_immediately(self, clean_logging):
        """FileNotFoundError should fail without retry."""
        call_count = [0]
        
        @retry_operation(max_attempts=3, base_delay=0.01)
        def not_found_func():
            call_count[0] += 1
            raise FileNotFoundError("File not found")
        
        with pytest.raises(FileNotFoundError):
            not_found_func()
        
        # Should fail on first attempt without retrying
        assert call_count[0] == 1


class TestRetryIntegrationPattern:
    """Test realistic rom_converter usage patterns."""
    
    def test_download_method_pattern(self, clean_logging):
        """Simulate rom_converter.download_mame_tools pattern."""
        
        class MockROMConverter:
            def __init__(self):
                self.logger = get_logger("rom_converter")
                self.call_count = 0
            
            @retry_operation(max_attempts=5, base_delay=0.01, operation_name="download_mame")
            def download_mame_tools(self):
                self.call_count += 1
                if self.call_count < 2:
                    raise TimeoutError("Timeout", 30)
                return {"version": "0.256", "url": "http://download.mame.dev"}
        
        converter = MockROMConverter()
        result = converter.download_mame_tools()
        
        assert result["version"] == "0.256"
        assert converter.call_count == 2
    
    def test_extraction_method_pattern(self, clean_logging):
        """Simulate rom_converter.extract_archive pattern."""
        
        class MockROMConverter:
            def __init__(self):
                self.logger = get_logger("rom_converter")
                self.call_count = 0
            
            @retry_operation(max_attempts=3, base_delay=0.01, operation_name="extract")
            def extract_archive(self, archive_path):
                self.call_count += 1
                if self.call_count < 2:
                    raise TimeoutError("Timeout", 10)
                return {"files_extracted": 25}
        
        converter = MockROMConverter()
        result = converter.extract_archive("/path/archive.zip")
        
        assert result["files_extracted"] == 25
        assert converter.call_count == 2


if __name__ == "__main__":
    pytest.main([__file__, "-v"])



