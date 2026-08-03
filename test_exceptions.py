"""
Unit tests for custom exception hierarchy.

Tests the exception classes to ensure:
- Exceptions are created with proper attributes
- Error categorization helpers work correctly
- User-friendly error messages are generated
"""

import pytest
import sys
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from exceptions import (
    ROMConverterError,
    ConversionError,
    TimeoutError as ROMConverterTimeoutError,
    ToolNotFoundError,
    ResourceError,
    OutOfMemoryError,
    DiskFullError,
    ConfigError,
    InvalidConfigError,
    MissingConfigError,
    ExtractionError,
    CorruptArchiveError,
    UnsupportedFormatError,
    DecryptionError,
    MissingKeysError,
    PlatformError,
    is_transient_error,
    is_recoverable_error,
    error_to_user_message,
)


@pytest.mark.unit
class TestBaseException:
    """Test the base ROMConverterError class."""
    
    def test_basic_exception_creation(self):
        """Test creating a basic exception."""
        error = ROMConverterError("Test error")
        
        assert error.message == "Test error"
        assert error.details is None
        assert str(error) == "Test error"
    
    def test_exception_with_details(self):
        """Test exception with technical details."""
        error = ROMConverterError(
            "Test error",
            details="Technical details here"
        )
        
        assert error.message == "Test error"
        assert error.details == "Technical details here"
        assert "Technical details" in str(error)
    
    def test_exception_inheritance(self):
        """Test that exceptions inherit from base class."""
        error = ConversionError("Conversion failed")
        
        assert isinstance(error, ROMConverterError)
        assert isinstance(error, Exception)


@pytest.mark.unit
class TestConversionError:
    """Test ConversionError and related exceptions."""
    
    def test_conversion_error_basic(self):
        """Test basic conversion error."""
        error = ConversionError("Conversion failed")
        
        assert error.message == "Conversion failed"
        assert error.stderr is None
        assert error.returncode is None
        assert error.retry_possible is False
    
    def test_conversion_error_with_stderr(self):
        """Test conversion error with stderr output."""
        error = ConversionError(
            "Conversion failed",
            stderr="Invalid file format",
            returncode=1
        )
        
        assert error.stderr == "Invalid file format"
        assert error.returncode == 1
    
    def test_conversion_error_transient(self):
        """Test conversion error marked as transient."""
        error = ConversionError(
            "Conversion failed",
            stderr="out of memory",
            retry_possible=True
        )
        
        assert error.retry_possible is True
        assert "may be transient" in str(error)
    
    def test_timeout_error(self):
        """Test timeout error."""
        error = ROMConverterTimeoutError(
            "Conversion took too long",
            timeout_seconds=600
        )
        
        assert error.message == "Conversion took too long (timeout after 600s)"
        assert error.timeout_seconds == 600
        assert error.retry_possible is True  # Timeouts are transient


@pytest.mark.unit
class TestResourceErrors:
    """Test resource-related errors."""
    
    def test_out_of_memory_error(self):
        """Test out of memory error."""
        error = OutOfMemoryError(required_mb=2048, available_mb=512)
        
        assert error.required_mb == 2048
        assert error.available_mb == 512
        assert "need 2048 MB" in str(error)
        assert "have 512 MB" in str(error)
    
    def test_disk_full_error(self):
        """Test disk full error."""
        error = DiskFullError(required_gb=5.5, available_gb=2.1)
        
        assert error.required_gb == 5.5
        assert error.available_gb == 2.1
        assert "5.5 GB" in str(error)
        assert "2.1 GB" in str(error)
    
    def test_resource_errors_inherit_from_base(self):
        """Test that resource errors inherit properly."""
        oom_error = OutOfMemoryError(required_mb=1024, available_mb=512)
        disk_error = DiskFullError(required_gb=10, available_gb=5)
        
        assert isinstance(oom_error, ResourceError)
        assert isinstance(oom_error, ROMConverterError)
        assert isinstance(disk_error, ResourceError)
        assert isinstance(disk_error, ROMConverterError)


@pytest.mark.unit
class TestToolErrors:
    """Test tool-related errors."""
    
    def test_tool_not_found_error(self):
        """Test tool not found error."""
        error = ToolNotFoundError("chdman.exe")
        
        assert error.tool_name == "chdman.exe"
        assert "chdman.exe" in str(error)
    
    def test_tool_not_found_with_search_paths(self):
        """Test tool not found error with search paths."""
        search_paths = ["/usr/bin", "/usr/local/bin", "/opt/tools"]
        error = ToolNotFoundError(
            "chdman",
            search_paths=search_paths
        )
        
        assert error.search_paths == search_paths
        for path in search_paths:
            assert path in str(error)


@pytest.mark.unit
class TestConfigErrors:
    """Test configuration-related errors."""
    
    def test_invalid_config_error(self):
        """Test invalid config value error."""
        error = InvalidConfigError(
            key="max_workers",
            value="abc",
            expected="integer between 1 and 16"
        )
        
        assert error.key == "max_workers"
        assert error.value == "abc"
        assert error.expected == "integer between 1 and 16"
        assert "max_workers" in str(error)
        assert "abc" in str(error)
    
    def test_missing_config_error(self):
        """Test missing config keys error."""
        missing = ["output_format", "max_workers"]
        error = MissingConfigError(missing_keys=missing)
        
        assert error.missing_keys == missing
        for key in missing:
            assert key in str(error)


@pytest.mark.unit
class TestExtractionErrors:
    """Test archive extraction errors."""
    
    def test_corrupt_archive_error(self):
        """Test corrupt archive error."""
        error = CorruptArchiveError("/path/to/archive.zip")
        
        assert error.archive_path == "/path/to/archive.zip"
        assert "archive.zip" in str(error)
    
    def test_unsupported_format_error(self):
        """Test unsupported archive format error."""
        error = UnsupportedFormatError(
            file_extension=".rar",
            archive_path="/path/to/archive.rar"
        )
        
        assert error.file_extension == ".rar"
        assert ".rar" in str(error)


@pytest.mark.unit
class TestDecryptionErrors:
    """Test 3DS decryption errors."""
    
    def test_missing_keys_error(self):
        """Test missing encryption keys error."""
        error = MissingKeysError("/home/user/3DS_AES_Keys.txt")
        
        assert error.keys_file_path == "/home/user/3DS_AES_Keys.txt"
        assert "3DS_AES_Keys.txt" in str(error)


@pytest.mark.unit
class TestErrorCategorization:
    """Test error categorization helper functions."""
    
    def test_is_transient_error_timeout(self):
        """Test that TimeoutError is marked transient."""
        error = ROMConverterTimeoutError("Timed out", timeout_seconds=600)
        assert is_transient_error(error) is True
    
    def test_is_transient_error_oom(self):
        """Test that OutOfMemoryError is marked transient."""
        error = OutOfMemoryError(required_mb=2048, available_mb=512)
        assert is_transient_error(error) is True
    
    def test_is_transient_error_disk_full(self):
        """Test that DiskFullError is marked transient."""
        error = DiskFullError(required_gb=10, available_gb=5)
        assert is_transient_error(error) is True
    
    def test_is_transient_error_conversion(self):
        """Test transient ConversionError detection."""
        transient = ConversionError(
            "Failed",
            retry_possible=True
        )
        non_transient = ConversionError("Failed")
        
        assert is_transient_error(transient) is True
        assert is_transient_error(non_transient) is False
    
    def test_is_recoverable_error_tool_not_found(self):
        """Test that ToolNotFoundError is recoverable."""
        error = ToolNotFoundError("chdman.exe")
        assert is_recoverable_error(error) is True
    
    def test_is_recoverable_error_missing_keys(self):
        """Test that MissingKeysError is recoverable."""
        error = MissingKeysError("/path/to/keys.txt")
        assert is_recoverable_error(error) is True
    
    def test_is_recoverable_error_invalid_config(self):
        """Test that InvalidConfigError is recoverable."""
        error = InvalidConfigError("max_workers", "abc", "integer")
        assert is_recoverable_error(error) is True
    
    def test_is_not_recoverable_for_other_errors(self):
        """Test that other errors are not marked recoverable."""
        error = ConversionError("Some error")
        assert is_recoverable_error(error) is False


@pytest.mark.unit
class TestErrorMessages:
    """Test user-friendly error message generation."""
    
    def test_user_message_rom_converter_error(self):
        """Test user message from ROMConverterError."""
        error = ROMConverterError("This is a user-friendly message")
        message = error_to_user_message(error)
        
        assert message == "This is a user-friendly message"
    
    def test_user_message_conversion_error(self):
        """Test user message from ConversionError."""
        error = ConversionError("Conversion failed due to invalid format")
        message = error_to_user_message(error)
        
        assert message == "Conversion failed due to invalid format"
    
    def test_user_message_generic_exception(self):
        """Test user message from generic exception."""
        error = ValueError("Generic error")
        message = error_to_user_message(error)
        
        assert "An error occurred" in message
        assert "Generic error" in message


@pytest.mark.unit
def test_exception_chain():
    """Test exception inheritance chain."""
    # Create various exceptions and verify inheritance
    errors = [
        ConversionError("test"),
        ROMConverterTimeoutError("test", 600),
        OutOfMemoryError(1024, 512),
        ToolNotFoundError("chdman.exe"),
        InvalidConfigError("key", "value", "expected"),
        CorruptArchiveError("/path.zip"),
        MissingKeysError("/path/keys.txt"),
    ]
    
    # All should inherit from ROMConverterError
    for error in errors:
        assert isinstance(error, ROMConverterError)
        assert isinstance(error, Exception)
