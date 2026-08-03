"""
Custom exception hierarchy for ROM Converter.

All exceptions in this module inherit from ROMConverterError and should be
caught at the top level for proper error handling and user feedback.
"""

from typing import Optional


class ROMConverterError(Exception):
    """Base exception for all ROM Converter errors.
    
    All other exceptions in this module should inherit from this class.
    """
    
    def __init__(self, message: str, details: Optional[str] = None):
        """Initialize exception with message and optional details.
        
        Args:
            message: User-friendly error message
            details: Technical details for logging/debugging
        """
        self.message = message
        self.details = details
        super().__init__(self.message)
    
    def __str__(self) -> str:
        if self.details:
            return f"{self.message}\n[Details: {self.details}]"
        return self.message


class ConversionError(ROMConverterError):
    """Raised when ROM conversion fails.
    
    This includes any error that occurs during the actual conversion process,
    such as the converter tool failing or producing invalid output.
    """
    
    def __init__(
        self,
        message: str,
        stderr: Optional[str] = None,
        returncode: Optional[int] = None,
        retry_possible: bool = False,
        details: Optional[str] = None
    ):
        """Initialize conversion error.
        
        Args:
            message: User-friendly error message
            stderr: Standard error output from converter tool
            returncode: Exit code from converter tool
            retry_possible: Whether retrying might succeed (transient error)
            details: Additional technical details
        """
        super().__init__(message, details)
        self.stderr = stderr
        self.returncode = returncode
        self.retry_possible = retry_possible
    
    def __str__(self) -> str:
        base = super().__str__()
        if self.retry_possible:
            base += "\n[This error may be transient - retry might succeed]"
        return base


class TimeoutError(ConversionError):
    """Raised when conversion tool times out.
    
    Indicates that the converter took too long and was forcibly terminated.
    May be a transient error if the system was just overloaded.
    """
    
    def __init__(self, message: str, timeout_seconds: int, details: Optional[str] = None):
        """Initialize timeout error.
        
        Args:
            message: User-friendly error message
            timeout_seconds: Timeout limit that was exceeded
            details: Additional technical details
        """
        super().__init__(
            f"{message} (timeout after {timeout_seconds}s)",
            retry_possible=True,
            details=details
        )
        self.timeout_seconds = timeout_seconds


class ToolNotFoundError(ROMConverterError):
    """Raised when a required external tool is not found or not executable.
    
    This can occur when:
    - chdman, maxcso, 7z, extract-xiso, etc. are missing
    - Tool is not in PATH and not in app directory
    - Tool exists but is not executable (permissions issue)
    """
    
    def __init__(self, tool_name: str, search_paths: Optional[list] = None, details: Optional[str] = None):
        """Initialize tool not found error.
        
        Args:
            tool_name: Name of the missing tool (e.g., 'chdman.exe')
            search_paths: List of paths that were searched
            details: Additional technical details
        """
        paths_msg = ""
        if search_paths:
            paths_msg = f"\nSearched in: {', '.join(search_paths)}"
        
        message = f"Required tool not found: {tool_name}{paths_msg}"
        super().__init__(message, details)
        self.tool_name = tool_name
        self.search_paths = search_paths or []


class ResourceError(ROMConverterError):
    """Raised when system resources (RAM, disk space, CPU) are exhausted.
    
    These errors are typically transient - retrying after the system
    recovers from the resource shortage should succeed.
    """
    
    def __init__(self, message: str, details: Optional[str] = None):
        """Initialize resource error.
        
        Args:
            message: User-friendly error message
            details: Additional technical details (e.g., "RAM: 92%, CPU: 98%")
        """
        super().__init__(message, details)


class OutOfMemoryError(ResourceError):
    """Raised when system runs out of available memory.
    
    Indicates the conversion process cannot continue until memory is freed.
    This is typically a transient error - waiting and retrying usually works.
    """
    
    def __init__(self, required_mb: int, available_mb: int, details: Optional[str] = None):
        """Initialize out of memory error.
        
        Args:
            required_mb: Memory needed for conversion
            available_mb: Memory currently available
            details: Additional technical details
        """
        message = f"Out of memory: need {required_mb} MB, have {available_mb} MB"
        super().__init__(message, details)
        self.required_mb = required_mb
        self.available_mb = available_mb


class DiskFullError(ResourceError):
    """Raised when output disk is full or nearly full.
    
    Indicates there is not enough disk space to complete the conversion.
    """
    
    def __init__(self, required_gb: float, available_gb: float, details: Optional[str] = None):
        """Initialize disk full error.
        
        Args:
            required_gb: Disk space needed for conversion
            available_gb: Disk space currently available
            details: Additional technical details
        """
        message = f"Disk full: need {required_gb:.1f} GB, have {available_gb:.1f} GB"
        super().__init__(message, details)
        self.required_gb = required_gb
        self.available_gb = available_gb


class ConfigError(ROMConverterError):
    """Raised when configuration is invalid or cannot be loaded/saved.
    
    This includes:
    - Corrupted config file (JSON parse error)
    - Invalid config values
    - File permission issues
    - Missing required config keys
    """
    
    def __init__(self, message: str, details: Optional[str] = None):
        """Initialize config error.
        
        Args:
            message: User-friendly error message
            details: Additional technical details
        """
        super().__init__(message, details)


class InvalidConfigError(ConfigError):
    """Raised when config values are invalid or out of range."""
    
    def __init__(self, key: str, value: str, expected: str, details: Optional[str] = None):
        """Initialize invalid config error.
        
        Args:
            key: Config key name (e.g., 'max_concurrent_conversions')
            value: Invalid value provided
            expected: Description of expected format
            details: Additional technical details
        """
        message = f"Invalid config value for '{key}': {value} (expected: {expected})"
        super().__init__(message, details)
        self.key = key
        self.value = value
        self.expected = expected


class MissingConfigError(ConfigError):
    """Raised when required config keys are missing."""
    
    def __init__(self, missing_keys: list, details: Optional[str] = None):
        """Initialize missing config error.
        
        Args:
            missing_keys: List of required keys that are missing
            details: Additional technical details
        """
        keys_str = ", ".join(missing_keys)
        message = f"Missing required config keys: {keys_str}"
        super().__init__(message, details)
        self.missing_keys = missing_keys


class ExtractionError(ROMConverterError):
    """Raised when archive extraction fails.
    
    This includes:
    - Corrupted or invalid archive
    - Unsupported archive format
    - Extraction tool (7z, etc.) errors
    - Permission issues on output directory
    """
    
    def __init__(self, message: str, archive_path: Optional[str] = None, details: Optional[str] = None):
        """Initialize extraction error.
        
        Args:
            message: User-friendly error message
            archive_path: Path to archive that failed to extract
            details: Additional technical details
        """
        if archive_path:
            message = f"{message}: {archive_path}"
        super().__init__(message, details)
        self.archive_path = archive_path


class CorruptArchiveError(ExtractionError):
    """Raised when archive file is corrupted or invalid."""
    
    def __init__(self, archive_path: str, details: Optional[str] = None):
        """Initialize corrupt archive error.
        
        Args:
            archive_path: Path to corrupted archive
            details: Additional technical details
        """
        message = "Archive is corrupted or invalid"
        super().__init__(message, archive_path, details)


class UnsupportedFormatError(ExtractionError):
    """Raised when archive format is not supported."""
    
    def __init__(self, file_extension: str, archive_path: Optional[str] = None, details: Optional[str] = None):
        """Initialize unsupported format error.
        
        Args:
            file_extension: File extension of unsupported format
            archive_path: Path to unsupported archive
            details: Additional technical details
        """
        message = f"Unsupported archive format: {file_extension}"
        super().__init__(message, archive_path, details)
        self.file_extension = file_extension


class DecryptionError(ROMConverterError):
    """Raised when 3DS ROM decryption fails.
    
    This includes:
    - Missing or invalid encryption keys
    - Corrupted encrypted ROM
    - NDecrypt tool errors
    """
    
    def __init__(self, message: str, rom_path: Optional[str] = None, details: Optional[str] = None):
        """Initialize decryption error.
        
        Args:
            message: User-friendly error message
            rom_path: Path to ROM that failed to decrypt
            details: Additional technical details
        """
        if rom_path:
            message = f"{message}: {rom_path}"
        super().__init__(message, details)
        self.rom_path = rom_path


class MissingKeysError(DecryptionError):
    """Raised when encryption keys are not found or invalid."""
    
    def __init__(self, keys_file_path: str, details: Optional[str] = None):
        """Initialize missing keys error.
        
        Args:
            keys_file_path: Expected path to keys file
            details: Additional technical details
        """
        message = f"Encryption keys not found: {keys_file_path}"
        super().__init__(message, details=details)
        self.keys_file_path = keys_file_path


class PlatformError(ROMConverterError):
    """Raised when platform-specific functionality fails.
    
    This includes:
    - Flatpak detection/mounting issues
    - Immutable distro incompatibilities
    - File permission issues
    """
    
    def __init__(self, message: str, platform: Optional[str] = None, details: Optional[str] = None):
        """Initialize platform error.
        
        Args:
            message: User-friendly error message
            platform: Detected platform (e.g., 'Fedora Atomic')
            details: Additional technical details
        """
        if platform:
            message = f"{message} (on {platform})"
        super().__init__(message, details)
        self.platform = platform


# Error categorization helpers

def is_transient_error(error: Exception) -> bool:
    """Check if an error is transient (retry might work).
    
    Transient errors include:
    - TimeoutError
    - OutOfMemoryError (might recover after waiting)
    - DiskFullError (might recover after freeing space)
    - ConversionError with retry_possible=True
    
    Args:
        error: Exception to check
    
    Returns:
        True if error is transient, False otherwise
    """
    if isinstance(error, TimeoutError):
        return True
    if isinstance(error, (OutOfMemoryError, DiskFullError)):
        return True
    if isinstance(error, ConversionError) and error.retry_possible:
        return True
    return False


def is_recoverable_error(error: Exception) -> bool:
    """Check if an error is recoverable (user action might fix it).
    
    Recoverable errors include:
    - ToolNotFoundError (user can download/install tool)
    - MissingKeysError (user can download keys)
    - InvalidConfigError (user can fix config)
    
    Args:
        error: Exception to check
    
    Returns:
        True if error is recoverable, False otherwise
    """
    if isinstance(error, ToolNotFoundError):
        return True
    if isinstance(error, MissingKeysError):
        return True
    if isinstance(error, InvalidConfigError):
        return True
    return False


def error_to_user_message(error: Exception) -> str:
    """Convert exception to user-friendly error message.
    
    Args:
        error: Exception to convert
    
    Returns:
        User-friendly error message
    """
    if isinstance(error, ROMConverterError):
        return error.message
    
    # Generic exception
    return f"An error occurred: {str(error)}"
