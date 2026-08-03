"""
Structured logging configuration for ROM Converter.

Provides file and console logging with:
- Multiple log levels (DEBUG, INFO, WARNING, ERROR, CRITICAL)
- Automatic log rotation
- Formatted output with timestamps and context
- Both synchronous and asynchronous logging
"""

import logging
import logging.handlers
from pathlib import Path
from datetime import datetime
from typing import Optional
import sys
import pytest


@pytest.fixture
def cleanup_logging():
    """Fixture to clean up logging handlers after tests."""
    yield
    # Clean up all handlers from all loggers
    for logger_name in list(logging.Logger.manager.loggerDict):
        logger = logging.getLogger(logger_name)
        for handler in logger.handlers[:]:
            handler.close()
            logger.removeHandler(handler)


class ColoredFormatter(logging.Formatter):
    """Format logs with color codes for console output."""
    
    # ANSI color codes
    COLORS = {
        'DEBUG': '\033[36m',      # Cyan
        'INFO': '\033[32m',       # Green
        'WARNING': '\033[33m',    # Yellow
        'ERROR': '\033[31m',      # Red
        'CRITICAL': '\033[35m',   # Magenta
    }
    RESET = '\033[0m'
    
    def format(self, record):
        # Add color to level name if on terminal
        if sys.stdout.isatty():
            level_color = self.COLORS.get(record.levelname, '')
            record.levelname = f"{level_color}{record.levelname}{self.RESET}"
        
        # Format the message
        return super().format(record)


class FileFormatter(logging.Formatter):
    """Format logs for file output with full details."""
    
    def format(self, record):
        # Ensure we have a timestamp
        if not hasattr(record, 'created') or record.created is None:
            record.created = datetime.now().timestamp()
        
        return super().format(record)


def setup_logging(
    log_dir: Optional[Path] = None,
    app_name: str = "rom_converter",
    console_level: int = logging.INFO,
    file_level: int = logging.DEBUG,
    max_bytes: int = 10 * 1024 * 1024,  # 10 MB
    backup_count: int = 5
) -> logging.Logger:
    """Set up logging for ROM Converter.
    
    Creates both file and console handlers with appropriate formatting.
    File logging includes automatic rotation when size limit is reached.
    
    Args:
        log_dir: Directory for log files. If None, uses app directory.
        app_name: Name for the logger and log files
        console_level: Logging level for console (default: INFO)
        file_level: Logging level for file (default: DEBUG)
        max_bytes: Maximum file size before rotation (default: 10 MB)
        backup_count: Number of backup files to keep (default: 5)
    
    Returns:
        Configured logger instance
    
    Example:
        logger = setup_logging(Path("./logs"))
        logger.info("Application started")
        logger.debug("Debug information")
        logger.error("An error occurred: %s", error_msg)
    """
    
    # Create logger
    logger = logging.getLogger(app_name)
    logger.setLevel(logging.DEBUG)  # Capture everything, handlers filter
    
    # Remove existing handlers (but keep caplog's handler in tests)
    # Count non-caplog handlers to avoid duplicate initialization
    non_caplog_handlers = [h for h in logger.handlers 
                          if not h.__class__.__name__ == 'CaptureHandler']
    if non_caplog_handlers:
        return logger
    
    # Determine log directory
    if log_dir is None:
        log_dir = Path.home() / ".rom_converter_logs"
    else:
        log_dir = Path(log_dir)
    
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # ===== FILE HANDLER =====
    log_file = log_dir / f"{app_name}_{datetime.now().strftime('%Y%m%d')}.log"
    file_handler = logging.handlers.RotatingFileHandler(
        filename=str(log_file),
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding='utf-8'
    )
    file_handler.setLevel(file_level)
    
    # File formatter with full details
    file_formatter = FileFormatter(
        fmt='[%(asctime)s] %(levelname)-8s [%(name)s:%(funcName)s:%(lineno)d] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)
    
    # ===== CONSOLE HANDLER =====
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(console_level)
    
    # Console formatter (more concise, with color)
    console_formatter = ColoredFormatter(
        fmt='%(levelname)-8s %(message)s',
        datefmt='%H:%M:%S'
    )
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)
    
    return logger


def get_logger(name: str = "rom_converter") -> logging.Logger:
    """Get logger instance (after setup_logging has been called).
    
    Args:
        name: Logger name (should match setup_logging app_name)
    
    Returns:
        Logger instance
    
    Example:
        logger = get_logger()
        logger.info("Message")
    """
    return logging.getLogger(name)


def log_exception(logger: logging.Logger, exception: Exception, message: str = "Exception occurred") -> None:
    """Log an exception with full traceback.
    
    Args:
        logger: Logger instance
        exception: Exception that occurred
        message: Custom message to prefix exception
    
    Example:
        try:
            risky_operation()
        except Exception as e:
            log_exception(logger, e, "Failed to process file")
    """
    logger.exception(f"{message}: {exception}")


def log_conversion_start(logger: logging.Logger, file_name: str, format_in: str, format_out: str) -> None:
    """Log conversion start with context.
    
    Args:
        logger: Logger instance
        file_name: Name of file being converted
        format_in: Input format (e.g., "ISO")
        format_out: Output format (e.g., "CHD")
    
    Example:
        log_conversion_start(logger, "game.iso", "ISO", "CHD")
    """
    logger.info(f"Starting conversion: {file_name} ({format_in} → {format_out})")


def log_conversion_complete(
    logger: logging.Logger,
    file_name: str,
    original_size: int,
    output_size: int,
    duration_seconds: float
) -> None:
    """Log successful conversion with metrics.
    
    Args:
        logger: Logger instance
        file_name: Name of converted file
        original_size: Original file size in bytes
        output_size: Output file size in bytes
        duration_seconds: Time taken in seconds
    
    Example:
        log_conversion_complete(logger, "game.chd", 2500000000, 1800000000, 45.5)
    """
    size_reduction = 100 * (1 - output_size / original_size) if original_size > 0 else 0
    logger.info(
        f"Conversion complete: {file_name} | "
        f"Size: {original_size:,} → {output_size:,} bytes "
        f"({size_reduction:.1f}% reduction) | "
        f"Time: {duration_seconds:.1f}s"
    )


def log_conversion_error(
    logger: logging.Logger,
    file_name: str,
    error_msg: str,
    returncode: Optional[int] = None
) -> None:
    """Log conversion failure.
    
    Args:
        logger: Logger instance
        file_name: Name of file that failed
        error_msg: Error message
        returncode: Tool exit code (optional)
    
    Example:
        log_conversion_error(logger, "game.iso", "Out of disk space", 1)
    """
    code_str = f" (exit code: {returncode})" if returncode else ""
    logger.error(f"Conversion failed: {file_name} | Error: {error_msg}{code_str}")


def log_resource_status(
    logger: logging.Logger,
    cpu_percent: float,
    ram_percent: float,
    disk_percent: float,
    active_workers: int
) -> None:
    """Log system resource status.
    
    Args:
        logger: Logger instance
        cpu_percent: CPU usage percentage (0-100)
        ram_percent: RAM usage percentage (0-100)
        disk_percent: Disk usage percentage (0-100)
        active_workers: Number of active conversion workers
    
    Example:
        log_resource_status(logger, 85.5, 72.3, 65.0, 3)
    """
    logger.debug(
        f"Resources: CPU={cpu_percent:.1f}% RAM={ram_percent:.1f}% "
        f"Disk={disk_percent:.1f}% Workers={active_workers}"
    )


# Convenience module-level functions for quick logging
_default_logger: Optional[logging.Logger] = None


def init(log_dir: Optional[Path] = None) -> logging.Logger:
    """Initialize logging globally.
    
    Args:
        log_dir: Directory for log files
    
    Returns:
        Logger instance
    
    Example:
        from logging_setup import init, info, error
        init(Path("./logs"))
        info("Application started")
    """
    global _default_logger
    _default_logger = setup_logging(log_dir)
    return _default_logger


def debug(msg: str, *args, **kwargs) -> None:
    """Log debug message (requires init() first)."""
    if _default_logger:
        _default_logger.debug(msg, *args, **kwargs)


def info(msg: str, *args, **kwargs) -> None:
    """Log info message (requires init() first)."""
    if _default_logger:
        _default_logger.info(msg, *args, **kwargs)


def warning(msg: str, *args, **kwargs) -> None:
    """Log warning message (requires init() first)."""
    if _default_logger:
        _default_logger.warning(msg, *args, **kwargs)


def error(msg: str, *args, **kwargs) -> None:
    """Log error message (requires init() first)."""
    if _default_logger:
        _default_logger.error(msg, *args, **kwargs)


def critical(msg: str, *args, **kwargs) -> None:
    """Log critical message (requires init() first)."""
    if _default_logger:
        _default_logger.critical(msg, *args, **kwargs)


def exception(msg: str, *args, **kwargs) -> None:
    """Log exception with traceback (requires init() first)."""
    if _default_logger:
        _default_logger.exception(msg, *args, **kwargs)
