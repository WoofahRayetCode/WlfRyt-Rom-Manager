"""
Logging integration for ROM Converter.

This module provides a bridge between the existing ROM Converter UI logging
and the structured logging system. It enables gradual migration to structured
logging without requiring a complete rewrite of the main application.

Usage:
    from logging_integration import setup_converter_logging, get_converter_logger
    
    # Initialize logging at app startup
    setup_converter_logging()
    
    # Get logger for current module/context
    logger = get_converter_logger()
    
    # Log events (works alongside existing UI logging)
    logger.info("Conversion started", extra={"file": "game.iso", "format": "CHD"})
    logger.error("Conversion failed", extra={"error": "Out of disk space", "code": 1})
"""

import logging
from pathlib import Path
from typing import Optional
from logging_setup import setup_logging, get_logger, log_conversion_start, log_conversion_complete, log_conversion_error


def setup_converter_logging(log_dir: Optional[Path] = None) -> logging.Logger:
    """Initialize logging for ROM Converter application.
    
    Sets up file and console logging. Call this once at application startup.
    
    Args:
        log_dir: Directory for log files. If None, uses user home directory.
    
    Returns:
        Configured logger instance
    
    Example:
        # In rom_converter.py __init__ or app startup
        self.logger = setup_converter_logging()
    """
    if log_dir is None:
        log_dir = Path.home() / ".rom_converter_logs"
    
    return setup_logging(
        log_dir=log_dir,
        app_name="rom_converter",
        console_level=logging.INFO,
        file_level=logging.DEBUG
    )


def get_converter_logger(name: str = "rom_converter") -> logging.Logger:
    """Get logger for ROM Converter.
    
    Args:
        name: Logger name (default: "rom_converter")
    
    Returns:
        Logger instance
    
    Example:
        logger = get_converter_logger()
        logger.info("Starting conversion")
    """
    return get_logger(name)


class ConversionLogger:
    """Convenience wrapper for conversion-specific logging.
    
    Simplifies logging of conversion lifecycle events with consistent formatting.
    
    Example:
        conv_logger = ConversionLogger()
        conv_logger.start("game.iso", "CHD")
        conv_logger.complete("game.chd", 2500000000, 1800000000, 45.5)
    """
    
    def __init__(self, logger: Optional[logging.Logger] = None):
        """Initialize conversion logger.
        
        Args:
            logger: Logger instance. If None, uses default rom_converter logger.
        """
        self.logger = logger or get_converter_logger()
    
    def start(self, filename: str, format_in: str, format_out: str) -> None:
        """Log conversion start.
        
        Args:
            filename: Name of file being converted
            format_in: Input format (e.g., "ISO")
            format_out: Output format (e.g., "CHD")
        """
        log_conversion_start(self.logger, filename, format_in, format_out)
    
    def complete(
        self,
        filename: str,
        original_size: int,
        output_size: int,
        duration_seconds: float
    ) -> None:
        """Log successful conversion.
        
        Args:
            filename: Name of converted file
            original_size: Original file size in bytes
            output_size: Output file size in bytes
            duration_seconds: Time taken in seconds
        """
        log_conversion_complete(
            self.logger,
            filename,
            original_size,
            output_size,
            duration_seconds
        )
    
    def error(
        self,
        filename: str,
        error_msg: str,
        returncode: Optional[int] = None
    ) -> None:
        """Log conversion failure.
        
        Args:
            filename: Name of file that failed
            error_msg: Error message
            returncode: Tool exit code (optional)
        """
        log_conversion_error(self.logger, filename, error_msg, returncode)
    
    def debug(self, message: str, **kwargs) -> None:
        """Log debug message.
        
        Args:
            message: Message to log
            **kwargs: Additional context (e.g., file, format, size)
        """
        if kwargs:
            context = " | ".join(f"{k}={v}" for k, v in kwargs.items())
            self.logger.debug(f"{message} | {context}")
        else:
            self.logger.debug(message)
    
    def info(self, message: str, **kwargs) -> None:
        """Log info message.
        
        Args:
            message: Message to log
            **kwargs: Additional context
        """
        if kwargs:
            context = " | ".join(f"{k}={v}" for k, v in kwargs.items())
            self.logger.info(f"{message} | {context}")
        else:
            self.logger.info(message)
    
    def warning(self, message: str, **kwargs) -> None:
        """Log warning message.
        
        Args:
            message: Message to log
            **kwargs: Additional context
        """
        if kwargs:
            context = " | ".join(f"{k}={v}" for k, v in kwargs.items())
            self.logger.warning(f"{message} | {context}")
        else:
            self.logger.warning(message)


# Module-level convenience functions for quick logging
_converter_logger: Optional[logging.Logger] = None


def init(log_dir: Optional[Path] = None) -> logging.Logger:
    """Initialize logging globally.
    
    Args:
        log_dir: Directory for log files
    
    Returns:
        Logger instance
    
    Example:
        from logging_integration import init, info, error
        init()
        info("Starting conversion")
    """
    global _converter_logger
    _converter_logger = setup_converter_logging(log_dir)
    return _converter_logger


def debug(msg: str, *args, **kwargs) -> None:
    """Log debug message."""
    if _converter_logger:
        _converter_logger.debug(msg, *args, **kwargs)


def info(msg: str, *args, **kwargs) -> None:
    """Log info message."""
    if _converter_logger:
        _converter_logger.info(msg, *args, **kwargs)


def warning(msg: str, *args, **kwargs) -> None:
    """Log warning message."""
    if _converter_logger:
        _converter_logger.warning(msg, *args, **kwargs)


def error(msg: str, *args, **kwargs) -> None:
    """Log error message."""
    if _converter_logger:
        _converter_logger.error(msg, *args, **kwargs)


def exception(msg: str, *args, **kwargs) -> None:
    """Log exception with traceback."""
    if _converter_logger:
        _converter_logger.exception(msg, *args, **kwargs)
