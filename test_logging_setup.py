"""
Unit tests for logging_setup module.

Tests the logging configuration, handlers, and formatting.
"""

import pytest
import logging
from logging_setup import (
    get_logger,
    log_exception,
    log_conversion_start,
    log_conversion_complete,
    log_conversion_error,
    log_resource_status,
    ColoredFormatter,
    FileFormatter,
)


@pytest.mark.unit
class TestLogFormatters:
    """Test log formatting."""
    
    def test_colored_formatter_creates_output(self):
        """Test ColoredFormatter produces output."""
        formatter = ColoredFormatter('%(levelname)s: %(message)s')
        record = logging.LogRecord(
            name="test",
            level=logging.INFO,
            pathname="test.py",
            lineno=10,
            msg="Test message",
            args=(),
            exc_info=None
        )
        
        result = formatter.format(record)
        assert "Test message" in result
    
    def test_file_formatter_includes_timestamp(self):
        """Test FileFormatter includes timestamp."""
        formatter = FileFormatter(
            '%(asctime)s - %(levelname)s - %(message)s'
        )
        record = logging.LogRecord(
            name="test",
            level=logging.INFO,
            pathname="test.py",
            lineno=10,
            msg="Test message",
            args=(),
            exc_info=None
        )
        
        result = formatter.format(record)
        assert "Test message" in result
        # Should have timestamp (time format)
        assert ":" in result  # Time separator


@pytest.mark.unit
class TestConversionLogging:
    """Test conversion-specific logging functions."""
    
    def test_log_conversion_start(self, caplog):
        """Test logging conversion start."""
        logger = get_logger("test_conv_start")
        
        with caplog.at_level(logging.INFO):
            log_conversion_start(logger, "game.iso", "ISO", "CHD")
        
        assert "Starting conversion" in caplog.text
        assert "game.iso" in caplog.text
        assert "ISO → CHD" in caplog.text
    
    def test_log_conversion_complete(self, caplog):
        """Test logging conversion completion."""
        logger = get_logger("test_conv_complete")
        
        with caplog.at_level(logging.INFO):
            log_conversion_complete(
                logger,
                "game.chd",
                original_size=2_500_000_000,
                output_size=1_800_000_000,
                duration_seconds=45.5
            )
        
        assert "Conversion complete" in caplog.text
        assert "game.chd" in caplog.text
        assert "28.0%" in caplog.text  # Size reduction
        assert "45.5s" in caplog.text
    
    def test_log_conversion_error(self, caplog):
        """Test logging conversion error."""
        logger = get_logger("test_conv_error")
        
        with caplog.at_level(logging.ERROR):
            log_conversion_error(
                logger,
                "game.iso",
                "Out of disk space",
                returncode=1
            )
        
        assert "Conversion failed" in caplog.text
        assert "game.iso" in caplog.text
        assert "Out of disk space" in caplog.text
        assert "exit code: 1" in caplog.text
    
    def test_log_resource_status(self, caplog):
        """Test logging resource status."""
        logger = get_logger("test_resource_status")
        
        with caplog.at_level(logging.DEBUG):
            log_resource_status(
                logger,
                cpu_percent=85.5,
                ram_percent=72.3,
                disk_percent=65.0,
                active_workers=3
            )
        
        assert "CPU=85.5%" in caplog.text
        assert "RAM=72.3%" in caplog.text
        assert "Disk=65.0%" in caplog.text
        assert "Workers=3" in caplog.text


@pytest.mark.unit
class TestLogException:
    """Test exception logging."""
    
    def test_log_exception_captures_traceback(self, caplog):
        """Test that log_exception captures exception info."""
        logger = get_logger("test_exception")
        
        try:
            1 / 0
        except ZeroDivisionError as e:
            with caplog.at_level(logging.ERROR):
                log_exception(logger, e, "Math error")
        
        assert "Math error" in caplog.text
        assert "ZeroDivisionError" in caplog.text


@pytest.mark.unit
class TestGetLogger:
    """Test get_logger convenience function."""
    
    def test_get_logger_returns_logger(self):
        """Test get_logger returns a logger."""
        logger = get_logger("test_app")
        
        assert logger is not None
        assert logger.name == "test_app"
    
    def test_get_logger_different_names(self):
        """Test get_logger with different logger names."""
        logger1 = get_logger("app1")
        logger2 = get_logger("app2")
        
        assert logger1.name == "app1"
        assert logger2.name == "app2"
        assert logger1 is not logger2


@pytest.mark.unit
class TestLogLevels:
    """Test logging at different levels."""
    
    def test_debug_logging(self, caplog):
        """Test DEBUG level logging."""
        logger = get_logger("test_debug_level")
        logger.setLevel(logging.DEBUG)
        
        with caplog.at_level(logging.DEBUG):
            logger.debug("Debug message")
        
        assert "Debug message" in caplog.text
    
    def test_info_logging(self, caplog):
        """Test INFO level logging."""
        logger = get_logger("test_info_level")
        
        with caplog.at_level(logging.INFO):
            logger.info("Info message")
        
        assert "Info message" in caplog.text
    
    def test_warning_logging(self, caplog):
        """Test WARNING level logging."""
        logger = get_logger("test_warning_level")
        
        with caplog.at_level(logging.WARNING):
            logger.warning("Warning message")
        
        assert "Warning message" in caplog.text
    
    def test_error_logging(self, caplog):
        """Test ERROR level logging."""
        logger = get_logger("test_error_level")
        
        with caplog.at_level(logging.ERROR):
            logger.error("Error message")
        
        assert "Error message" in caplog.text
    
    def test_critical_logging(self, caplog):
        """Test CRITICAL level logging."""
        logger = get_logger("test_critical_level")
        
        with caplog.at_level(logging.CRITICAL):
            logger.critical("Critical message")
        
        assert "Critical message" in caplog.text


@pytest.mark.unit
class TestLoggingFunctionality:
    """Test general logging functionality."""
    
    def test_logger_exists(self):
        """Test that loggers can be created."""
        logger = get_logger("test_exists")
        assert logger is not None
        assert hasattr(logger, 'debug')
        assert hasattr(logger, 'info')
        assert hasattr(logger, 'warning')
        assert hasattr(logger, 'error')
        assert hasattr(logger, 'critical')
    
    def test_formatter_attributes(self):
        """Test formatter has required attributes."""
        formatter = ColoredFormatter('%(levelname)s - %(message)s')
        assert hasattr(formatter, 'format')
        assert hasattr(formatter, '_fmt')
    
    def test_file_formatter_attributes(self):
        """Test file formatter has required attributes."""
        formatter = FileFormatter('%(asctime)s - %(message)s')
        assert hasattr(formatter, 'format')
        assert hasattr(formatter, '_fmt')
