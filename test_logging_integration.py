"""
Unit tests for logging_integration module.

Tests the bridge between UI logging and structured logging system.
"""

import pytest
import logging
from logging_integration import (
    setup_converter_logging,
    get_converter_logger,
    ConversionLogger,
    init,
    info,
    warning,
    error,
    debug,
    exception,
)


@pytest.mark.unit
class TestLoggingIntegrationSetup:
    """Test logging initialization."""
    
    def test_setup_converter_logging_returns_logger(self):
        """Test setup_converter_logging creates a logger."""
        logger = get_converter_logger("test_setup")
        
        assert logger is not None
        assert logger.name == "test_setup"
    
    def test_get_converter_logger_default(self):
        """Test get_converter_logger with default name."""
        logger = get_converter_logger()
        
        assert logger is not None
        assert logger.name == "rom_converter"


@pytest.mark.unit
class TestConversionLogger:
    """Test ConversionLogger convenience wrapper."""
    
    def test_conversion_logger_creation(self):
        """Test creating a ConversionLogger."""
        logger = get_converter_logger("test_conv")
        conv_logger = ConversionLogger(logger)
        
        assert conv_logger is not None
        assert conv_logger.logger is logger
    
    def test_conversion_logger_default_logger(self):
        """Test ConversionLogger with default logger."""
        conv_logger = ConversionLogger()
        
        assert conv_logger.logger is not None
    
    def test_conversion_logger_start(self, caplog):
        """Test logging conversion start."""
        logger = get_converter_logger("test_start")
        conv_logger = ConversionLogger(logger)
        
        with caplog.at_level(logging.INFO):
            conv_logger.start("game.iso", "ISO", "CHD")
        
        assert "Starting conversion" in caplog.text
        assert "game.iso" in caplog.text
    
    def test_conversion_logger_complete(self, caplog):
        """Test logging conversion completion."""
        logger = get_converter_logger("test_complete")
        conv_logger = ConversionLogger(logger)
        
        with caplog.at_level(logging.INFO):
            conv_logger.complete("game.chd", 2_500_000_000, 1_800_000_000, 45.5)
        
        assert "Conversion complete" in caplog.text
        assert "28.0%" in caplog.text  # Size reduction
    
    def test_conversion_logger_error(self, caplog):
        """Test logging conversion error."""
        logger = get_converter_logger("test_error")
        conv_logger = ConversionLogger(logger)
        
        with caplog.at_level(logging.ERROR):
            conv_logger.error("game.iso", "Out of disk space", 1)
        
        assert "Conversion failed" in caplog.text
        assert "game.iso" in caplog.text
    
    def test_conversion_logger_debug(self, caplog):
        """Test debug logging with context."""
        logger = get_converter_logger("test_debug")
        conv_logger = ConversionLogger(logger)
        
        with caplog.at_level(logging.DEBUG):
            conv_logger.debug("Processing file", file="game.iso", format="ISO")
        
        assert "Processing file" in caplog.text
        assert "file=game.iso" in caplog.text
        assert "format=ISO" in caplog.text
    
    def test_conversion_logger_info(self, caplog):
        """Test info logging with context."""
        logger = get_converter_logger("test_info")
        conv_logger = ConversionLogger(logger)
        
        with caplog.at_level(logging.INFO):
            conv_logger.info("Batch started", count=5, total=20)
        
        assert "Batch started" in caplog.text
        assert "count=5" in caplog.text
        assert "total=20" in caplog.text
    
    def test_conversion_logger_warning(self, caplog):
        """Test warning logging."""
        logger = get_converter_logger("test_warning")
        conv_logger = ConversionLogger(logger)
        
        with caplog.at_level(logging.WARNING):
            conv_logger.warning("Low disk space", available_gb=2)
        
        assert "Low disk space" in caplog.text
        assert "available_gb=2" in caplog.text


@pytest.mark.unit
class TestModuleLogFunctions:
    """Test module-level logging functions."""
    
    def test_init_function(self):
        """Test init() function."""
        logger = init()
        
        assert logger is not None
    
    def test_info_function(self, caplog):
        """Test info() function."""
        init()
        
        with caplog.at_level(logging.INFO):
            info("Test message")
        
        assert "Test message" in caplog.text
    
    def test_warning_function(self, caplog):
        """Test warning() function."""
        init()
        
        with caplog.at_level(logging.WARNING):
            warning("Test warning")
        
        assert "Test warning" in caplog.text
    
    def test_error_function(self, caplog):
        """Test error() function."""
        init()
        
        with caplog.at_level(logging.ERROR):
            error("Test error")
        
        assert "Test error" in caplog.text
    
    def test_debug_function(self, caplog):
        """Test debug() function."""
        init()
        
        with caplog.at_level(logging.DEBUG):
            debug("Test debug")
        
        assert "Test debug" in caplog.text
    
    def test_exception_function(self, caplog):
        """Test exception() function."""
        init()
        
        try:
            raise ValueError("Test error")
        except ValueError:
            with caplog.at_level(logging.ERROR):
                exception("Caught exception")
        
        assert "Caught exception" in caplog.text
        assert "ValueError" in caplog.text


@pytest.mark.unit
def test_logging_integration_documentation():
    """Test that logging integration module has proper documentation."""
    from logging_integration import setup_converter_logging, ConversionLogger, init
    
    # Verify docstrings exist
    assert setup_converter_logging.__doc__ is not None
    assert ConversionLogger.__doc__ is not None
    assert ConversionLogger.start.__doc__ is not None
    assert ConversionLogger.complete.__doc__ is not None
    assert ConversionLogger.error.__doc__ is not None
    assert init.__doc__ is not None
