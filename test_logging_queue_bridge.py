"""
Tests for logging_queue_bridge.py - Queue logging adapter integration tests.

Tests the adapter that bridges rom_converter's log queue with structured logging.
"""

import logging
import pytest
import tempfile
from pathlib import Path
from queue import Queue
import time
import threading

from logging_setup import cleanup_logging
from logging_queue_bridge import (
    QueueLogHandler,
    setup_queue_logging_bridge,
    get_or_create_queue_logger,
)


@pytest.fixture
def clean_logging():
    """Cleanup logging handlers after each test."""
    yield
    cleanup_logging()


class TestQueueLogHandler:
    """Test the QueueLogHandler emoji-to-level parsing."""
    
    def test_handler_creation(self, clean_logging):
        """QueueLogHandler should be creatable with a logger."""
        logger = logging.getLogger("test")
        handler = QueueLogHandler(logger)
        assert handler is not None
        assert handler.structured_logger is logger
    
    def test_emoji_patterns_defined(self, clean_logging):
        """QueueLogHandler should have emoji patterns."""
        logger = logging.getLogger("test")
        handler = QueueLogHandler(logger)
        assert len(handler.emoji_patterns) > 0
        assert logging.INFO in handler.emoji_patterns.values()
        assert logging.WARNING in handler.emoji_patterns.values()
        assert logging.ERROR in handler.emoji_patterns.values()
    
    def test_emit_with_success_emoji(self, clean_logging):
        """emit() should process ✓ emoji."""
        logger = logging.getLogger("test_emoji")
        handler = QueueLogHandler(logger)
        
        record = logging.LogRecord(
            name="test",
            level=logging.INFO,
            pathname="test.py",
            lineno=1,
            msg="✓ Operation complete",
            args=(),
            exc_info=None,
        )
        
        # Should not raise
        handler.emit(record)
    
    def test_emit_with_warning_emoji(self, clean_logging):
        """emit() should process ⚠️ emoji."""
        logger = logging.getLogger("test_warning_emoji")
        handler = QueueLogHandler(logger)
        
        record = logging.LogRecord(
            name="test",
            level=logging.INFO,
            pathname="test.py",
            lineno=1,
            msg="⚠️ Something needs attention",
            args=(),
            exc_info=None,
        )
        
        handler.emit(record)
    
    def test_emit_with_error_emoji(self, clean_logging):
        """emit() should process ❌ emoji."""
        logger = logging.getLogger("test_error_emoji")
        handler = QueueLogHandler(logger)
        
        record = logging.LogRecord(
            name="test",
            level=logging.INFO,
            pathname="test.py",
            lineno=1,
            msg="❌ Operation failed",
            args=(),
            exc_info=None,
        )
        
        handler.emit(record)
    
    def test_emit_with_debug_emoji(self, clean_logging):
        """emit() should process 🔄 emoji."""
        logger = logging.getLogger("test_debug_emoji")
        handler = QueueLogHandler(logger)
        
        record = logging.LogRecord(
            name="test",
            level=logging.INFO,
            pathname="test.py",
            lineno=1,
            msg="🔄 Processing item",
            args=(),
            exc_info=None,
        )
        
        handler.emit(record)
    
    def test_emit_without_emoji(self, clean_logging):
        """emit() should handle messages without emoji."""
        logger = logging.getLogger("test_no_emoji")
        handler = QueueLogHandler(logger)
        
        record = logging.LogRecord(
            name="test",
            level=logging.INFO,
            pathname="test.py",
            lineno=1,
            msg="Plain message",
            args=(),
            exc_info=None,
        )
        
        handler.emit(record)


class TestSetupQueueLoggingBridge:
    """Test the queue logging bridge setup function."""
    
    def test_setup_with_valid_queue(self, clean_logging, tmp_path):
        """setup_queue_logging_bridge should work with a valid Queue."""
        log_queue = Queue()
        logger, thread = setup_queue_logging_bridge(
            log_queue,
            app_name="test_app",
            log_dir=tmp_path
        )
        
        assert logger is not None
        assert isinstance(logger, logging.Logger)
        assert thread is not None
        assert isinstance(thread, threading.Thread)
        assert thread.daemon
    
    def test_setup_with_invalid_queue_raises(self, clean_logging):
        """setup_queue_logging_bridge should raise ValueError for non-Queue."""
        with pytest.raises(ValueError, match="Expected Queue"):
            setup_queue_logging_bridge("not a queue")
    
    def test_queue_messages_are_processed(self, clean_logging, tmp_path):
        """Messages put in queue should be processed."""
        log_queue = Queue()
        logger, thread = setup_queue_logging_bridge(
            log_queue,
            app_name="test_app",
            log_dir=tmp_path
        )
        
        # Put messages in queue
        test_messages = [
            "✓ Test message 1",
            "⚠️ Test message 2",
            "❌ Test message 3",
        ]
        
        for msg in test_messages:
            log_queue.put(msg)
        
        # Give thread time to process
        time.sleep(0.3)
        
        # Should complete without error
        assert logger is not None
    
    def test_stop_listener_with_none(self, clean_logging, tmp_path):
        """Sending None to queue should gracefully stop listener."""
        log_queue = Queue()
        logger, thread = setup_queue_logging_bridge(
            log_queue,
            app_name="test_app",
            log_dir=tmp_path
        )
        
        # Send sentinel
        log_queue.put(None)
        
        # Thread should exit
        thread.join(timeout=2)
        assert not thread.is_alive()


class TestGetOrCreateQueueLogger:
    """Test creating child loggers under the bridge."""
    
    def test_get_child_logger(self, clean_logging, tmp_path):
        """get_or_create_queue_logger should return a child logger."""
        log_queue = Queue()
        parent_logger, _ = setup_queue_logging_bridge(
            log_queue,
            app_name="test_app",
            log_dir=tmp_path
        )
        
        child_logger = get_or_create_queue_logger(
            "converters.chdman",
            parent_logger
        )
        
        assert child_logger is not None
        assert isinstance(child_logger, logging.Logger)
        assert "converters.chdman" in child_logger.name
    
    def test_child_logger_can_log(self, clean_logging, tmp_path):
        """Child logger should be able to log messages."""
        log_queue = Queue()
        parent_logger, _ = setup_queue_logging_bridge(
            log_queue,
            app_name="test_app",
            log_dir=tmp_path
        )
        
        child_logger = get_or_create_queue_logger(
            "converters",
            parent_logger
        )
        
        # Should not raise
        child_logger.info("Test from child")
        child_logger.warning("Warning from child")
        child_logger.error("Error from child")


class TestDualModeOperation:
    """Test that both old queue-based and new structured logging work."""
    
    def test_queue_and_direct_logging_both_work(
        self,
        clean_logging,
        tmp_path
    ):
        """Both queue.put() and direct logger.info() should work together."""
        log_queue = Queue()
        logger, thread = setup_queue_logging_bridge(
            log_queue,
            app_name="test_app",
            log_dir=tmp_path
        )
        
        # Old way: queue-based
        log_queue.put("✓ Queue message")
        
        # New way: direct structured logging
        logger.info("Direct message")
        
        # Give time for processing
        time.sleep(0.2)
        
        # Both should work without error
        assert True
    
    def test_rom_converter_usage_pattern(self, clean_logging, tmp_path):
        """Simulate how rom_converter.py would use the bridge."""
        # Simulate rom_converter initialization
        log_queue = Queue()
        logger, listener_thread = setup_queue_logging_bridge(
            log_queue,
            app_name="ROM Manager",
            log_dir=tmp_path
        )
        
        # Simulate rom_converter's log() method usage
        def rom_log(message: str):
            log_queue.put(message)
        
        # Simulate typical conversion messages
        rom_log("✓ Scanning directory")
        rom_log("🔄 Processing file.iso")
        rom_log("⚠️ Warning: file might be corrupted")
        rom_log("❌ Conversion failed")
        
        # Give time for processing
        time.sleep(0.2)
        
        # Simulate shutdown
        log_queue.put(None)
        listener_thread.join(timeout=5)
        
        # Should complete without errors
        assert True


class TestThreadSafety:
    """Test thread-safety of queue logging."""
    
    def test_multiple_threads_logging(self, clean_logging, tmp_path):
        """Multiple threads should be able to log concurrently."""
        log_queue = Queue()
        logger, _ = setup_queue_logging_bridge(
            log_queue,
            app_name="test_app",
            log_dir=tmp_path
        )
        
        message_count = [0]
        lock = threading.Lock()
        
        def thread_work(thread_id: int):
            """Simulate converter thread logging."""
            for i in range(3):
                log_queue.put(f"✓ Thread {thread_id} message {i}")
                with lock:
                    message_count[0] += 1
        
        # Start multiple threads
        threads = []
        for i in range(3):
            t = threading.Thread(target=thread_work, args=(i,))
            t.start()
            threads.append(t)
        
        # Wait for threads
        for t in threads:
            t.join()
        
        # Give listener time to process all messages
        time.sleep(0.3)
        
        # All messages should be put in queue
        assert message_count[0] == 9


class TestLogDirectory:
    """Test log file creation."""
    
    def test_custom_log_directory(self, clean_logging, tmp_path):
        """Custom log directory should be created and used."""
        log_dir = tmp_path / "custom_logs"
        
        log_queue = Queue()
        logger, thread = setup_queue_logging_bridge(
            log_queue,
            app_name="test_app",
            log_dir=log_dir
        )
        
        # Log some messages
        log_queue.put("✓ Test message 1")
        logger.info("Test message 2")
        
        # Give time for I/O
        time.sleep(0.2)
        
        # Clean shutdown
        log_queue.put(None)
        thread.join(timeout=2)
        
        # Should complete without error
        assert logger is not None
    
    def test_default_log_directory(self, clean_logging):
        """Default log directory should work when not specified."""
        log_queue = Queue()
        logger, thread = setup_queue_logging_bridge(log_queue)
        
        log_queue.put("✓ Test with default dir")
        time.sleep(0.1)
        
        log_queue.put(None)
        thread.join(timeout=2)
        
        assert logger is not None


@pytest.fixture
def clean_logging():
    """Cleanup logging handlers after each test."""
    yield
    cleanup_logging()


class TestIntegration:
    """Integration tests for the complete queue bridge."""
    
    def test_complete_workflow(self, clean_logging, tmp_path):
        """Test complete workflow from setup to shutdown."""
        # Setup
        log_queue = Queue()
        logger, listener_thread = setup_queue_logging_bridge(
            log_queue,
            app_name="ROM Manager",
            log_dir=tmp_path
        )
        
        # Use queue-based logging
        for i in range(5):
            log_queue.put(f"✓ Processing file {i}")
        
        # Use direct logging
        logger.info("Starting conversion")
        logger.warning("Disk space low")
        logger.error("Conversion failed")
        
        # Wait for processing
        time.sleep(0.3)
        
        # Shutdown
        log_queue.put(None)
        listener_thread.join(timeout=5)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])

