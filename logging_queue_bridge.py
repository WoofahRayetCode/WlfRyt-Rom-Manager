"""
Logging Queue Bridge - Adapter between rom_converter's queue-based logging and structured logging.

This module provides an adapter that allows the existing rom_converter.py logging queue
to feed into the structured logging system from logging_setup.py without requiring 
immediate replacement of all logging calls.

Architecture:
    rom_converter.py: self.log_queue.put(message_string)
         ↓
    log_queue_listener thread (new) - reads strings from queue
         ↓
    QueueLogHandler (new) - converts strings to proper log levels
         ↓
    logging_setup.py structured loggers
         ↓
    File + Console output (structured)

This maintains dual-mode operation: old log queue works, but output is structured.
"""

import logging
import threading
from queue import Queue, Empty
from typing import Optional, Tuple
from pathlib import Path
import re

from logging_setup import (
    setup_logging,
    get_logger,
    cleanup_logging,
)


class QueueLogHandler(logging.Handler):
    """Handler that processes emoji-prefixed messages from rom_converter's log queue.
    
    Parses emoji prefixes to determine log level:
    - ✓/✅ → INFO (success)
    - ⚠️/❌ → WARNING (alert)
    - ❌/ERROR → ERROR (failure)
    - 🔄 → DEBUG (processing)
    - → INFO (default, no emoji)
    
    Attributes:
        structured_logger: The target logger from logging_setup
    """
    
    def __init__(self, structured_logger: logging.Logger):
        """Initialize the queue handler.
        
        Args:
            structured_logger: Target logger from logging_setup to write to
        """
        super().__init__()
        self.structured_logger = structured_logger
        
        # Emoji to log level mapping
        self.emoji_patterns = {
            r'^✓': logging.INFO,      # Success
            r'^✅': logging.INFO,      # Success
            r'^⚠': logging.WARNING,    # Alert
            r'^❌': logging.ERROR,      # Error
            r'^🔄': logging.DEBUG,      # Processing
        }
    
    def emit(self, record: logging.LogRecord) -> None:
        """Process a log record from the queue.
        
        Parses emoji prefix, determines appropriate log level, and writes to
        structured logger. Strips emoji from message for cleaner file logging.
        
        Args:
            record: Log record from queue
        """
        try:
            msg = record.getMessage()
            level = logging.INFO  # Default
            
            # Check for emoji prefix and map to level
            for pattern, mapped_level in self.emoji_patterns.items():
                if re.match(pattern, msg):
                    level = mapped_level
                    break
            
            # Strip emoji prefix for cleaner file logging
            clean_msg = re.sub(r'^[✓✅⚠❌🔄]\s*', '', msg).strip()
            
            # Write to structured logger at detected level
            self.structured_logger.log(level, clean_msg)
            
        except Exception:
            self.handleError(record)


def _queue_listener_thread(
    log_queue: Queue,
    structured_logger: logging.Logger,
    stop_event: threading.Event,
) -> None:
    """Background thread that processes messages from the log queue.
    
    Reads string messages from rom_converter's log queue and processes them
    through the handler.
    
    Args:
        log_queue: Queue to read messages from
        structured_logger: Logger to write processed messages to
        stop_event: Threading event to signal shutdown
    """
    handler = QueueLogHandler(structured_logger)
    
    while not stop_event.is_set():
        try:
            # Non-blocking get with short timeout
            msg = log_queue.get(timeout=0.1)
            if msg is None:  # Sentinel value for shutdown
                break
            
            # Convert string message to LogRecord
            record = logging.LogRecord(
                name=structured_logger.name,
                level=logging.INFO,  # Will be re-mapped by handler
                pathname="<rom_converter>",
                lineno=0,
                msg=str(msg),
                args=(),
                exc_info=None,
            )
            
            # Process through handler
            handler.emit(record)
            
        except Empty:
            # No message available, continue loop
            continue
        except Exception as e:
            structured_logger.error(f"Error processing queue message: {e}")


def setup_queue_logging_bridge(
    log_queue: Queue,
    app_name: str = "ROM Manager",
    log_dir: Optional[Path] = None,
) -> Tuple[logging.Logger, threading.Thread]:
    """Set up the queue logging bridge.
    
    Creates structured logging infrastructure and starts a background thread
    that processes messages from rom_converter's log_queue.
    
    Dual-mode operation:
    - rom_converter continues calling self.log_queue.put(message)
    - Messages are processed in background thread and written to structured logs
    - Both console (with colors) and file (structured) output
    
    Args:
        log_queue: The Queue object from rom_converter.log_queue
        app_name: Application name for log formatting
        log_dir: Directory for log files (default: ./logs)
    
    Returns:
        Tuple of (structured_logger, listener_thread)
        - structured_logger: Logger for direct use
        - listener_thread: Background thread processing queue (caller can join)
    
    Raises:
        ValueError: If log_queue is not a Queue
    
    Example:
        In rom_converter.py __init__:
        
            from logging_queue_bridge import setup_queue_logging_bridge
            
            self.logger, self.listener_thread = setup_queue_logging_bridge(
                self.log_queue,
                app_name="ROM Manager"
            )
            
            # rom_converter continues using self.log_queue.put() normally
            # Output is now structured (file + console)
            
            # On shutdown:
            # self.log_queue.put(None)  # Sentinel
            # self.listener_thread.join(timeout=5)
    """
    if not isinstance(log_queue, Queue):
        raise ValueError(f"Expected Queue, got {type(log_queue)}")
    
    # Set up structured logging
    structured_logger = setup_logging(app_name=app_name, log_dir=log_dir)
    
    # Create and start the queue listener thread
    stop_event = threading.Event()
    listener_thread = threading.Thread(
        target=_queue_listener_thread,
        args=(log_queue, structured_logger, stop_event),
        daemon=True,
        name="LogQueueListener",
    )
    listener_thread.start()
    
    # Return logger and thread for caller to manage
    return structured_logger, listener_thread


def get_or_create_queue_logger(
    name: str,
    parent_logger: logging.Logger,
) -> logging.Logger:
    """Get or create a child logger under the queue-integrated logger.
    
    Args:
        name: Logger name (e.g., "rom_converter.converters.chdman")
        parent_logger: The parent logger from setup_queue_logging_bridge
    
    Returns:
        A logger with structured output
    """
    return parent_logger.getChild(name)
