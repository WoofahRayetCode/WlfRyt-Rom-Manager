"""
Base Extractor Class

Abstract base class for all archive extraction implementations.
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, List
import logging


@dataclass
class ExtractionResult:
    """Result of an extraction operation."""
    success: bool
    archive_path: Path
    output_dir: Optional[Path] = None
    file_count: int = 0
    total_size: int = 0
    duration_seconds: float = 0.0
    error_message: Optional[str] = None
    tool_used: Optional[str] = None

    def __str__(self) -> str:
        if self.success:
            return (
                f"✅ Extracted {self.file_count} files to {self.output_dir.name} "
                f"({self.total_size / 1024 / 1024:.1f} MB, {self.duration_seconds:.1f}s)"
            )
        else:
            return f"❌ {self.archive_path.name}: {self.error_message}"


class BaseExtractor(ABC):
    """Base class for all archive extractors.
    
    Each extractor handles a specific archive format (ZIP, TAR, 7Z, etc.)
    """

    def __init__(
        self,
        archive_path: Path,
        output_dir: Optional[Path] = None,
        logger: Optional[logging.Logger] = None,
        log_callback=None,
    ):
        """Initialize extractor.
        
        Args:
            archive_path: Path to archive file
            output_dir: Output directory (defaults to archive parent/stem)
            logger: Optional logging.Logger instance
            log_callback: Optional callback for logging
        """
        self.archive_path = Path(archive_path)
        self.output_dir = output_dir or (archive_path.parent / archive_path.stem)
        self.logger = logger
        self.log_callback = log_callback

    def log(self, message: str) -> None:
        """Log a message."""
        if self.logger:
            self.logger.info(message)
        if self.log_callback:
            self.log_callback(message)

    @abstractmethod
    def can_extract(self) -> bool:
        """Check if extractor can handle this file.
        
        Returns:
            True if extractor can process this file
        """
        pass

    @abstractmethod
    def extract(self) -> ExtractionResult:
        """Perform the extraction.
        
        Returns:
            ExtractionResult with success status and details
        """
        pass

    @abstractmethod
    def get_file_count(self) -> int:
        """Get number of files in archive without extracting.
        
        Returns:
            Number of files, or -1 if unable to determine
        """
        pass
