"""
ZIP Extractor

Handles extraction of ZIP archives using Python's built-in zipfile module.
"""

from pathlib import Path
from typing import Optional
import logging
import zipfile
import time

from extractors.base import BaseExtractor, ExtractionResult


class ZipExtractor(BaseExtractor):
    """Extract ZIP archives.
    
    Uses Python's built-in zipfile module for maximum compatibility.
    Supports standard ZIP and most common variants.
    """

    def can_extract(self) -> bool:
        """Check if file is a valid ZIP archive."""
        if not self.archive_path.exists() or self.archive_path.suffix.lower() != '.zip':
            return False
        
        try:
            with zipfile.ZipFile(self.archive_path, 'r') as zf:
                return zf.testzip() is None
        except Exception:
            return False

    def extract(self) -> ExtractionResult:
        """Extract ZIP archive.
        
        Returns:
            ExtractionResult with extraction status
        """
        if not self.can_extract():
            return ExtractionResult(
                success=False,
                archive_path=self.archive_path,
                error_message="Invalid or corrupted ZIP file",
            )

        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.log(f"📦 Extracting ZIP: {self.archive_path.name}")
        
        start_time = time.time()
        file_count = 0
        total_size = 0

        try:
            with zipfile.ZipFile(self.archive_path, 'r') as zf:
                for info in zf.infolist():
                    zf.extract(info, self.output_dir)
                    file_count += 1
                    total_size += info.file_size
            
            duration = time.time() - start_time
            self.log(f"✅ Extracted {file_count} files to {self.output_dir.name}/")
            
            return ExtractionResult(
                success=True,
                archive_path=self.archive_path,
                output_dir=self.output_dir,
                file_count=file_count,
                total_size=total_size,
                duration_seconds=duration,
                tool_used='zipfile',
            )

        except Exception as e:
            return ExtractionResult(
                success=False,
                archive_path=self.archive_path,
                error_message=f"Extraction failed: {e}",
                tool_used='zipfile',
            )

    def get_file_count(self) -> int:
        """Get number of files in ZIP archive.
        
        Returns:
            Number of files, or -1 if unable to determine
        """
        try:
            with zipfile.ZipFile(self.archive_path, 'r') as zf:
                return len(zf.infolist())
        except Exception:
            return -1
