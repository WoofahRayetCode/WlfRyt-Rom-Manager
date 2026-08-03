"""
TAR Extractor

Handles extraction of TAR archives (TAR, TAR.GZ, TAR.BZ2, TAR.XZ).
"""

from pathlib import Path
from typing import Optional
import logging
import tarfile
import time

from extractors.base import BaseExtractor, ExtractionResult


class TarExtractor(BaseExtractor):
    """Extract TAR archives in various compression formats.
    
    Supports:
    - .tar (uncompressed)
    - .tar.gz / .tgz (gzip compressed)
    - .tar.bz2 / .tbz2 (bzip2 compressed)
    - .tar.xz / .txz (xz compressed)
    """

    def can_extract(self) -> bool:
        """Check if file is a valid TAR archive."""
        if not self.archive_path.exists():
            return False
        
        # Check filename
        name = self.archive_path.name.lower()
        if not any(name.endswith(ext) for ext in ['.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2', '.tar.xz', '.txz']):
            return False
        
        # Try to open as TAR
        try:
            with tarfile.open(self.archive_path, 'r:*') as tf:
                return True
        except Exception:
            return False

    def extract(self) -> ExtractionResult:
        """Extract TAR archive.
        
        Returns:
            ExtractionResult with extraction status
        """
        if not self.can_extract():
            return ExtractionResult(
                success=False,
                archive_path=self.archive_path,
                error_message="Invalid or corrupted TAR file",
            )

        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.log(f"📦 Extracting TAR: {self.archive_path.name}")
        
        start_time = time.time()
        file_count = 0
        total_size = 0

        try:
            with tarfile.open(self.archive_path, 'r:*') as tf:
                members = tf.getmembers()
                for member in members:
                    tf.extract(member, self.output_dir)
                    if member.isfile():
                        file_count += 1
                        total_size += member.size
            
            duration = time.time() - start_time
            self.log(f"✅ Extracted {file_count} files to {self.output_dir.name}/")
            
            return ExtractionResult(
                success=True,
                archive_path=self.archive_path,
                output_dir=self.output_dir,
                file_count=file_count,
                total_size=total_size,
                duration_seconds=duration,
                tool_used='tarfile',
            )

        except Exception as e:
            return ExtractionResult(
                success=False,
                archive_path=self.archive_path,
                error_message=f"Extraction failed: {e}",
                tool_used='tarfile',
            )

    def get_file_count(self) -> int:
        """Get number of files in TAR archive.
        
        Returns:
            Number of files, or -1 if unable to determine
        """
        try:
            with tarfile.open(self.archive_path, 'r:*') as tf:
                return len([m for m in tf.getmembers() if m.isfile()])
        except Exception:
            return -1
