"""
7Z/RAR Extractor

Handles extraction of 7Z and RAR archives via external 7-Zip tool.
"""

from pathlib import Path
from typing import Optional
import logging
import subprocess
import time

from extractors.base import BaseExtractor, ExtractionResult


class SevenZipExtractor(BaseExtractor):
    """Extract 7Z and RAR archives via 7-Zip.
    
    Supports:
    - .7z (7-Zip format)
    - .rar / .rar5 (RAR archives)
    - Legacy formats (.ace, .arj, .cab, .chm, .cpio, .deb, .dmg, .iso, .lzh, .lzma, .msi, .nsis, .rpm, .udf, .wim, .xar, .z)
    
    Requires 7-Zip command-line tool to be installed and in PATH.
    """

    def __init__(
        self,
        archive_path: Path,
        seven_zip_path: Optional[Path] = None,
        output_dir: Optional[Path] = None,
        logger: Optional[logging.Logger] = None,
        log_callback=None,
    ):
        """Initialize 7Z/RAR extractor.
        
        Args:
            archive_path: Path to archive file
            seven_zip_path: Path to 7z executable (defaults to searching PATH)
            output_dir: Output directory
            logger: Optional logger
            log_callback: Optional log callback
            
        Raises:
            ValueError: If 7-Zip tool not found
        """
        super().__init__(archive_path, output_dir, logger, log_callback)
        self.seven_zip_path = seven_zip_path or self._find_7zip()
        
        if not self.seven_zip_path or not self.seven_zip_path.exists():
            raise ValueError("7-Zip not found. Please install 7-Zip or provide path to 7z executable.")

    @staticmethod
    def _find_7zip() -> Optional[Path]:
        """Find 7z executable in common locations."""
        # Try to find 7z in PATH via 'where' command
        try:
            result = subprocess.run(
                ['where', '7z'] if Path('C:\\').exists() else ['which', '7z'],
                capture_output=True,
                text=True,
            )
            if result.returncode == 0:
                return Path(result.stdout.strip())
        except Exception:
            pass
        
        # Common Windows paths
        common_paths = [
            Path('C:\\Program Files\\7-Zip\\7z.exe'),
            Path('C:\\Program Files (x86)\\7-Zip\\7z.exe'),
        ]
        for path in common_paths:
            if path.exists():
                return path
        
        return None

    def can_extract(self) -> bool:
        """Check if file is a supported 7Z/RAR archive."""
        if not self.archive_path.exists():
            return False
        
        ext = self.archive_path.suffix.lower()
        supported = {'.7z', '.rar', '.rar5'}
        return ext in supported

    def extract(self) -> ExtractionResult:
        """Extract 7Z/RAR archive.
        
        Returns:
            ExtractionResult with extraction status
        """
        if not self.can_extract():
            return ExtractionResult(
                success=False,
                archive_path=self.archive_path,
                error_message="Unsupported archive format for 7-Zip",
            )

        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.log(f"📦 Extracting {self.archive_path.suffix.upper()}: {self.archive_path.name}")
        
        # Build 7z command: 7z x -o<output> archive
        cmd = [
            str(self.seven_zip_path),
            'x',                           # Extract with full paths
            f'-o{self.output_dir}',        # Output directory
            str(self.archive_path),
        ]

        start_time = time.time()

        try:
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            
            stdout, stderr = process.communicate(timeout=3600)
            
            if process.returncode != 0:
                return ExtractionResult(
                    success=False,
                    archive_path=self.archive_path,
                    error_message=f"7-Zip extraction failed: {stderr}",
                    tool_used='7-Zip',
                )
            
            # Count extracted files
            file_count = self.get_file_count()
            duration = time.time() - start_time
            
            self.log(f"✅ Extracted to {self.output_dir.name}/")
            
            return ExtractionResult(
                success=True,
                archive_path=self.archive_path,
                output_dir=self.output_dir,
                file_count=file_count if file_count > 0 else -1,
                duration_seconds=duration,
                tool_used='7-Zip',
            )

        except subprocess.TimeoutExpired:
            return ExtractionResult(
                success=False,
                archive_path=self.archive_path,
                error_message="Extraction timeout (exceeded 1 hour)",
                tool_used='7-Zip',
            )
        except Exception as e:
            return ExtractionResult(
                success=False,
                archive_path=self.archive_path,
                error_message=f"Extraction failed: {e}",
                tool_used='7-Zip',
            )

    def get_file_count(self) -> int:
        """Get number of files in archive via 7z list command.
        
        Returns:
            Number of files, or -1 if unable to determine
        """
        try:
            cmd = [str(self.seven_zip_path), 'l', str(self.archive_path)]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            
            if result.returncode != 0:
                return -1
            
            # Count files from output (last line shows "Files: N")
            for line in result.stdout.split('\n'):
                if 'Files:' in line:
                    parts = line.split()
                    try:
                        return int(parts[-2])  # Number before closing paren
                    except (ValueError, IndexError):
                        pass
            return -1
        except Exception:
            return -1
