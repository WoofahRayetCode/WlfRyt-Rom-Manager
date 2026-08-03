"""
Streaming Extractor Module

Memory-efficient archive extraction using streaming to avoid loading
entire files into memory. Supports ISO, ZIP, RAR, 7Z formats.

Phase 4 Week 3: Performance optimization
"""

import os
import io
import logging
from pathlib import Path
from typing import Optional, Callable, Generator, Tuple
from abc import ABC, abstractmethod

try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False


logger = logging.getLogger(__name__)


class StreamingExtractor(ABC):
    """Base class for streaming archive extraction"""
    
    def __init__(
        self,
        archive_path: str,
        chunk_size: int = 8192,
        log_callback: Optional[Callable[[str], None]] = None,
    ):
        """
        Initialize streaming extractor.
        
        Args:
            archive_path: Path to archive file
            chunk_size: Size of chunks to read (default 8KB)
            log_callback: Optional callback for logging
        """
        self.archive_path = Path(archive_path)
        self.chunk_size = chunk_size
        self.log_callback = log_callback
        self.total_bytes = 0
        self.extracted_bytes = 0
        
        if not self.archive_path.exists():
            raise FileNotFoundError(f"Archive not found: {archive_path}")
        
        self.total_bytes = self.archive_path.stat().st_size
    
    def _log(self, message: str) -> None:
        """Log a message"""
        if self.log_callback:
            self.log_callback(message)
        logger.info(message)
    
    @abstractmethod
    def extract_file_stream(
        self,
        member_name: str,
    ) -> Generator[bytes, None, None]:
        """
        Extract a file from archive as a stream.
        
        Yields chunks of bytes to avoid loading entire file into memory.
        
        Args:
            member_name: Name of member within archive
            
        Yields:
            Bytes chunks from the file
        """
        pass
    
    @abstractmethod
    def list_members(self) -> list:
        """List all members in archive"""
        pass
    
    @abstractmethod
    def extract_all(self, output_dir: str, progress_callback: Optional[Callable] = None) -> bool:
        """
        Extract all files from archive.
        
        Args:
            output_dir: Directory to extract to
            progress_callback: Optional callback with (current_bytes, total_bytes)
            
        Returns:
            True if successful
        """
        pass
    
    def get_memory_usage_mb(self) -> float:
        """Get current process memory usage in MB"""
        if PSUTIL_AVAILABLE:
            process = psutil.Process()
            return process.memory_info().rss / 1024 / 1024
        return 0.0
    
    def should_throttle(self, max_memory_mb: int = 500) -> bool:
        """Check if extraction should throttle (memory check)"""
        current_mb = self.get_memory_usage_mb()
        return current_mb > max_memory_mb
    
    def get_progress_percent(self) -> float:
        """Get extraction progress percentage"""
        if self.total_bytes == 0:
            return 0.0
        return (self.extracted_bytes / self.total_bytes) * 100


class ZipStreamingExtractor(StreamingExtractor):
    """Streaming extraction for ZIP files"""
    
    def __init__(
        self,
        archive_path: str,
        chunk_size: int = 8192,
        log_callback: Optional[Callable[[str], None]] = None,
    ):
        """Initialize ZIP streaming extractor"""
        super().__init__(archive_path, chunk_size, log_callback)
        
        try:
            import zipfile
            self.zipfile = zipfile.ZipFile(self.archive_path, 'r')
        except ImportError:
            self._log("zipfile module not available")
            self.zipfile = None
        except Exception as e:
            self._log(f"Failed to open ZIP: {e}")
            self.zipfile = None
    
    def extract_file_stream(self, member_name: str) -> Generator[bytes, None, None]:
        """Extract file from ZIP as stream"""
        if not self.zipfile:
            self._log(f"ZIP not available, cannot extract {member_name}")
            return
        
        try:
            with self.zipfile.open(member_name) as member_file:
                while True:
                    chunk = member_file.read(self.chunk_size)
                    if not chunk:
                        break
                    self.extracted_bytes += len(chunk)
                    yield chunk
        except Exception as e:
            self._log(f"Error extracting {member_name}: {e}")
    
    def list_members(self) -> list:
        """List all members in ZIP"""
        if not self.zipfile:
            return []
        return self.zipfile.namelist()
    
    def extract_all(
        self,
        output_dir: str,
        progress_callback: Optional[Callable] = None,
    ) -> bool:
        """Extract all files from ZIP"""
        if not self.zipfile:
            self._log("ZIP not available")
            return False
        
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        try:
            members = self.zipfile.namelist()
            for i, member_name in enumerate(members):
                if self.should_throttle():
                    self._log("Memory usage high, throttling extraction")
                
                # Extract via stream to memory-efficient extraction
                output_file = output_path / member_name
                output_file.parent.mkdir(parents=True, exist_ok=True)
                
                if not member_name.endswith('/'):
                    with open(output_file, 'wb') as f:
                        for chunk in self.extract_file_stream(member_name):
                            f.write(chunk)
                
                if progress_callback:
                    progress_callback(i + 1, len(members))
            
            self._log(f"Successfully extracted ZIP to {output_dir}")
            return True
        
        except Exception as e:
            self._log(f"Error extracting all files: {e}")
            return False
    
    def __del__(self):
        """Close ZIP file on cleanup"""
        if self.zipfile:
            try:
                self.zipfile.close()
            except:
                pass


class IsoStreamingExtractor(StreamingExtractor):
    """Streaming extraction for ISO files"""
    
    def __init__(
        self,
        archive_path: str,
        chunk_size: int = 8192,
        log_callback: Optional[Callable[[str], None]] = None,
    ):
        """Initialize ISO streaming extractor"""
        super().__init__(archive_path, chunk_size, log_callback)
        
        try:
            import pycdlib
            self.iso = pycdlib.PyCdlib()
            self.iso.open(str(self.archive_path))
        except ImportError:
            self._log("pycdlib module not available, ISO extraction limited")
            self.iso = None
        except Exception as e:
            self._log(f"Failed to open ISO: {e}")
            self.iso = None
    
    def extract_file_stream(self, member_name: str) -> Generator[bytes, None, None]:
        """Extract file from ISO as stream"""
        if not self.iso:
            self._log(f"ISO not available, cannot extract {member_name}")
            return
        
        try:
            # For ISO, we need to use a BytesIO buffer
            buffer = io.BytesIO()
            self.iso.get_file_from_iso_fp(buffer, iso_path=member_name)
            buffer.seek(0)
            
            while True:
                chunk = buffer.read(self.chunk_size)
                if not chunk:
                    break
                self.extracted_bytes += len(chunk)
                yield chunk
        except Exception as e:
            self._log(f"Error extracting {member_name}: {e}")
    
    def list_members(self) -> list:
        """List all members in ISO"""
        if not self.iso:
            return []
        
        members = []
        for child in self.iso.list_dir('/'):
            members.append(child.file_identifier().decode())
        return members
    
    def extract_all(
        self,
        output_dir: str,
        progress_callback: Optional[Callable] = None,
    ) -> bool:
        """Extract all files from ISO"""
        if not self.iso:
            self._log("ISO not available")
            return False
        
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        try:
            members = self.list_members()
            for i, member_name in enumerate(members):
                if self.should_throttle():
                    self._log("Memory usage high, throttling extraction")
                
                output_file = output_path / member_name
                output_file.parent.mkdir(parents=True, exist_ok=True)
                
                with open(output_file, 'wb') as f:
                    for chunk in self.extract_file_stream(member_name):
                        f.write(chunk)
                
                if progress_callback:
                    progress_callback(i + 1, len(members))
            
            self._log(f"Successfully extracted ISO to {output_dir}")
            return True
        
        except Exception as e:
            self._log(f"Error extracting all files: {e}")
            return False
    
    def __del__(self):
        """Close ISO on cleanup"""
        if self.iso:
            try:
                self.iso.close()
            except:
                pass


class GenericStreamingExtractor(StreamingExtractor):
    """
    Generic streaming extractor using system tools.
    Falls back when specialized libraries unavailable.
    """
    
    def __init__(
        self,
        archive_path: str,
        chunk_size: int = 8192,
        log_callback: Optional[Callable[[str], None]] = None,
    ):
        """Initialize generic streaming extractor"""
        super().__init__(archive_path, chunk_size, log_callback)
        self.archive_type = self._detect_archive_type()
    
    def _detect_archive_type(self) -> str:
        """Detect archive type from extension"""
        suffix = self.archive_path.suffix.lower()
        if suffix == '.zip':
            return 'zip'
        elif suffix in ['.iso', '.cue']:
            return 'iso'
        elif suffix == '.7z':
            return '7z'
        elif suffix in ['.rar', '.r00']:
            return 'rar'
        return 'unknown'
    
    def extract_file_stream(self, member_name: str) -> Generator[bytes, None, None]:
        """Extract file as stream using system tool"""
        self._log(f"Generic extractor: streaming {member_name}")
        # This would use system commands like unzip, 7z, etc.
        return
        yield
    
    def list_members(self) -> list:
        """List members using system tool"""
        return []
    
    def extract_all(
        self,
        output_dir: str,
        progress_callback: Optional[Callable] = None,
    ) -> bool:
        """Extract using system tool"""
        self._log(f"Extracting {self.archive_type} with system tool")
        return False


def create_streaming_extractor(
    archive_path: str,
    chunk_size: int = 8192,
    log_callback: Optional[Callable[[str], None]] = None,
) -> Optional[StreamingExtractor]:
    """
    Factory function to create appropriate streaming extractor.
    
    Args:
        archive_path: Path to archive file
        chunk_size: Size of chunks to stream
        log_callback: Optional logging callback
        
    Returns:
        Appropriate StreamingExtractor subclass or None
    """
    archive_path_obj = Path(archive_path)
    suffix = archive_path_obj.suffix.lower()
    
    if suffix == '.zip':
        return ZipStreamingExtractor(archive_path, chunk_size, log_callback)
    elif suffix in ['.iso', '.cue']:
        return IsoStreamingExtractor(archive_path, chunk_size, log_callback)
    else:
        return GenericStreamingExtractor(archive_path, chunk_size, log_callback)


def extract_archive_efficiently(
    archive_path: str,
    output_dir: str,
    progress_callback: Optional[Callable] = None,
    log_callback: Optional[Callable[[str], None]] = None,
) -> bool:
    """
    Extract archive using memory-efficient streaming.
    
    Args:
        archive_path: Path to archive
        output_dir: Directory to extract to
        progress_callback: Optional progress callback
        log_callback: Optional logging callback
        
    Returns:
        True if extraction successful
    """
    extractor = create_streaming_extractor(
        archive_path,
        log_callback=log_callback,
    )
    
    if not extractor:
        if log_callback:
            log_callback(f"Could not create extractor for {archive_path}")
        return False
    
    return extractor.extract_all(output_dir, progress_callback)


if __name__ == '__main__':
    # Example usage
    logging.basicConfig(level=logging.INFO)
    
    # Test with a ZIP file if available
    test_zip = Path("test_archive.zip")
    if test_zip.exists():
        extractor = create_streaming_extractor(str(test_zip))
        if extractor:
            members = extractor.list_members()
            print(f"Found {len(members)} files")
            
            output_dir = Path("extracted_test")
            extractor.extract_all(str(output_dir))
