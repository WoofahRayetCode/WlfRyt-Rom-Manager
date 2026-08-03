"""
Registry for creating and dispatching extractors based on archive format
"""

from pathlib import Path
from typing import Optional, Tuple, Callable
from extractors import ZipExtractor, TarExtractor, SevenZipExtractor, BaseExtractor, ExtractionResult


class ExtractionRegistry:
    """Registry for creating appropriate extractor instances based on archive format"""
    
    # Map of file extensions to extractor classes
    EXTRACTORS = {
        '.zip': ZipExtractor,
        '.tar': TarExtractor,
        '.tar.gz': TarExtractor,
        '.tgz': TarExtractor,
        '.tar.bz2': TarExtractor,
        '.tbz2': TarExtractor,
        '.tar.xz': TarExtractor,
        '.txz': TarExtractor,
        '.7z': SevenZipExtractor,
        '.rar': SevenZipExtractor,
    }
    
    def __init__(self, seven_zip_path: Optional[Path] = None,
                 log_callback: Optional[Callable[[str], None]] = None):
        """
        Initialize ExtractionRegistry with tool paths
        
        Args:
            seven_zip_path: Path to 7-Zip executable (required for 7Z/RAR)
            log_callback: Optional logging callback function
        """
        self.seven_zip_path = seven_zip_path
        self.log_callback = log_callback
    
    def _log(self, message: str) -> None:
        """Log a message if callback is provided"""
        if self.log_callback:
            self.log_callback(message)
    
    def _get_file_format(self, archive_path: Path) -> str:
        """
        Detect archive format from file extension and name
        
        Args:
            archive_path: Path to archive file
        
        Returns:
            File format string (e.g., '.zip', '.tar.gz')
        """
        name = archive_path.name.lower()
        
        # Check multi-part extensions first
        if name.endswith('.tar.gz'):
            return '.tar.gz'
        elif name.endswith('.tar.bz2'):
            return '.tar.bz2'
        elif name.endswith('.tar.xz'):
            return '.tar.xz'
        elif name.endswith('.tgz'):
            return '.tgz'
        elif name.endswith('.tbz2'):
            return '.tbz2'
        elif name.endswith('.txz'):
            return '.txz'
        else:
            # Single extension
            return archive_path.suffix.lower()
    
    def is_supported(self, archive_path: Path) -> bool:
        """
        Check if archive format is supported
        
        Args:
            archive_path: Path to archive file
        
        Returns:
            True if supported, False otherwise
        """
        fmt = self._get_file_format(archive_path)
        return fmt in self.EXTRACTORS
    
    def get_extractor(self, archive_path: Path, 
                      output_dir: Optional[Path] = None) -> Optional[BaseExtractor]:
        """
        Get appropriate extractor for the given archive
        
        Args:
            archive_path: Path to archive file
            output_dir: Optional output directory (defaults to archive stem)
        
        Returns:
            Extractor instance or None if format not supported
        """
        fmt = self._get_file_format(archive_path)
        
        if fmt not in self.EXTRACTORS:
            self._log(f"Unsupported archive format: {fmt}")
            return None
        
        extractor_class = self.EXTRACTORS[fmt]
        
        # Special handling for 7Z/RAR (requires external tool)
        if fmt in ['.7z', '.rar']:
            if not self.seven_zip_path:
                self._log(f"7-Zip not configured, cannot extract {fmt}")
                return None
            return extractor_class(
                archive_path=archive_path,
                output_dir=output_dir,
                seven_zip_path=self.seven_zip_path,
                log_callback=self.log_callback
            )
        
        # Native extractors (ZIP, TAR)
        return extractor_class(
            archive_path=archive_path,
            output_dir=output_dir,
            log_callback=self.log_callback
        )
    
    def extract(self, archive_path: Path, 
                output_dir: Optional[Path] = None) -> Tuple[bool, Optional[Path]]:
        """
        Extract an archive to the specified output directory
        
        Args:
            archive_path: Path to archive file
            output_dir: Optional output directory (defaults to archive stem)
        
        Returns:
            Tuple of (success: bool, output_dir: Path or None)
        """
        # Determine output directory
        if output_dir is None:
            archive_path = Path(archive_path)
            stem = archive_path.stem
            
            # Handle multi-part extensions (e.g., archive.tar.gz -> archive)
            name_lower = archive_path.name.lower()
            if name_lower.endswith('.tar.gz') or name_lower.endswith('.tgz'):
                stem = archive_path.name.replace('.tar.gz', '').replace('.tgz', '')
            elif name_lower.endswith('.tar.bz2') or name_lower.endswith('.tbz2'):
                stem = archive_path.name.replace('.tar.bz2', '').replace('.tbz2', '')
            elif name_lower.endswith('.tar.xz') or name_lower.endswith('.txz'):
                stem = archive_path.name.replace('.tar.xz', '').replace('.txz', '')
            
            output_dir = archive_path.parent / stem
        
        # Get appropriate extractor
        extractor = self.get_extractor(archive_path, output_dir)
        if not extractor:
            return False, None
        
        # Extract archive
        try:
            result = extractor.extract()
            if result.success:
                return True, result.output_dir
            else:
                self._log(f"Extraction failed for: {archive_path.name}")
                return False, None
        except Exception as e:
            self._log(f"Extraction error: {e}")
            return False, None
    
    def extract_all(self, directory: Path, recursive: bool = True,
                    delete_archives: bool = False,
                    on_extract: Optional[Callable[[Path, bool], None]] = None) -> list[Path]:
        """
        Find and extract all supported archives in directory
        
        Args:
            directory: Root directory to search
            recursive: Whether to search recursively
            delete_archives: Whether to delete archives after successful extraction
            on_extract: Optional callback called with (archive_path, success) after each extraction
        
        Returns:
            List of directories that were extracted
        """
        directory = Path(directory)
        extracted_dirs = []
        
        # Find all supported archives
        archives = self._find_archives(directory, recursive)
        
        if not archives:
            self._log("No supported archives found")
            return extracted_dirs
        
        self._log(f"Found {len(archives)} archive(s) to extract:")
        for archive in archives:
            self._log(f"  - {archive.name}")
        self._log("")
        
        # Extract each archive
        for archive_path in archives:
            success, output_dir = self.extract(archive_path)
            
            if success and output_dir:
                extracted_dirs.append(output_dir)
                self._log(f"Extracted to: {output_dir.name}/")
                
                # Delete archive if requested
                if delete_archives:
                    try:
                        archive_path.unlink()
                        self._log(f"Deleted archive: {archive_path.name}")
                    except Exception as e:
                        self._log(f"Could not delete archive: {e}")
            else:
                self._log(f"Failed to extract: {archive_path.name}")
            
            if on_extract:
                on_extract(archive_path, success)
        
        return extracted_dirs
    
    def _find_archives(self, directory: Path, recursive: bool = True) -> list[Path]:
        """
        Find all supported archives in directory
        
        Args:
            directory: Root directory to search
            recursive: Whether to search recursively
        
        Returns:
            List of archive paths
        """
        directory = Path(directory)
        archives = []
        
        if not directory.exists():
            return archives
        
        # Build glob pattern
        pattern = '**/*' if recursive else '*'
        
        # Search for each supported extension
        for ext in self.EXTRACTORS.keys():
            # Handle special multi-part extensions
            if ext in ['.tar.gz', '.tar.bz2', '.tar.xz', '.tgz', '.tbz2', '.txz']:
                search_pattern = f"{pattern}{ext.split('.')[0]}*{'.'.join(ext.split('.')[1:])}"
            else:
                search_pattern = f"{pattern}{ext}"
            
            archives.extend(directory.glob(search_pattern))
        
        # Remove duplicates and sort
        return sorted(set(archives))
