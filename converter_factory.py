"""
Factory for creating and dispatching converters based on file type and system
"""

from pathlib import Path
from typing import Optional, Callable
from converters import (
    PS1Converter, PS2Converter, PS3Converter, PSPConverter, 
    XboxConverter, NintendoConverter, BaseConverter
)


class ConverterFactory:
    """Factory for creating appropriate converter instances based on file type and user preferences"""
    
    def __init__(self, 
                 chdman_path: Optional[Path] = None,
                 maxcso_path: Optional[Path] = None,
                 ps3_dumper_path: Optional[Path] = None,
                 extract_xiso_path: Optional[Path] = None,
                 ndecrypt_path: Optional[Path] = None,
                 log_callback: Optional[Callable[[str], None]] = None):
        """
        Initialize ConverterFactory with tool paths
        
        Args:
            chdman_path: Path to chdman executable
            maxcso_path: Path to maxcso executable
            ps3_dumper_path: Path to ps3-disc-dumper executable
            extract_xiso_path: Path to extract-xiso executable
            ndecrypt_path: Path to NDecrypt executable
            log_callback: Optional logging callback function
        """
        self.chdman_path = chdman_path
        self.maxcso_path = maxcso_path
        self.ps3_dumper_path = ps3_dumper_path
        self.extract_xiso_path = extract_xiso_path
        self.ndecrypt_path = ndecrypt_path
        self.log_callback = log_callback
    
    def _log(self, message: str) -> None:
        """Log a message if callback is provided"""
        if self.log_callback:
            self.log_callback(message)
    
    def get_converter(self, file_path: Path, output_format: str, **kwargs) -> Optional[BaseConverter]:
        """
        Get appropriate converter for the given file based on extension and preferences
        
        Args:
            file_path: Path to file to convert
            output_format: Desired output format (CHD, CSO, ZSO, ISO, etc.)
            **kwargs: Additional arguments for converter (process_ps1_cues, process_ps2_cues, etc.)
        
        Returns:
            Converter instance or None if file type not supported
        """
        ext = file_path.suffix.lower()
        
        # PlayStation 1 CUE files
        if ext == '.cue':
            if not self.chdman_path:
                self._log(f"chdman not configured for: {file_path.name}")
                return None
            return PS1Converter(
                input_path=file_path,
                chdman_path=self.chdman_path,
                log_callback=self.log_callback,
                **{k: v for k, v in kwargs.items() if k in ['max_processors']}
            )
        
        # CHD files (reverse conversion - extract to ISO/BIN/CUE)
        elif ext == '.chd':
            if not self.chdman_path:
                self._log(f"chdman not configured for: {file_path.name}")
                return None
            # Determine source system from preferences or file detection
            source_system = kwargs.get('source_system', 'auto')
            return PS1Converter(
                input_path=file_path,
                chdman_path=self.chdman_path,
                log_callback=self.log_callback,
                **{k: v for k, v in kwargs.items() if k in ['max_processors']}
            )
        
        # ISO files (multi-system detection)
        elif ext == '.iso':
            iso_size = file_path.stat().st_size
            system_guess = kwargs.get('system_guess', 'PS2')
            
            # PlayStation 3 detection and processing
            if system_guess == 'PlayStation 3' or kwargs.get('force_ps3', False):
                if not self.ps3_dumper_path:
                    self._log(f"PS3 Dumper not configured for: {file_path.name}")
                    return None
                return PS3Converter(
                    input_path=file_path,
                    ps3_dumper_path=self.ps3_dumper_path,
                    log_callback=self.log_callback
                )
            
            # Xbox detection and processing
            if system_guess == 'Xbox' or kwargs.get('force_xbox', False):
                if not self.extract_xiso_path:
                    self._log(f"extract-xiso not configured for: {file_path.name}")
                    return None
                return XboxConverter(
                    input_path=file_path,
                    extract_xiso_path=self.extract_xiso_path,
                    log_callback=self.log_callback
                )
            
            # Determine PSP vs PS2 based on size or explicit preference
            treat_as_psp = kwargs.get('treat_as_psp', False)
            if treat_as_psp or (iso_size < 1.4e9 and not kwargs.get('force_ps2', False)):
                # PSP ISO conversion
                if not self.maxcso_path:
                    self._log(f"maxcso not configured for PSP ISO: {file_path.name}")
                    return None
                return PSPConverter(
                    input_path=file_path,
                    maxcso_path=self.maxcso_path,
                    output_format=output_format,
                    log_callback=self.log_callback,
                    **{k: v for k, v in kwargs.items() if k in ['max_threads']}
                )
            else:
                # PS2 ISO conversion
                if output_format == 'CHD' and not self.chdman_path:
                    self._log(f"chdman not configured for PS2 ISO to CHD: {file_path.name}")
                    return None
                if output_format in ['CSO', 'ZSO'] and not self.maxcso_path:
                    self._log(f"maxcso not configured for PS2 ISO to {output_format}: {file_path.name}")
                    return None
                
                return PS2Converter(
                    input_path=file_path,
                    chdman_path=self.chdman_path,
                    maxcso_path=self.maxcso_path,
                    output_format=output_format,
                    log_callback=self.log_callback,
                    **{k: v for k, v in kwargs.items() if k in ['max_processors', 'max_threads']}
                )
        
        # Nintendo ROM detection and processing
        elif ext in ['.gb', '.gbc', '.gba', '.nds', '.3ds', '.cia', '.nes', '.sfc', '.smc', '.snes', '.n64', '.z64', '.v64']:
            return NintendoConverter(
                input_path=file_path,
                ndecrypt_path=self.ndecrypt_path if ext in ['.3ds', '.cia'] else None,
                log_callback=self.log_callback
            )
        
        else:
            self._log(f"Unsupported file type: {ext}")
            return None
    
    def convert(self, file_path: Path, output_format: str, **kwargs) -> bool:
        """
        Convert a file to the specified output format
        
        Args:
            file_path: Path to file to convert
            output_format: Desired output format
            **kwargs: Additional arguments for converter selection
        
        Returns:
            True if conversion successful, False otherwise
        """
        converter = self.get_converter(file_path, output_format, **kwargs)
        if not converter:
            self._log(f"Could not create converter for: {file_path.name}")
            return False
        
        result = converter.convert(output_format)
        return result.success
    
    def batch_convert(self, file_paths: list[Path], output_format: str, 
                      on_complete: Optional[Callable] = None, **kwargs) -> dict:
        """
        Convert multiple files to the specified output format
        
        Args:
            file_paths: List of file paths to convert
            output_format: Desired output format for all files
            on_complete: Optional callback called with (file_path, success) after each conversion
            **kwargs: Additional arguments for converter selection
        
        Returns:
            Dictionary with results: {'total': int, 'successful': int, 'failed': int}
        """
        results = {'total': len(file_paths), 'successful': 0, 'failed': 0}
        
        for file_path in file_paths:
            success = self.convert(file_path, output_format, **kwargs)
            if success:
                results['successful'] += 1
            else:
                results['failed'] += 1
            
            if on_complete:
                on_complete(file_path, success)
        
        return results
