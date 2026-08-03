"""
PlayStation 1 (PS1) Converter

Handles conversion of PS1 CUE/BIN disc images to CHD (MAME Compressed Hunks of Data) format.
"""

from pathlib import Path
from typing import Optional, List
import logging
import time
import subprocess

from converters.base import BaseConverter, ConversionResult


class PS1Converter(BaseConverter):
    """Convert PS1 CUE/BIN disc images to CHD format.
    
    PS1 games are typically distributed as CUE files (cue sheet) with associated BIN files
    (raw binary data). This converter creates CHD files which are compressed disc images.
    
    Input formats: .cue (with .bin files)
    Output format: .chd
    Tool: chdman (from MAME toolset)
    """

    def __init__(
        self,
        input_path: Path,
        chdman_path: Path,
        max_processors: int = 1,
        logger: Optional[logging.Logger] = None,
        log_callback=None,
    ):
        """Initialize PS1 converter.
        
        Args:
            input_path: Path to .cue file
            chdman_path: Path to chdman executable
            max_processors: Max processors for chdman (limits RAM usage)
            logger: Optional logger
            log_callback: Optional log callback for GUI
            
        Raises:
            ValueError: If chdman_path doesn't exist or input is not .cue
        """
        super().__init__(input_path, logger, log_callback)
        
        if not input_path.suffix.lower() == '.cue':
            raise ValueError(f"PS1Converter requires .cue file, got {input_path.suffix}")
        
        self.chdman_path = Path(chdman_path)
        if not self.chdman_path.exists():
            raise ValueError(f"chdman not found at {chdman_path}")
        
        self.max_processors = max(1, min(max_processors, 4))  # Clamp to 1-4

    def can_convert(self) -> bool:
        """Check if input is a valid PS1 CUE file.
        
        Returns:
            True if input exists and is .cue format
        """
        return (
            self.input_path.exists()
            and self.input_path.suffix.lower() == '.cue'
            and self.input_path.stat().st_size > 0
        )

    def get_output_formats(self) -> List[str]:
        """Get supported output formats for PS1.
        
        Returns:
            List of supported formats
        """
        return ['CHD']

    def convert(self, output_format: str = 'CHD') -> ConversionResult:
        """Convert PS1 CUE/BIN to CHD format.
        
        Args:
            output_format: Target format (must be 'CHD')
            
        Returns:
            ConversionResult with conversion status
        """
        if output_format.upper() != 'CHD':
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message=f"PS1 only supports CHD format, got {output_format}",
            )

        output_path = self.input_path.with_suffix('.chd')

        # Check if output already exists
        if output_path.exists():
            self.log(f"⚠️  CHD already exists, skipping: {output_path.name}")
            return ConversionResult(
                success=True,
                input_path=self.input_path,
                output_path=output_path,
                original_size=output_path.stat().st_size,
                output_size=output_path.stat().st_size,
                tool_used='chdman',
            )

        # Calculate original size (CUE + all BIN files)
        try:
            bin_files = self._get_bin_files()
            original_size = sum(f.stat().st_size for f in bin_files)
            original_size += self.input_path.stat().st_size
        except Exception as e:
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message=f"Failed to calculate size: {e}",
            )

        # Build conversion command
        cmd = [
            str(self.chdman_path),
            'createcd',
            '-np', str(self.max_processors),  # Number of processors
            '-i', str(self.input_path),       # Input CUE file
            '-o', str(output_path),           # Output CHD file
        ]

        self.log(f"Converting (CD → CHD): {self.input_path.name}")
        start_time = time.time()

        try:
            # Run chdman with output capture for progress reporting
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            
            stdout, stderr = process.communicate(timeout=3600)
            
            if process.returncode != 0:
                return ConversionResult(
                    success=False,
                    input_path=self.input_path,
                    error_message=f"chdman failed: {stderr}",
                    tool_used='chdman',
                )
            
            # Conversion successful
            duration = time.time() - start_time
            output_size = output_path.stat().st_size
            
            self.log(
                f"✅ Converted to CHD: {self.input_path.name} → {output_path.name} "
                f"({output_size / original_size:.1%} compression, {duration:.1f}s)"
            )
            
            return ConversionResult(
                success=True,
                input_path=self.input_path,
                output_path=output_path,
                original_size=original_size,
                output_size=output_size,
                duration_seconds=duration,
                tool_used='chdman',
            )

        except subprocess.TimeoutExpired:
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message="Conversion timeout (exceeded 1 hour)",
                tool_used='chdman',
            )
        except Exception as e:
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message=f"Conversion failed: {e}",
                tool_used='chdman',
            )

    def _get_bin_files(self) -> List[Path]:
        """Get all BIN files referenced by the CUE sheet.
        
        Returns:
            List of Path objects for BIN files
        """
        bin_files = []
        try:
            with open(self.input_path, 'r', encoding='utf-8', errors='ignore') as f:
                for line in f:
                    line = line.strip()
                    if line.lower().startswith('file'):
                        # Parse: FILE "filename" BINARY
                        parts = line.split('"')
                        if len(parts) >= 2:
                            filename = parts[1]
                            bin_path = self.input_path.parent / filename
                            if bin_path.exists():
                                bin_files.append(bin_path)
        except Exception as e:
            self.log(f"⚠️  Failed to parse CUE file: {e}")
        
        return bin_files
