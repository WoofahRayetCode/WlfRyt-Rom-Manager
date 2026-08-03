"""
PlayStation Portable (PSP) Converter

Handles conversion of PSP ISO disc images to CSO/ZSO (compressed) formats.
"""

from pathlib import Path
from typing import Optional, List
import logging
import time
import subprocess

from converters.base import BaseConverter, ConversionResult


class PSPConverter(BaseConverter):
    """Convert PSP ISO disc images to CSO/ZSO formats.
    
    PSP games are distributed as ISO files. This converter compresses them
    to CSO or ZSO format for reduced storage.
    
    Input format: .iso
    Output formats: .cso, .zso
    Tool: maxcso
    """

    def __init__(
        self,
        input_path: Path,
        maxcso_path: Path,
        maxcso_threads: int = 4,
        logger: Optional[logging.Logger] = None,
        log_callback=None,
    ):
        """Initialize PSP converter.
        
        Args:
            input_path: Path to .iso file
            maxcso_path: Path to maxcso executable
            maxcso_threads: Number of threads for maxcso
            logger: Optional logger
            log_callback: Optional log callback for GUI
        """
        super().__init__(input_path, logger, log_callback)
        self.maxcso_path = Path(maxcso_path)
        self.maxcso_threads = max(1, maxcso_threads)
        
        if not self.maxcso_path.exists():
            raise ValueError(f"maxcso not found at {maxcso_path}")

    def can_convert(self) -> bool:
        """Check if input is a valid PSP ISO file."""
        return (
            self.input_path.exists()
            and self.input_path.suffix.lower() == '.iso'
            and self.input_path.stat().st_size > 0
        )

    def get_output_formats(self) -> List[str]:
        """Get supported output formats for PSP."""
        return ['CSO', 'ZSO']

    def convert(self, output_format: str = 'CSO') -> ConversionResult:
        """Convert PSP ISO to CSO or ZSO format.
        
        Args:
            output_format: Target format (CSO or ZSO)
            
        Returns:
            ConversionResult with conversion status
        """
        output_format = output_format.upper()
        
        if output_format == 'CSO':
            return self._convert_to_cso()
        elif output_format == 'ZSO':
            return self._convert_to_zso()
        else:
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message=f"Unsupported PSP format: {output_format}",
            )

    def _convert_to_cso(self) -> ConversionResult:
        """Convert PSP ISO to CSO format."""
        output_path = self.input_path.with_suffix('.cso')
        
        if output_path.exists():
            self.log(f"⚠️  CSO already exists, skipping: {output_path.name}")
            return ConversionResult(
                success=True,
                input_path=self.input_path,
                output_path=output_path,
                original_size=self.input_path.stat().st_size,
                output_size=output_path.stat().st_size,
                tool_used='maxcso',
            )

        original_size = self.input_path.stat().st_size
        cmd = [
            str(self.maxcso_path),
            '--threads', str(self.maxcso_threads),
            str(self.input_path),
            '-o', str(output_path),
        ]

        self.log(f"Converting (ISO → CSO): {self.input_path.name}")
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
                return ConversionResult(
                    success=False,
                    input_path=self.input_path,
                    error_message=f"maxcso failed: {stderr}",
                    tool_used='maxcso',
                )
            
            duration = time.time() - start_time
            output_size = output_path.stat().st_size
            
            self.log(f"✅ Converted to CSO: {output_path.name} ({output_size / original_size:.1%}, {duration:.1f}s)")
            
            return ConversionResult(
                success=True,
                input_path=self.input_path,
                output_path=output_path,
                original_size=original_size,
                output_size=output_size,
                duration_seconds=duration,
                tool_used='maxcso',
            )

        except subprocess.TimeoutExpired:
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message="Conversion timeout",
                tool_used='maxcso',
            )
        except Exception as e:
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message=f"Conversion failed: {e}",
                tool_used='maxcso',
            )

    def _convert_to_zso(self) -> ConversionResult:
        """Convert PSP ISO to ZSO format."""
        output_path = self.input_path.with_suffix('.zso')
        
        if output_path.exists():
            self.log(f"⚠️  ZSO already exists, skipping: {output_path.name}")
            return ConversionResult(
                success=True,
                input_path=self.input_path,
                output_path=output_path,
                original_size=self.input_path.stat().st_size,
                output_size=output_path.stat().st_size,
                tool_used='maxcso',
            )

        original_size = self.input_path.stat().st_size
        cmd = [
            str(self.maxcso_path),
            '--zso',
            '--threads', str(self.maxcso_threads),
            str(self.input_path),
            '-o', str(output_path),
        ]

        self.log(f"Converting (ISO → ZSO): {self.input_path.name}")
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
                return ConversionResult(
                    success=False,
                    input_path=self.input_path,
                    error_message=f"maxcso failed: {stderr}",
                    tool_used='maxcso',
                )
            
            duration = time.time() - start_time
            output_size = output_path.stat().st_size
            
            self.log(f"✅ Converted to ZSO: {output_path.name} ({output_size / original_size:.1%}, {duration:.1f}s)")
            
            return ConversionResult(
                success=True,
                input_path=self.input_path,
                output_path=output_path,
                original_size=original_size,
                output_size=output_size,
                duration_seconds=duration,
                tool_used='maxcso',
            )

        except subprocess.TimeoutExpired:
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message="Conversion timeout",
                tool_used='maxcso',
            )
        except Exception as e:
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message=f"Conversion failed: {e}",
                tool_used='maxcso',
            )
