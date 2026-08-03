"""
PlayStation 2 (PS2) Converter

Handles conversion of PS2 ISO disc images to multiple formats (CHD, CSO, ZSO).
"""

from pathlib import Path
from typing import Optional, List
import logging
import time
import subprocess

from converters.base import BaseConverter, ConversionResult


class PS2Converter(BaseConverter):
    """Convert PS2 ISO disc images to CHD/CSO/ZSO formats.
    
    PS2 games are typically distributed as ISO files (disc images).
    This converter can create CHD (compressed), CSO, or ZSO formats.
    
    Input formats: .iso
    Output formats: .chd, .cso, .zso
    Tools: chdman (for CHD), maxcso (for CSO/ZSO)
    """

    def __init__(
        self,
        input_path: Path,
        chdman_path: Optional[Path] = None,
        maxcso_path: Optional[Path] = None,
        chdman_max_processors: int = 1,
        maxcso_threads: int = 4,
        logger: Optional[logging.Logger] = None,
        log_callback=None,
    ):
        """Initialize PS2 converter.
        
        Args:
            input_path: Path to .iso file
            chdman_path: Path to chdman executable (for CHD conversion)
            maxcso_path: Path to maxcso executable (for CSO/ZSO conversion)
            chdman_max_processors: Max processors for chdman
            maxcso_threads: Number of threads for maxcso
            logger: Optional logger
            log_callback: Optional log callback for GUI
            
        Raises:
            ValueError: If input is not .iso or tools not found
        """
        super().__init__(input_path, logger, log_callback)
        
        if not input_path.suffix.lower() == '.iso':
            raise ValueError(f"PS2Converter requires .iso file, got {input_path.suffix}")
        
        self.chdman_path = Path(chdman_path) if chdman_path else None
        self.maxcso_path = Path(maxcso_path) if maxcso_path else None
        self.chdman_max_processors = max(1, min(chdman_max_processors, 4))
        self.maxcso_threads = max(1, maxcso_threads)

    def can_convert(self) -> bool:
        """Check if input is a valid PS2 ISO file.
        
        Returns:
            True if input exists and is .iso format
        """
        return (
            self.input_path.exists()
            and self.input_path.suffix.lower() == '.iso'
            and self.input_path.stat().st_size > 0
        )

    def get_output_formats(self) -> List[str]:
        """Get supported output formats for PS2.
        
        Returns:
            List of supported formats (CHD, CSO, ZSO depending on tools available)
        """
        formats = []
        if self.chdman_path and self.chdman_path.exists():
            formats.append('CHD')
        if self.maxcso_path and self.maxcso_path.exists():
            formats.extend(['CSO', 'ZSO'])
        return formats if formats else ['CHD', 'CSO', 'ZSO']

    def convert(self, output_format: str = 'CHD') -> ConversionResult:
        """Convert PS2 ISO to specified format.
        
        Args:
            output_format: Target format (CHD, CSO, or ZSO)
            
        Returns:
            ConversionResult with conversion status
        """
        output_format = output_format.upper()
        
        if output_format == 'CHD':
            return self._convert_to_chd()
        elif output_format == 'CSO':
            return self._convert_to_cso()
        elif output_format == 'ZSO':
            return self._convert_to_zso()
        else:
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message=f"Unsupported PS2 format: {output_format}",
            )

    def _convert_to_chd(self) -> ConversionResult:
        """Convert PS2 ISO to CHD format.
        
        Returns:
            ConversionResult with conversion status
        """
        if not self.chdman_path or not self.chdman_path.exists():
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message="chdman not available",
            )

        output_path = self.input_path.with_suffix('.chd')
        
        # Check if output already exists
        if output_path.exists():
            self.log(f"⚠️  CHD already exists, skipping: {output_path.name}")
            return ConversionResult(
                success=True,
                input_path=self.input_path,
                output_path=output_path,
                original_size=self.input_path.stat().st_size,
                output_size=output_path.stat().st_size,
                tool_used='chdman',
            )

        original_size = self.input_path.stat().st_size
        
        # Build command: chdman createdvd -np <procs> -i input.iso -o output.chd
        cmd = [
            str(self.chdman_path),
            'createdvd',
            '-np', str(self.cdhman_max_processors),
            '-i', str(self.input_path),
            '-o', str(output_path),
        ]

        self.log(f"Converting (ISO → CHD): {self.input_path.name}")
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
                    error_message=f"chdman failed: {stderr}",
                    tool_used='chdman',
                )
            
            duration = time.time() - start_time
            output_size = output_path.stat().st_size
            
            self.log(
                f"✅ Converted to CHD: {output_path.name} "
                f"({output_size / original_size:.1%}, {duration:.1f}s)"
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
                error_message="Conversion timeout",
                tool_used='chdman',
            )
        except Exception as e:
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message=f"Conversion failed: {e}",
                tool_used='chdman',
            )

    def _convert_to_cso(self) -> ConversionResult:
        """Convert PS2 ISO to CSO (compressed ISO) format.
        
        Returns:
            ConversionResult with conversion status
        """
        if not self.maxcso_path or not self.maxcso_path.exists():
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message="maxcso not available",
            )

        output_path = self.input_path.with_suffix('.cso')
        
        # Check if output already exists
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
        
        # Build command: maxcso --threads <n> input.iso -o output.cso
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
            
            self.log(
                f"✅ Converted to CSO: {output_path.name} "
                f"({output_size / original_size:.1%}, {duration:.1f}s)"
            )
            
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
        """Convert PS2 ISO to ZSO (compressed ISO with zlib) format.
        
        Returns:
            ConversionResult with conversion status
        """
        if not self.maxcso_path or not self.maxcso_path.exists():
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message="maxcso not available",
            )

        output_path = self.input_path.with_suffix('.zso')
        
        # Check if output already exists
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
        
        # Build command: maxcso --ziso --threads <n> input.iso -o output.zso
        cmd = [
            str(self.maxcso_path),
            '--ziso',
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
            
            self.log(
                f"✅ Converted to ZSO: {output_path.name} "
                f"({output_size / original_size:.1%}, {duration:.1f}s)"
            )
            
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
