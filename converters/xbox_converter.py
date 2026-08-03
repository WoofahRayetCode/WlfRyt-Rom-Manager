"""
Xbox / Xbox 360 Converter

Handles extraction of Xbox and Xbox 360 ISO images.
"""

from pathlib import Path
from typing import Optional, List
import logging
import subprocess

from converters.base import BaseConverter, ConversionResult


class XboxConverter(BaseConverter):
    """Extract Xbox and Xbox 360 ISO disc images.
    
    Xbox games are distributed as ISO files. This converter extracts them
    to individual files using extract-xiso tool.
    
    Input format: .iso
    Output: Extracted files
    Tool: extract-xiso
    """

    def __init__(
        self,
        input_path: Path,
        extract_xiso_path: Optional[Path] = None,
        logger: Optional[logging.Logger] = None,
        log_callback=None,
    ):
        """Initialize Xbox converter.
        
        Args:
            input_path: Path to .iso file
            extract_xiso_path: Path to extract-xiso executable (optional)
            logger: Optional logger
            log_callback: Optional log callback for GUI
        """
        super().__init__(input_path, logger, log_callback)
        self.extract_xiso_path = Path(extract_xiso_path) if extract_xiso_path else None

    def can_convert(self) -> bool:
        """Check if input is a valid Xbox ISO file."""
        return (
            self.input_path.exists()
            and self.input_path.suffix.lower() == '.iso'
            and self.input_path.stat().st_size > 0
        )

    def get_output_formats(self) -> List[str]:
        """Get supported output formats for Xbox."""
        return ['EXTRACTED']

    def convert(self, output_format: str = 'EXTRACTED') -> ConversionResult:
        """Extract Xbox ISO image.
        
        Args:
            output_format: Must be 'EXTRACTED'
            
        Returns:
            ConversionResult with extraction status
        """
        if output_format.upper() != 'EXTRACTED':
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message="Xbox only supports extraction",
            )

        # Create output directory
        output_dir = self.input_path.parent / self.input_path.stem
        
        try:
            output_dir.mkdir(exist_ok=True)
            
            if self.extract_xiso_path and self.extract_xiso_path.exists():
                # Use extract-xiso tool if available
                cmd = [str(self.extract_xiso_path), str(self.input_path), str(output_dir)]
                
                self.log(f"Extracting Xbox ISO: {self.input_path.name}")
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
                        error_message=f"Extraction failed: {stderr}",
                        tool_used='extract-xiso',
                    )
                
                self.log(f"✅ Extracted Xbox ISO to: {output_dir.name}/")
                
                return ConversionResult(
                    success=True,
                    input_path=self.input_path,
                    output_path=output_dir,
                    original_size=self.input_path.stat().st_size,
                    tool_used='extract-xiso',
                )
            else:
                self.log(f"⚠️  extract-xiso not configured. Please extract manually: {self.input_path.name}")
                return ConversionResult(
                    success=False,
                    input_path=self.input_path,
                    error_message="extract-xiso not available",
                )

        except subprocess.TimeoutExpired:
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message="Extraction timeout",
                tool_used='extract-xiso',
            )
        except Exception as e:
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message=f"Extraction failed: {e}",
            )
