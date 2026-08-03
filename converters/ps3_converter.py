"""
PlayStation 3 (PS3) Converter

Handles PS3 ISO decryption via ps3-disc-dumper tool.
"""

from pathlib import Path
from typing import Optional, List
import logging
import subprocess

from converters.base import BaseConverter, ConversionResult


class PS3Converter(BaseConverter):
    """Handle PS3 ISO decryption.
    
    PS3 games are encrypted. This converter launches ps3-disc-dumper GUI
    to allow users to decrypt ISOs interactively.
    
    Input format: .iso
    Output: Decrypted ISO (user decrypts via GUI tool)
    Tool: ps3-disc-dumper
    """

    def __init__(
        self,
        input_path: Path,
        ps3_dumper_path: Path,
        logger: Optional[logging.Logger] = None,
        log_callback=None,
    ):
        """Initialize PS3 converter.
        
        Args:
            input_path: Path to encrypted .iso file
            ps3_dumper_path: Path to ps3-disc-dumper executable
            logger: Optional logger
            log_callback: Optional log callback for GUI
        """
        super().__init__(input_path, logger, log_callback)
        self.ps3_dumper_path = Path(ps3_dumper_path)
        
        if not self.ps3_dumper_path.exists():
            raise ValueError(f"ps3-disc-dumper not found at {ps3_dumper_path}")

    def can_convert(self) -> bool:
        """Check if input is a valid PS3 ISO file."""
        return (
            self.input_path.exists()
            and self.input_path.suffix.lower() == '.iso'
            and self.input_path.stat().st_size > 0
        )

    def get_output_formats(self) -> List[str]:
        """Get supported output formats for PS3."""
        return ['ISO']

    def convert(self, output_format: str = 'ISO') -> ConversionResult:
        """Launch PS3 disc dumper for decryption.
        
        Args:
            output_format: Must be 'ISO'
            
        Returns:
            ConversionResult (user completes decryption in GUI)
        """
        if output_format.upper() != 'ISO':
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message=f"PS3 only supports ISO format",
            )

        self.log(f"🎮 Launching PS3 Disc Dumper for {self.input_path.name}...")
        
        try:
            # Launch dumper GUI (non-blocking)
            subprocess.Popen([str(self.ps3_dumper_path), str(self.input_path)])
            self.log("✅ PS3 Disc Dumper launched. Please complete decryption in the dumper window.")
            
            return ConversionResult(
                success=True,
                input_path=self.input_path,
                tool_used='ps3-disc-dumper',
            )
        except Exception as e:
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message=f"Failed to launch PS3 Disc Dumper: {e}",
                tool_used='ps3-disc-dumper',
            )
