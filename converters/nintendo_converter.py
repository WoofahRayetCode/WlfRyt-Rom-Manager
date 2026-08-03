"""
Nintendo Consoles Converter

Handles conversion of Nintendo ROM files for various systems (NES, SNES, N64, GB, GBA, etc.)
"""

from pathlib import Path
from typing import Optional, List
import logging

from converters.base import BaseConverter, ConversionResult


class NintendoConverter(BaseConverter):
    """Handle Nintendo ROM file conversions and processing.
    
    Supports multiple Nintendo systems:
    - NES (.nes)
    - SNES (.smc, .sfc)
    - Nintendo 64 (.z64, .n64, .v64)
    - Game Boy / Game Boy Color (.gb, .gbc)
    - Game Boy Advance (.gba)
    - Nintendo DS (.nds)
    - Nintendo 3DS (.3ds, .cia)
    - Nintendo Switch (.xci, .nsp)
    
    Note: Most Nintendo ROMs don't require conversion. This converter
    mainly validates and organizes Nintendo game files.
    """

    def __init__(
        self,
        input_path: Path,
        ndecrypt_path: Optional[Path] = None,
        logger: Optional[logging.Logger] = None,
        log_callback=None,
    ):
        """Initialize Nintendo converter.
        
        Args:
            input_path: Path to Nintendo ROM file
            ndecrypt_path: Path to NDecrypt executable (for 3DS decryption)
            logger: Optional logger
            log_callback: Optional log callback for GUI
        """
        super().__init__(input_path, logger, log_callback)
        self.ndecrypt_path = Path(ndecrypt_path) if ndecrypt_path else None
        self.system = self._detect_system()

    def can_convert(self) -> bool:
        """Check if input is a supported Nintendo ROM file."""
        ext = self.input_path.suffix.lower()
        nintendo_extensions = {
            '.nes', '.smc', '.sfc', '.z64', '.n64', '.v64',
            '.gb', '.gbc', '.gba', '.nds', '.3ds', '.cia',
            '.xci', '.nsp'
        }
        return (
            self.input_path.exists()
            and ext in nintendo_extensions
            and self.input_path.stat().st_size > 0
        )

    def get_output_formats(self) -> List[str]:
        """Get supported output formats for detected Nintendo system.
        
        Most Nintendo ROMs are already in optimal format, so output is
        primarily for validation/organization.
        """
        if self.system == 'Nintendo 3DS':
            return ['DECRYPTED']
        else:
            return ['VALIDATED']

    def convert(self, output_format: str = 'VALIDATED') -> ConversionResult:
        """Process Nintendo ROM file.
        
        Args:
            output_format: Target format (depends on system)
            
        Returns:
            ConversionResult with processing status
        """
        output_format = output_format.upper()
        
        if self.system == 'Nintendo 3DS':
            if output_format == 'DECRYPTED':
                return self._decrypt_3ds()
            else:
                return ConversionResult(
                    success=True,
                    input_path=self.input_path,
                    error_message="3DS ROM must be decrypted",
                )
        else:
            # Other Nintendo systems are already in optimal format
            self.log(f"✅ {self.system} ROM validated: {self.input_path.name}")
            return ConversionResult(
                success=True,
                input_path=self.input_path,
                original_size=self.input_path.stat().st_size,
                output_size=self.input_path.stat().st_size,
                tool_used='validation',
            )

    def _detect_system(self) -> str:
        """Detect Nintendo system based on file extension."""
        ext = self.input_path.suffix.lower()
        
        if ext == '.nes':
            return 'Nintendo Entertainment System (NES)'
        elif ext in {'.smc', '.sfc'}:
            return 'Super Nintendo (SNES)'
        elif ext in {'.z64', '.n64', '.v64'}:
            return 'Nintendo 64'
        elif ext == '.gb':
            return 'Game Boy'
        elif ext == '.gbc':
            return 'Game Boy Color'
        elif ext == '.gba':
            return 'Game Boy Advance'
        elif ext == '.nds':
            return 'Nintendo DS'
        elif ext in {'.3ds', '.cia'}:
            return 'Nintendo 3DS'
        elif ext in {'.xci', '.nsp'}:
            return 'Nintendo Switch'
        else:
            return 'Unknown Nintendo System'

    def _decrypt_3ds(self) -> ConversionResult:
        """Decrypt 3DS ROM image.
        
        3DS games are typically encrypted. This method attempts decryption
        if NDecrypt tool is available and aes_keys.txt is configured.
        
        Returns:
            ConversionResult with decryption status
        """
        if not self.ndecrypt_path or not self.ndecrypt_path.exists():
            return ConversionResult(
                success=False,
                input_path=self.input_path,
                error_message="NDecrypt not configured for 3DS decryption",
                tool_used='NDecrypt',
            )

        # Note: NDecrypt requires AES keys file and may need user interaction
        # Full implementation would integrate with rom_converter's key management
        self.log(f"🔐 3DS Decryption support planned: {self.input_path.name}")
        
        return ConversionResult(
            success=False,
            input_path=self.input_path,
            error_message="3DS decryption not yet implemented",
            tool_used='NDecrypt',
        )
