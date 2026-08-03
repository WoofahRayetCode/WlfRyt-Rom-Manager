"""
ROM Converters Package

Modularized converter implementations extracted from rom_converter.py.
Each converter handles a specific game system and format conversion.

Supported Systems:
- PlayStation (PS1): CUE/BIN → CHD
- PlayStation 2 (PS2): ISO → CHD/CSO/ZSO
- PlayStation 3 (PS3): ISO → Decryption via ps3-disc-dumper
- PSP: ISO → CSO/ZSO
- Xbox/Xbox 360: ISO → Extraction via extract-xiso
- GameCube/Wii: ISO/GCM → GCZ (via dolphin-tool)
- Nintendo Handhelds: GB/GBA/NDS/3DS → Various formats
"""

from converters.base import BaseConverter, ConversionResult
from converters.ps1_converter import PS1Converter
from converters.ps2_converter import PS2Converter
from converters.ps3_converter import PS3Converter
from converters.psp_converter import PSPConverter
from converters.xbox_converter import XboxConverter
from converters.nintendo_converter import NintendoConverter

__all__ = [
    'BaseConverter',
    'ConversionResult',
    'PS1Converter',
    'PS2Converter',
    'PS3Converter',
    'PSPConverter',
    'XboxConverter',
    'NintendoConverter',
]
