"""
Extractors Package

Modularized archive extraction implementations for multiple formats.

Supported Formats:
- ZIP (.zip)
- TAR (.tar, .tar.gz, .tgz, .tar.bz2, .tar.xz)
- 7-Zip (.7z)
- RAR (.rar, .rar5)
"""

from extractors.base import BaseExtractor, ExtractionResult
from extractors.zip_extractor import ZipExtractor
from extractors.tar_extractor import TarExtractor
from extractors.sevenzip_extractor import SevenZipExtractor

__all__ = [
    'BaseExtractor',
    'ExtractionResult',
    'ZipExtractor',
    'TarExtractor',
    'SevenZipExtractor',
]
