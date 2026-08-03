"""
Tests for Converters Package

Tests the modularized converter implementations.
"""

import pytest
from pathlib import Path
from unittest.mock import Mock, MagicMock, patch
import tempfile
import logging

from converters.base import BaseConverter, ConversionResult
from converters.ps1_converter import PS1Converter
from converters.ps2_converter import PS2Converter
from converters.ps3_converter import PS3Converter
from converters.psp_converter import PSPConverter
from converters.xbox_converter import XboxConverter
from converters.nintendo_converter import NintendoConverter


class TestConversionResult:
    """Test ConversionResult dataclass"""
    
    def test_successful_conversion(self):
        """Test successful conversion result."""
        result = ConversionResult(
            success=True,
            input_path=Path("game.iso"),
            output_path=Path("game.chd"),
            original_size=4_000_000_000,
            output_size=2_000_000_000,
            duration_seconds=120.5,
            tool_used='chdman',
        )
        
        assert result.success
        assert result.compression_ratio == 0.5
        assert "✅" in str(result)
        assert "50.0%" in str(result)

    def test_failed_conversion(self):
        """Test failed conversion result."""
        result = ConversionResult(
            success=False,
            input_path=Path("game.iso"),
            error_message="Disk full",
        )
        
        assert not result.success
        assert "❌" in str(result)
        assert "Disk full" in str(result)


class TestPS1Converter:
    """Test PlayStation 1 converter"""
    
    def test_ps1_converter_creation(self):
        """Test PS1Converter initialization."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            cue_file = tmpdir / "game.cue"
            cue_file.write_text("FILE \"game.bin\" BINARY\n")
            
            chdman = tmpdir / "chdman.exe"
            chdman.write_text("")
            
            converter = PS1Converter(
                input_path=cue_file,
                chdman_path=chdman,
            )
            
            assert converter.input_path == cue_file
            assert converter.chdman_path == chdman
            assert converter.can_convert()

    def test_ps1_requires_cue_file(self):
        """Test PS1 converter rejects non-CUE files."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            iso_file = tmpdir / "game.iso"
            iso_file.write_text("")
            
            chdman = tmpdir / "chdman.exe"
            chdman.write_text("")
            
            with pytest.raises(ValueError, match="requires .cue file"):
                PS1Converter(iso_file, chdman)

    def test_ps1_output_formats(self):
        """Test PS1 supported output formats."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            cue_file = tmpdir / "game.cue"
            cue_file.write_text("")
            
            chdman = tmpdir / "chdman.exe"
            chdman.write_text("")
            
            converter = PS1Converter(cue_file, chdman)
            assert converter.get_output_formats() == ['CHD']


class TestPS2Converter:
    """Test PlayStation 2 converter"""
    
    def test_ps2_converter_creation(self):
        """Test PS2Converter initialization."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            iso_file = tmpdir / "game.iso"
            iso_file.write_bytes(b'\x00' * 1000)  # Create non-empty file
            
            chdman = tmpdir / "chdman.exe"
            chdman.write_text("")
            
            converter = PS2Converter(
                input_path=iso_file,
                chdman_path=chdman,
            )
            
            assert converter.input_path == iso_file
            assert converter.can_convert()

    def test_ps2_output_formats(self):
        """Test PS2 supported output formats."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            iso_file = tmpdir / "game.iso"
            iso_file.write_bytes(b'\x00' * 1000)  # Create non-empty file
            
            chdman = tmpdir / "chdman.exe"
            chdman.write_text("")
            
            maxcso = tmpdir / "maxcso.exe"
            maxcso.write_text("")
            
            converter = PS2Converter(
                iso_file,
                chdman_path=chdman,
                maxcso_path=maxcso,
            )
            
            formats = converter.get_output_formats()
            assert 'CHD' in formats
            assert 'CSO' in formats
            assert 'ZSO' in formats


class TestPS3Converter:
    """Test PlayStation 3 converter"""
    
    def test_ps3_converter_creation(self):
        """Test PS3Converter initialization."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            iso_file = tmpdir / "game.iso"
            iso_file.write_bytes(b'\x00' * 1000)  # Create non-empty file
            
            dumper = tmpdir / "ps3-disc-dumper.exe"
            dumper.write_text("")
            
            converter = PS3Converter(
                input_path=iso_file,
                ps3_dumper_path=dumper,
            )
            
            assert converter.input_path == iso_file
            assert converter.can_convert()


class TestPSPConverter:
    """Test PSP converter"""
    
    def test_psp_converter_creation(self):
        """Test PSPConverter initialization."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            iso_file = tmpdir / "game.iso"
            iso_file.write_bytes(b'\x00' * 1000)  # Create non-empty file
            
            maxcso = tmpdir / "maxcso.exe"
            maxcso.write_text("")
            
            converter = PSPConverter(
                input_path=iso_file,
                maxcso_path=maxcso,
            )
            
            assert converter.input_path == iso_file
            assert converter.can_convert()

    def test_psp_output_formats(self):
        """Test PSP supported output formats."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            iso_file = tmpdir / "game.iso"
            iso_file.write_bytes(b'\x00' * 1000)  # Create non-empty file
            
            maxcso = tmpdir / "maxcso.exe"
            maxcso.write_text("")
            
            converter = PSPConverter(iso_file, maxcso)
            assert converter.get_output_formats() == ['CSO', 'ZSO']


class TestXboxConverter:
    """Test Xbox converter"""
    
    def test_xbox_converter_creation(self):
        """Test XboxConverter initialization."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            iso_file = tmpdir / "game.iso"
            iso_file.write_bytes(b'\x00' * 1000)  # Create non-empty file
            
            converter = XboxConverter(input_path=iso_file)
            assert converter.input_path == iso_file
            assert converter.can_convert()

    def test_xbox_output_formats(self):
        """Test Xbox supported output formats."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            iso_file = tmpdir / "game.iso"
            iso_file.write_bytes(b'\x00' * 1000)  # Create non-empty file
            
            converter = XboxConverter(iso_file)
            assert converter.get_output_formats() == ['EXTRACTED']


class TestNintendoConverter:
    """Test Nintendo converter"""
    
    @pytest.mark.parametrize("ext,system", [
        ('.nes', 'Nintendo Entertainment System (NES)'),
        ('.smc', 'Super Nintendo (SNES)'),
        ('.gba', 'Game Boy Advance'),
        ('.3ds', 'Nintendo 3DS'),
    ])
    def test_nintendo_system_detection(self, ext, system):
        """Test Nintendo system detection."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            rom_file = tmpdir / f"game{ext}"
            rom_file.write_bytes(b'\x00' * 1000)  # Create non-empty file
            
            converter = NintendoConverter(rom_file)
            assert converter.system == system

    def test_nintendo_converter_validation(self):
        """Test Nintendo converter can validate ROMs."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            rom_file = tmpdir / "game.nes"
            rom_file.write_bytes(b'\x00' * 1000)  # Create non-empty file
            
            converter = NintendoConverter(rom_file)
            assert converter.can_convert()

    def test_nintendo_output_formats(self):
        """Test Nintendo supported output formats."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            rom_file = tmpdir / "game.nes"
            rom_file.write_bytes(b'\x00' * 1000)  # Create non-empty file
            
            converter = NintendoConverter(rom_file)
            assert converter.get_output_formats() == ['VALIDATED']


class TestConversionResultCompression:
    """Test compression ratio calculation"""
    
    def test_compression_ratio_calculation(self):
        """Test compression ratio is calculated correctly."""
        result = ConversionResult(
            success=True,
            input_path=Path("game.iso"),
            output_path=Path("game.chd"),
            original_size=1000,
            output_size=500,
        )
        assert result.compression_ratio == 0.5

    def test_compression_ratio_zero_input(self):
        """Test compression ratio with zero input size."""
        result = ConversionResult(
            success=True,
            input_path=Path("game.iso"),
            output_path=Path("game.chd"),
            original_size=0,
            output_size=0,
        )
        assert result.compression_ratio == 0.0


class TestBaseConverter:
    """Test BaseConverter abstract class"""
    
    def test_base_converter_is_abstract(self):
        """Test BaseConverter cannot be instantiated directly."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            rom_file = tmpdir / "game.iso"
            rom_file.write_text("")
            
            with pytest.raises(TypeError, match="abstract"):
                BaseConverter(rom_file)
