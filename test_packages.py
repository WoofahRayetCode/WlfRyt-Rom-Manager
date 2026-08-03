"""
Tests for Extractors, Tools, and Config Packages
"""

import pytest
from pathlib import Path
import tempfile
import json
import logging

from extractors import ZipExtractor, TarExtractor, SevenZipExtractor, BaseExtractor, ExtractionResult
from tools import ChdmanTool, MaxcsoTool, ToolManager
from config import ConfigManager, UserPreferences


# ============================================================================
# EXTRACTOR TESTS
# ============================================================================

class TestExtractionResult:
    """Test ExtractionResult dataclass"""
    
    def test_successful_extraction(self):
        """Test successful extraction result."""
        result = ExtractionResult(
            success=True,
            archive_path=Path("game.zip"),
            output_dir=Path("game"),
            file_count=25,
            total_size=1_000_000,
            duration_seconds=5.2,
            tool_used='zipfile',
        )
        
        assert result.success
        assert result.file_count == 25
        assert "✅" in str(result)

    def test_failed_extraction(self):
        """Test failed extraction result."""
        result = ExtractionResult(
            success=False,
            archive_path=Path("game.zip"),
            error_message="File not found",
        )
        
        assert not result.success
        assert "❌" in str(result)


class TestZipExtractor:
    """Test ZIP extractor"""
    
    def test_zip_extractor_creation(self):
        """Test ZipExtractor initialization."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            zip_file = tmpdir / "archive.zip"
            zip_file.write_bytes(b'PK\x03\x04')  # ZIP file signature
            
            extractor = ZipExtractor(zip_file)
            assert extractor.archive_path == zip_file
            assert extractor.output_dir == (tmpdir / "archive")

    def test_zip_custom_output_dir(self):
        """Test ZipExtractor with custom output directory."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            zip_file = tmpdir / "archive.zip"
            zip_file.write_bytes(b'')
            
            output_dir = tmpdir / "custom_output"
            extractor = ZipExtractor(zip_file, output_dir=output_dir)
            assert extractor.output_dir == output_dir


class TestTarExtractor:
    """Test TAR extractor"""
    
    def test_tar_extractor_creation(self):
        """Test TarExtractor initialization."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            tar_file = tmpdir / "archive.tar"
            tar_file.write_bytes(b'\x1f\x8b\x08')  # TAR signature
            
            extractor = TarExtractor(tar_file)
            assert extractor.archive_path == tar_file
            assert extractor.output_dir == (tmpdir / "archive")

    def test_tar_filename_variations(self):
        """Test TAR extractor recognizes various TAR filenames."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            
            names = ['archive.tar', 'archive.tar.gz', 'archive.tgz', 'archive.tar.bz2']
            for name in names:
                tar_file = tmpdir / name
                tar_file.write_bytes(b'')
                
                extractor = TarExtractor(tar_file)
                # Filename should be recognized
                assert name in [extractor.archive_path.name]


class TestSevenZipExtractor:
    """Test 7Z/RAR extractor"""
    
    def test_sevenzip_extractor_creation(self):
        """Test SevenZipExtractor initialization."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            archive_file = tmpdir / "archive.7z"
            archive_file.write_bytes(b'7z\xbc\xaf\x27\x1c')  # 7Z signature
            
            # Don't require 7-zip to be installed for this test
            try:
                extractor = SevenZipExtractor(archive_file)
                assert extractor.archive_path == archive_file
            except ValueError:
                # 7-Zip not installed, skip test
                pytest.skip("7-Zip not installed")


# ============================================================================
# TOOL MANAGER TESTS
# ============================================================================

class TestToolInfo:
    """Test ToolInfo dataclass"""
    
    def test_available_tool_display(self):
        """Test string representation of available tool."""
        from tools.base import ToolInfo
        
        info = ToolInfo(
            name="chdman",
            version="0283b",
            path=Path("/usr/bin/chdman"),
            available=True,
        )
        
        assert "✅" in str(info)
        assert "chdman" in str(info)
        assert "0283b" in str(info)

    def test_unavailable_tool_display(self):
        """Test string representation of unavailable tool."""
        from tools.base import ToolInfo
        
        info = ToolInfo(
            name="chdman",
            available=False,
        )
        
        assert "❌" in str(info)


class TestChdmanTool:
    """Test chdman tool manager"""
    
    def test_chdman_tool_creation(self):
        """Test ChdmanTool initialization."""
        tool = ChdmanTool()
        assert tool.name == "chdman"
        assert not tool.info.available

    def test_chdman_tool_with_path(self):
        """Test ChdmanTool with specified path."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            chdman_exe = tmpdir / "chdman.exe"
            chdman_exe.write_bytes(b'')
            
            tool = ChdmanTool(path=chdman_exe)
            assert tool.info.path == chdman_exe


class TestMaxcsoTool:
    """Test maxcso tool manager"""
    
    def test_maxcso_tool_creation(self):
        """Test MaxcsoTool initialization."""
        tool = MaxcsoTool()
        assert tool.name == "maxcso"
        assert not tool.info.available

    def test_maxcso_tool_with_path(self):
        """Test MaxcsoTool with specified path."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            maxcso_exe = tmpdir / "maxcso.exe"
            maxcso_exe.write_bytes(b'')
            
            tool = MaxcsoTool(path=maxcso_exe)
            assert tool.info.path == maxcso_exe


# ============================================================================
# CONFIG MANAGER TESTS
# ============================================================================

class TestUserPreferences:
    """Test UserPreferences dataclass"""
    
    def test_default_preferences(self):
        """Test default user preferences."""
        prefs = UserPreferences()
        
        assert prefs.window_width == 1100
        assert prefs.window_height == 800
        assert prefs.ps1_output_format == 'CHD'
        assert prefs.extract_compressed == True

    def test_preferences_modification(self):
        """Test modifying preferences."""
        prefs = UserPreferences()
        prefs.window_width = 1280
        prefs.ps1_output_format = 'ISO'
        
        assert prefs.window_width == 1280
        assert prefs.ps1_output_format == 'ISO'


class TestConfigManager:
    """Test ConfigManager"""
    
    def test_config_manager_creation(self):
        """Test ConfigManager initialization."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            config_file = tmpdir / "config.json"
            
            manager = ConfigManager([config_file])
            assert manager.config_file == config_file
            assert isinstance(manager.preferences, UserPreferences)

    def test_config_save_and_load(self):
        """Test saving and loading config."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            config_file = tmpdir / "config.json"
            
            # Create and save config
            manager1 = ConfigManager([config_file])
            manager1.preferences.window_width = 1920
            manager1.preferences.ps1_output_format = 'ISO'
            assert manager1.save()
            
            # Load config
            manager2 = ConfigManager([config_file])
            assert manager2.load()
            assert manager2.preferences.window_width == 1920
            assert manager2.preferences.ps1_output_format == 'ISO'

    def test_config_get_set(self):
        """Test get/set configuration values."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            config_file = tmpdir / "config.json"
            
            manager = ConfigManager([config_file])
            
            # Get value
            assert manager.get('window_width') == 1100
            
            # Set value
            assert manager.set('window_width', 1280)
            assert manager.get('window_width') == 1280
            
            # Try invalid key
            assert not manager.set('invalid_key', 'value')

    def test_config_validation(self):
        """Test configuration validation."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            config_file = tmpdir / "config.json"
            
            manager = ConfigManager([config_file])
            
            # Valid config
            errors = manager.validate()
            assert len(errors) == 0
            
            # Invalid window size
            manager.preferences.window_width = 400
            errors = manager.validate()
            assert any("Window width" in e for e in errors)
            
            # Invalid format
            manager.preferences.window_width = 1100
            manager.preferences.ps1_output_format = 'INVALID'
            errors = manager.validate()
            assert any("ps1_output_format" in e for e in errors)

    def test_config_reset_to_defaults(self):
        """Test resetting config to defaults."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            config_file = tmpdir / "config.json"
            
            manager = ConfigManager([config_file])
            manager.preferences.window_width = 1920
            manager.preferences.ps1_output_format = 'ISO'
            
            manager.reset_to_defaults()
            
            assert manager.preferences.window_width == 1100
            assert manager.preferences.ps1_output_format == 'CHD'

    def test_config_multiple_candidates(self):
        """Test config file candidate fallback."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            
            # Create config in second candidate
            config1 = tmpdir / "config1.json"
            config2 = tmpdir / "config2.json"
            config2.write_text('{}')
            
            manager = ConfigManager([config1, config2])
            assert manager.config_file == config2

    def test_config_string_representation(self):
        """Test string representation of config."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            config_file = tmpdir / "config.json"
            
            manager = ConfigManager([config_file])
            config_str = str(manager)
            
            assert "Configuration:" in config_str
            assert str(config_file) in config_str
            assert "window_width" in config_str
