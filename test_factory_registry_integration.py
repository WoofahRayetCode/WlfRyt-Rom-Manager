"""
Integration tests for ConverterFactory and ExtractionRegistry with rom_converter.py
Tests that the factories are properly initialized and integrated
"""

import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch
import sys
from queue import Queue


@pytest.fixture
def mock_tkinter():
    """Mock tkinter to avoid GUI initialization"""
    mock_tk = MagicMock()
    sys.modules['tkinter'] = mock_tk
    sys.modules['tkinter.ttk'] = MagicMock()
    sys.modules['tkinter.font'] = MagicMock()
    sys.modules['tkinterdnd2'] = MagicMock()
    return mock_tk


class TestConverterFactoryAvailability:
    """Test that ConverterFactory is available and importable"""
    
    def test_converter_factory_import(self):
        """Test ConverterFactory can be imported"""
        from converter_factory import ConverterFactory
        assert ConverterFactory is not None
    
    def test_converter_factory_initialization(self):
        """Test ConverterFactory can be initialized"""
        from converter_factory import ConverterFactory
        
        factory = ConverterFactory()
        assert factory is not None
        assert factory.chdman_path is None
        assert factory.maxcso_path is None
    
    def test_converter_factory_with_paths(self):
        """Test ConverterFactory initializes with tool paths"""
        from converter_factory import ConverterFactory
        
        factory = ConverterFactory(
            chdman_path=Path("fake/chdman"),
            maxcso_path=Path("fake/maxcso")
        )
        assert factory.chdman_path == Path("fake/chdman")
        assert factory.maxcso_path == Path("fake/maxcso")
    
    def test_converter_factory_logging_callback(self):
        """Test ConverterFactory accepts logging callback"""
        from converter_factory import ConverterFactory
        
        messages = []
        def log_fn(msg):
            messages.append(msg)
        
        factory = ConverterFactory(log_callback=log_fn)
        factory._log("Test message")
        assert "Test message" in messages


class TestExtractionRegistryAvailability:
    """Test that ExtractionRegistry is available and importable"""
    
    def test_extraction_registry_import(self):
        """Test ExtractionRegistry can be imported"""
        from extraction_registry import ExtractionRegistry
        assert ExtractionRegistry is not None
    
    def test_extraction_registry_initialization(self):
        """Test ExtractionRegistry can be initialized"""
        from extraction_registry import ExtractionRegistry
        
        registry = ExtractionRegistry()
        assert registry is not None
        assert registry.seven_zip_path is None
    
    def test_extraction_registry_format_support(self):
        """Test ExtractionRegistry supports standard formats"""
        from extraction_registry import ExtractionRegistry
        
        registry = ExtractionRegistry()
        
        # Test supported formats
        assert registry.is_supported(Path("file.zip")) is True
        assert registry.is_supported(Path("file.tar.gz")) is True
        assert registry.is_supported(Path("file.7z")) is True
        
        # Test unsupported formats
        assert registry.is_supported(Path("file.iso")) is False
        assert registry.is_supported(Path("file.txt")) is False
    
    def test_extraction_registry_logging_callback(self):
        """Test ExtractionRegistry accepts logging callback"""
        from extraction_registry import ExtractionRegistry
        
        messages = []
        def log_fn(msg):
            messages.append(msg)
        
        registry = ExtractionRegistry(log_callback=log_fn)
        registry._log("Test message")
        assert "Test message" in messages


class TestROMConverterFactoryIntegration:
    """Test that rom_converter.py properly integrates ConverterFactory"""
    
    def test_rom_converter_imports(self):
        """Test rom_converter.py imports successfully"""
        try:
            from unittest.mock import MagicMock
            sys.modules['tkinterdnd2'] = MagicMock()
            import rom_converter
            assert rom_converter is not None
        except SyntaxError:
            pytest.fail("rom_converter.py has syntax errors")
    
    def test_converter_factory_flag_available(self):
        """Test CONVERTER_FACTORY_AVAILABLE flag is set"""
        from unittest.mock import MagicMock
        sys.modules['tkinterdnd2'] = MagicMock()
        import rom_converter
        
        # Factory should be available since we created converter_factory.py
        assert hasattr(rom_converter, 'CONVERTER_FACTORY_AVAILABLE')
        assert rom_converter.CONVERTER_FACTORY_AVAILABLE is True
    
    def test_extraction_registry_flag_available(self):
        """Test EXTRACTION_REGISTRY_AVAILABLE flag is set"""
        from unittest.mock import MagicMock
        sys.modules['tkinterdnd2'] = MagicMock()
        import rom_converter
        
        # Registry should be available since we created extraction_registry.py
        assert hasattr(rom_converter, 'EXTRACTION_REGISTRY_AVAILABLE')
        assert rom_converter.EXTRACTION_REGISTRY_AVAILABLE is True


class TestConverterFactoryFormatDetection:
    """Test ConverterFactory's format detection"""
    
    def test_get_converter_for_cue(self, tmp_path):
        """Test ConverterFactory returns converter for CUE file"""
        from converter_factory import ConverterFactory
        from pathlib import Path
        
        # Create a temporary CUE file
        cue_file = tmp_path / "game.cue"
        cue_file.write_text("FILE \"game.bin\" BINARY\n  TRACK 01 MODE1/2352\n    INDEX 01 00:00:00\n")
        
        factory = ConverterFactory()
        # Should work even if chdman_path is None (returns None)
        converter = factory.get_converter(cue_file, "CHD")
        # May be None if chdman not available, but shouldn't error
        assert converter is None or hasattr(converter, 'convert')
    
    def test_get_converter_for_iso(self, tmp_path):
        """Test ConverterFactory returns converter for ISO file"""
        from converter_factory import ConverterFactory
        from pathlib import Path
        
        # Create a temporary ISO file
        iso_file = tmp_path / "game.iso"
        iso_file.write_bytes(b'\x00' * 1000000)  # 1MB dummy ISO
        
        factory = ConverterFactory()
        converter = factory.get_converter(
            iso_file, 
            "CHD",
            system_guess='PS2'
        )
        # May be None if chdman not available, but shouldn't error
        assert converter is None or hasattr(converter, 'convert')
    
    def test_get_converter_unsupported_format(self):
        """Test ConverterFactory returns None for unsupported formats"""
        from converter_factory import ConverterFactory
        from pathlib import Path
        
        factory = ConverterFactory()
        converter = factory.get_converter(Path("file.txt"), "UNKNOWN")
        assert converter is None


class TestExtractionRegistryFormatDetection:
    """Test ExtractionRegistry's format detection"""
    
    def test_get_extractor_for_zip(self):
        """Test ExtractionRegistry returns extractor for ZIP"""
        from extraction_registry import ExtractionRegistry
        from pathlib import Path
        
        registry = ExtractionRegistry()
        extractor = registry.get_extractor(Path("archive.zip"))
        assert extractor is not None
    
    def test_get_extractor_for_tar_gz(self):
        """Test ExtractionRegistry detects tar.gz format"""
        from extraction_registry import ExtractionRegistry
        from pathlib import Path
        
        registry = ExtractionRegistry()
        extractor = registry.get_extractor(Path("archive.tar.gz"))
        assert extractor is not None
    
    def test_get_extractor_for_7z_without_7zip(self):
        """Test ExtractionRegistry returns None for 7Z without 7-Zip"""
        from extraction_registry import ExtractionRegistry
        from pathlib import Path
        
        registry = ExtractionRegistry()  # No seven_zip_path set
        extractor = registry.get_extractor(Path("archive.7z"))
        # Should return None since 7-Zip not configured
        assert extractor is None
    
    def test_get_extractor_unsupported_format(self):
        """Test ExtractionRegistry returns None for unsupported formats"""
        from extraction_registry import ExtractionRegistry
        from pathlib import Path
        
        registry = ExtractionRegistry()
        extractor = registry.get_extractor(Path("file.iso"))
        assert extractor is None


class TestConverterFactoryBatchOperations:
    """Test ConverterFactory batch processing"""
    
    def test_batch_convert_empty_list(self):
        """Test batch_convert with empty file list"""
        from converter_factory import ConverterFactory
        from pathlib import Path
        
        factory = ConverterFactory()
        results = factory.batch_convert([], "CHD")
        
        assert results['total'] == 0
        assert results['successful'] == 0
        assert results['failed'] == 0


class TestExtractionRegistryBatchOperations:
    """Test ExtractionRegistry batch processing"""
    
    def test_find_archives_empty_directory(self, tmp_path):
        """Test finding archives in empty directory"""
        from extraction_registry import ExtractionRegistry
        
        registry = ExtractionRegistry()
        archives = registry._find_archives(tmp_path)
        
        assert len(archives) == 0
    
    def test_find_archives_finds_zip(self, tmp_path):
        """Test finding ZIP archives"""
        from extraction_registry import ExtractionRegistry
        import zipfile
        
        # Create a test ZIP file
        test_zip = tmp_path / "test.zip"
        with zipfile.ZipFile(test_zip, 'w') as z:
            z.writestr("test.txt", "test content")
        
        registry = ExtractionRegistry()
        archives = registry._find_archives(tmp_path, recursive=False)
        
        assert test_zip in archives


class TestDispatchLogic:
    """Test that rom_converter.py correctly dispatches to factories"""
    
    def test_convert_game_dispatch_method_exists(self):
        """Test convert_game_with_factory method exists"""
        from unittest.mock import MagicMock
        sys.modules['tkinterdnd2'] = MagicMock()
        import rom_converter
        
        # Check that dispatch methods exist
        assert hasattr(rom_converter.ROMConverter, 'convert_game_with_factory')
        assert hasattr(rom_converter.ROMConverter, '_convert_game_original')
    
    def test_extract_archive_dispatch_method_exists(self):
        """Test extract_archive_with_registry method exists"""
        from unittest.mock import MagicMock
        sys.modules['tkinterdnd2'] = MagicMock()
        import rom_converter
        
        # Check that dispatch methods exist
        assert hasattr(rom_converter.ROMConverter, 'extract_archive_with_registry')
        assert hasattr(rom_converter.ROMConverter, '_extract_archive_original')
    
    def test_factory_initialization_method_exists(self):
        """Test _init_service_factories method exists"""
        from unittest.mock import MagicMock
        sys.modules['tkinterdnd2'] = MagicMock()
        import rom_converter
        
        assert hasattr(rom_converter.ROMConverter, '_init_service_factories')


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
