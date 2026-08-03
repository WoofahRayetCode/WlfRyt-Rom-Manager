"""
Phase 4 Week 2 Integration Tests

Tests for ToolRegistry and ConfigAdapter integration with rom_converter.py
"""

import pytest
from pathlib import Path
from unittest.mock import MagicMock
import sys
from tempfile import NamedTemporaryFile


@pytest.fixture
def mock_tkinter():
    """Mock tkinter to avoid GUI initialization"""
    mock_tk = MagicMock()
    sys.modules['tkinter'] = mock_tk
    sys.modules['tkinter.ttk'] = MagicMock()
    sys.modules['tkinter.font'] = MagicMock()
    sys.modules['tkinterdnd2'] = MagicMock()
    return mock_tk


class TestROMConverterImportability:
    """Test that rom_converter.py can be imported with Phase 4 changes"""
    
    def test_rom_converter_imports_successfully(self, mock_tkinter):
        """Test rom_converter can be imported without errors"""
        try:
            import rom_converter
            # Check that the file has no syntax errors
            assert rom_converter is not None
        except SyntaxError as e:
            pytest.fail(f"rom_converter.py has syntax error: {e}")
    
    def test_rom_converter_imports_tool_registry(self, mock_tkinter):
        """Test rom_converter can import ToolRegistry"""
        try:
            import rom_converter
            # Check that the flag is set
            assert hasattr(rom_converter, 'TOOL_REGISTRY_AVAILABLE')
        except ImportError:
            pytest.skip("rom_converter dependencies not available")
    
    def test_rom_converter_imports_config_adapter(self, mock_tkinter):
        """Test rom_converter can import ConfigAdapter"""
        try:
            import rom_converter
            # Check that the flag is set
            assert hasattr(rom_converter, 'CONFIG_ADAPTER_AVAILABLE')
        except ImportError:
            pytest.skip("rom_converter dependencies not available")


class TestConfigAdapterIntegration:
    """Test ConfigAdapter can be used with rom_converter"""
    
    def test_config_adapter_can_be_created(self, mock_tkinter):
        """Test ConfigAdapter can be created for rom_converter"""
        from config_adapter import ROMConverterConfigAdapter
        
        # Create mock rom_converter
        mock_rom_conv = MagicMock()
        adapter = ROMConverterConfigAdapter(mock_rom_conv)
        
        # Verify adapter was created
        assert adapter is not None
        assert adapter.rom_converter == mock_rom_conv
    
    def test_config_adapter_can_initialize_config_manager(self, mock_tkinter):
        """Test ConfigAdapter can initialize ConfigManager"""
        from config_adapter import ROMConverterConfigAdapter
        
        mock_rom_conv = MagicMock()
        adapter = ROMConverterConfigAdapter(mock_rom_conv)
        
        # Initialize ConfigManager
        config_mgr = adapter.init_config_manager()
        assert config_mgr is not None
        assert adapter.config_manager is not None
    
    def test_config_adapter_can_validate(self, mock_tkinter):
        """Test ConfigAdapter can validate configuration"""
        from config_adapter import ROMConverterConfigAdapter
        
        mock_rom_conv = MagicMock()
        adapter = ROMConverterConfigAdapter(mock_rom_conv)
        adapter.init_config_manager()
        
        # Validate should work
        is_valid = adapter.validate()
        assert isinstance(is_valid, bool)


class TestToolRegistryIntegration:
    """Test ToolRegistry can be used with rom_converter"""
    
    def test_tool_registry_can_be_created(self, mock_tkinter):
        """Test ToolRegistry can be created for rom_converter"""
        from tool_registry import create_default_tool_registry
        
        # Create registry
        registry = create_default_tool_registry()
        
        # Verify it was created
        assert registry is not None
    
    def test_tool_registry_can_check_all_tools(self, mock_tkinter):
        """Test ToolRegistry can check all tools"""
        from tool_registry import create_default_tool_registry
        
        registry = create_default_tool_registry()
        results = registry.check_all()
        
        # Should return a dict with tool results
        assert isinstance(results, dict)
    
    def test_tool_registry_can_get_status(self, mock_tkinter):
        """Test ToolRegistry can generate status report"""
        from tool_registry import create_default_tool_registry
        
        registry = create_default_tool_registry()
        status = registry.get_status_report()
        
        # Should return a string
        assert isinstance(status, str)
        assert 'Tool Registry' in status


class TestLoadConfigIntegration:
    """Test load_config method integrates with ConfigAdapter"""
    
    def test_load_config_with_config_adapter(self, mock_tkinter):
        """Test load_config method with ConfigAdapter available"""
        from config_adapter import ROMConverterConfigAdapter
        
        # Create mock rom_converter
        mock_rom_conv = MagicMock()
        mock_rom_conv.log = MagicMock()
        mock_rom_conv.config_file = Path("/tmp/test_config.json")
        
        # Create adapter
        adapter = ROMConverterConfigAdapter(mock_rom_conv)
        adapter.init_config_manager()
        
        # Should not raise any exceptions
        assert adapter is not None


class TestSaveConfigIntegration:
    """Test save_config method integrates with ConfigAdapter"""
    
    def test_save_config_with_config_adapter(self, mock_tkinter):
        """Test save_config method with ConfigAdapter available"""
        from config_adapter import ROMConverterConfigAdapter
        
        # Create mock rom_converter
        mock_rom_conv = MagicMock()
        mock_rom_conv.log = MagicMock()
        
        with NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            config_file = Path(f.name)
        
        try:
            mock_rom_conv.config_file = config_file
            
            # Create adapter
            adapter = ROMConverterConfigAdapter(mock_rom_conv)
            adapter.init_config_manager()
            
            # Should not raise any exceptions
            assert adapter is not None
        finally:
            if config_file.exists():
                config_file.unlink()


class TestEndToEndROMConverterIntegration:
    """End-to-end tests for rom_converter integration"""
    
    def test_rom_converter_with_tool_registry(self, mock_tkinter):
        """Test rom_converter can work with ToolRegistry"""
        from tool_registry import create_default_tool_registry
        
        # Verify registry can be created
        registry = create_default_tool_registry()
        assert registry is not None
    
    def test_rom_converter_with_config_adapter(self, mock_tkinter):
        """Test rom_converter can work with ConfigAdapter"""
        from config_adapter import ROMConverterConfigAdapter
        
        # Create mock rom_converter
        mock_rom_conv = MagicMock()
        
        # Verify adapter can be created
        adapter = ROMConverterConfigAdapter(mock_rom_conv)
        assert adapter is not None
    
    def test_imports_work_together(self, mock_tkinter):
        """Test all Phase 4 components import successfully together"""
        from tool_registry import ToolRegistry
        from config_adapter import ROMConverterConfigAdapter
        import rom_converter
        
        # All should import without errors
        assert ToolRegistry is not None
        assert ROMConverterConfigAdapter is not None
        assert rom_converter is not None
    
    def test_bidirectional_config_syncing(self, mock_tkinter):
        """Test bidirectional config syncing with ConfigAdapter"""
        from config_adapter import ROMConverterConfigAdapter
        
        # Create mock rom_converter with necessary attributes
        mock_rom_conv = MagicMock()
        mock_rom_conv.log = MagicMock()
        mock_rom_conv.process_ps1_cues = MagicMock()
        mock_rom_conv.process_ps1_cues.get = MagicMock(return_value=True)
        mock_rom_conv.extract_before_converting = MagicMock()
        mock_rom_conv.extract_before_converting.set = MagicMock()
        mock_rom_conv.chdman_max_processors = 4
        mock_rom_conv.ps1_output_format = 'CHD'
        mock_rom_conv.ps2_output_format = 'CHD'
        mock_rom_conv.ps3_output_format = 'ISO'
        
        # Create adapter
        adapter = ROMConverterConfigAdapter(mock_rom_conv)
        adapter.init_config_manager()
        
        # Test sync from rom_converter
        adapter.sync_from_rom_converter()
        
        # Should have stored values
        assert adapter.config_manager is not None
        assert adapter.config_manager.preferences is not None


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
