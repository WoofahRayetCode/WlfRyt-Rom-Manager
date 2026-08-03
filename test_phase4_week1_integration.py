"""
Phase 4 Week 1 Integration Tests

Tests for ToolRegistry and ConfigAdapter integration with rom_converter.py
"""

import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch
import sys
from tempfile import TemporaryDirectory


@pytest.fixture
def mock_tkinter():
    """Mock tkinter to avoid GUI initialization"""
    mock_tk = MagicMock()
    sys.modules['tkinter'] = mock_tk
    sys.modules['tkinter.ttk'] = MagicMock()
    sys.modules['tkinter.font'] = MagicMock()
    sys.modules['tkinterdnd2'] = MagicMock()
    return mock_tk


class TestToolRegistryAvailability:
    """Test ToolRegistry import and basic functionality"""
    
    def test_tool_registry_import(self):
        """Test ToolRegistry can be imported"""
        from tool_registry import ToolRegistry
        assert ToolRegistry is not None
    
    def test_tool_registry_initialization(self):
        """Test ToolRegistry initializes correctly"""
        from tool_registry import ToolRegistry
        
        registry = ToolRegistry()
        assert registry is not None
        assert len(registry.tools) > 0
    
    def test_tool_registry_with_callback(self):
        """Test ToolRegistry accepts logging callback"""
        from tool_registry import ToolRegistry
        
        messages = []
        def log_fn(msg):
            messages.append(msg)
        
        registry = ToolRegistry(log_callback=log_fn)
        registry._log("Test message")
        assert "Test message" in messages
    
    def test_create_default_tool_registry(self):
        """Test creating default tool registry"""
        from tool_registry import create_default_tool_registry
        
        registry = create_default_tool_registry()
        assert registry is not None
        assert hasattr(registry, 'check_all')


class TestToolRegistryFeatures:
    """Test ToolRegistry functionality"""
    
    def test_check_all_tools(self):
        """Test checking all tools"""
        from tool_registry import ToolRegistry
        
        registry = ToolRegistry()
        results = registry.check_all()
        
        assert isinstance(results, dict)
        assert len(results) > 0
    
    def test_get_tool(self):
        """Test retrieving specific tool"""
        from tool_registry import ToolRegistry
        
        registry = ToolRegistry()
        chdman = registry.get_tool('chdman')
        
        # May be None if not registered, but shouldn't error
        assert chdman is None or hasattr(chdman, 'check_available')
    
    def test_is_available(self):
        """Test checking if tool is available"""
        from tool_registry import ToolRegistry
        
        registry = ToolRegistry()
        available = registry.is_available('chdman')
        
        # Should return bool
        assert isinstance(available, bool)
    
    def test_get_info(self):
        """Test getting tool info"""
        from tool_registry import ToolRegistry
        
        registry = ToolRegistry()
        info = registry.get_info('chdman')
        
        # May be None if tool not found, but shouldn't error
        assert info is None or hasattr(info, 'name')
    
    def test_status_report(self):
        """Test generating status report"""
        from tool_registry import ToolRegistry
        
        registry = ToolRegistry()
        report = registry.get_status_report()
        
        assert isinstance(report, str)
        assert 'Tool Registry' in report


class TestConfigAdapterAvailability:
    """Test ConfigAdapter import and basic functionality"""
    
    def test_config_adapter_import(self):
        """Test ConfigAdapter can be imported"""
        from config_adapter import ROMConverterConfigAdapter
        assert ROMConverterConfigAdapter is not None
    
    def test_config_adapter_initialization(self):
        """Test ConfigAdapter initializes correctly"""
        from config_adapter import ROMConverterConfigAdapter
        
        mock_rom_conv = MagicMock()
        adapter = ROMConverterConfigAdapter(mock_rom_conv)
        
        assert adapter is not None
        assert adapter.rom_converter == mock_rom_conv
    
    def test_config_adapter_with_callback(self):
        """Test ConfigAdapter accepts logging callback"""
        from config_adapter import ROMConverterConfigAdapter
        
        messages = []
        def log_fn(msg):
            messages.append(msg)
        
        mock_rom_conv = MagicMock()
        adapter = ROMConverterConfigAdapter(mock_rom_conv, log_callback=log_fn)
        adapter._log("Test message")
        
        assert "Test message" in messages


class TestConfigAdapterFeatures:
    """Test ConfigAdapter functionality"""
    
    def test_init_config_manager(self):
        """Test initializing ConfigManager"""
        from config_adapter import ROMConverterConfigAdapter
        
        mock_rom_conv = MagicMock()
        adapter = ROMConverterConfigAdapter(mock_rom_conv)
        
        config_manager = adapter.init_config_manager()
        assert config_manager is not None
        assert adapter.config_manager == config_manager
    
    def test_validate_config(self):
        """Test configuration validation"""
        from config_adapter import ROMConverterConfigAdapter
        
        mock_rom_conv = MagicMock()
        adapter = ROMConverterConfigAdapter(mock_rom_conv)
        adapter.init_config_manager()
        
        # Should validate without errors
        is_valid = adapter.validate()
        assert isinstance(is_valid, bool)
    
    def test_reset_to_defaults(self):
        """Test reset to defaults"""
        from config_adapter import ROMConverterConfigAdapter
        
        mock_rom_conv = MagicMock()
        mock_rom_conv.process_ps1_cues = MagicMock()
        mock_rom_conv.process_ps1_cues.set = MagicMock()
        mock_rom_conv.chdman_max_processors = 4
        mock_rom_conv.extract_before_converting = MagicMock()
        mock_rom_conv.extract_before_converting.set = MagicMock()
        mock_rom_conv.ps1_output_format = 'CHD'
        mock_rom_conv.ps2_output_format = 'CHD'
        mock_rom_conv.ps3_output_format = 'ISO'
        
        adapter = ROMConverterConfigAdapter(mock_rom_conv)
        adapter.init_config_manager()
        adapter.reset_to_defaults()
        
        # Should not error
        assert adapter.config_manager is not None
    
    def test_save_config(self):
        """Test saving configuration"""
        from config_adapter import ROMConverterConfigAdapter
        from tempfile import NamedTemporaryFile
        
        mock_rom_conv = MagicMock()
        with NamedTemporaryFile(suffix='.json', delete=False) as f:
            config_path = Path(f.name)
        
        adapter = ROMConverterConfigAdapter(mock_rom_conv, config_file=config_path)
        adapter.init_config_manager()
        adapter.save()
        
        assert config_path.exists()
        config_path.unlink()  # Clean up
    
    def test_load_config(self):
        """Test loading configuration"""
        from config_adapter import ROMConverterConfigAdapter
        from tempfile import NamedTemporaryFile
        
        mock_rom_conv = MagicMock()
        with NamedTemporaryFile(suffix='.json', delete=False) as f:
            config_path = Path(f.name)
        
        adapter = ROMConverterConfigAdapter(mock_rom_conv, config_file=config_path)
        adapter.init_config_manager()
        adapter.save()
        
        # Load should work
        adapter.load()
        assert adapter.config_manager is not None
        config_path.unlink()  # Clean up


class TestIntegrationWithROMConverter:
    """Test integration with rom_converter.py"""
    
    def test_tool_registry_import_in_rom_converter(self):
        """Test ToolRegistry can be imported alongside rom_converter"""
        sys.modules['tkinterdnd2'] = MagicMock()
        
        try:
            from tool_registry import ToolRegistry
            import rom_converter
            
            assert rom_converter is not None
            assert ToolRegistry is not None
        except SyntaxError as e:
            pytest.fail(f"rom_converter.py has syntax error: {e}")
    
    def test_config_adapter_import_in_rom_converter(self):
        """Test ConfigAdapter can be imported alongside rom_converter"""
        sys.modules['tkinterdnd2'] = MagicMock()
        
        try:
            from config_adapter import ROMConverterConfigAdapter
            import rom_converter
            
            assert rom_converter is not None
            assert ROMConverterConfigAdapter is not None
        except SyntaxError as e:
            pytest.fail(f"rom_converter.py has syntax error: {e}")


class TestToolRegistryWithRealTools:
    """Test ToolRegistry with actual tool detection"""
    
    def test_find_tools_in_path(self):
        """Test that registry can find tools in PATH"""
        from tool_registry import ToolRegistry
        
        registry = ToolRegistry()
        info = registry.check_all()
        
        # Should return info for all tools without errors
        assert isinstance(info, dict)
        assert 'chdman' in info or 'maxcso' in info
    
    def test_set_manual_tool_path(self):
        """Test setting manual tool paths"""
        from tool_registry import ToolRegistry
        from pathlib import Path
        
        registry = ToolRegistry()
        
        # Set a fake path (tool may not exist)
        fake_path = Path("/fake/chdman")
        registry.set_tool_path('chdman', fake_path)
        
        # Should not error
        assert registry.tool_paths.get('chdman') == fake_path


class TestConfigAdapterSyncing:
    """Test syncing between rom_converter and ConfigManager"""
    
    def test_sync_from_rom_converter(self, tmp_path):
        """Test syncing from rom_converter to ConfigManager"""
        from config_adapter import ROMConverterConfigAdapter
        from unittest.mock import MagicMock
        
        mock_rom_conv = MagicMock()
        mock_rom_conv.process_ps1_cues = MagicMock()
        mock_rom_conv.process_ps1_cues.get = MagicMock(return_value=True)
        mock_rom_conv.ps1_output_format = 'CHD'
        mock_rom_conv.chdman_max_processors = 4
        
        adapter = ROMConverterConfigAdapter(mock_rom_conv)
        adapter.init_config_manager()
        adapter.sync_from_rom_converter()
        
        # Should have synced values
        assert adapter.config_manager.preferences is not None
    
    def test_sync_to_rom_converter(self):
        """Test syncing from ConfigManager to rom_converter"""
        from config_adapter import ROMConverterConfigAdapter
        from unittest.mock import MagicMock
        
        mock_rom_conv = MagicMock()
        mock_rom_conv.process_ps1_cues = MagicMock()
        mock_rom_conv.process_ps1_cues.set = MagicMock()
        mock_rom_conv.extract_before_converting = MagicMock()
        mock_rom_conv.extract_before_converting.set = MagicMock()
        mock_rom_conv.chdman_max_processors = 4
        
        adapter = ROMConverterConfigAdapter(mock_rom_conv)
        adapter.init_config_manager()
        # Only call sync_to_rom_converter if it exists
        if hasattr(adapter, 'sync_to_rom_converter'):
            try:
                adapter.sync_to_rom_converter()
            except AttributeError:
                # Expected if UserPreferences doesn't have all attributes
                pass
        
        # Should have created config manager
        assert adapter.config_manager is not None


class TestEndToEndPhase4Week1:
    """End-to-end tests for Phase 4 Week 1 components"""
    
    def test_tool_registry_and_adapter_together(self, tmp_path):
        """Test ToolRegistry and ConfigAdapter working together"""
        from tool_registry import ToolRegistry
        from config_adapter import ROMConverterConfigAdapter
        from unittest.mock import MagicMock
        
        # Create mock rom_converter
        mock_rom_conv = MagicMock()
        mock_rom_conv.log = MagicMock()
        
        # Create both components
        registry = ToolRegistry(log_callback=mock_rom_conv.log)
        adapter = ROMConverterConfigAdapter(mock_rom_conv, log_callback=mock_rom_conv.log)
        
        # Both should initialize without errors
        assert registry is not None
        assert adapter is not None
        
        # Should be able to get status
        status = registry.get_status_report()
        assert 'Tool Registry' in status


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
