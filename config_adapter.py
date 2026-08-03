"""
ROM Converter Configuration Adapter

Provides a bridge between rom_converter.py's existing configuration 
and the new ConfigManager interface for gradual migration.
"""

from pathlib import Path
from typing import Optional, Dict, Any, Callable
from config import ConfigManager, UserPreferences


class ROMConverterConfigAdapter:
    """Adapter that bridges rom_converter.py config with ConfigManager
    
    Allows gradual migration from manual config handling to ConfigManager
    while maintaining backward compatibility.
    """
    
    def __init__(self, 
                 rom_converter_instance,
                 config_file: Optional[Path] = None,
                 log_callback: Optional[Callable[[str], None]] = None):
        """Initialize config adapter
        
        Args:
            rom_converter_instance: Reference to ROMConverter instance
            config_file: Path to config file (default: ~/.rom_converter_config.json)
            log_callback: Optional logging callback
        """
        self.rom_converter = rom_converter_instance
        self.config_file = config_file
        self.log_callback = log_callback
        self.config_manager = None
    
    def _log(self, message: str) -> None:
        """Log a message"""
        if self.log_callback:
            self.log_callback(message)
        if self.rom_converter and hasattr(self.rom_converter, 'log'):
            self.rom_converter.log(message)
    
    def init_config_manager(self) -> ConfigManager:
        """Initialize ConfigManager with existing rom_converter config
        
        Returns:
            Configured ConfigManager instance
        """
        if self.config_manager:
            return self.config_manager
        
        # Determine config file location
        config_file = self.config_file
        if not config_file and self.rom_converter:
            config_file = getattr(self.rom_converter, 'config_file', None)
        
        # Create ConfigManager with candidates
        candidates = [config_file] if config_file else []
        if Path.home().exists():
            candidates.append(Path.home() / ".rom_converter_config.json")
        
        self.config_manager = ConfigManager(candidates)
        
        # Load existing config if it exists
        try:
            self.config_manager.load()
            self._log("✅ ConfigManager initialized and config loaded")
        except Exception as e:
            self._log(f"⚠️  ConfigManager load failed: {e}, using defaults")
            self.config_manager.preferences = UserPreferences()
        
        return self.config_manager
    
    def sync_from_rom_converter(self) -> None:
        """Sync rom_converter.py settings into ConfigManager
        
        Reads current values from rom_converter instance and updates ConfigManager.
        """
        if not self.config_manager:
            self.init_config_manager()
        
        rom_conv = self.rom_converter
        prefs = self.config_manager.preferences
        
        # Sync UI preferences
        if hasattr(rom_conv, 'master'):
            try:
                geom = rom_conv.master.geometry()
                if 'x' in geom:
                    w, h = geom.split('x')[0], geom.split('x')[1].split('+')[0]
                    prefs.window_width = int(w) if w else 1100
                    prefs.window_height = int(h) if h else 800
            except:
                pass
        
        # Sync conversion preferences
        if hasattr(rom_conv, 'source_dir'):
            prefs.source_dir = rom_conv.source_dir or ""
        if hasattr(rom_conv, 'process_ps1_cues'):
            prefs.process_ps1_cues = rom_conv.process_ps1_cues.get()
        if hasattr(rom_conv, 'process_ps2_cues'):
            prefs.process_ps2_cues = rom_conv.process_ps2_cues.get()
        if hasattr(rom_conv, 'process_ps2_isos'):
            prefs.process_ps2_isos = rom_conv.process_ps2_isos.get()
        if hasattr(rom_conv, 'process_psp_isos'):
            prefs.process_psp_isos = rom_conv.process_psp_isos.get()
        if hasattr(rom_conv, 'process_ps3_isos'):
            prefs.process_ps3_isos = rom_conv.process_ps3_isos.get()
        if hasattr(rom_conv, 'process_xbox_isos'):
            prefs.process_xbox_isos = rom_conv.process_xbox_isos.get()
        
        # Sync output formats
        if hasattr(rom_conv, 'ps1_output_format'):
            prefs.ps1_output_format = rom_conv.ps1_output_format
        if hasattr(rom_conv, 'ps2_output_format'):
            prefs.ps2_output_format = rom_conv.ps2_output_format
        if hasattr(rom_conv, 'psp_output_format'):
            prefs.psp_output_format = rom_conv.psp_output_format
        
        # Sync extraction preferences
        if hasattr(rom_conv, 'extract_compressed'):
            prefs.extract_compressed = rom_conv.extract_compressed.get()
        if hasattr(rom_conv, 'delete_archives_after_extract'):
            prefs.delete_archives_after_extract = rom_conv.delete_archives_after_extract.get()
        
        # Sync performance preferences
        if hasattr(rom_conv, 'max_concurrent_conversions'):
            prefs.max_workers = rom_conv.max_concurrent_conversions
        if hasattr(rom_conv, 'ram_threshold_percent'):
            prefs.ram_threshold_percent = rom_conv.ram_threshold_percent
            prefs.memory_threshold_pct = rom_conv.ram_threshold_percent
        if hasattr(rom_conv, 'disk_write_throttle_mb_s'):
            prefs.disk_write_throttle_mb_s = rom_conv.disk_write_throttle_mb_s
        
        self._log("✅ Synced rom_converter settings to ConfigManager")
    
    def sync_to_rom_converter(self) -> None:
        """Sync ConfigManager settings into rom_converter.py
        
        Updates rom_converter instance with ConfigManager values.
        """
        if not self.config_manager:
            return
        
        rom_conv = self.rom_converter
        prefs = self.config_manager.preferences
        
        # Sync conversion preferences
        if hasattr(rom_conv, 'source_dir'):
            rom_conv.source_dir = prefs.source_dir or ""
        if hasattr(rom_conv, 'process_ps1_cues'):
            rom_conv.process_ps1_cues.set(prefs.process_ps1_cues)
        if hasattr(rom_conv, 'process_ps2_cues'):
            rom_conv.process_ps2_cues.set(prefs.process_ps2_cues)
        if hasattr(rom_conv, 'process_ps2_isos'):
            rom_conv.process_ps2_isos.set(prefs.process_ps2_isos)
        if hasattr(rom_conv, 'process_psp_isos'):
            rom_conv.process_psp_isos.set(prefs.process_psp_isos)
        if hasattr(rom_conv, 'process_ps3_isos'):
            rom_conv.process_ps3_isos.set(prefs.process_ps3_isos)
        if hasattr(rom_conv, 'process_xbox_isos'):
            rom_conv.process_xbox_isos.set(prefs.process_xbox_isos)
        
        # Sync output formats
        rom_conv.ps1_output_format = prefs.ps1_output_format
        rom_conv.ps2_output_format = prefs.ps2_output_format
        rom_conv.psp_output_format = prefs.psp_output_format
        
        # Sync extraction preferences
        if hasattr(rom_conv, 'extract_compressed'):
            rom_conv.extract_compressed.set(prefs.extract_compressed)
        if hasattr(rom_conv, 'delete_archives_after_extract'):
            rom_conv.delete_archives_after_extract.set(prefs.delete_archives_after_extract)
        
        # Sync performance preferences
        if hasattr(rom_conv, 'chdman_max_processors'):
            rom_conv.chdman_max_processors = prefs.max_workers
        if hasattr(rom_conv, 'maxcso_threads'):
            rom_conv.maxcso_threads = prefs.max_workers
        if hasattr(rom_conv, 'max_concurrent_conversions'):
            rom_conv.max_concurrent_conversions = prefs.max_workers
        if hasattr(rom_conv, 'ram_threshold_percent'):
            rom_conv.ram_threshold_percent = prefs.memory_threshold_pct
        if hasattr(rom_conv, 'disk_write_throttle_mb_s'):
            rom_conv.disk_write_throttle_mb_s = prefs.disk_write_throttle_mb_s
        
        self._log("✅ Synced ConfigManager settings to rom_converter")
    
    def save(self) -> bool:
        """Save configuration to file
        
        Returns:
            True if save successful
        """
        if not self.config_manager:
            return False
        
        try:
            self.sync_from_rom_converter()  # Get latest values
            self.config_manager.save()
            self._log("✅ Configuration saved")
            return True
        except Exception as e:
            self._log(f"❌ Save failed: {e}")
            return False
    
    def load(self) -> bool:
        """Load configuration from file
        
        Returns:
            True if load successful
        """
        if not self.config_manager:
            self.init_config_manager()
        
        try:
            self.config_manager.load()
            self.sync_to_rom_converter()
            self._log("✅ Configuration loaded")
            return True
        except Exception as e:
            self._log(f"⚠️  Load failed: {e}, using defaults")
            return False
    
    def validate(self) -> bool:
        """Validate current configuration
        
        Returns:
            True if configuration is valid
        """
        if not self.config_manager:
            return False
        
        errors = self.config_manager.validate()
        if errors:
            for error in errors:
                self._log(f"❌ Config error: {error}")
            return False
        
        self._log("✅ Configuration is valid")
        return True
    
    def reset_to_defaults(self) -> None:
        """Reset all preferences to defaults"""
        if self.config_manager:
            self.config_manager.reset_to_defaults()
            self.sync_to_rom_converter()
            self._log("✅ Reset configuration to defaults")

    def get_preferences(self) -> UserPreferences:
        """Get current user preferences, initializing config manager if needed."""
        if not self.config_manager:
            self.init_config_manager()
        return self.config_manager.preferences

    def update_preferences(self, **kwargs: Any) -> bool:
        """Update preferences from UI panels and persist them.

        Args:
            **kwargs: Preference key/value pairs

        Returns:
            True if update and save succeeded, False otherwise
        """
        if not self.config_manager:
            self.init_config_manager()

        prefs = self.config_manager.preferences
        updated_any = False

        # Backward-compatible alias support.
        key_aliases = {
            "max_processes": "max_workers",
            "max_threads": "max_workers",
            "ram_threshold_percent": "memory_threshold_pct",
        }

        for key, value in kwargs.items():
            target_key = key_aliases.get(key, key)
            if hasattr(prefs, target_key):
                setattr(prefs, target_key, value)
                updated_any = True
            else:
                self._log(f"⚠️  Ignoring unknown preference key: {key}")

        if not updated_any:
            return False

        self.sync_to_rom_converter()
        if self.config_manager.save():
            self._log("✅ Updated preferences saved")
            return True

        self._log("❌ Failed to save updated preferences")
        return False
    
    def get_config_manager(self) -> Optional[ConfigManager]:
        """Get the underlying ConfigManager instance
        
        Returns:
            ConfigManager or None if not initialized
        """
        return self.config_manager
    
    def __str__(self) -> str:
        """String representation"""
        if self.config_manager:
            return str(self.config_manager)
        return "ConfigAdapter (uninitialized)"
