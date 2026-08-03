"""
Configuration Management

Handles loading, saving, and validating ROM Manager configuration.
"""

import json
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional, Dict, Any
import logging


@dataclass
class UserPreferences:
    """User preferences for ROM Manager."""
    # Window settings
    window_width: int = 1100
    window_height: int = 800
    
    # Conversion settings
    process_ps1_cues: bool = False
    process_ps2_cues: bool = False
    process_ps2_isos: bool = False
    process_psp_isos: bool = False
    process_ps3_isos: bool = False
    process_xbox_isos: bool = False
    process_nes_roms: bool = False
    process_snes_roms: bool = False
    process_n64_roms: bool = False
    
    # Output formats
    ps1_output_format: str = 'CHD'
    ps2_output_format: str = 'CHD'
    psp_output_format: str = 'CSO'
    
    # Archive handling
    extract_compressed: bool = True
    delete_archives_after_extract: bool = False
    source_dir: str = ""
    
    # System settings
    delete_originals: bool = False
    move_to_backup: bool = True
    recursive: bool = True
    
    # Advanced
    max_workers: int = 4
    ram_threshold_percent: int = 80
    disk_write_throttle_mb_s: int = 500
    chunk_size_mb: int = 8
    memory_threshold_pct: int = 80
    cpu_threshold_pct: int = 95
    retry_count: int = 3
    circuit_breaker_threshold: int = 5
    temp_dir: str = ""


class ConfigManager:
    """Manage ROM Manager configuration.
    
    Handles loading, saving, and validating configuration from JSON files.
    Supports multiple config file locations with fallback.
    """

    def __init__(
        self,
        config_candidates: list[Path],
        logger: Optional[logging.Logger] = None,
    ):
        """Initialize config manager.
        
        Args:
            config_candidates: List of config file paths to try (in order)
            logger: Optional logger for debug output
        """
        self.config_candidates = [Path(p) for p in config_candidates]
        self.logger = logger
        self.config_file: Optional[Path] = None
        self.preferences = UserPreferences()
        
        # Find existing config or use first candidate as default
        for candidate in self.config_candidates:
            if candidate.exists():
                self.config_file = candidate
                break
        
        if not self.config_file:
            self.config_file = self.config_candidates[0]

    def log(self, message: str) -> None:
        """Log a message if logger available."""
        if self.logger:
            self.logger.debug(message)

    def load(self) -> bool:
        """Load configuration from file.
        
        Returns:
            True if loaded successfully, False otherwise
        """
        if not self.config_file or not self.config_file.exists():
            self.log(f"Config file not found: {self.config_file}")
            return False

        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Merge loaded data with defaults
            for key, value in data.items():
                if hasattr(self.preferences, key):
                    setattr(self.preferences, key, value)
            
            self.log(f"Config loaded from {self.config_file}")
            return True

        except json.JSONDecodeError as e:
            self.log(f"Config parse error: {e}")
            return False
        except Exception as e:
            self.log(f"Failed to load config: {e}")
            return False

    def save(self) -> bool:
        """Save configuration to file.
        
        Returns:
            True if saved successfully, False otherwise
        """
        try:
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(asdict(self.preferences), f, indent=2)
            
            self.log(f"Config saved to {self.config_file}")
            return True

        except Exception as e:
            self.log(f"Failed to save config: {e}")
            return False

    def get(self, key: str, default: Any = None) -> Any:
        """Get configuration value.
        
        Args:
            key: Configuration key
            default: Default value if key not found
            
        Returns:
            Configuration value, or default if not found
        """
        return getattr(self.preferences, key, default)

    def set(self, key: str, value: Any) -> bool:
        """Set configuration value.
        
        Args:
            key: Configuration key
            value: New value
            
        Returns:
            True if set successfully
        """
        if not hasattr(self.preferences, key):
            self.log(f"Unknown config key: {key}")
            return False
        
        setattr(self.preferences, key, value)
        return True

    def reset_to_defaults(self) -> None:
        """Reset all preferences to defaults."""
        self.preferences = UserPreferences()
        self.log("Config reset to defaults")

    def validate(self) -> list[str]:
        """Validate configuration values.
        
        Returns:
            List of validation error messages (empty if valid)
        """
        errors = []

        # Window size bounds
        if self.preferences.window_width < 800:
            errors.append("Window width must be at least 800")
        if self.preferences.window_height < 600:
            errors.append("Window height must be at least 600")

        # Percentage bounds
        if not (0 <= self.preferences.ram_threshold_percent <= 100):
            errors.append("RAM threshold must be 0-100")

        # Worker bounds
        if self.preferences.max_workers < 1:
            errors.append("Max workers must be at least 1")
        if self.preferences.chunk_size_mb < 1:
            errors.append("Chunk size must be at least 1 MB")
        if not (0 <= self.preferences.memory_threshold_pct <= 100):
            errors.append("Memory threshold must be 0-100")
        if not (0 <= self.preferences.cpu_threshold_pct <= 100):
            errors.append("CPU threshold must be 0-100")
        if self.preferences.retry_count < 0:
            errors.append("Retry count must be at least 0")
        if self.preferences.circuit_breaker_threshold < 1:
            errors.append("Circuit breaker threshold must be at least 1")

        # Format validation
        valid_formats = {'CHD', 'CSO', 'ZSO', 'ISO'}
        for fmt_key in ['ps1_output_format', 'ps2_output_format', 'psp_output_format']:
            fmt = getattr(self.preferences, fmt_key)
            if fmt not in valid_formats:
                errors.append(f"{fmt_key} must be one of {valid_formats}, got {fmt}")

        return errors

    def __str__(self) -> str:
        """Return string representation of current config."""
        config_dict = asdict(self.preferences)
        lines = ['Configuration:', f'  File: {self.config_file}']
        for key, value in config_dict.items():
            lines.append(f'  {key}: {value}')
        return '\n'.join(lines)
