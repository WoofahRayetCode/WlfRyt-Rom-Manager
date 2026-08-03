"""
Base Tool Manager Class

Abstract base class for external tool management.
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
import logging
import subprocess


@dataclass
class ToolInfo:
    """Information about an external tool."""
    name: str
    version: Optional[str] = None
    path: Optional[Path] = None
    available: bool = False
    last_check: Optional[float] = None

    def __str__(self) -> str:
        if self.available:
            version_str = f" v{self.version}" if self.version else ""
            return f"✅ {self.name}{version_str} ({self.path})"
        else:
            return f"❌ {self.name} not found"


class ToolManager(ABC):
    """Base class for external tool managers.
    
    Each tool manager handles detection, version checking, and configuration
    for an external tool (chdman, maxcso, 7-zip, etc.)
    """

    def __init__(
        self,
        name: str,
        logger: Optional[logging.Logger] = None,
        log_callback=None,
    ):
        """Initialize tool manager.
        
        Args:
            name: Human-readable tool name
            logger: Optional logger
            log_callback: Optional log callback
        """
        self.name = name
        self.logger = logger
        self.log_callback = log_callback
        self.info = ToolInfo(name=name)

    def log(self, message: str) -> None:
        """Log a message."""
        if self.logger:
            self.logger.info(message)
        if self.log_callback:
            self.log_callback(message)

    @abstractmethod
    def check_available(self) -> bool:
        """Check if tool is available on system.
        
        Returns:
            True if tool found and working
        """
        pass

    @abstractmethod
    def get_version(self) -> Optional[str]:
        """Get installed tool version.
        
        Returns:
            Version string, or None if unable to determine
        """
        pass

    @abstractmethod
    def download(self, install_dir: Path) -> bool:
        """Download and install tool to specified directory.
        
        Args:
            install_dir: Directory to install tool to
            
        Returns:
            True if installation successful
        """
        pass

    def get_info(self) -> ToolInfo:
        """Get current tool information.
        
        Returns:
            ToolInfo with name, version, path, availability
        """
        self.info.available = self.check_available()
        if self.info.available:
            self.info.version = self.get_version()
        return self.info

    def _run_command(self, cmd: list, timeout: int = 10) -> Optional[str]:
        """Run command and capture output.
        
        Args:
            cmd: Command and arguments
            timeout: Command timeout in seconds
            
        Returns:
            Command output, or None if failed
        """
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=timeout,
            )
            return result.stdout if result.returncode == 0 else None
        except Exception:
            return None
