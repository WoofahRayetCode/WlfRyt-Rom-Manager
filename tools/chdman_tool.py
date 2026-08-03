"""
CHDman Tool Manager

Manages chdman (MAME disc image compression tool).
"""

from pathlib import Path
from typing import Optional
import logging
import subprocess
import re

from tools.base import ToolManager, ToolInfo


class ChdmanTool(ToolManager):
    """Manage chdman external tool.
    
    chdman is the MAME compression utility used for creating CHD
    (Compressed Hunks of Data) disc images from ISO/CUE files.
    
    Source: https://github.com/mamedev/mame
    """

    def __init__(
        self,
        path: Optional[Path] = None,
        logger: Optional[logging.Logger] = None,
        log_callback=None,
    ):
        """Initialize chdman tool manager.
        
        Args:
            path: Path to chdman executable (auto-detect if None)
            logger: Optional logger
            log_callback: Optional log callback
        """
        super().__init__("chdman", logger, log_callback)
        self.info.path = path

    def check_available(self) -> bool:
        """Check if chdman is available."""
        if self.info.path and self.info.path.exists():
            return True
        
        # Try to find in PATH
        try:
            result = subprocess.run(
                ['chdman', '--version'],
                capture_output=True,
                text=True,
                timeout=5,
            )
            return result.returncode == 0
        except Exception:
            return False

    def get_version(self) -> Optional[str]:
        """Get chdman version.
        
        Returns:
            Version string like "0283b", or None
        """
        try:
            cmd = [str(self.info.path)] if self.info.path else ['chdman']
            result = subprocess.run(
                cmd + ['--version'],
                capture_output=True,
                text=True,
                timeout=5,
            )
            
            if result.returncode != 0:
                return None
            
            # Parse version from output (e.g., "MAME chdman (0283b)")
            match = re.search(r'chdman\s+\(([^)]+)\)', result.stdout)
            if match:
                return match.group(1)
            
            # Fallback: search for any version-like string
            match = re.search(r'\d+[a-z]?', result.stdout)
            if match:
                return match.group(0)
            
            return None
        except Exception:
            return None

    def download(self, install_dir: Path) -> bool:
        """Download chdman (not implemented - requires manual download).
        
        Args:
            install_dir: Directory to install to
            
        Returns:
            False (manual download required)
        """
        self.log(
            "⚠️  chdman must be downloaded manually from MAME GitHub releases:\n"
            "    https://github.com/mamedev/mame/releases\n"
            f"    Place extracted chdman.exe in {install_dir}"
        )
        return False
