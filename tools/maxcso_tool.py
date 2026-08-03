"""
Maxcso Tool Manager

Manages maxcso (PSP ISO compression tool).
"""

from pathlib import Path
from typing import Optional
import logging
import subprocess
import re

from tools.base import ToolManager, ToolInfo


class MaxcsoTool(ToolManager):
    """Manage maxcso external tool.
    
    maxcso compresses PSP and PS2 ISO files to CSO/ZSO format,
    reducing file size significantly.
    
    Source: https://github.com/unknownbrackets/maxcso
    """

    def __init__(
        self,
        path: Optional[Path] = None,
        logger: Optional[logging.Logger] = None,
        log_callback=None,
    ):
        """Initialize maxcso tool manager.
        
        Args:
            path: Path to maxcso executable (auto-detect if None)
            logger: Optional logger
            log_callback: Optional log callback
        """
        super().__init__("maxcso", logger, log_callback)
        self.info.path = path

    def check_available(self) -> bool:
        """Check if maxcso is available."""
        if self.info.path and self.info.path.exists():
            return True
        
        # Try to find in PATH
        try:
            result = subprocess.run(
                ['maxcso', '--version'],
                capture_output=True,
                text=True,
                timeout=5,
            )
            return result.returncode == 0
        except Exception:
            return False

    def get_version(self) -> Optional[str]:
        """Get maxcso version.
        
        Returns:
            Version string, or None
        """
        try:
            cmd = [str(self.info.path)] if self.info.path else ['maxcso']
            result = subprocess.run(
                cmd + ['--version'],
                capture_output=True,
                text=True,
                timeout=5,
            )
            
            if result.returncode != 0:
                return None
            
            # Parse version from output
            match = re.search(r'maxcso\s+(?:v)?(\d+\.\d+[a-z]*)', result.stdout, re.IGNORECASE)
            if match:
                return match.group(1)
            
            # Try maxcso (version X) format
            match = re.search(r'\(?(?:version\s+)?(\d+\.\d+[a-z]*)\)?', result.stdout, re.IGNORECASE)
            if match:
                return match.group(1)
            
            return None
        except Exception:
            return None

    def download(self, install_dir: Path) -> bool:
        """Download maxcso (not implemented - requires manual download).
        
        Args:
            install_dir: Directory to install to
            
        Returns:
            False (manual download required)
        """
        self.log(
            "⚠️  maxcso must be downloaded manually from GitHub:\n"
            "    https://github.com/unknownbrackets/maxcso/releases\n"
            f"    Place extracted maxcso.exe in {install_dir}"
        )
        return False
