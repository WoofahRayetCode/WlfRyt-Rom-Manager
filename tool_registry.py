"""
Tool Registry for centralized tool management

Provides a central registry for managing external tool detection and configuration.
Coordinates tool availability, version tracking, and provides unified interface.
"""

from pathlib import Path
from typing import Optional, Dict, Callable
import logging
from tools import ChdmanTool, MaxcsoTool, ToolManager, ToolInfo


class ToolRegistry:
    """Central registry for managing external tools
    
    Provides unified interface for:
    - Tool detection and availability checking
    - Version tracking across all tools
    - Centralized tool configuration
    - Tool status monitoring
    """
    
    # Known tools and their fallback search names
    KNOWN_TOOLS = {
        'chdman': ('chdman.exe', 'chdman'),
        'maxcso': ('maxcso.exe', 'maxcso'),
        '7zip': ('7z.exe', '7z'),
        'ps3-dumper': ('ps3-disc-dumper.exe', 'ps3-disc-dumper'),
        'extract-xiso': ('extract-xiso.exe', 'extract-xiso'),
        'ndecrypt': ('NDecrypt.exe', 'NDecrypt'),
    }
    
    def __init__(self, 
                 logger: Optional[logging.Logger] = None,
                 log_callback: Optional[Callable[[str], None]] = None):
        """Initialize Tool Registry
        
        Args:
            logger: Optional logger instance
            log_callback: Optional logging callback function
        """
        self.logger = logger
        self.log_callback = log_callback
        self.tools: Dict[str, ToolManager] = {}
        self.tool_paths: Dict[str, Optional[Path]] = {}
        
        # Initialize known tools
        self._init_tools()
    
    def _log(self, message: str) -> None:
        """Log a message"""
        if self.logger:
            self.logger.info(message)
        if self.log_callback:
            self.log_callback(message)
    
    def _init_tools(self) -> None:
        """Initialize tool managers for known tools"""
        # Add chdman
        self.tools['chdman'] = ChdmanTool(
            logger=self.logger,
            log_callback=self.log_callback
        )
        
        # Add maxcso
        self.tools['maxcso'] = MaxcsoTool(
            logger=self.logger,
            log_callback=self.log_callback
        )
        
        # TODO: Add other tools when ToolManager subclasses are created
        # self.tools['7zip'] = SevenZipTool(...)
        # self.tools['ps3-dumper'] = PS3DumperTool(...)
        # self.tools['extract-xiso'] = ExtractXisoTool(...)
        # self.tools['ndecrypt'] = NDecryptTool(...)
    
    def register_tool(self, tool_id: str, tool_manager: ToolManager) -> None:
        """Register a tool manager
        
        Args:
            tool_id: Unique identifier for the tool
            tool_manager: ToolManager instance
        """
        self.tools[tool_id] = tool_manager
        self._log(f"Registered tool: {tool_id}")
    
    def set_tool_path(self, tool_id: str, path: Optional[Path]) -> None:
        """Set the path for a specific tool
        
        Args:
            tool_id: Tool identifier
            path: Path to tool executable, or None to auto-detect
        """
        if tool_id not in self.tools:
            self._log(f"Warning: Tool '{tool_id}' not registered")
            return
        
        self.tool_paths[tool_id] = path
        tool = self.tools[tool_id]
        
        # If tool has a set_path method, use it
        if hasattr(tool, 'set_path'):
            tool.set_path(path)
        
        if path:
            self._log(f"Set {tool_id} path: {path}")
        else:
            self._log(f"Cleared {tool_id} path (will auto-detect)")
    
    def check_all(self) -> Dict[str, ToolInfo]:
        """Check availability of all registered tools
        
        Returns:
            Dictionary mapping tool_id to ToolInfo
        """
        results = {}
        for tool_id, tool in self.tools.items():
            available = tool.check_available()
            version = tool.get_version() if available else None
            info = ToolInfo(
                name=tool.name,
                version=version,
                path=getattr(tool, 'get_path', lambda: None)(),
                available=available
            )
            results[tool_id] = info
            self._log(f"Tool check {tool_id}: {info}")
        
        return results
    
    def get_tool(self, tool_id: str) -> Optional[ToolManager]:
        """Get a tool manager by ID
        
        Args:
            tool_id: Tool identifier
        
        Returns:
            ToolManager instance or None if not registered
        """
        return self.tools.get(tool_id)
    
    def is_available(self, tool_id: str) -> bool:
        """Check if a specific tool is available
        
        Args:
            tool_id: Tool identifier
        
        Returns:
            True if tool is available
        """
        tool = self.get_tool(tool_id)
        if not tool:
            return False
        
        return tool.check_available()
    
    def get_info(self, tool_id: str) -> Optional[ToolInfo]:
        """Get information about a tool
        
        Args:
            tool_id: Tool identifier
        
        Returns:
            ToolInfo or None if tool not found
        """
        tool = self.get_tool(tool_id)
        if not tool:
            return None
        
        available = tool.check_available()
        version = tool.get_version() if available else None
        
        return ToolInfo(
            name=tool.name,
            version=version,
            path=getattr(tool, 'get_path', lambda: None)(),
            available=available
        )
    
    def get_status_report(self) -> str:
        """Generate a human-readable status report of all tools
        
        Returns:
            Formatted status report
        """
        lines = ["Tool Registry Status Report", "=" * 40]
        
        for tool_id, tool in self.tools.items():
            info = self.get_info(tool_id)
            if info:
                lines.append(f"  {info}")
            else:
                lines.append(f"  ❌ {tool_id}: Not found")
        
        lines.append("=" * 40)
        return "\n".join(lines)
    
    def export_paths(self) -> Dict[str, str]:
        """Export tool paths for saving to config
        
        Returns:
            Dictionary of tool_id -> path_str
        """
        result = {}
        for tool_id, path in self.tool_paths.items():
            if path:
                result[tool_id] = str(path)
        return result
    
    def import_paths(self, paths: Dict[str, str]) -> None:
        """Import tool paths from config
        
        Args:
            paths: Dictionary of tool_id -> path_str
        """
        for tool_id, path_str in paths.items():
            self.set_tool_path(tool_id, Path(path_str) if path_str else None)
    
    def __str__(self) -> str:
        """String representation showing all tool statuses"""
        return self.get_status_report()


# Convenience function for creating a default registry
def create_default_tool_registry(log_callback: Optional[Callable[[str], None]] = None) -> ToolRegistry:
    """Create a ToolRegistry with standard configuration
    
    Args:
        log_callback: Optional logging callback
    
    Returns:
        Configured ToolRegistry instance
    """
    return ToolRegistry(log_callback=log_callback)
