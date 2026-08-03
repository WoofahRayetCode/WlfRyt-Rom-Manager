"""
Tools Package

Modularized tool managers for external dependencies.

Supported Tools:
- chdman (MAME disc compression)
- maxcso (PSP/PS2 ISO compression)
- 7-Zip (archive extraction)
- NDecrypt (3DS decryption)
- ps3-disc-dumper (PS3 decryption)
"""

from tools.base import ToolManager, ToolInfo
from tools.chdman_tool import ChdmanTool
from tools.maxcso_tool import MaxcsoTool

__all__ = [
    'ToolManager',
    'ToolInfo',
    'ChdmanTool',
    'MaxcsoTool',
]
