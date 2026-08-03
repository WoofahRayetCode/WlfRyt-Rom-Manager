"""
Base Converter Class

Abstract base class for all ROM converter implementations.
Defines the interface that specific converters must implement.
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, List
import subprocess
import logging


@dataclass
class ConversionResult:
    """Result of a conversion operation."""
    success: bool
    input_path: Path
    output_path: Optional[Path] = None
    original_size: int = 0
    output_size: int = 0
    duration_seconds: float = 0.0
    error_message: Optional[str] = None
    attempt_count: int = 1
    tool_used: Optional[str] = None

    @property
    def compression_ratio(self) -> float:
        """Calculate compression ratio (output/original)."""
        if self.original_size == 0:
            return 0.0
        return self.output_size / self.original_size

    def __str__(self) -> str:
        if self.success:
            ratio = f"{self.compression_ratio:.1%}"
            return (
                f"✅ {self.input_path.name} → {self.output_path.name} "
                f"({ratio} compression, {self.duration_seconds:.1f}s)"
            )
        else:
            return f"❌ {self.input_path.name}: {self.error_message}"


class BaseConverter(ABC):
    """Base class for all ROM format converters.
    
    Each converter handles conversion for a specific game system and format.
    Subclasses implement the specific logic for their system.
    """

    def __init__(
        self,
        input_path: Path,
        logger: Optional[logging.Logger] = None,
        log_callback=None,
    ):
        """Initialize converter.
        
        Args:
            input_path: Path to input ROM/ISO file
            logger: Optional logging.Logger instance
            log_callback: Optional callback for logging (for tk GUI integration)
        """
        self.input_path = Path(input_path)
        self.logger = logger
        self.log_callback = log_callback

    def log(self, message: str) -> None:
        """Log a message via logger or callback."""
        if self.logger:
            self.logger.info(message)
        if self.log_callback:
            self.log_callback(message)

    @abstractmethod
    def can_convert(self) -> bool:
        """Check if this converter can handle the input file.
        
        Returns:
            True if converter can process this file, False otherwise
        """
        pass

    @abstractmethod
    def get_output_formats(self) -> List[str]:
        """Get list of supported output formats.
        
        Returns:
            List of format strings (e.g., ['CHD', 'CSO', 'ZSO'])
        """
        pass

    @abstractmethod
    def convert(self, output_format: str) -> ConversionResult:
        """Perform the conversion.
        
        Args:
            output_format: Target format (must be in get_output_formats())
            
        Returns:
            ConversionResult with success status and details
        """
        pass

    def _run_command(
        self,
        cmd: List[str],
        timeout: int = 3600,
        capture_output: bool = False,
    ) -> subprocess.CompletedProcess:
        """Run a system command with error handling.
        
        Args:
            cmd: Command and arguments as list
            timeout: Command timeout in seconds
            capture_output: Whether to capture stdout/stderr
            
        Returns:
            subprocess.CompletedProcess
            
        Raises:
            subprocess.TimeoutExpired: If command exceeds timeout
            subprocess.CalledProcessError: If command returns non-zero
        """
        try:
            return subprocess.run(
                cmd,
                timeout=timeout,
                capture_output=capture_output,
                text=True,
                check=True,
            )
        except subprocess.TimeoutExpired as e:
            self.log(f"❌ Command timeout after {timeout}s: {' '.join(cmd)}")
            raise
        except subprocess.CalledProcessError as e:
            self.log(f"❌ Command failed: {e}")
            if e.stderr:
                self.log(f"   Error output: {e.stderr}")
            raise
