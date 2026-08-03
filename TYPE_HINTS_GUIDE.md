"""
Type hints reference for ROM Converter.

This module and companion documents provide guidance for adding type hints
to the ROM Converter codebase, with examples and best practices.

Current Status:
- New modules (logging_setup.py, logging_integration.py, retry_logic.py) ✅ Fully typed
- Exception hierarchy (exceptions.py) ✅ Fully typed
- Core rom_converter.py: Partial (candidate for gradual migration)

Type Hints Strategy:
1. New modules are created with full type hints from the start
2. Existing large files (rom_converter.py) are candidates for gradual typing
3. Type checking performed with mypy on new modules
4. Tests include mypy validation in CI/CD

Type Hints Benefits:
- IDE autocompletion and navigation
- Early error detection (mypy catches bugs before runtime)
- Better code documentation
- Easier refactoring with confidence
- Gradual typing allows incremental improvements

Adding Type Hints to Existing Code:

Example 1: Simple function with type hints
    # Before:
    def convert_rom(input_file, output_format):
        return converter.convert(input_file, output_format)
    
    # After:
    from pathlib import Path
    from typing import Optional
    
    def convert_rom(input_file: Path, output_format: str) -> dict:
        \"\"\"Convert ROM file to target format.
        
        Args:
            input_file: Path to input ROM file
            output_format: Target format (CHD, CSO, etc.)
        
        Returns:
            Dictionary with conversion results
        
        Raises:
            FileNotFoundError: If input file does not exist
            ValueError: If format is not supported
        \"\"\"
        return converter.convert(input_file, output_format)

Example 2: Function with Optional return value
    from typing import Optional
    
    def find_tool(tool_name: str) -> Optional[Path]:
        \"\"\"Find external tool in system PATH.
        
        Args:
            tool_name: Name of tool to find
        
        Returns:
            Path to tool if found, None otherwise
        \"\"\"
        # Implementation
        pass

Example 3: Function with multiple return types
    from typing import Union
    
    def get_conversion_status() -> Union[str, dict]:
        \"\"\"Get current conversion status.
        
        Returns:
            Status string or detailed status dict
        \"\"\"
        pass

Example 4: Generic collections
    from typing import List, Dict, Set
    
    def find_rom_files(directory: Path) -> List[Path]:
        \"\"\"Find all ROM files in directory.
        
        Args:
            directory: Directory to search
        
        Returns:
            List of Path objects for each ROM file found
        \"\"\"
        pass
    
    def get_system_config() -> Dict[str, any]:
        \"\"\"Load system configuration.
        
        Returns:
            Dictionary of configuration settings
        \"\"\"
        pass

Example 5: Custom type aliases
    from typing import Tuple
    from pathlib import Path
    
    # Type alias for conversion result
    ConversionResult = Tuple[Path, bool, Optional[str]]
    
    def convert_rom(input_file: Path) -> ConversionResult:
        \"\"\"Convert ROM and return result tuple.
        
        Returns:
            Tuple of (output_path, success, error_message)
        \"\"\"
        pass

Example 6: Protocol/Interface definition (Python 3.8+)
    from typing import Protocol
    
    class Converter(Protocol):
        \"\"\"Interface for ROM converters.\"\"\"
        
        def convert(self, input_file: Path, output_format: str) -> Path:
            \"\"\"Convert ROM file.
            
            Args:
                input_file: Input file path
                output_format: Target format
            
            Returns:
                Path to converted file
            \"\"\"
            ...

Example 7: Class with type hints
    from typing import Optional, List
    from pathlib import Path
    
    class ROMConverter:
        \"\"\"Main ROM converter application.\"\"\"
        
        def __init__(self, config_file: Path) -> None:
            \"\"\"Initialize converter.
            
            Args:
                config_file: Path to configuration file
            \"\"\"
            self.config_file = config_file
            self.conversion_queue: List[Path] = []
        
        def add_file(self, file_path: Path) -> None:
            \"\"\"Add file to conversion queue.
            
            Args:
                file_path: Path to file to convert
            
            Raises:
                FileNotFoundError: If file does not exist
                ValueError: If file is not a valid ROM
            \"\"\"
            if not file_path.exists():
                raise FileNotFoundError(f"File not found: {file_path}")
            self.conversion_queue.append(file_path)
        
        def get_queue_size(self) -> int:
            \"\"\"Get number of files in queue.
            
            Returns:
                Number of files pending conversion
            \"\"\"
            return len(self.conversion_queue)

Running Type Checking:
    # Check single file
    mypy logging_setup.py
    
    # Check multiple files
    mypy logging_setup.py logging_integration.py retry_logic.py
    
    # Check with strict mode (recommended for new code)
    mypy --strict logging_setup.py
    
    # Generate type stubs for external libraries
    mypy --install-types

Common Type Hints Issues:

Issue: "Type X is not defined"
    Solution: Import from typing module
    from typing import Optional, List, Dict, Union

Issue: "Module X has no attribute Y"
    Solution: Add py.typed marker or use TYPE_CHECKING
    from typing import TYPE_CHECKING
    
    if TYPE_CHECKING:
        from expensive_module import SomeClass

Issue: "Any" type warnings
    Solution: Provide explicit types instead of Any
    # Bad:
    def process(data: Any) -> Any:
        pass
    
    # Good:
    def process(data: Dict[str, str]) -> List[str]:
        pass

Gradual Typing Strategy for rom_converter.py:

Phase 1 (Current): New modules fully typed
- logging_setup.py ✅
- logging_integration.py ✅
- retry_logic.py ✅
- exceptions.py ✅

Phase 2: Type hints for module interfaces
- Add type hints to public methods
- Add type hints to critical internal functions
- Document interfaces

Phase 3: Modularization
- Split rom_converter.py into focused modules
- Each new module gets full type hints
- Gradual replacement of monolithic class

Phase 4: Full typing coverage
- Add type hints to remaining code
- Achieve strict mypy compliance
- Maintain typing in all future code

Type Hints Checklist:
- [ ] Function parameters have types
- [ ] Function return values have types
- [ ] Class attributes are typed
- [ ] Methods are documented with Args/Returns
- [ ] Imports include type hints (from typing import ...)
- [ ] No bare 'Any' types (use specific types)
- [ ] Optional/Union types used appropriately
- [ ] mypy runs without errors
- [ ] Type stubs available for external deps

Resources:
- https://docs.python.org/3/library/typing.html
- https://mypy.readthedocs.io/
- https://pep484.pycqa.org/ - Type hints PEP
- https://pep544.pycqa.org/ - Protocols
"""

# This file is for documentation purposes.
# See logging_setup.py, logging_integration.py, and retry_logic.py
# for examples of fully typed Python modules.
