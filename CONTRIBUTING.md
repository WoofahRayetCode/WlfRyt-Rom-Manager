# Contributing to ROM Converter

Thank you for your interest in contributing to ROM Converter! This document provides guidelines for contributing code, reporting issues, and suggesting improvements.

## Table of Contents

1. [Getting Started](#getting-started)
2. [Development Setup](#development-setup)
3. [Code Style](#code-style)
4. [Testing](#testing)
5. [Submitting Changes](#submitting-changes)
6. [Adding New Converters](#adding-new-converters)
7. [Reporting Issues](#reporting-issues)

---

## Getting Started

### Prerequisites

- Python 3.8+
- Git
- Basic familiarity with ROM conversion concepts
- Tkinter (included with Python)

### Fork and Clone

```bash
# Fork the repository on GitHub
# Clone your fork
git clone https://github.com/YOUR_USERNAME/WlfRyt-Rom-Manager.git
cd WlfRyt-Rom-Manager

# Add upstream remote
git remote add upstream https://github.com/WoofahRayetCode/WlfRyt-Rom-Manager.git
```

---

## Development Setup

### Install Dependencies

```bash
# Create a virtual environment (recommended)
python -m venv venv

# Activate it
# On Windows:
venv\Scripts\activate
# On Linux/macOS:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Install development dependencies
pip install pytest pytest-cov mypy black flake8
```

### Verify Setup

```bash
# Run the app
python rom_converter.py

# Run tests
pytest --verbose

# Check code style
flake8 rom_converter.py
mypy rom_converter.py --ignore-missing-imports
```

---

## Code Style

### Conventions

- **Formatting**: Use [Black](https://github.com/psf/black) for consistent formatting
- **Type Hints**: Add type hints to all public methods
- **Docstrings**: Use Google-style docstrings for classes and methods
- **Comments**: Comment "why", not "what" - code should be self-documenting
- **Line Length**: 100 characters (Black's default)
- **Imports**: Organize as: stdlib → third-party → local

### Example: Well-Styled Code

```python
"""Module for PS2 ROM conversion.

Handles conversion of PlayStation 2 disc images to various output formats
(CHD, ISO, CSO, ZSO) based on emulator recommendations.
"""

from typing import Optional
from pathlib import Path
import subprocess
import logging

logger = logging.getLogger(__name__)


class PS2Converter:
    """Convert PS2 disc images to CHD, ISO, CSO, or ZSO format.
    
    Attributes:
        chdman_path: Path to chdman executable
        max_processors: Maximum parallel processors for chdman
    """
    
    def __init__(self, chdman_path: str, max_processors: int = 2) -> None:
        """Initialize PS2 converter.
        
        Args:
            chdman_path: Full path to chdman executable
            max_processors: Max processors (1-8, default 2)
        
        Raises:
            FileNotFoundError: If chdman not found
            ValueError: If max_processors out of range
        """
        if max_processors < 1 or max_processors > 8:
            raise ValueError(f"max_processors must be 1-8, got {max_processors}")
        
        self.chdman_path = chdman_path
        self.max_processors = max_processors
    
    def convert(
        self,
        source: Path,
        output_dir: Path,
        output_format: str = "CHD"
    ) -> bool:
        """Convert PS2 ISO to target format.
        
        Args:
            source: Path to source ISO file
            output_dir: Directory for output file
            output_format: Target format (CHD, ISO, CSO, ZSO)
        
        Returns:
            True if conversion succeeded, False otherwise
        
        Raises:
            ValueError: If output_format not supported
            FileNotFoundError: If source file doesn't exist
        """
        if output_format not in ("CHD", "ISO", "CSO", "ZSO"):
            raise ValueError(f"Unsupported format: {output_format}")
        
        if not source.exists():
            raise FileNotFoundError(f"Source file not found: {source}")
        
        # Perform conversion
        return self._run_conversion(source, output_dir, output_format)
```

### Linting & Type Checking

Before committing, run:

```bash
# Format code
black rom_converter.py

# Check for style issues
flake8 rom_converter.py

# Check types
mypy rom_converter.py --ignore-missing-imports

# Run tests
pytest -v
```

---

## Testing

### Test Structure

```
tests/
├── conftest.py           # Shared fixtures
├── test_converters.py    # Converter unit tests
├── test_extractors.py    # Extractor unit tests
├── test_config.py        # Config management tests
├── test_file_scanner.py  # File discovery tests
└── integration/
    └── test_workflow.py  # Full pipeline tests
```

### Writing Tests

```python
"""Tests for new feature."""

import pytest
from pathlib import Path


@pytest.mark.unit
class TestNewFeature:
    """Test suite for new feature."""
    
    def test_basic_functionality(self, temp_rom_dir):
        """Test that feature works correctly."""
        # Setup
        feature = NewFeature(temp_rom_dir)
        
        # Execute
        result = feature.do_something()
        
        # Assert
        assert result.success is True
    
    def test_error_handling(self):
        """Test proper error handling."""
        with pytest.raises(FeatureError):
            feature = NewFeature(None)
            feature.do_something()
    
    @pytest.mark.slow
    def test_large_input(self, large_rom_directory):
        """Test with large input (slow test)."""
        # This test takes >1 second
        pass
```

### Running Tests

```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=. --cov-report=html

# Run only unit tests
pytest -m unit

# Run specific test
pytest test_converters.py::TestPS2Converter::test_to_chd

# Run and show print output
pytest -s
```

### Test Markers

- `@pytest.mark.unit` - Fast unit tests (no external tools)
- `@pytest.mark.integration` - Tests that invoke external tools
- `@pytest.mark.slow` - Tests taking >1 second
- `@pytest.mark.requires_tool` - Tests requiring external tools

---

## Submitting Changes

### Branch Naming

Use descriptive branch names:

```
feature/add-xbox-support
fix/handle-corrupted-archives
docs/update-readme
refactor/modularize-ui
```

### Commit Messages

Follow conventional commits:

```
<type>(<scope>): <subject>

<body>

<footer>
```

**Types**: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `perf`

**Example**:
```
feat(converters): add support for Xbox 360 ISOs

Add Xbox360Converter class to convert Xbox 360 disc images to CHD format.
Includes automatic system detection via sector analysis.

Fixes #42
```

### Pull Request Process

1. **Update your branch** with latest upstream changes:
   ```bash
   git fetch upstream
   git rebase upstream/main
   ```

2. **Push to your fork**:
   ```bash
   git push origin feature/my-feature
   ```

3. **Open PR on GitHub** with:
   - Clear description of changes
   - Reference to related issues (#42)
   - Test results (coverage, all tests pass)
   - Screenshots if UI changes

4. **Address feedback** - maintainers may request changes

5. **Merge** - maintainers merge when approved

---

## Adding New Converters

### Step-by-Step Guide

#### 1. Create Converter Class

```python
# converters/xbox_converter.py
from pathlib import Path
from .base import BaseConverter, ConversionResult


class XboxConverter(BaseConverter):
    """Convert Xbox/Xbox 360 disc images to CHD format."""
    
    def can_convert(self, file: Path) -> bool:
        """Check if file is Xbox format."""
        # Check file extension and/or magic bytes
        return file.suffix.lower() in ('.xiso', '.iso')
    
    def convert(
        self,
        source: Path,
        output_dir: Path,
        **kwargs
    ) -> ConversionResult:
        """Perform the actual conversion."""
        # Implementation
        pass
```

#### 2. Register Converter

```python
# In core/converter_engine.py
from converters.xbox_converter import XboxConverter

def _load_converters(self):
    """Load all available converters."""
    return {
        'PS1': PS1Converter(...),
        'PS2': PS2Converter(...),
        'Xbox': XboxConverter(...),  # Add here
    }
```

#### 3. Add System Detection

```python
# In constants.py
SYSTEM_EXTENSIONS = {
    # ... existing ...
    '.xiso': 'Xbox',
    '.iso': 'Xbox/Xbox 360',  # Update detection logic
}

XBOX_ID_PATTERNS = ('xbox_id', 'xbx')  # Example patterns
```

#### 4. Add Tests

```python
# test_converters.py
def test_xbox_converter_to_chd(temp_rom_dir):
    """Test Xbox ISO to CHD conversion."""
    converter = XboxConverter(chdman_path="/mock/chdman")
    
    result = converter.convert(
        source=Path("game.xiso"),
        output_dir=Path("/output")
    )
    
    assert result.success
    assert result.output_file.suffix == ".chd"
```

#### 5. Update UI

```python
# In ui/main_window.py
self.process_xbox_isos = BooleanVar(value=False)  # Add checkbox
self.xbox_output_format = 'CHD'  # Add format option
```

#### 6. Update Documentation

```markdown
# In README.md
| Xbox | .xiso, .iso | CHD | Xbox discs |
```

---

## Reporting Issues

### Bug Reports

Include:

1. **System Info**:
   - OS (Windows 10, Ubuntu 20.04, etc.)
   - Python version (`python --version`)
   - ROM Converter version

2. **Steps to Reproduce**:
   ```
   1. Open ROM Converter
   2. Select source folder with Xbox ISOs
   3. Select output format: CHD
   4. Click "Start Conversion"
   ```

3. **Expected vs Actual**:
   - Expected: Conversion completes successfully
   - Actual: Error "chdman not found"

4. **Log Output**:
   - Full error message
   - Last 50 lines from `.rom_converter_progress.json`
   - Check `~/.rom_converter_*.log` if it exists

### Feature Requests

Describe:
- **What**: Feature you want to add
- **Why**: Why it's useful
- **Example**: Mockup or usage scenario

Example:
```
**Feature**: Batch presets

**Why**: Users often convert for specific emulators (PCSX2, OPL, 
Citra) with different optimal settings.

**Example**: 
- Click "Presets" dropdown
- Select "PCSX2 Optimized" (CHD, max quality)
- Conversion starts with recommended settings
```

---

## Questions?

- **Issues & Bugs**: [GitHub Issues](../../issues)
- **Discussions**: [GitHub Discussions](../../discussions)
- **Email**: (contact info)

---

**Thank you for contributing!** 🙏

Your effort helps make ROM Converter better for everyone.
