"""
Pytest configuration and shared fixtures for ROM Converter tests.

This file provides:
- Mock external tools (chdman, maxcso, 7z, etc.)
- Temporary directories for test data
- Mock system resource data
"""

import pytest
import tempfile
from pathlib import Path
from unittest.mock import Mock, MagicMock, patch
import json
import logging


@pytest.fixture
def temp_rom_dir():
    """Create a temporary directory with mock ROM files.
    
    Returns a Path to a temporary directory containing:
    - PS1 ROMs: game1.iso, game2.cue, game2.bin
    - PS2 ROMs: ps2_game.iso
    - PSP ROMs: psp_game.iso
    - Nintendo: nes_game.nes, snes_game.smc
    - Archives: archive.zip, archive.7z
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        temp_path = Path(tmpdir)
        
        # Create mock ROM files (empty)
        (temp_path / "game1.iso").touch()
        (temp_path / "game2.cue").touch()
        (temp_path / "game2.bin").touch()
        (temp_path / "ps2_game.iso").touch()
        (temp_path / "psp_game.iso").touch()
        (temp_path / "nes_game.nes").touch()
        (temp_path / "snes_game.smc").touch()
        
        # Create subdirectory with more ROMs
        subdir = temp_path / "subdir"
        subdir.mkdir()
        (subdir / "game3.iso").touch()
        (subdir / "game4.cue").touch()
        
        # Create mock archives
        (temp_path / "archive.zip").touch()
        (temp_path / "archive.7z").touch()
        
        yield temp_path


@pytest.fixture
def temp_output_dir():
    """Create a temporary output directory for conversions."""
    with tempfile.TemporaryDirectory() as tmpdir:
        yield Path(tmpdir)


@pytest.fixture
def mock_chdman():
    """Mock chdman executable that simulates successful conversion.
    
    Returns a Mock object that behaves like a subprocess.run result.
    """
    mock = Mock()
    mock.returncode = 0
    mock.stdout = "CHD file created successfully"
    mock.stderr = ""
    return mock


@pytest.fixture
def mock_maxcso():
    """Mock maxcso executable that simulates successful compression."""
    mock = Mock()
    mock.returncode = 0
    mock.stdout = "CSO file created successfully"
    mock.stderr = ""
    return mock


@pytest.fixture
def mock_system_resources():
    """Mock system resource information (CPU, RAM, disk).
    
    Returns mocked psutil data showing:
    - 8 CPU cores
    - 16 GB total RAM, 10 GB available
    - 100 GB disk free
    """
    with patch('psutil.cpu_count', return_value=8):
        with patch('psutil.virtual_memory') as mock_mem:
            mock_mem.return_value = MagicMock(
                total=16 * 1024**3,  # 16 GB
                available=10 * 1024**3,  # 10 GB available
                percent=37.5
            )
            with patch('psutil.disk_usage') as mock_disk:
                mock_disk.return_value = MagicMock(
                    total=500 * 1024**3,  # 500 GB
                    used=400 * 1024**3,  # 400 GB used
                    free=100 * 1024**3,  # 100 GB free
                    percent=80.0
                )
                yield {
                    'cpu_count': 8,
                    'memory_total_gb': 16,
                    'memory_available_gb': 10,
                    'disk_free_gb': 100
                }


@pytest.fixture
def temp_config_file():
    """Create a temporary config file for testing.
    
    Returns a Path to a temporary JSON config file with default settings.
    """
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
        config = {
            "_version": 2,
            "source_dir": "",
            "ps1_output_format": "CHD",
            "ps2_output_format": "CHD",
            "psp_output_format": "CSO",
            "ps2_emulator": "PCSX2",
            "current_theme": "PS2",
            "delete_originals": False,
            "move_to_backup": True,
            "extract_compressed": True,
            "delete_archives_after_extract": False,
            "max_concurrent_conversions": None,
            "chdman_max_processors": None,
            "maxcso_threads": None
        }
        json.dump(config, f)
        config_path = Path(f.name)
    
    yield config_path
    
    # Cleanup
    config_path.unlink()


@pytest.fixture
def mock_subprocess_run():
    """Mock subprocess.run for external tool execution.
    
    By default, simulates successful command execution.
    Can be configured per test to simulate failures, timeouts, etc.
    """
    with patch('subprocess.run') as mock_run:
        mock_run.return_value = MagicMock(
            returncode=0,
            stdout="Success",
            stderr=""
        )
        yield mock_run


@pytest.fixture
def cleanup_logging():
    """Cleanup logging handlers after each test.
    
    Removes file handlers that hold open files, allowing temp directories
    to be cleaned up properly after tests.
    """
    yield
    # Clean up all handlers from all loggers to avoid PermissionError on Windows
    for logger_name in list(logging.Logger.manager.loggerDict):
        logger = logging.getLogger(logger_name)
        for handler in logger.handlers[:]:
            handler.close()
            logger.removeHandler(handler)


# Pytest configuration

def pytest_configure(config):
    """Register custom markers for test organization."""
    config.addinivalue_line(
        "markers", "unit: unit tests that don't require external tools"
    )
    config.addinivalue_line(
        "markers", "integration: integration tests that may invoke external tools"
    )
    config.addinivalue_line(
        "markers", "slow: slow tests that take >1 second to run"
    )
    config.addinivalue_line(
        "markers", "requires_tool: tests that require external tools (chdman, 7z, etc)"
    )


# Helper functions for tests

def assert_rom_file_detected(scanner, directory: Path, expected_files: list):
    """Helper to assert that specific ROM files were detected.
    
    Args:
        scanner: FileScanner instance
        directory: Directory to scan
        expected_files: List of expected filenames (e.g., ['game1.iso', 'game2.cue'])
    """
    found_files = scanner.find_game_files(directory)
    found_names = {f.name for f in found_files}
    expected_names = set(expected_files)
    
    assert found_names == expected_names, \
        f"Expected {expected_names}, but got {found_names}"


def assert_system_detected(detector, file_path: Path, expected_system: str):
    """Helper to assert that a file is detected as a specific system.
    
    Args:
        detector: SystemDetector instance
        file_path: Path to ROM file
        expected_system: Expected system name (e.g., 'PlayStation 2')
    """
    detected = detector.detect_system(file_path)
    assert detected == expected_system, \
        f"Expected {expected_system}, but got {detected}"
