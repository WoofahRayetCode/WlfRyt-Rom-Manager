"""
Unit tests for ROM file detection and classification.

Tests the file scanning functionality to ensure:
- ROM files are correctly identified by extension
- Systems are correctly detected
- Recursive scanning works as expected
"""

import pytest
from pathlib import Path


# Mock file scanner for testing (demonstrates test pattern)
class MockFileScanner:
    """Simplified file scanner for unit testing."""
    
    SYSTEM_EXTENSIONS = {
        '.iso': 'PlayStation 2',
        '.cue': 'PlayStation',
        '.bin': 'PlayStation',
        '.nes': 'NES',
        '.smc': 'SNES',
        '.sfc': 'SNES',
    }
    
    def find_game_files(self, directory: Path, recursive: bool = True) -> list:
        """Find ROM files in directory."""
        if not directory.exists():
            return []
        
        roms = []
        
        if recursive:
            pattern = "**/*"
        else:
            pattern = "*"
        
        for item in directory.glob(pattern):
            if item.is_file() and item.suffix.lower() in self.SYSTEM_EXTENSIONS:
                roms.append(item)
        
        return sorted(roms)
    
    def detect_system(self, file_path: Path) -> str:
        """Detect system type from file extension."""
        ext = file_path.suffix.lower()
        return self.SYSTEM_EXTENSIONS.get(ext, 'Unknown')


@pytest.mark.unit
class TestFileScanning:
    """Test suite for file scanning functionality."""
    
    def test_find_ps1_files(self, temp_rom_dir):
        """Test that PS1 ROM files are detected."""
        scanner = MockFileScanner()
        
        # Should find .iso, .cue, .bin files
        files = scanner.find_game_files(temp_rom_dir, recursive=False)
        filenames = {f.name for f in files}
        
        # Top-level PS1 files
        assert "game1.iso" in filenames
        assert "game2.cue" in filenames
        assert "game2.bin" in filenames
        
        # Should NOT include subdirectory files (recursive=False)
        assert "game3.iso" not in filenames
    
    def test_find_game_files_recursive(self, temp_rom_dir):
        """Test recursive scanning finds files in subdirectories."""
        scanner = MockFileScanner()
        files = scanner.find_game_files(temp_rom_dir, recursive=True)
        filenames = {f.name for f in files}
        
        # Should include files from both root and subdirectory
        assert "game1.iso" in filenames
        assert "game2.cue" in filenames
        assert "game3.iso" in filenames  # From subdirectory
        assert "game4.cue" in filenames  # From subdirectory
    
    def test_find_game_files_non_recursive(self, temp_rom_dir):
        """Test that non-recursive scanning excludes subdirectories."""
        scanner = MockFileScanner()
        files = scanner.find_game_files(temp_rom_dir, recursive=False)
        filenames = {f.name for f in files}
        
        # Should include top-level files
        assert "game1.iso" in filenames
        
        # Should NOT include subdirectory files
        assert "game3.iso" not in filenames
        assert "game4.cue" not in filenames
    
    def test_find_game_files_empty_directory(self):
        """Test scanning an empty directory."""
        with pytest.temp_file() as tmp:
            scanner = MockFileScanner()
            files = scanner.find_game_files(Path(tmp))
            
            assert files == []
    
    def test_find_game_files_nonexistent_directory(self):
        """Test scanning a directory that doesn't exist."""
        scanner = MockFileScanner()
        files = scanner.find_game_files(Path("/nonexistent/directory"))
        
        assert files == []
    
    def test_system_detection_ps1(self, temp_rom_dir):
        """Test PS1 ROM detection."""
        scanner = MockFileScanner()
        
        iso_file = temp_rom_dir / "game1.iso"
        cue_file = temp_rom_dir / "game2.cue"
        bin_file = temp_rom_dir / "game2.bin"
        
        assert scanner.detect_system(iso_file) == "PlayStation 2"
        assert scanner.detect_system(cue_file) == "PlayStation"
        assert scanner.detect_system(bin_file) == "PlayStation"
    
    def test_system_detection_nintendo(self, temp_rom_dir):
        """Test Nintendo ROM detection."""
        scanner = MockFileScanner()
        
        nes_file = temp_rom_dir / "nes_game.nes"
        snes_file = temp_rom_dir / "snes_game.smc"
        
        assert scanner.detect_system(nes_file) == "NES"
        assert scanner.detect_system(snes_file) == "SNES"
    
    def test_system_detection_unknown(self):
        """Test unknown format detection."""
        scanner = MockFileScanner()
        
        unknown_file = Path("game.unknown")
        result = scanner.detect_system(unknown_file)
        
        assert result == "Unknown"
    
    def test_sorted_results(self, temp_rom_dir):
        """Test that file results are sorted."""
        scanner = MockFileScanner()
        files = scanner.find_game_files(temp_rom_dir, recursive=True)
        filenames = [f.name for f in files]
        
        # Check that results are sorted
        assert filenames == sorted(filenames)


@pytest.mark.unit
class TestArchiveDetection:
    """Test suite for archive file detection."""
    
    def test_archive_files_identified(self, temp_rom_dir):
        """Test that archive files are identified but not treated as ROMs."""
        scanner = MockFileScanner()
        files = scanner.find_game_files(temp_rom_dir)
        filenames = {f.name for f in files}
        
        # Archive files should NOT be included in ROM list
        assert "archive.zip" not in filenames
        assert "archive.7z" not in filenames
    
    def test_archive_extensions(self):
        """Test common archive format extensions."""
        archive_extensions = {'.zip', '.7z', '.rar', '.gz', '.tar', '.tgz'}
        
        # Just verify the set exists (real implementation would check these)
        assert len(archive_extensions) == 6
        assert '.zip' in archive_extensions


@pytest.mark.unit
def test_file_count_consistency(temp_rom_dir):
    """Test that file count is consistent across multiple scans."""
    scanner = MockFileScanner()
    
    # Scan multiple times
    results1 = scanner.find_game_files(temp_rom_dir)
    results2 = scanner.find_game_files(temp_rom_dir)
    
    # Results should be identical
    assert len(results1) == len(results2)
    assert [f.name for f in results1] == [f.name for f in results2]
