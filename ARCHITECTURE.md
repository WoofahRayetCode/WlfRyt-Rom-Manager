# Architecture Guide - ROM Converter

This document describes the system architecture, design decisions, and key components of ROM Converter.

## Table of Contents

1. [High-Level Overview](#high-level-overview)
2. [Current Architecture (Monolithic)](#current-architecture-monolithic)
3. [Proposed Architecture (Modular)](#proposed-architecture-modular)
4. [Key Components](#key-components)
5. [Conversion Pipeline](#conversion-pipeline)
6. [Threading & Concurrency](#threading--concurrency)
7. [Resource Management](#resource-management)
8. [State Management](#state-management)

---

## High-Level Overview

ROM Converter is a bulk ROM conversion tool that:
- **Scans** directories for game files (ROMs)
- **Detects** file types and target systems
- **Converts** disc images to optimized formats
- **Extracts** compressed archives
- **Monitors** system resources to prevent overload
- **Tracks** progress for crash recovery

```
User Input
   ↓
File Discovery (scan for ROMs)
   ↓
Archive Extraction (if enabled)
   ↓
Format Detection (PS1? PS2? etc)
   ↓
Parallel Conversion (multiple files at once)
   ↓
Resource Monitoring (CPU, RAM, disk I/O)
   ↓
Completion & Cleanup
```

---

## Current Architecture (Monolithic)

### Structure

```
rom_converter.py (436 KB, 12,000+ lines)
├── Imports & Configuration
├── UI Constants (themes, colors, extensions)
├── CollapsibleFrame (custom widget)
└── ROMConverter Class (195 methods)
    ├── __init__() - Setup & dependency checking
    ├── UI Methods - setup_ui(), init_fonts(), set_theme_colors()
    ├── File Discovery - find_game_files(), detect_iso_system()
    ├── Archive Extraction - extract_all_archives(), extract_zip(), extract_7z()
    ├── Conversion - convert_ps1(), convert_ps2(), convert_psp(), etc.
    ├── Tool Management - check_chdman(), download_mame_tools()
    ├── Configuration - load_config(), save_config()
    ├── 3DS Workflow - setup_3ds_workflow_window(), decrypt_3ds_roms()
    ├── Threading - conversion_thread(), process_single_file()
    ├── Resource Monitoring - check_system_resources(), metrics_thread()
    └── Logging - log(), log_async()
```

### Pros
- ✅ Simple for single-developer projects
- ✅ Easy to understand overall flow
- ✅ Fast initial development
- ✅ No circular dependency issues

### Cons
- ❌ Hard to navigate (12,000 lines in one file)
- ❌ Difficult to test individual components
- ❌ Merge conflicts in collaborative development
- ❌ High cyclomatic complexity
- ❌ Tightly coupled concerns (UI, conversion, I/O)

---

## Proposed Architecture (Modular)

### Recommended Structure

```
rom_converter/
├── __main__.py                      # Entry point
├── __init__.py
│
├── ui/
│   ├── __init__.py
│   ├── main_window.py               # Main ROMConverter class (UI only)
│   ├── themes.py                    # THEME_PRESETS, color constants
│   ├── widgets.py                   # CollapsibleFrame, custom widgets
│   └── dialogs/
│       ├── __init__.py
│       ├── threeds_workflow.py      # 3DS workflow window
│       ├── settings_dialog.py       # Settings dialog
│       └── about_dialog.py          # About dialog
│
├── core/
│   ├── __init__.py
│   ├── converter_engine.py          # Main conversion orchestration
│   ├── file_scanner.py              # find_game_files, detect_iso_system
│   ├── resource_monitor.py          # System resource tracking
│   └── progress_tracker.py          # Crash recovery, progress saving
│
├── converters/
│   ├── __init__.py
│   ├── base.py                      # BaseConverter abstract class
│   ├── ps_converters.py             # PS1/PS2/PSP/PS3/PSVita converters
│   ├── nintendo_converters.py       # NES/SNES/N64/GC/Wii converters
│   ├── xbox_converter.py            # Xbox/Xbox 360
│   └── other_converters.py          # Sega, Atari, etc.
│
├── extractors/
│   ├── __init__.py
│   ├── base.py                      # BaseExtractor abstract class
│   ├── zip_extractor.py             # ZIP files
│   ├── sevenzip_extractor.py        # 7Z/RAR files
│   ├── tar_extractor.py             # TAR/TAR.GZ files
│   └── archive_factory.py           # Auto-detect & create extractors
│
├── tools/
│   ├── __init__.py
│   ├── dependency_manager.py        # check_chdman, download_mame_tools
│   ├── platform_detector.py         # Flatpak, immutable distros, etc.
│   ├── ia_client.py                 # Archive.org S3 integration
│   └── threeds/
│       ├── __init__.py
│       ├── decryptor.py             # 3DS decryption (NDecrypt wrapper)
│       └── key_manager.py           # AES key handling
│
├── config/
│   ├── __init__.py
│   ├── schema.py                    # ConversionConfig dataclass
│   ├── loader.py                    # Config load/save with versioning
│   └── migrator.py                  # Config version migrations
│
├── logging_setup.py                 # Structured logging configuration
├── constants.py                     # System extensions, ID patterns, formats
└── exceptions.py                    # Custom exception hierarchy
```

### Migration Benefits

| Aspect | Monolithic | Modular |
|--------|-----------|---------|
| **File Size** | 436 KB (1 file) | ~40 KB average (11 files) |
| **Class Size** | 195 methods | 5-15 methods each |
| **Testability** | Hard | Easy (isolated units) |
| **Reusability** | Limited | High (import converters, extractors) |
| **IDE Navigation** | Slow | Fast (organized packages) |
| **Merge Conflicts** | Common | Rare |
| **Feature Additions** | Risky | Safe |

---

## Key Components

### 1. ROMConverter (UI Class) - `ui/main_window.py`

**Responsibility**: Handle all UI interactions and updates

```python
class ROMConverter:
    def __init__(self, master):
        self.converter_engine = ConverterEngine()
        self.progress_tracker = ProgressTracker()
        self.setup_ui()
    
    def on_start_conversion(self):
        """User clicked 'Start' - delegate to engine"""
        self.converter_engine.convert_async(
            source_dir=self.source_dir.get(),
            settings=self.get_settings(),
            progress_callback=self.update_progress_bar,
            log_callback=self.append_to_log
        )
    
    def update_progress_bar(self, percentage: int):
        """Called by engine as files complete"""
        self.progress_var.set(percentage)
        self.master.update_idletasks()
```

### 2. ConverterEngine (Core Logic) - `core/converter_engine.py`

**Responsibility**: Orchestrate the conversion pipeline

```python
class ConverterEngine:
    def __init__(self):
        self.file_scanner = FileScanner()
        self.resource_monitor = ResourceMonitor()
        self.converters = self._load_converters()
    
    def convert_async(self, source_dir, settings, progress_callback, log_callback):
        """Run conversion in background thread"""
        def run_conversion():
            files = self.file_scanner.find_game_files(source_dir)
            with ThreadPoolExecutor(max_workers=self._optimal_workers()) as executor:
                for file in files:
                    self._convert_file(file, settings)
                    progress_callback(self._calc_progress())
        
        threading.Thread(target=run_conversion, daemon=True).start()
    
    def _convert_file(self, file: Path, settings: ConversionConfig) -> ConversionResult:
        """Convert a single file using appropriate converter"""
        file_type = self._detect_system(file)
        converter = self.converters[file_type]
        return converter.convert(file, settings)
```

### 3. Converter Classes - `converters/`

**Responsibility**: Implement conversion logic for each system

```python
class BaseConverter(ABC):
    """Abstract base for all system converters"""
    
    @abstractmethod
    def can_convert(self, file: Path) -> bool:
        """Check if this converter can handle the file"""
        pass
    
    @abstractmethod
    def convert(self, source: Path, output_dir: Path, **kwargs) -> ConversionResult:
        """Perform the actual conversion"""
        pass

class PS2Converter(BaseConverter):
    def __init__(self, chdman_path: str, max_processors: int = 2):
        self.chdman_path = chdman_path
        self.max_processors = max_processors
    
    def convert(self, source: Path, output_dir: Path, output_format: str = "CHD") -> ConversionResult:
        if output_format == "CHD":
            return self._convert_to_chd(source, output_dir)
        elif output_format == "CSO":
            return self._convert_to_cso(source, output_dir)
        else:
            return ConversionResult(success=False, error="Unsupported format")
    
    def _convert_to_chd(self, source: Path, output_dir: Path) -> ConversionResult:
        output = output_dir / source.stem + ".chd"
        cmd = [
            self.chdman_path, "createcd",
            "-i", str(source),
            "-o", str(output),
            "-f"  # Force overwrite
        ]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=1800)
            if result.returncode == 0:
                return ConversionResult(
                    success=True,
                    output_file=output,
                    original_size=source.stat().st_size,
                    output_size=output.stat().st_size
                )
        except subprocess.TimeoutExpired:
            return ConversionResult(success=False, error="Conversion timeout")
        except Exception as e:
            return ConversionResult(success=False, error=str(e))
```

### 4. Extractor Classes - `extractors/`

**Responsibility**: Extract different archive formats

```python
class BaseExtractor(ABC):
    @abstractmethod
    def can_extract(self, archive: Path) -> bool:
        pass
    
    @abstractmethod
    def extract(self, archive: Path, output_dir: Path) -> ExtractResult:
        pass

class SevenZipExtractor(BaseExtractor):
    def __init__(self, sevenzip_path: str):
        self.sevenzip_path = sevenzip_path
    
    def extract(self, archive: Path, output_dir: Path) -> ExtractResult:
        output_dir.mkdir(parents=True, exist_ok=True)
        cmd = [self.sevenzip_path, "x", str(archive), f"-o{output_dir}"]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
            if result.returncode == 0:
                extracted_files = list(output_dir.rglob("*"))
                return ExtractResult(success=True, files=extracted_files)
            else:
                return ExtractResult(success=False, error=result.stderr)
        except Exception as e:
            return ExtractResult(success=False, error=str(e))
```

### 5. Configuration Management - `config/`

**Responsibility**: Load/save settings with versioning

```python
from dataclasses import dataclass, asdict
from typing import Optional

@dataclass
class ConversionConfig:
    VERSION = 2
    
    # Paths
    source_dir: str = ""
    output_dir: str = ""
    
    # Formats
    ps1_output_format: str = "CHD"
    ps2_output_format: str = "CHD"
    psp_output_format: str = "CSO"
    
    # Processing
    max_concurrent_conversions: Optional[int] = None
    delete_originals: bool = False
    extract_compressed: bool = True
    
    @classmethod
    def from_file(cls, path: Path) -> "ConversionConfig":
        if not path.exists():
            return cls()
        
        with open(path, 'r') as f:
            data = json.load(f)
        
        file_version = data.pop("_version", 1)
        config = cls(**data)
        
        # Apply migrations if needed
        if file_version < cls.VERSION:
            config = migrate_config(file_version, config)
        
        return config
    
    def save(self, path: Path):
        data = asdict(self)
        data["_version"] = self.VERSION
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, 'w') as f:
            json.dump(data, f, indent=2)
```

---

## Conversion Pipeline

### Step-by-Step Flow

```
1. User selects source folder & settings
                ↓
2. File Discovery
   - Scan directory (recursive if enabled)
   - Identify ROM files by extension
   - Filter by selected systems
   - Detect ISO type (PS1 vs PS2 vs Xbox)
                ↓
3. Archive Extraction (if enabled)
   - Identify .zip, .7z, .rar, .tar files
   - Extract to temporary directory
   - Add extracted files to conversion queue
                ↓
4. Conversion Preparation
   - Match files to converters
   - Calculate total size
   - Estimate time
                ↓
5. Parallel Conversion
   - ThreadPoolExecutor with N workers
   - Each worker processes one file at a time
   - Monitor progress and update UI
   - Check system resources before each file
                ↓
6. Resource Monitoring
   - RAM: Throttle at 85%, pause at 92%
   - CPU: Log if exceeding 95%
   - Disk I/O: Limit to 500 MB/s
                ↓
7. Completion & Cleanup
   - Move/delete originals (if requested)
   - Save progress file
   - Display summary (files converted, space saved)
```

### Example: PS2 ISO → CHD Conversion

```
1. Detect File Type
   - File: game.iso (2.5 GB)
   - System: Detect via internal structure → PS2
   - Format: ISO (full disc dump)

2. Select Converter
   - PS2Converter with CHD output format

3. Run Conversion
   $ chdman.exe createcd -i game.iso -o game.chd -f
   
4. Monitor Progress
   - Every 100MB written, update progress bar
   - Estimate: 15 minutes (based on system speed)
   
5. Completion
   - Original: game.iso (2.5 GB)
   - Output: game.chd (1.8 GB) ← 28% smaller!
```

---

## Threading & Concurrency

### Architecture

```
Main UI Thread
    ↓
User clicks "Start Conversion"
    ↓
conversion_thread() spawns in background
    ↓
ThreadPoolExecutor(max_workers=N)
    ├── Worker 1 → Process file 1
    ├── Worker 2 → Process file 2
    ├── Worker 3 → Process file 3
    └── Worker N → Process file N
    
Results queued back to UI thread
    ↓
UI thread updates progress bar via master.after()
```

### Worker Logic

```python
def process_single_file(self, file_path: Path, file_number: int, total: int) -> ProcessResult:
    """Called by thread pool worker"""
    try:
        # Check resources before starting
        if not self.resource_monitor.can_proceed():
            time.sleep(1)  # Wait and retry
            return self.process_single_file(file_path, file_number, total)
        
        # Perform conversion
        converter = self.get_converter_for(file_path)
        result = converter.convert(file_path, self.output_dir, **self.settings)
        
        # Update progress
        self.progress_tracker.mark_complete(file_path)
        
        # Log result
        self.log_queue.put(f"✅ {file_path.name}: {result.output_size} bytes")
        
        return result
    
    except ResourceError:
        # Re-queue for retry
        self.log_queue.put(f"⏸️  {file_path.name}: Waiting for resources...")
        return self.process_single_file(file_path, file_number, total)
    
    except ConversionError as e:
        self.log_queue.put(f"❌ {file_path.name}: {e.message}")
        return ProcessResult(success=False, error=e)
```

### Thread Safety

- **Log Queue**: Thread-safe Queue for logging from workers
- **Progress File**: Atomic writes (write to temp, then rename)
- **Resource Monitor**: Locks for reading system stats
- **Converter Instances**: Stateless (can be shared safely)

---

## Resource Management

### System Monitoring

```
Every 1 second:

┌─────────────────────┐
│ Check RAM Usage     │
└──────┬──────────────┘
       ↓
    < 80%? Continue
    80-85%? Throttle (reduce workers)
    > 85%? Pause new jobs
    > 92%? Force wait

┌─────────────────────┐
│ Check CPU Usage     │
└──────┬──────────────┘
       ↓
    < 95%? Normal speed
    > 95%? Log warning

┌─────────────────────┐
│ Check Disk I/O      │
└──────┬──────────────┘
       ↓
    < 500 MB/s? Normal
    > 500 MB/s? Throttle writes
```

### Auto-Detection

```python
def detect_optimal_workers() -> int:
    """Determine worker count based on system specs"""
    
    total_cores = cpu_count()
    total_ram_gb = virtual_memory().total / (1024 ** 3)
    
    # Each converter uses ~500MB-1GB
    workers_by_ram = total_ram_gb / 1.5
    
    # Use minimum of RAM and CPU limits
    optimal = min(workers_by_ram, total_cores - 1)
    
    return max(1, int(optimal))

# Examples:
#  2-core, 4GB RAM  → 1 worker
#  4-core, 8GB RAM  → 3 workers
#  8-core, 16GB RAM → 6 workers
# 16-core, 32GB RAM → 8 workers (capped)
```

---

## State Management

### Configuration Persistence

```
On Startup:
  1. Check .rom_converter_config.json (portable path)
  2. Check ~/.rom_converter_config.json (home directory)
  3. Use first found, or create new

On Settings Change:
  1. Update ConversionConfig object in memory
  2. Save to appropriate location
  3. Reload UI from config

On Upgrade:
  1. Load config version
  2. Apply migrations if needed
  3. Save with new version number
```

### Progress Tracking (Crash Recovery)

```
During Conversion:
  - Every 5 files: Save progress to .rom_converter_progress.json
  - Track which files completed
  
On Resume:
  1. Load progress file
  2. Filter completed files from queue
  3. Log "Resume mode: Skipping X files"
  4. Continue with remaining files
```

### Progress File Format

```json
{
  "batch_id": "uuid-here",
  "timestamp": "2026-08-02T10:30:00",
  "completed_files": [
    "/path/to/game1.iso",
    "/path/to/game2.iso"
  ],
  "settings": {
    "output_format": "CHD",
    "delete_originals": false
  }
}
```

---

## Error Handling Strategy

### Exception Hierarchy

```python
ROMConverterError (base)
├── ConversionError
│   ├── TimeoutError
│   └── ToolNotFoundError
├── ResourceError
│   ├── OutOfMemoryError
│   └── DiskFullError
├── ConfigError
│   ├── InvalidConfigError
│   └── MissingConfigError
└── ExtractionError
    ├── CorruptArchiveError
    └── UnsupportedFormatError
```

### Error Recovery

| Error Type | Action | Recovery |
|-----------|--------|----------|
| **Timeout** | Log & pause | Retry with extended timeout |
| **Out of Memory** | Pause new jobs | Resume when memory available |
| **Disk Full** | Stop conversion | Ask user to free space |
| **Tool not found** | Stop & offer download | Auto-download & retry |
| **Corrupt archive** | Skip file | Log and continue |

---

## Future Enhancements

### Potential Improvements

1. **Plugin System**: Load converters dynamically
2. **Batch Presets**: "PCSX2 Optimized", "Handheld Best"
3. **Web UI**: Remote control via browser
4. **GPU Acceleration**: Use FFmpeg GPU encoding
5. **Cloud Upload**: Auto-upload to Archive.org
6. **Comparison Mode**: Before/after size/speed analysis

---

## References

- **Threading**: Python `threading`, `concurrent.futures.ThreadPoolExecutor`
- **Resource Monitoring**: `psutil` library
- **UI**: `tkinter` (Python's standard GUI toolkit)
- **External Tools**: `chdman`, `maxcso`, `7z`, `extract-xiso`

---

**Document Version**: 1.0  
**Last Updated**: 2026-08-02  
**Status**: Architecture guide for refactoring planning
