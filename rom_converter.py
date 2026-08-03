#!/usr/bin/env python3
"""
ROM Converter
A GUI tool for bulk converting disc images to CHD format (PS1, PS2, and more)
"""

import os
import sys
import subprocess
import shutil
import json
import zipfile
import tarfile
import urllib.request
import urllib.error
from tkinter import simpledialog
import xml.etree.ElementTree as ET
import gzip
from pathlib import Path
from tkinter import Tk, Frame, Label, Button, Entry, Text, Scrollbar, Checkbutton, BooleanVar, Radiobutton, StringVar, filedialog, messagebox, Toplevel
from tkinter import ttk
import tkinter.font as tkfont
import threading
import re
import multiprocessing
from concurrent.futures import ThreadPoolExecutor, as_completed
from queue import Queue
import time
import threading as pythread
import gc  # For memory management
try:
    import psutil  # For CPU, memory, disk usage
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

try:
    import docx
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# Phase 3: Robustness improvements
try:
    from retry_integration import make_download_retryable, make_extraction_retryable, make_conversion_retryable
    RETRY_INTEGRATION_AVAILABLE = True
except ImportError:
    RETRY_INTEGRATION_AVAILABLE = False

try:
    from logging_queue_bridge import setup_queue_logging_bridge
    LOGGING_QUEUE_BRIDGE_AVAILABLE = True
except ImportError:
    LOGGING_QUEUE_BRIDGE_AVAILABLE = False

# Phase 3 Week 4: Service modularization
try:
    from converter_factory import ConverterFactory
    CONVERTER_FACTORY_AVAILABLE = True
except ImportError:
    CONVERTER_FACTORY_AVAILABLE = False

try:
    from extraction_registry import ExtractionRegistry
    EXTRACTION_REGISTRY_AVAILABLE = True
except ImportError:
    EXTRACTION_REGISTRY_AVAILABLE = False

# Phase 4 Week 1: Tool and config management
try:
    from tool_registry import ToolRegistry, create_default_tool_registry
    TOOL_REGISTRY_AVAILABLE = True
except ImportError:
    TOOL_REGISTRY_AVAILABLE = False

try:
    from config_adapter import ROMConverterConfigAdapter
    CONFIG_ADAPTER_AVAILABLE = True
except ImportError:
    CONFIG_ADAPTER_AVAILABLE = False

# Phase 4 Week 3: Performance optimization
try:
    from performance_optimizer import PerformanceOptimizer, create_performance_optimizer
    PERFORMANCE_OPTIMIZER_AVAILABLE = True
except ImportError:
    PERFORMANCE_OPTIMIZER_AVAILABLE = False

# Phase 4 Week 4: Streaming & error handling
try:
    from streaming_converter import StreamingConverter, create_streaming_converter
    STREAMING_CONVERTER_AVAILABLE = True
except ImportError:
    STREAMING_CONVERTER_AVAILABLE = False

try:
    from conversion_error_handler import ConversionErrorHandler, create_error_handler
    ERROR_HANDLER_AVAILABLE = True
except ImportError:
    ERROR_HANDLER_AVAILABLE = False

# Phase 4 Week 5: Metrics persistence & performance analysis
try:
    from metrics_store import MetricsStore, create_metrics_store
    METRICS_STORE_AVAILABLE = True
except ImportError:
    METRICS_STORE_AVAILABLE = False

try:
    from performance_analyzer import PerformanceAnalyzer, create_performance_analyzer
    PERFORMANCE_ANALYZER_AVAILABLE = True
except ImportError:
    PERFORMANCE_ANALYZER_AVAILABLE = False

# Phase 4 Week 6: Performance status panel (UI widget — import deferred to setup_ui)
PERFORMANCE_PANEL_AVAILABLE = False
try:
    import performance_status_panel as _psp_mod  # noqa: F401
    PERFORMANCE_PANEL_AVAILABLE = True
except ImportError:
    pass

# Phase 5: UI track
try:
    from batch_progress_panel import BatchProgressPanel, FileStatus, create_batch_progress_panel
    BATCH_PROGRESS_AVAILABLE = True
except ImportError:
    BATCH_PROGRESS_AVAILABLE = False

try:
    from performance_settings_panel import PerformanceSettingsPanel, create_performance_settings_panel
    PERF_SETTINGS_PANEL_AVAILABLE = True
except ImportError:
    PERF_SETTINGS_PANEL_AVAILABLE = False

try:
    from session_history_panel import SessionHistoryPanel, create_session_history_panel
    SESSION_HISTORY_PANEL_AVAILABLE = True
except ImportError:
    SESSION_HISTORY_PANEL_AVAILABLE = False

# Phase 5: Format track
try:
    from chd_verifier import CHDVerifier, create_chd_verifier
    CHD_VERIFIER_AVAILABLE = True
except ImportError:
    CHD_VERIFIER_AVAILABLE = False

try:
    from cue_parser import CueParser, CueDisc, DiscSystem, parse_cue
    CUE_PARSER_AVAILABLE = True
except ImportError:
    CUE_PARSER_AVAILABLE = False

try:
    from gamecube_converter import GameCubeConverter, detect_disc_type as gcn_detect_disc_type, is_gamecube_or_wii, create_gamecube_converter
    GAMECUBE_CONVERTER_AVAILABLE = True
except ImportError:
    GAMECUBE_CONVERTER_AVAILABLE = False

try:
    from dat_matcher import DatMatcher, MatchStatus, create_dat_matcher
    DAT_MATCHER_AVAILABLE = True
except ImportError:
    DAT_MATCHER_AVAILABLE = False

# Phase 5 Week 2 - Queue Management (drag-drop)
try:
    from dragdrop_queue_panel import DragDropQueuePanel, create_dragdrop_queue_panel, TKDND_AVAILABLE as DRAGDROP_AVAILABLE
    DRAGDROP_PANEL_AVAILABLE = True
except ImportError:
    DRAGDROP_PANEL_AVAILABLE = False

# Phase 4 Week 1 - Config Manager UI
try:
    from config_manager_ui import (
        ToolStatusDisplay, ConfigValidationDisplay, ConfigExportImport,
        create_tool_status_display, create_config_validation_display, create_config_export_import,
    )
    CONFIG_MANAGER_UI_AVAILABLE = True
except ImportError:
    CONFIG_MANAGER_UI_AVAILABLE = False

# MAME download configuration
MAME_RELEASE_URL = "https://www.mamedev.org/release.html"
MAME_GITHUB_RELEASES_API = "https://api.github.com/repos/mamedev/mame/releases/latest"

# NDecrypt download configuration (for 3DS ROM decryption)
NDECRYPT_GITHUB_RELEASES_API = "https://api.github.com/repos/SabreTools/NDecrypt/releases/latest"

# PS3 Disc Dumper download configuration
PS3_DISC_DUMPER_GITHUB_RELEASES_API = "https://api.github.com/repos/13xforever/ps3-disc-dumper/releases/latest"

# AutoDecryptPS3 download configuration
AUTO_DECRYPT_PS3_GITHUB_RELEASES_API = "https://api.github.com/repos/ChrOn0os/AutoDecryptPS3/releases/latest"

# Supported compressed file extensions
COMPRESSED_EXTENSIONS = {'.zip', '.7z', '.rar', '.gz', '.tar', '.tar.gz', '.tgz'}

# ROM extension to system mapping for archive scanning
SYSTEM_EXTENSIONS = {
    # Sony PlayStation
    '.cue': 'PlayStation',
    '.bin': 'PlayStation',  # Could be PS1/PS2, will be refined by context
    '.iso': 'PlayStation 2',
    '.img': 'PlayStation 2',
    '.chd': 'PlayStation',
    '.psx': 'PlayStation',
    '.pbp': 'PSP',
    '.cso': 'PSP',
    '.zso': 'PSP',
    '.iso': 'PlayStation 3', # Will be refined by detect_iso_system
    # Nintendo Handhelds
    '.gb': 'Game Boy',
    '.gbc': 'Game Boy Color',
    '.gba': 'Game Boy Advance',
    '.sgb': 'Super Game Boy',
    '.nds': 'Nintendo DS',
    '.3ds': 'Nintendo 3DS',
    '.cia': 'Nintendo 3DS',
    # Nintendo Home Consoles
    '.nes': 'NES',
    '.snes': 'SNES',
    '.sfc': 'SNES',
    '.smc': 'SNES',
    '.n64': 'Nintendo 64',
    '.z64': 'Nintendo 64',
    '.v64': 'Nintendo 64',
    '.gcm': 'GameCube',
    '.gcz': 'GameCube',
    '.rvz': 'GameCube/Wii',
    '.wbfs': 'Wii',
    '.wad': 'Wii',
    # Nintendo Switch
    '.xci': 'Nintendo Switch',
    '.nsp': 'Nintendo Switch',
    # Xbox
    '.xiso': 'Xbox',
    '.iso': 'Xbox/Xbox 360', # Added for detection refinement
    # Sega
    '.md': 'Sega Genesis',
    '.gen': 'Sega Genesis',
    '.smd': 'Sega Genesis',
    '.32x': 'Sega 32X',
    '.sms': 'Sega Master System',
    '.gg': 'Sega Game Gear',
    '.cdi': 'Dreamcast',
    '.gdi': 'Dreamcast',
    # Atari
    '.a26': 'Atari 2600',
    '.a78': 'Atari 7800',
    '.lnx': 'Atari Lynx',
    # Other
    '.pce': 'PC Engine',
    '.ngp': 'Neo Geo Pocket',
    '.ngc': 'Neo Geo Pocket Color',
    '.ws': 'WonderSwan',
    '.wsc': 'WonderSwan Color',
    '.vb': 'Virtual Boy',
}

PSP_ID_PATTERNS = ('ulus', 'ules', 'uljm', 'uljs', 'ucus', 'uces', 'uckr', 'ulks')
PS2_ID_PATTERNS = ('slus', 'sles', 'scus', 'sces', 'slpm', 'slps', 'scps')
PS3_ID_PATTERNS = ('blus', 'bles', 'bljm', 'bljs', 'bcus', 'bces', 'bcjs', 'bcas', 'mrts')

# Supported PS1 output formats
PS1_OUTPUT_FORMATS = ['CHD', 'ISO']

# Supported PS2 output formats
PS2_OUTPUT_FORMATS = ['CHD', 'ISO', 'CSO', 'ZSO']

# PS2 emulator presets and their recommended formats
PS2_EMULATORS = ['PCSX2', 'AetherSX2', 'OPL (PS2)']
PS2_EMULATOR_RECOMMENDATIONS = {
    'PCSX2': 'CHD',        # Fast, low CPU overhead
    'AetherSX2': 'CHD',    # Mobile-friendly; CHD generally better than CSO
    'OPL (PS2)': 'ZSO',    # OPL supports ZSO; good size savings with low CPU hit
}

# Theme presets for different PlayStation eras
THEME_PRESETS = {
    'PS1': {
        'bg_dark': '#0e0e15', 'bg_medium': '#161728', 'bg_light': '#212236', 'bg_input': '#1b1c2e',
        'text_primary': '#f7f9ff', 'text_secondary': '#b5deff', 'text_muted': '#9aa4b8',
        'accent_pink': '#ff65a3', 'accent_purple': '#c792ea', 'accent_yellow': '#ffd166',
        'accent_orange': '#ff9f1c', 'accent_red': '#ff5c8d', 'button_green': '#6ce37e',
        'button_blue': '#6fa8ff', 'scanline': '#ffffff08',
        'font_body': 'Courier New', 'font_heading': 'Courier New', 'font_mono': 'Courier New'
    },
    'PS2': {
        'bg_dark': '#0a0b14', 'bg_medium': '#131526', 'bg_light': '#1c1f33', 'bg_input': '#16243a',
        'text_primary': '#d9ffe8', 'text_secondary': '#9ae6ff', 'text_muted': '#7fa4b5',
        'accent_pink': '#ff00aa', 'accent_purple': '#9945ff', 'accent_yellow': '#ffdd00',
        'accent_orange': '#ff6b35', 'accent_red': '#ff3366', 'button_green': '#00cc66',
        'button_blue': '#0088ff', 'scanline': '#ffffff08',
        'font_body': 'Consolas', 'font_heading': 'Consolas', 'font_mono': 'Consolas'
    },
    'PS3': {
        'bg_dark': '#0c0d14', 'bg_medium': '#161826', 'bg_light': '#1f2132', 'bg_input': '#1a1d2b',
        'text_primary': '#eef2ff', 'text_secondary': '#83ceff', 'text_muted': '#96a0b3',
        'accent_pink': '#ff6fa1', 'accent_purple': '#8c7bff', 'accent_yellow': '#ffd479',
        'accent_orange': '#ff8f5a', 'accent_red': '#ff5f6d', 'button_green': '#4ade80',
        'button_blue': '#3ea7ff', 'scanline': '#ffffff08',
        'font_body': 'Segoe UI', 'font_heading': 'Segoe UI Semibold', 'font_mono': 'Consolas'
    },
    'PS4': {
        'bg_dark': '#101937', 'bg_medium': '#1d2744', 'bg_light': '#283758', 'bg_input': '#202d4a',
        'text_primary': '#f7fbff', 'text_secondary': '#9ecbff', 'text_muted': '#8a97ad',
        'accent_pink': '#ff7bba', 'accent_purple': '#8d7bff', 'accent_yellow': '#ffd166',
        'accent_orange': '#f8961e', 'accent_red': '#ef476f', 'button_green': '#5be7a9',
        'button_blue': '#3a86ff', 'scanline': '#ffffff08',
        'font_body': 'Segoe UI', 'font_heading': 'Segoe UI Semibold', 'font_mono': 'Consolas'
    },
    'PS5': {
        'bg_dark': '#111827', 'bg_medium': '#1a2233', 'bg_light': '#242f43', 'bg_input': '#1e2738',
        'text_primary': '#f1f5ff', 'text_secondary': '#88b7ff', 'text_muted': '#8b94a6',
        'accent_pink': '#ec4899', 'accent_purple': '#a855f7', 'accent_yellow': '#fbbf24',
        'accent_orange': '#f97316', 'accent_red': '#ef4444', 'button_green': '#10b981',
        'button_blue': '#3b82f6', 'scanline': '#ffffff08',
        'font_body': 'Segoe UI', 'font_heading': 'Segoe UI Semibold', 'font_mono': 'Consolas'
    },
    'PSP': {
        'bg_dark': '#0e1119', 'bg_medium': '#181c28', 'bg_light': '#222837', 'bg_input': '#1c2232',
        'text_primary': '#f7fbff', 'text_secondary': '#7fe0ff', 'text_muted': '#93a0b3',
        'accent_pink': '#ff6fb7', 'accent_purple': '#9d7bff', 'accent_yellow': '#ffd166',
        'accent_orange': '#ff9f1c', 'accent_red': '#ff5f6d', 'button_green': '#4ade80',
        'button_blue': '#4aa3ff', 'scanline': '#ffffff08',
        'font_body': 'Tahoma', 'font_heading': 'Tahoma', 'font_mono': 'Consolas'
    },
    'PSVita': {
        'bg_dark': '#0d1230', 'bg_medium': '#151c3f', 'bg_light': '#1e2850', 'bg_input': '#182342',
        'text_primary': '#f5f8ff', 'text_secondary': '#9ad5ff', 'text_muted': '#93a1b8',
        'accent_pink': '#ff7ba5', 'accent_purple': '#8b7bff', 'accent_yellow': '#ffd479',
        'accent_orange': '#ff9e5a', 'accent_red': '#ff6b81', 'button_green': '#5de0a4',
        'button_blue': '#4d9dff', 'scanline': '#ffffff08',
        'font_body': 'Segoe UI', 'font_heading': 'Segoe UI Semibold', 'font_mono': 'Consolas'
    },
}

# Active colors (will be set from selected theme)
COLORS = dict(THEME_PRESETS['PS2'])


class CollapsibleFrame(Frame):
    """A frame with a clickable header that expands/collapses its content."""
    def __init__(self, parent, title="", font=None, fg="white", bg=None, collapsed=False, **kwargs):
        bg = bg or COLORS.get('bg_medium', '#1a1a2e')
        super().__init__(parent, bg=bg, bd=0, highlightthickness=0, **kwargs)
        self._collapsed = collapsed
        self._header_bg = bg
        self._title = title
        self._header_font = font
        self._header_fg = fg

        self._header_frame = Frame(self, bg=bg, cursor="hand2")
        self._header_frame.pack(fill="x")

        self._indicator = Label(self._header_frame, text="▼" if not collapsed else "▶",
                                font=font, fg=fg, bg=bg, width=2)
        self._indicator.pack(side="left", padx=(4, 0))

        self._title_label = Label(self._header_frame, text=title, font=font, fg=fg, bg=bg, anchor="w")
        self._title_label.pack(side="left", fill="x", expand=True)

        self._content_frame = Frame(self, bg=bg)
        if not collapsed:
            self._content_frame.pack(fill="x", padx=8, pady=(0, 4))

        self._header_frame.bind("<Button-1>", self._toggle)
        self._indicator.bind("<Button-1>", self._toggle)
        self._title_label.bind("<Button-1>", self._toggle)

    @property
    def content(self):
        return self._content_frame

    @property
    def collapsed(self):
        return self._collapsed

    def _toggle(self, event=None):
        if self._collapsed:
            self._content_frame.pack(fill="x", padx=8, pady=(0, 4))
            self._indicator.config(text="▼")
        else:
            self._content_frame.pack_forget()
            self._indicator.config(text="▶")
        self._collapsed = not self._collapsed

    def expand(self):
        if self._collapsed:
            self._toggle()

    def collapse(self):
        if not self._collapsed:
            self._toggle()

    def set_colors(self, header_bg, title_fg):
        self._header_bg = header_bg
        self._header_frame.configure(bg=header_bg)
        self._indicator.configure(bg=header_bg, fg=title_fg)
        self._title_label.configure(bg=header_bg, fg=title_fg)
        self._content_frame.configure(bg=header_bg)
        self.configure(bg=header_bg)


class ROMConverter:
    def __init__(self, master):
        self.master = master
        master.title("⚡ ROM CONVERTER ⚡")
        # Load persistent window size from config, fallback to 1100x800
        win_w, win_h = self._load_window_size()
        master.geometry(f"{win_w}x{win_h}")
        master.minsize(950, 650)
        master.resizable(True, True)
        master.configure(bg=COLORS['bg_dark'])
        master.bind("<Configure>", self._on_window_resize)

        # Drag-and-drop ROM directory (requires tkinterdnd2; silently skip if unavailable)
        try:
            master.drop_target_register("DND_Files")
            master.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass

        # Log panel visibility tracking
        self.log_visible = False

        # Maximize window on startup (cross-platform)
        try:
            master.state('zoomed')  # Windows
        except:
            master.attributes('-zoomed', True)  # Linux
        
        # Config file location (portable: lives beside the app)
        # Handle PyInstaller's temporary folder for bundled resources
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            # Check if running from AppImage (read-only mount)
            appimage_path = os.environ.get('APPIMAGE')
            if appimage_path:
                # AppImage: use directory containing the .AppImage file
                self.script_dir = Path(appimage_path).parent.resolve()
            else:
                self.script_dir = Path(sys.executable).parent.resolve()
            # PyInstaller extracts bundled files to sys._MEIPASS
            self.bundle_dir = Path(getattr(sys, '_MEIPASS', self.script_dir))
        else:
            # Running as script
            self.script_dir = Path(__file__).parent.resolve()
            self.bundle_dir = self.script_dir
        
        self.config_candidates = [
            self.script_dir / ".rom_converter_config.json",
            Path.home() / ".rom_converter_config.json"
        ]
        self.config_file = next((p for p in self.config_candidates if p.exists()), self.config_candidates[0])
        
        # Variables
        self.source_dir = ""
        self.delete_originals = BooleanVar(value=False)
        self.move_to_backup = BooleanVar(value=True)
        self.recursive = BooleanVar(value=True)
        self.is_converting = False
        # Keep one CPU core free for system responsiveness
        total_cores = multiprocessing.cpu_count()
        self.cpu_cores = max(1, total_cores - 1)
        self.max_workers = self.cpu_cores  # Dynamic worker count
        self.max_concurrent_conversions = self._detect_optimal_workers()  # Auto-detect based on system specs
        self.ram_threshold_percent = 80  # Throttle if RAM usage exceeds this
        self.ram_critical_percent = 85  # Pause new conversions if RAM exceeds this
        self.ram_hard_limit_percent = 92  # Absolute max - force wait if exceeded
        self.cpu_threshold_percent = 95  # Throttle if CPU usage exceeds this
        self.disk_write_throttle_mb_s = 500  # Throttle disk writes if exceeding this rate (MB/s) - raised for NVMe
        self.disk_io_check_interval = 0.5  # How often to check disk I/O (seconds)
        self.last_disk_throttle_check = 0  # Timestamp of last disk throttle check
        self.chdman_max_processors = self._detect_chdman_processors()  # Auto-detect based on RAM
        self.maxcso_threads = self._detect_maxcso_threads()  # Auto-detect based on CPU/RAM
        self.conversion_semaphore = None  # Will be initialized when conversions start
        self.log_queue = Queue()
        self.total_original_size = 0
        self.total_chd_size = 0
        self.process_ps1_cues = BooleanVar(value=False)  # Toggle for PS1 CUE processing
        self.process_ps2_cues = BooleanVar(value=False)  # Toggle for PS2 BIN/CUE processing (CD-based games)
        self.process_ps2_isos = BooleanVar(value=False)  # Toggle for PS2 ISO processing
        self.process_psp_isos = BooleanVar(value=False)  # Toggle for PSP ISO processing
        self.process_ps3_isos = BooleanVar(value=False)  # Toggle for PS3 ISO processing
        self.process_xbox_isos = BooleanVar(value=False)  # Toggle for Xbox ISO processing
        self.process_nes_roms = BooleanVar(value=False)  # Toggle for NES ROM processing
        self.process_snes_roms = BooleanVar(value=False)  # Toggle for SNES ROM processing
        self.process_n64_roms = BooleanVar(value=False)  # Toggle for N64 ROM processing
        self.extract_compressed = BooleanVar(value=True)  # Toggle for extracting compressed files
        self.delete_archives_after_extract = BooleanVar(value=False)  # Delete archives after extraction
        self.seven_zip_path = None  # Path to 7z executable for .7z and .rar files
        self.maxcso_path = None  # Path to maxcso executable for CSO/ZSO
        self.ndecrypt_path = None  # Path to NDecrypt executable for 3DS decryption
        self.ps3_dumper_path = None  # Path to ps3-disc-dumper executable
        self.auto_decrypt_ps3_path = None  # Path to AutoDecrypt.ps1
        self.extract_xiso_path = None  # Path to extract-xiso executable
        self.ps1_output_format = 'CHD'  # Default PS1 output format
        self.ps2_output_format = 'CHD'  # Default PS2 output format
        self.psp_output_format = 'CSO'  # Default PSP output format
        self.ps2_emulator = 'PCSX2'  # Default emulator preference
        self.current_theme = 'PS2'  # Default UI theme
        self.system_extract_dirs = {}  # Persistent mapping of system -> extraction directory
        # 3DS workflow settings
        self.threeds_backup_original = True
        self.threeds_delete_archives = False
        self.threeds_delete_after_move = False
        self.threeds_auto_clean_names = True
        self.threeds_source_dir = ""
        self.threeds_dest_dir = ""
        self.font_body_family = None
        self.font_heading_family = None
        self.font_mono_family = None
        # Metrics / ETA tracking
        self.total_jobs = 0
        self.completed_jobs = 0
        self.file_start_times = {}
        self.file_durations = []
        self.conversion_start_time = None
        self.initial_disk_write_bytes = 0
        self.last_disk_write_bytes = 0
        self.metrics_running = False
        self.metrics_lock = pythread.Lock()
        self.last_ui_update = 0  # Throttle UI updates
        self.chdman_path = None  # Will store path to chdman executable
        self.build_timestamp = self.get_build_timestamp()
        
        # Archive.org session
        self.ia_session = None
        self.ia_username = ""
        self.ia_s3_keys = None # (access, secret)
        
        # Progress tracking for crash recovery
        self.progress_file = self.script_dir / ".rom_converter_progress.json"
        self.completed_files = set()  # Track completed conversions
        self.current_batch_id = None
        
        # Phase 4 Week 2: Initialize ToolRegistry for centralized tool management
        self.tool_registry = None
        self.config_adapter = None
        
        if TOOL_REGISTRY_AVAILABLE:
            try:
                # Use print for now since log method isn't set up yet
                self.tool_registry = create_default_tool_registry(log_callback=print)
                self.tool_registry.check_all()
                print("✅ ToolRegistry initialized successfully")
            except Exception as e:
                print(f"⚠️  Failed to initialize ToolRegistry: {e}")
                self.tool_registry = None
        else:
            print("⚠️  ToolRegistry not available - using fallback tool detection")
        
        # Phase 4 Week 2: Initialize ConfigAdapter for centralized config management
        if CONFIG_ADAPTER_AVAILABLE:
            try:
                self.config_adapter = ROMConverterConfigAdapter(
                    self,
                    config_file=self.config_file,
                    log_callback=print
                )
                self.config_adapter.init_config_manager()
                print("✅ ConfigAdapter initialized successfully")
            except Exception as e:
                print(f"⚠️  Failed to initialize ConfigAdapter: {e}")
                self.config_adapter = None
        else:
            print("⚠️  ConfigAdapter not available - using fallback config handling")
        
        # Phase 4 Week 3: Initialize PerformanceOptimizer for parallel conversions
        self.performance_optimizer = None
        if PERFORMANCE_OPTIMIZER_AVAILABLE:
            try:
                self.performance_optimizer = create_performance_optimizer(
                    log_callback=print
                )
                print("✅ PerformanceOptimizer initialized successfully")
            except Exception as e:
                print(f"⚠️  Failed to initialize PerformanceOptimizer: {e}")
                self.performance_optimizer = None
        else:
            print("⚠️  PerformanceOptimizer not available - using single-threaded conversion")

        # Phase 4 Week 4: Initialize StreamingConverter
        self.streaming_converter = None
        if STREAMING_CONVERTER_AVAILABLE:
            try:
                self.streaming_converter = create_streaming_converter(log_callback=print)
                print("✅ StreamingConverter initialized successfully")
            except Exception as e:
                print(f"⚠️  Failed to initialize StreamingConverter: {e}")
        else:
            print("⚠️  StreamingConverter not available - using standard I/O")

        # Phase 4 Week 4: Initialize ConversionErrorHandler
        self.error_handler = None
        if ERROR_HANDLER_AVAILABLE:
            try:
                self.error_handler = create_error_handler(log_callback=print)
                print("✅ ConversionErrorHandler initialized successfully")
            except Exception as e:
                print(f"⚠️  Failed to initialize ConversionErrorHandler: {e}")
        else:
            print("⚠️  ConversionErrorHandler not available - using basic error handling")

        # Phase 4 Week 5: Initialize MetricsStore and PerformanceAnalyzer
        self.metrics_store = None
        if METRICS_STORE_AVAILABLE:
            try:
                self.metrics_store = create_metrics_store(log_callback=print)
                print("✅ MetricsStore initialized successfully")
            except Exception as e:
                print(f"⚠️  Failed to initialize MetricsStore: {e}")
        else:
            print("⚠️  MetricsStore not available - metrics will not be persisted")

        self.performance_analyzer = None
        if PERFORMANCE_ANALYZER_AVAILABLE and self.metrics_store is not None:
            try:
                self.performance_analyzer = create_performance_analyzer(
                    store=self.metrics_store, log_callback=print
                )
                print("✅ PerformanceAnalyzer initialized successfully")
            except Exception as e:
                print(f"⚠️  Failed to initialize PerformanceAnalyzer: {e}")
        elif PERFORMANCE_ANALYZER_AVAILABLE:
            try:
                self.performance_analyzer = create_performance_analyzer(log_callback=print)
                print("✅ PerformanceAnalyzer initialized (in-memory store)")
            except Exception as e:
                print(f"⚠️  Failed to initialize PerformanceAnalyzer: {e}")

        # Phase 5: Initialize CHDVerifier (format track)
        self.chd_verifier = None
        if CHD_VERIFIER_AVAILABLE:
            try:
                self.chd_verifier = create_chd_verifier(
                    chdman_path=getattr(self, "chdman_path", None),
                    log_callback=print,
                )
                print("✅ CHDVerifier initialized successfully")
            except Exception as e:
                print(f"⚠️  Failed to initialize CHDVerifier: {e}")
        else:
            print("⚠️  CHDVerifier not available - CHD output will not be verified")

        # Phase 5: CueParser available (stateless, no init needed)
        self._cue_parser_available = CUE_PARSER_AVAILABLE

        # Phase 5 Week 2: GameCube converter
        self.gamecube_converter = None
        if GAMECUBE_CONVERTER_AVAILABLE:
            try:
                self.gamecube_converter = create_gamecube_converter(log_callback=print)
                print("✅ GameCubeConverter initialized successfully")
            except Exception as e:
                print(f"⚠️  Failed to initialize GameCubeConverter: {e}")

        # Phase 5 Week 2: DAT matcher
        self.dat_matcher = None
        if DAT_MATCHER_AVAILABLE:
            try:
                self.dat_matcher = create_dat_matcher(log_callback=print)
                print("✅ DatMatcher initialized successfully")
            except Exception as e:
                print(f"⚠️  Failed to initialize DatMatcher: {e}")

        # Phase 5 Week 2: UI panels (created in setup_ui)
        self._perf_settings_panel = None
        self._session_history_panel = None
        self._dragdrop_queue_panel = None

        # Phase 4 Week 1: Config Manager UI panels (created in setup_ui)
        self._tool_status_display = None
        self._config_validation_display = None
        self._config_export_import = None

        # Phase 5: BatchProgressPanel created in setup_ui
        self._batch_progress_panel = None

        # Load saved configuration
        self.load_config()
        self.load_progress()

        # Apply theme colors before UI construction
        self.set_theme_colors(self.current_theme)
        self.init_fonts()
        
        # Check for 7-Zip (for .7z and .rar support)
        self.check_7zip()

        # Auto-pull essential dependencies (3DS keys, maxcso)
        self.auto_pull_dependencies()
        
        # Detect and auto-login with IA S3 keys
        self.detect_ia_keys()

        # Check for maxcso (for CSO/ZSO output)
        self.check_maxcso()
        
        # Check for extract-xiso (for Xbox/Xbox 360 extraction)
        self.check_extract_xiso()
        
        # Check for ps3-disc-dumper
        if not self.check_ps3_dumper():
            # Automatically download PS3 Disc Dumper if missing
            self.download_ps3_dumper()
        else:
            # Check for updates
            self.check_for_ps3_dumper_update()

        # Check for AutoDecryptPS3
        if not self.check_auto_decrypt_ps3():
            self.download_auto_decrypt_ps3()
        else:
            self.check_for_auto_decrypt_ps3_update()

        # Check for chdman
        if not self.check_chdman():
            # Hide main window and force dialog to front on Linux
            master.withdraw()
            master.update()
            
            # Offer to download MAME tools or manual selection
            response = messagebox.askyesnocancel(
                "chdman Not Found",
                "chdman not found!\n\n"
                "Would you like to download MAME tools automatically?\n\n"
                "Yes = Download from mamedev.org\n"
                "No = Manually select chdman.exe\n",
                parent=master
            )
            
            # Show main window again
            master.deiconify()
            master.update()
            
            if response is True:  # Yes - download
                if self.download_mame_tools():
                    if not self.check_chdman():
                        messagebox.showerror("Error", "Failed to find chdman after download", parent=master)
                        master.destroy()
                        return
                else:
                    master.destroy()
                    return
            elif response is False:  # No - manual selection
                self.browse_chdman()
                if not self.chdman_path:
                    master.destroy()
                    return
            else:  # Cancel
                master.destroy()
                return
        else:
            # chdman found - check for updates
            self.check_for_chdman_update()
        
        # Phase 3 Week 4: Initialize service factories
        self._init_service_factories()
        
        self.setup_ui()
        
        # Phase 3: Setup logging queue bridge for robustness
        self._setup_retry_wrappers()

    def _setup_retry_wrappers(self):
        """Setup retry logic wrappers for critical methods.
        
        This wraps download, extraction, and conversion methods with exponential backoff
        to handle transient errors (timeouts, network issues, temporary resource constraints).
        """
        if not RETRY_INTEGRATION_AVAILABLE:
            self.log("⚠️  Retry integration not available - skipping robustness setup")
            return
        
        # Wrap download methods with 5 attempts, 2s base delay, 120s max
        self.download_mame_tools = make_download_retryable(
            self.download_mame_tools, 
            name="download_mame_tools",
            logger=None  # Will use built-in logger via self.log
        )
        self.download_7zip = make_download_retryable(
            self.download_7zip,
            name="download_7zip",
            logger=None
        )
        self.download_maxcso = make_download_retryable(
            self.download_maxcso,
            name="download_maxcso",
            logger=None
        )
        self.download_ndecrypt = make_download_retryable(
            self.download_ndecrypt,
            name="download_ndecrypt",
            logger=None
        )
        self.download_ps3_dumper = make_download_retryable(
            self.download_ps3_dumper,
            name="download_ps3_dumper",
            logger=None
        )
        self.download_auto_decrypt_ps3 = make_download_retryable(
            self.download_auto_decrypt_ps3,
            name="download_auto_decrypt_ps3",
            logger=None
        )
        self.download_extract_xiso = make_download_retryable(
            self.download_extract_xiso,
            name="download_extract_xiso",
            logger=None
        )
        
        # Wrap extraction methods with 3 attempts, 1s base delay, 60s max
        self.extract_archive = make_extraction_retryable(
            self.extract_archive,
            name="extract_archive",
            logger=None
        )
        self.extract_xbox_iso = make_extraction_retryable(
            self.extract_xbox_iso,
            name="extract_xbox_iso",
            logger=None
        )
        
        # Wrap conversion methods with 2 attempts, 5s base delay, 120s max
        self.convert_game = make_conversion_retryable(
            self.convert_game,
            name="convert_game",
            logger=None
        )
        
        self.log("✅ Retry logic wrappers installed - robustness enabled")

    def _init_service_factories(self):
        """Initialize ConverterFactory and ExtractionRegistry for modularized services"""
        if not CONVERTER_FACTORY_AVAILABLE:
            self.log("⚠️  ConverterFactory not available - using fallback convert_game")
            self.converter_factory = None
        else:
            self.converter_factory = ConverterFactory(
                chdman_path=Path(self.chdman_path) if self.chdman_path else None,
                maxcso_path=Path(self.maxcso_path) if self.maxcso_path else None,
                ps3_dumper_path=Path(self.ps3_dumper_path) if self.ps3_dumper_path else None,
                extract_xiso_path=Path(self.extract_xiso_path) if self.extract_xiso_path else None,
                ndecrypt_path=Path(self.ndecrypt_path) if self.ndecrypt_path else None,
                log_callback=self.log
            )
            self.log("✅ ConverterFactory initialized")
        
        if not EXTRACTION_REGISTRY_AVAILABLE:
            self.log("⚠️  ExtractionRegistry not available - using fallback extract_archive")
            self.extraction_registry = None
        else:
            self.extraction_registry = ExtractionRegistry(
                seven_zip_path=Path(self.seven_zip_path) if self.seven_zip_path else None,
                log_callback=self.log
            )
            self.log("✅ ExtractionRegistry initialized")

    def get_build_timestamp(self):
        """Return build timestamp for About dialog."""
        env_ts = os.environ.get("BUILD_TIMESTAMP")
        if env_ts:
            cleaned = env_ts.strip()
            if cleaned.isdigit():
                try:
                    ts = int(cleaned)
                    return time.strftime("%Y-%m-%d %H:%M:%S %Z", time.localtime(ts))
                except Exception:
                    return cleaned
            return cleaned
        try:
            target_path = Path(sys.executable if getattr(sys, "frozen", False) else __file__)
            mtime = target_path.stat().st_mtime
            return time.strftime("%Y-%m-%d %H:%M:%S %Z", time.localtime(mtime))
        except Exception:
            return "Unknown"
    
    def auto_pull_dependencies(self):
        """Automatically download essential project files if missing from app directory.
        Platform-aware: downloads .exe on Windows, native binaries on Linux."""
        base_url = "https://raw.githubusercontent.com/WoofahRayetCode/WlfRyt-Rom-Manager/main"

        # 3DS Keys filename has spaces - URL encode them as %20
        deps = [
            ("3DS AES Keys.txt", f"{base_url}/3DS%20AES%20Keys.txt"),
        ]

        if sys.platform == "win32":
            deps.append(("maxcso.exe", f"{base_url}/maxcso.exe"))
            deps.append(("extract-xiso.exe", f"{base_url}/extract-xiso.exe"))
        else:
            deps.append(("maxcso", f"{base_url}/maxcso"))
            deps.append(("extract-xiso", f"{base_url}/extract-xiso"))
        
        missing = []
        for filename, url in deps:
            local_path = self.script_dir / filename
            if not local_path.exists():
                missing.append((filename, url, local_path))
        
        if not missing:
            return

        # Perform downloads in a background thread to not block UI startup
        def run_downloads():
            for filename, url, local_path in missing:
                try:
                    # Use requests if available, fallback to urllib
                    if REQUESTS_AVAILABLE:
                        response = requests.get(url, timeout=30)
                        response.raise_for_status()
                        with open(local_path, "wb") as f:
                            f.write(response.content)
                    else:
                        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                        with urllib.request.urlopen(req, timeout=30) as response:
                            with open(local_path, "wb") as f:
                                shutil.copyfileobj(response, f)
                    
                    # Make binaries executable on Linux
                    if sys.platform != "win32" and filename in ("maxcso", "extract-xiso"):
                        try:
                            os.chmod(str(local_path), 0o755)
                        except Exception:
                            pass

                    # Log success if UI is ready, otherwise just console
                    msg = f"✅ Auto-pulled missing dependency: {filename}"
                    print(msg)
                    try:
                        self.log(msg)
                    except:
                        pass
                except Exception as e:
                    msg = f"⚠️ Failed to auto-pull {filename}: {e}"
                    print(msg)
                    try:
                        self.log(msg)
                    except:
                        pass
                        
        # Using a daemon thread so it doesn't block closing the app if it hangs
        threading.Thread(target=run_downloads, daemon=True).start()

    def detect_ia_keys(self):
        """Scan application directory for .txt or .docx files containing S3 keys"""
        try:
            # Check .txt files
            for txt_file in self.script_dir.glob("*.txt"):
                if txt_file.name == "3DS AES Keys.txt": continue
                try:
                    with open(txt_file, 'r', encoding='utf-8', errors='ignore') as f:
                        content = f.read()
                    if self._parse_and_login_ia(content):
                        self.log(f"🔑 Auto-logged into Archive.org using {txt_file.name}")
                        return True
                except:
                    pass
            
            # Check .docx files
            if DOCX_AVAILABLE:
                for docx_file in self.script_dir.glob("*.docx"):
                    try:
                        doc = docx.Document(docx_file)
                        content = "\n".join([p.text for p in doc.paragraphs])
                        if self._parse_and_login_ia(content):
                            self.log(f"🔑 Auto-logged into Archive.org using {docx_file.name}")
                            return True
                    except:
                        pass
        except Exception as e:
            print(f"Error during IA key detection: {e}")
        return False

    def _parse_and_login_ia(self, text):
        """Extract keys from text and attempt login"""
        # Look for access_key: [key] and secret_key: [key]
        access_match = re.search(r'access_key:\s*([a-zA-Z0-9]+)', text, re.IGNORECASE)
        secret_match = re.search(r'secret_key:\s*([a-zA-Z0-9]+)', text, re.IGNORECASE)
        
        if access_match and secret_match:
            access = access_match.group(1)
            secret = secret_match.group(1)
            
            # Basic verification (quiet)
            headers = {"Authorization": f"LOW {access}:{secret}"}
            try:
                res = requests.get("https://archive.org/metadata/@me", headers=headers, timeout=10)
                if res.status_code == 200:
                    self.ia_s3_keys = (access, secret)
                    self.ia_username = res.json().get('username', 'S3 User')
                    return True
            except:
                pass
        return False

    def check_chdman(self):
        """Check if chdman is available"""
        # Determine platform-specific binary name
        chdman_name = "chdman.exe" if sys.platform == "win32" else "chdman"
        
        # First check bundled resources (PyInstaller)
        bundled_chdman = self.bundle_dir / chdman_name
        if bundled_chdman.exists():
            self.chdman_path = str(bundled_chdman)
            return True
        
        # Then check for chdman directly next to the executable/script
        direct_chdman = self.script_dir / chdman_name
        if direct_chdman.exists():
            self.chdman_path = str(direct_chdman)
            return True
        
        # Check PATH as fallback
        chdman = shutil.which("chdman")
        if chdman:
            self.chdman_path = chdman
            return True
        
        # On Linux, check for chdman inside Flatpak MAME installation
        if sys.platform != "win32":
            flatpak_chdman = self._find_flatpak_chdman()
            if flatpak_chdman:
                self.chdman_path = flatpak_chdman
                return True
        
        return False
    
    def _find_flatpak_chdman(self):
        """Search for chdman inside Flatpak MAME installations"""
        import glob
        
        # Try using 'flatpak info' to locate the MAME install
        try:
            result = subprocess.run(
                ['flatpak', 'info', '--show-location', 'org.mamedev.MAME'],
                capture_output=True, text=True, timeout=10
            )
            if result.returncode == 0:
                location = result.stdout.strip()
                chdman_path = os.path.join(location, 'files', 'bin', 'chdman')
                if os.path.isfile(chdman_path) and os.access(chdman_path, os.X_OK):
                    return chdman_path
        except (FileNotFoundError, subprocess.TimeoutExpired, Exception):
            pass
        
        # Fallback: search common Flatpak install directories
        flatpak_search_dirs = [
            '/var/lib/flatpak/app/org.mamedev.MAME',
            os.path.expanduser('~/.local/share/flatpak/app/org.mamedev.MAME'),
        ]
        for base_dir in flatpak_search_dirs:
            pattern = os.path.join(base_dir, '**', 'files', 'bin', 'chdman')
            matches = glob.glob(pattern, recursive=True)
            for match in matches:
                if os.path.isfile(match) and os.access(match, os.X_OK):
                    return match
        
        return None
    
    def _is_immutable_distro(self):
        """Detect if running on an immutable/atomic Linux distro (Bazzite, SteamOS, Kinoite, etc.)"""
        try:
            with open('/etc/os-release', 'r') as f:
                os_release = f.read().lower()
            # Check for known immutable distros
            immutable_ids = ['bazzite', 'steamos', 'silverblue', 'kinoite', 'aurora', 'bluefin', 'ublue']
            for name in immutable_ids:
                if name in os_release:
                    return True
            # Check for rpm-ostree (generic Fedora Atomic indicator)
            if shutil.which('rpm-ostree'):
                return True
        except Exception:
            pass
        return False
    
    def _get_distro_id(self):
        """Get the distro ID from os-release"""
        try:
            with open('/etc/os-release', 'r') as f:
                for line in f:
                    if line.startswith('ID='):
                        return line.strip().split('=', 1)[1].strip('"').lower()
                    if line.startswith('ID_LIKE='):
                        return line.strip().split('=', 1)[1].strip('"').lower()
        except Exception:
            pass
        return ''
    
    def _get_chdman_install_instructions(self):
        """Generate OS-aware chdman installation instructions"""
        msg = "Could not download chdman automatically.\n\n"
        
        if self._is_immutable_distro():
            # Immutable distro (Bazzite, SteamOS, Kinoite, etc.)
            msg += (
                "Install via Flatpak (recommended):\n"
                "   flatpak install --user flathub org.mamedev.MAME\n\n"
                "Or install EmuDeck (includes chdman).\n"
            )
        else:
            distro = self._get_distro_id()
            msg += "Install chdman using your package manager:\n\n"
            if 'arch' in distro or 'manjaro' in distro:
                msg += "   sudo pacman -S mame-tools\n"
            elif 'debian' in distro or 'ubuntu' in distro or 'mint' in distro:
                msg += "   sudo apt install mame-tools\n"
            elif 'fedora' in distro or 'rhel' in distro or 'centos' in distro:
                msg += "   sudo dnf install mame-tools\n"
            else:
                msg += (
                    "   sudo pacman -S mame-tools  (Arch)\n"
                    "   sudo apt install mame-tools  (Debian/Ubuntu)\n"
                    "   sudo dnf install mame-tools  (Fedora)\n"
                )
            msg += (
                "\nOr install via Flatpak:\n"
                "   flatpak install flathub org.mamedev.MAME\n"
            )
        
        msg += "\nThen restart this application."
        return msg
    
    def get_installed_chdman_version(self):
        """Get the version of the currently installed chdman"""
        if not self.chdman_path:
            return None
        
        try:
            result = subprocess.run(
                [self.chdman_path, '--version'],
                capture_output=True, text=True, timeout=10
            )
            # Output typically like: "chdman - MAME Compressed Hunks of Data (CHD) manager 0.271 (mame0271)"
            # or "chdman - MAME ... 0.283 (mame0283)"
            output = result.stdout + result.stderr
            
            # Look for version pattern like "0.283" or "(mame0283)"
            version_match = re.search(r'\(mame(\d{4})\)', output)
            if version_match:
                return version_match.group(1)
            
            # Alternative: look for version number like "0.283"
            version_match = re.search(r'(\d+\.\d+)', output)
            if version_match:
                ver = version_match.group(1).replace('.', '')
                return ver.zfill(4)
                
        except Exception as e:
            print(f"Error getting chdman version: {e}")
        
        return None
    
    def check_for_chdman_update(self):
        """Check if a newer version of chdman is available"""
        try:
            installed_version = self.get_installed_chdman_version()
            if not installed_version:
                return  # Can't determine installed version, skip update check
            
            latest_version = self.get_latest_mame_version()
            if not latest_version:
                return  # Can't fetch latest version, skip update check
            
            # Compare versions (they're 4-digit strings like "0283")
            if int(latest_version) > int(installed_version):
                # Ensure dialog appears in front
                self.master.lift()
                self.master.focus_force()
                response = messagebox.askyesno(
                    "chdman Update Available",
                    f"A newer version of chdman is available!\n\n"
                    f"Installed: MAME {installed_version[0]}.{installed_version[1:]}\n"
                    f"Latest: MAME {latest_version[0]}.{latest_version[1:]}\n\n"
                    "Would you like to update now?",
                    parent=self.master
                )
                if response:
                    self.download_mame_tools()
        except Exception as e:
            print(f"Error checking for chdman update: {e}")
    
    def get_latest_mame_version(self):
        """Fetch the latest MAME version from mamedev.org"""
        try:
            # Parse the release page to get version number
            req = urllib.request.Request(
                MAME_RELEASE_URL,
                headers={'User-Agent': 'Mozilla/5.0'}
            )
            with urllib.request.urlopen(req, timeout=30) as response:
                html = response.read().decode('utf-8')
            
            # Look for version pattern like "mame0283" or "MAME 0.283"
            version_match = re.search(r'mame(\d{4})', html, re.IGNORECASE)
            if version_match:
                return version_match.group(1)
            
            # Alternative pattern
            version_match = re.search(r'MAME\s+(\d+\.\d+)', html)
            if version_match:
                # Convert 0.283 to 0283
                ver = version_match.group(1).replace('.', '')
                return ver.zfill(4)
            
        except Exception as e:
            print(f"Error fetching MAME version: {e}")
        
        return None
    
    def download_mame_tools(self):
        """Download and extract MAME tools"""
        try:
            # On Linux, use a different approach - download prebuilt or guide user
            if sys.platform != "win32":
                return self.download_mame_tools_linux()
            
            # Windows: Download from MAME releases
            # Get latest version
            version = self.get_latest_mame_version()
            if not version:
                messagebox.showerror("Error", "Could not determine latest MAME version", parent=self.master)
                return False
            
            # Construct download URL for Windows x64 binary
            # Format: mame0283b_x64.exe (self-extracting 7z)
            download_url = f"https://github.com/mamedev/mame/releases/download/mame{version}/mame{version}b_64bit.exe"
            
            # Alternative URL patterns to try
            alt_urls = [
                f"https://github.com/mamedev/mame/releases/download/mame{version}/mame{version}b_x64.exe",
                f"https://github.com/mamedev/mame/releases/download/mame{version}/mame{version}b_64bit.exe",
            ]
            
            temp_dir = self.script_dir / "mame_temp"
            temp_dir.mkdir(exist_ok=True)
            
            download_path = temp_dir / f"mame{version}.exe"
            
            # Show progress dialog - ensure it appears in front
            progress_window = Toplevel(self.master)
            progress_window.title("Downloading MAME Tools")
            progress_window.geometry("400x150")
            progress_window.resizable(False, False)
            progress_window.transient(self.master)
            progress_window.grab_set()
            progress_window.lift()
            progress_window.focus_force()
            
            Label(progress_window, text=f"Downloading MAME {version} tools...", 
                  font=("Arial", 10)).pack(pady=10)
            Label(progress_window, text="This may take a few minutes (file is ~96 MB)", 
                  font=("Arial", 9)).pack()
            
            progress_bar = ttk.Progressbar(progress_window, mode='indeterminate', length=350)
            progress_bar.pack(pady=20)
            progress_bar.start(10)
            
            status_label = Label(progress_window, text="Connecting...")
            status_label.pack()
            
            progress_window.update()
            
            # Try downloading
            downloaded = False
            for url in [download_url] + alt_urls:
                try:
                    status_label.config(text=f"Trying: {url.split('/')[-1]}")
                    progress_window.update()
                    
                    req = urllib.request.Request(
                        url,
                        headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROM Converter'}
                    )
                    
                    with urllib.request.urlopen(req, timeout=300) as response:
                        total_size = int(response.headers.get('content-length', 0))
                        
                        with open(download_path, 'wb') as f:
                            downloaded_size = 0
                            block_size = 1024 * 1024  # 1MB blocks
                            
                            while True:
                                buffer = response.read(block_size)
                                if not buffer:
                                    break
                                f.write(buffer)
                                downloaded_size += len(buffer)
                                
                                if total_size > 0:
                                    percent = (downloaded_size / total_size) * 100
                                    mb_downloaded = downloaded_size / (1024 * 1024)
                                    mb_total = total_size / (1024 * 1024)
                                    status_label.config(text=f"Downloaded: {mb_downloaded:.1f} / {mb_total:.1f} MB ({percent:.1f}%)")
                                else:
                                    mb_downloaded = downloaded_size / (1024 * 1024)
                                    status_label.config(text=f"Downloaded: {mb_downloaded:.1f} MB")
                                progress_window.update()
                    
                    downloaded = True
                    break
                    
                except urllib.error.HTTPError as e:
                    if e.code == 404:
                        continue  # Try next URL
                    raise
                except Exception:
                    continue
            
            if not downloaded:
                progress_window.destroy()
                messagebox.showerror("Error", "Could not download MAME tools from any mirror", parent=self.master)
                return False
            
            # Extract using 7-Zip or the self-extracting exe
            status_label.config(text="Extracting chdman.exe...")
            progress_window.update()
            
            # The MAME exe is a self-extracting 7z archive
            # We need 7-Zip to extract just chdman.exe, or run with specific args
            extracted = False
            
            # Build list of all possible 7zip paths to try
            seven_zip_paths = []
            
            # Re-check for 7-Zip in case it was installed after app startup
            if not self.seven_zip_path:
                self.check_7zip()
            
            # Add current path if set
            if self.seven_zip_path:
                seven_zip_paths.append(self.seven_zip_path)
            
            # Add all fallback paths
            fallback_paths = [
                self.script_dir / "7za.exe",
                r"C:\Program Files\7-Zip\7z.exe",
                r"C:\Program Files (x86)\7-Zip\7z.exe",
                r"C:\Program Files\PeaZip\res\bin\7z\7z.exe",
                r"C:\Program Files (x86)\PeaZip\res\bin\7z\7z.exe",
                r"C:\Program Files\PeaZip\res\bin\7z\x64\7z.exe",
                r"C:\Program Files (x86)\PeaZip\res\bin\7z\x64\7z.exe",
            ]
            for path in fallback_paths:
                path_str = str(path)
                if path_str not in seven_zip_paths and os.path.exists(path_str):
                    seven_zip_paths.append(path_str)
            
            # Also check PATH
            seven_zip_in_path = shutil.which("7z")
            if seven_zip_in_path and seven_zip_in_path not in seven_zip_paths:
                seven_zip_paths.append(seven_zip_in_path)
            
            # Auto-download 7-Zip if no paths found
            if not seven_zip_paths:
                status_label.config(text="Downloading 7-Zip...")
                progress_window.update()
                if self.download_7zip():
                    seven_zip_paths.append(self.seven_zip_path)
            
            # Try each 7zip path until one works
            for sz_path in seven_zip_paths:
                if extracted:
                    break
                try:
                    status_label.config(text=f"Extracting with {Path(sz_path).name}...")
                    progress_window.update()
                    cmd = [
                        sz_path, 'e', str(download_path),
                        '-o' + str(self.script_dir),
                        'chdman.exe',
                        '-y'
                    ]
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
                    if result.returncode == 0 and (self.script_dir / "chdman.exe").exists():
                        extracted = True
                        self.seven_zip_path = sz_path  # Remember working path
                except Exception as e:
                    print(f"7-Zip extraction error with {sz_path}: {e}")
            
            if not extracted:
                # Try running as self-extracting archive with output directory
                try:
                    cmd = [str(download_path), '-o' + str(temp_dir), '-y']
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
                    # Move chdman.exe to script directory
                    temp_chdman = temp_dir / "chdman.exe"
                    if temp_chdman.exists():
                        shutil.move(str(temp_chdman), str(self.script_dir / "chdman.exe"))
                        extracted = True
                except Exception as e:
                    print(f"Self-extraction error: {e}")
            
            # If still not extracted, let user choose an extractor
            if not extracted:
                extracted = self.prompt_user_select_extractor(download_path, self.script_dir)
            
            progress_window.destroy()
            
            # Clean up the temp directory and download file
            if temp_dir.exists():
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass
            
            if extracted and (self.script_dir / "chdman.exe").exists():
                self.chdman_path = str(self.script_dir / "chdman.exe")
                self.save_config()
                messagebox.showinfo("Success", f"chdman.exe downloaded successfully!\n\nLocation:\n{self.chdman_path}", parent=self.master)
                return True
            else:
                messagebox.showerror(
                    "Extraction Failed",
                    "Could not extract chdman.exe from MAME package.\n\n"
                    "Please install 7-Zip and try again, or manually download MAME from:\n"
                    "https://www.mamedev.org/release.html",
                    parent=self.master
                )
                return False
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to download MAME tools:\n{e}", parent=self.master)
            return False
    
    def download_mame_tools_linux(self):
        """Download chdman for Linux systems (SteamOS, etc.)"""
        try:
            # Try to download prebuilt chdman from a reliable source
            # Using mame-tools from various Linux package mirrors or building from source
            
            temp_dir = self.script_dir / "mame_temp"
            temp_dir.mkdir(exist_ok=True)
            
            script_dir = self.script_dir
            chdman_dest = script_dir / "chdman"
            
            # Show progress dialog - ensure it appears in front
            progress_window = Toplevel(self.master)
            progress_window.title("Downloading chdman")
            progress_window.geometry("450x180")
            progress_window.resizable(False, False)
            progress_window.transient(self.master)
            progress_window.grab_set()
            progress_window.lift()
            progress_window.focus_force()
            
            Label(progress_window, text="Downloading chdman for Linux...", 
                  font=("Arial", 10)).pack(pady=10)
            
            progress_bar = ttk.Progressbar(progress_window, mode='indeterminate', length=350)
            progress_bar.pack(pady=10)
            progress_bar.start(10)
            
            status_label = Label(progress_window, text="Connecting...", wraplength=400)
            status_label.pack(pady=5)
            
            progress_window.update()
            
            # Try multiple sources for prebuilt chdman
            download_sources = [
                # EmuDeck's prebuilt chdman for Steam Deck
                ("https://raw.githubusercontent.com/dragoonDorise/EmuDeck/main/tools/chdconv/chdman", "direct"),
                # Arch Linux package extraction (mame-tools)
                ("https://archive.archlinux.org/packages/m/mame-tools/", "arch"),
            ]
            
            downloaded = False
            
            for source_url, source_type in download_sources:
                try:
                    status_label.config(text=f"Trying: {source_type}...")
                    progress_window.update()
                    
                    if source_type == "direct":
                        # Direct binary download
                        req = urllib.request.Request(
                            source_url,
                            headers={'User-Agent': 'Mozilla/5.0 ROM Converter'}
                        )
                        
                        with urllib.request.urlopen(req, timeout=60) as response:
                            with open(chdman_dest, 'wb') as f:
                                f.write(response.read())
                        
                        # Make executable
                        os.chmod(chdman_dest, 0o755)
                        
                        # Verify it works
                        result = subprocess.run([str(chdman_dest), '--help'], 
                                              capture_output=True, timeout=5)
                        if result.returncode == 0 or b'chdman' in result.stdout or b'chdman' in result.stderr:
                            downloaded = True
                            break
                        else:
                            chdman_dest.unlink(missing_ok=True)
                            
                except Exception as e:
                    print(f"Download from {source_type} failed: {e}")
                    continue
            
            # If direct download failed, try installing MAME via Flatpak
            if not downloaded:
                try:
                    status_label.config(text="Trying Flatpak install of MAME...")
                    progress_window.update()
                    
                    # Check if flatpak is available
                    if shutil.which('flatpak'):
                        # Try --user install first (works on immutable distros like Bazzite),
                        # then fall back to system-wide install
                        flatpak_cmds = [
                            ['flatpak', 'install', '-y', '--noninteractive', '--user', 'flathub', 'org.mamedev.MAME'],
                            ['flatpak', 'install', '-y', '--noninteractive', 'flathub', 'org.mamedev.MAME'],
                        ]
                        for flatpak_cmd in flatpak_cmds:
                            try:
                                install_type = '--user' if '--user' in flatpak_cmd else 'system'
                                status_label.config(text=f"Trying Flatpak install ({install_type})...")
                                progress_window.update()
                                result = subprocess.run(
                                    flatpak_cmd,
                                    capture_output=True, text=True, timeout=300
                                )
                                if result.returncode == 0:
                                    # Find chdman inside the Flatpak installation
                                    flatpak_chdman = self._find_flatpak_chdman()
                                    if flatpak_chdman:
                                        self.chdman_path = flatpak_chdman
                                        self.save_config()
                                        progress_window.destroy()
                                        if temp_dir.exists():
                                            try:
                                                shutil.rmtree(temp_dir)
                                            except:
                                                pass
                                        messagebox.showinfo("Success",
                                            f"MAME installed via Flatpak!\n\nchdman found at:\n{flatpak_chdman}",
                                            parent=self.master)
                                        return True
                            except (subprocess.TimeoutExpired, Exception) as e:
                                print(f"Flatpak install ({install_type}) failed: {e}")
                                continue
                except (FileNotFoundError, Exception) as e:
                    print(f"Flatpak install failed: {e}")
            
            progress_window.destroy()
            
            # Clean up temp directory
            if temp_dir.exists():
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass
            
            if downloaded and chdman_dest.exists():
                self.chdman_path = str(chdman_dest)
                self.save_config()
                messagebox.showinfo("Success", f"chdman downloaded successfully!\n\nLocation:\n{self.chdman_path}", parent=self.master)
                return True
            else:
                # Provide OS-aware instructions
                install_msg = self._get_chdman_install_instructions()
                messagebox.showwarning(
                    "Manual Installation Required",
                    install_msg,
                    parent=self.master
                )
                return False
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to download MAME tools:\n{e}", parent=self.master)
            return False
    
    def check_7zip(self):
        """Check if 7-Zip is available (including PeaZip's bundled 7z)"""
        # Check common locations on Windows
        common_paths = [
            r"C:\Program Files\7-Zip\7z.exe",
            r"C:\Program Files (x86)\7-Zip\7z.exe",
            # PeaZip bundles 7z.exe
            r"C:\Program Files\PeaZip\res\bin\7z\7z.exe",
            r"C:\Program Files (x86)\PeaZip\res\bin\7z\7z.exe",
        ]
        
        # Check for our downloaded 7za.exe first
        local_7za = self.script_dir / "7za.exe"
        if local_7za.exists():
            self.seven_zip_path = str(local_7za)
            return True
        
        # Check PATH for 7z
        seven_zip = shutil.which("7z")
        if seven_zip:
            self.seven_zip_path = seven_zip
            return True
        
        # Check PATH for peazip (in case it's there)
        peazip = shutil.which("peazip")
        if peazip:
            # PeaZip's 7z is relative to the peazip executable
            peazip_dir = Path(peazip).parent
            peazip_7z = peazip_dir / "res" / "bin" / "7z" / "7z.exe"
            if peazip_7z.exists():
                self.seven_zip_path = str(peazip_7z)
                return True
        
        # Check common install locations
        for path in common_paths:
            if os.path.exists(path):
                self.seven_zip_path = path
                return True
        
        return False

    def download_7zip(self):
        """Download and install 7-Zip portable for extraction"""
        if sys.platform != "win32":
            return False
        
        try:
            script_dir = self.script_dir
            temp_dir = script_dir / "temp_7zip"
            temp_dir.mkdir(exist_ok=True)
            
            # First download 7zr.exe (minimal 7z extractor)
            url_7zr = "https://www.7-zip.org/a/7zr.exe"
            seven_zr_path = temp_dir / "7zr.exe"
            
            response = requests.get(url_7zr, timeout=30)
            response.raise_for_status()
            with open(seven_zr_path, 'wb') as f:
                f.write(response.content)
            
            # Download 7-Zip Extra package (contains 7za.exe with full format support)
            url_extra = "https://www.7-zip.org/a/7z2409-extra.7z"
            extra_path = temp_dir / "7z-extra.7z"
            
            response = requests.get(url_extra, timeout=60)
            response.raise_for_status()
            with open(extra_path, 'wb') as f:
                f.write(response.content)
            
            # Use 7zr.exe to extract 7za.exe from the extra package
            cmd = [str(seven_zr_path), 'e', str(extra_path), '-o' + str(script_dir), '7za.exe', '-y']
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            
            seven_zip_path = script_dir / "7za.exe"
            if seven_zip_path.exists():
                self.seven_zip_path = str(seven_zip_path)
                self.save_config()
                # Clean up temp files
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass
                return True
            
            # Fallback: just use 7zr.exe if extraction failed
            fallback_path = script_dir / "7zr.exe"
            shutil.copy(seven_zr_path, fallback_path)
            if fallback_path.exists():
                self.seven_zip_path = str(fallback_path)
                self.save_config()
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass
                return True
                
        except Exception as e:
            print(f"Failed to download 7-Zip: {e}")
        
        return False

    def get_all_extractor_paths(self):
        """Get all detected extractor paths on the system"""
        extractors = []
        
        # All possible paths to check
        possible_paths = [
            (self.script_dir / "7za.exe", "7za.exe (Local)"),
            (r"C:\Program Files\7-Zip\7z.exe", "7-Zip (Program Files)"),
            (r"C:\Program Files (x86)\7-Zip\7z.exe", "7-Zip (Program Files x86)"),
            (r"C:\Program Files\PeaZip\res\bin\7z\7z.exe", "PeaZip 7z"),
            (r"C:\Program Files (x86)\PeaZip\res\bin\7z\7z.exe", "PeaZip 7z (x86)"),
            (r"C:\Program Files\PeaZip\res\bin\7z\x64\7z.exe", "PeaZip 7z x64"),
            (r"C:\Program Files (x86)\PeaZip\res\bin\7z\x64\7z.exe", "PeaZip 7z x64 (x86)"),
        ]
        
        for path, name in possible_paths:
            path_str = str(path)
            if os.path.exists(path_str):
                extractors.append((path_str, name))
        
        # Check PATH
        seven_zip_in_path = shutil.which("7z")
        if seven_zip_in_path:
            extractors.append((seven_zip_in_path, f"7z (PATH: {seven_zip_in_path})"))
        
        return extractors

    def prompt_user_select_extractor(self, archive_path, output_dir):
        """Show dialog for user to select an extractor when automatic extraction fails"""
        extractors = self.get_all_extractor_paths()
        
        if not extractors:
            # No extractors found, offer to browse
            result = messagebox.askyesno(
                "No Extractor Found",
                "No 7-Zip or compatible extractor was detected.\n\n"
                "Would you like to browse for a 7z.exe manually?",
                parent=self.master
            )
            if result:
                filetypes = [("Executable files", "*.exe"), ("All files", "*.*")]
                selected = filedialog.askopenfilename(
                    title="Select 7z.exe or compatible extractor",
                    filetypes=filetypes,
                    parent=self.master
                )
                if selected:
                    extractors = [(selected, "User selected")]
        
        if not extractors:
            return False
        
        # Create selection dialog
        dialog = Toplevel(self.master)
        dialog.title("Select Extractor")
        dialog.transient(self.master)
        dialog.grab_set()
        
        # Center dialog
        dialog.geometry("450x300")
        dialog.resizable(False, False)
        
        Label(dialog, text="Automatic extraction failed.\nPlease select an extractor to try:",
              font=("Consolas", 10)).pack(pady=10)
        
        # Listbox with extractors
        listbox_frame = Frame(dialog)
        listbox_frame.pack(fill="both", expand=True, padx=20, pady=5)
        
        scrollbar = Scrollbar(listbox_frame)
        scrollbar.pack(side="right", fill="y")
        
        listbox = Listbox(listbox_frame, font=("Consolas", 9), yscrollcommand=scrollbar.set)
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)
        
        for path, name in extractors:
            listbox.insert("end", f"{name}")
        
        if extractors:
            listbox.selection_set(0)
        
        result = {"extracted": False, "selected_path": None}
        
        def try_extract():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("No Selection", "Please select an extractor.", parent=dialog)
                return
            
            selected_path = extractors[selection[0]][0]
            result["selected_path"] = selected_path
            
            try:
                cmd = [selected_path, 'e', str(archive_path), '-o' + str(output_dir), 'chdman.exe', '-y']
                proc_result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
                if proc_result.returncode == 0 and (output_dir / "chdman.exe").exists():
                    result["extracted"] = True
                    self.seven_zip_path = selected_path
                    dialog.destroy()
                else:
                    messagebox.showerror(
                        "Extraction Failed",
                        f"Extraction failed with selected tool.\n\nError: {proc_result.stderr[:200] if proc_result.stderr else 'Unknown error'}",
                        parent=dialog
                    )
            except Exception as e:
                messagebox.showerror("Error", f"Extraction error: {e}", parent=dialog)
        
        def browse_custom():
            filetypes = [("Executable files", "*.exe"), ("All files", "*.*")]
            selected = filedialog.askopenfilename(
                title="Select 7z.exe or compatible extractor",
                filetypes=filetypes,
                parent=dialog
            )
            if selected:
                extractors.append((selected, f"Custom: {Path(selected).name}"))
                listbox.insert("end", f"Custom: {Path(selected).name}")
                listbox.selection_clear(0, "end")
                listbox.selection_set("end")
        
        def cancel():
            dialog.destroy()
        
        button_frame = Frame(dialog)
        button_frame.pack(pady=10)
        
        Button(button_frame, text="Try Selected", command=try_extract, width=12).pack(side="left", padx=5)
        Button(button_frame, text="Browse...", command=browse_custom, width=12).pack(side="left", padx=5)
        Button(button_frame, text="Cancel", command=cancel, width=12).pack(side="left", padx=5)
        
        dialog.wait_window()
        return result["extracted"]

    def check_maxcso(self):
        """Check if maxcso is available for CSO/ZSO output"""
        # Determine platform-specific binary name
        maxcso_name = "maxcso.exe" if sys.platform == "win32" else "maxcso"
        
        # First check bundled resources (PyInstaller)
        bundled_maxcso = self.bundle_dir / maxcso_name
        if bundled_maxcso.exists():
            self.maxcso_path = str(bundled_maxcso)
            return True
        
        # Then check for maxcso directly next to the executable/script
        direct_maxcso = self.script_dir / maxcso_name
        if direct_maxcso.exists():
            self.maxcso_path = str(direct_maxcso)
            return True

        # Check PATH as fallback
        maxcso = shutil.which("maxcso")
        if maxcso:
            self.maxcso_path = maxcso
            return True

        return False

    def check_extract_xiso(self):
        """Find extract-xiso executable. Platform-aware: checks .exe on Windows, binary on Linux."""
        xiso_name = "extract-xiso.exe" if sys.platform == "win32" else "extract-xiso"

        # Check bundled resources (PyInstaller)
        bundled_xiso = self.bundle_dir / xiso_name
        if bundled_xiso.exists():
            self.extract_xiso_path = str(bundled_xiso)
            return True

        # Check script directory
        app_xiso = self.script_dir / xiso_name
        if app_xiso.exists():
            self.extract_xiso_path = str(app_xiso)
            return True

        # Check PATH
        sys_xiso = shutil.which("extract-xiso")
        if sys_xiso:
            self.extract_xiso_path = sys_xiso
            return True

        # Linux: check common locations
        if sys.platform != "win32":
            for candidate in ["/usr/bin/extract-xiso", "/usr/local/bin/extract-xiso"]:
                if os.path.isfile(candidate) and os.access(candidate, os.X_OK):
                    self.extract_xiso_path = candidate
                    return True

        return False

    def extract_xbox_iso(self, iso_path, dest_folder):
        """Extract Xbox/Xbox 360 ISO using extract-xiso. Supports XGD2/XGD3 (Xbox 360) images."""
        if not self.extract_xiso_path:
            self.log("⚠️  extract-xiso not found. Use Settings → Tool Paths to download or set path.")
            return False

        iso_path = Path(iso_path)
        game_name = iso_path.stem
        target_dir = Path(dest_folder) / game_name
        target_dir.mkdir(parents=True, exist_ok=True)

        try:
            # Detect ISO size for logging
            iso_size_gb = iso_path.stat().st_size / (1024 ** 3) if iso_path.exists() else 0
            iso_type = "Xbox"
            if iso_size_gb > 7.0:
                iso_type = "Xbox 360 (XGD3)" if iso_size_gb > 8.0 else "Xbox 360 (XGD2)"

            self.log(f"🎮 Extracting {iso_type} ISO: {iso_path.name} ({iso_size_gb:.1f} GB)")

            # extract-xiso -x -d [directory] [iso]
            # -x = extract, -d = destination directory
            cmd = [self.extract_xiso_path, "-x", "-d", str(target_dir), str(iso_path)]

            creationflags = 0x00000040 if sys.platform == "win32" else 0  # IDLE_PRIORITY_CLASS

            result = subprocess.run(cmd, capture_output=True, text=True,
                                   creationflags=creationflags, timeout=600)
            # Check output for success indicators
            output = (result.stdout + result.stderr).lower()
            if result.returncode == 0 or "extraction complete" in output or "done" in output:
                # Verify files were actually extracted
                extracted_files = list(target_dir.glob("*"))
                if extracted_files:
                    self.log(f"✅ Extracted {iso_type} ISO: {iso_path.name} → {len(extracted_files)} files in {target_dir.name}/")
                    return True
                else:
                    self.log(f"⚠️  Extraction appeared to succeed but no files found for {iso_path.name}")
                    return False
            else:
                err_msg = result.stderr.strip() or result.stdout.strip()
                self.log(f"❌ Failed to extract {iso_path.name}: {err_msg[:200]}")
                return False
        except subprocess.TimeoutExpired:
            self.log(f"❌ Timed out extracting {iso_path.name} (over 10 minutes)")
            return False
        except Exception as e:
            self.log(f"❌ Error extracting Xbox ISO {iso_path.name}: {e}")
            return False

    def browse_7zip(self):
        """Allow user to manually select 7z executable"""
        filetypes = [("Executable files", "*.exe"), ("All files", "*.*")]
        seven_zip_file = filedialog.askopenfilename(
            title="Select 7z executable",
            filetypes=filetypes
        )
        if seven_zip_file:
            try:
                result = subprocess.run(
                    [seven_zip_file],
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                if "7-zip" in result.stdout.lower() or "7-zip" in result.stderr.lower():
                    self.seven_zip_path = seven_zip_file
                    self.save_config()
                    self.log(f"7-Zip location set to: {seven_zip_file}")
                    if hasattr(self, 'seven_zip_label'):
                        self.seven_zip_label.config(text=self.seven_zip_path or "Not set")
                    self._update_tool_status()
                    messagebox.showinfo("Success", f"7-Zip location set to:\n{seven_zip_file}")
                else:
                    messagebox.showerror("Error", "Selected file does not appear to be 7-Zip")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to verify 7-Zip:\n{e}")

    def browse_maxcso(self):
        """Allow user to manually select maxcso executable"""
        filetypes = [("Executable files", "*.exe"), ("All files", "*.*")]
        maxcso_file = filedialog.askopenfilename(
            title="Select maxcso executable",
            filetypes=filetypes
        )
        if maxcso_file:
            try:
                result = subprocess.run(
                    [maxcso_file, "--help"],
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                output = (result.stdout + result.stderr).lower()
                if "maxcso" in output:
                    self.maxcso_path = maxcso_file
                    self.save_config()
                    self.log(f"maxcso location set to: {maxcso_file}")
                    if hasattr(self, 'maxcso_label'):
                        self.maxcso_label.config(text=self.maxcso_path)
                    self._update_tool_status()
                    messagebox.showinfo("Success", f"maxcso location set to:\n{maxcso_file}")
                else:
                    messagebox.showerror("Error", "Selected file does not appear to be maxcso")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to verify maxcso:\n{e}")
    
    def download_maxcso(self):
        """Download the latest maxcso binary from GitHub releases with progress dialog."""
        progress_window = Toplevel(self.master)
        progress_window.title("Downloading maxcso")
        progress_window.geometry("450x180")
        progress_window.resizable(False, False)
        progress_window.transient(self.master)
        progress_window.grab_set()
        progress_window.configure(bg=COLORS["bg_dark"])
        progress_window.lift()
        progress_window.focus_force()

        Label(progress_window, text="Downloading maxcso...", font=self.font_title,
              fg=COLORS["text_primary"], bg=COLORS["bg_dark"]).pack(pady=10)
        progress_bar = ttk.Progressbar(progress_window, mode='indeterminate', length=380)
        progress_bar.pack(pady=8)
        progress_bar.start(10)
        status_label = Label(progress_window, text="Fetching latest release...",
                            font=self.font_small, fg=COLORS["text_secondary"],
                            bg=COLORS["bg_dark"], wraplength=420)
        status_label.pack(pady=5)
        progress_window.update()

        def do_download():
            try:
                # Fetch latest release info from GitHub API
                api_url = "https://api.github.com/repos/unknownbrackets/maxcso/releases/latest"
                req = urllib.request.Request(api_url, headers={
                    'User-Agent': 'ROM-Converter',
                    'Accept': 'application/vnd.github.v3+json'
                })
                with urllib.request.urlopen(req, timeout=30) as resp:
                    release_data = json.loads(resp.read().decode())

                # Find the correct asset for this platform
                asset = None
                if sys.platform == "win32":
                    for a in release_data.get("assets", []):
                        name = a["name"].lower()
                        if name.endswith(".exe") and ("windows" in name or "win" in name or "x86" in name):
                            asset = a
                            break
                    if not asset:
                        for a in release_data.get("assets", []):
                            if a["name"].lower().endswith(".exe"):
                                asset = a
                                break
                else:
                    for a in release_data.get("assets", []):
                        name = a["name"].lower()
                        if "linux" in name and not name.endswith(".exe"):
                            asset = a
                            break
                    if not asset:
                        for a in release_data.get("assets", []):
                            if not a["name"].lower().endswith(".exe") and "source" not in a["name"].lower():
                                asset = a
                                break

                if not asset:
                    status_label.config(text="No prebuilt binary found. Downloading source fallback...")
                    progress_window.update()
                    # Fallback: download from project repo (bundled binary)
                    fallback_url = "https://raw.githubusercontent.com/WoofahRayetCode/WlfRyt-Rom-Manager/main/maxcso.exe"
                    self._download_file(fallback_url, self.script_dir / "maxcso.exe", status_label, progress_window)
                else:
                    download_url = asset["browser_download_url"]
                    fname = asset["name"] if sys.platform == "win32" else "maxcso"
                    dest = self.script_dir / fname
                    status_label.config(text=f"Downloading {asset['name']} ({asset['size']/1024/1024:.1f} MB)...")
                    progress_window.update()
                    self._download_file(download_url, dest, status_label, progress_window)

                    # Make executable on Linux
                    if sys.platform != "win32":
                        os.chmod(str(dest), 0o755)

                # Re-check to pick up the newly downloaded binary
                self.check_maxcso()
                if self.maxcso_path:
                    self.save_config()
                    if hasattr(self, 'maxcso_label'):
                        self.maxcso_label.config(text=self.maxcso_path, fg=COLORS["text_secondary"])
                    self._update_tool_status()
                    self.log(f"✅ maxcso downloaded successfully: {self.maxcso_path}")
                    progress_window.after(500, progress_window.destroy)
                    self.master.after(600, lambda: messagebox.showinfo(
                        "Success", f"maxcso downloaded to:\n{self.maxcso_path}", parent=self.master))
                else:
                    self.log("⚠️  maxcso downloaded but not detected. Check Settings → Tool Paths.")
                    progress_window.after(500, progress_window.destroy)
                    self.master.after(600, lambda: messagebox.showwarning(
                        "Manual Setup Needed",
                        "maxcso was downloaded but could not be auto-detected.\n\n"
                        "Use Settings → Tool Paths → maxcso → [Set] to locate it manually.",
                        parent=self.master))
            except urllib.error.HTTPError as e:
                self.log(f"❌ GitHub API error downloading maxcso: {e}")
                progress_window.after(500, progress_window.destroy)
                self.master.after(600, lambda: messagebox.showerror(
                    "Download Failed",
                    f"Could not fetch maxcso release info (HTTP {e.code}).\n\n"
                    "Visit https://github.com/unknownbrackets/maxcso/releases\n"
                    "to download manually.",
                    parent=self.master))
            except Exception as e:
                self.log(f"❌ Failed to download maxcso: {e}")
                progress_window.after(500, progress_window.destroy)
                self.master.after(600, lambda: messagebox.showerror(
                    "Download Failed", f"Error: {e}", parent=self.master))

        threading.Thread(target=do_download, daemon=True).start()

    def _download_file(self, url, dest_path, status_label, progress_window):
        """Download a file with progress updates to the status label."""
        req = urllib.request.Request(url, headers={'User-Agent': 'ROM-Converter/1.0'})
        with urllib.request.urlopen(req, timeout=300) as response:
            total = int(response.headers.get('content-length', 0))
            downloaded = 0
            with open(str(dest_path), 'wb') as f:
                while True:
                    chunk = response.read(1024 * 1024)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total > 0:
                        pct = downloaded / total * 100
                        status_label.config(
                            text=f"Downloading... {downloaded/1024/1024:.1f} / {total/1024/1024:.1f} MB ({pct:.0f}%)")
                        progress_window.update()

    def download_extract_xiso(self):
        """Download or install extract-xiso with platform-aware logic."""
        # Linux: check package managers first
        if sys.platform != "win32":
            # Try system package manager
            for pm_cmd, pkg_name, install_cmd in [
                (shutil.which("dnf"), "extract-xiso", "sudo dnf install -y extract-xiso"),
                (shutil.which("apt"), "extract-xiso", "sudo apt install -y extract-xiso"),
                (shutil.which("pacman"), "extract-xiso", "sudo pacman -S --noconfirm extract-xiso"),
                (shutil.which("zypper"), "extract-xiso", "sudo zypper install -y extract-xiso"),
            ]:
                if pm_cmd:
                    response = messagebox.askyesno(
                        "Install extract-xiso",
                        f"extract-xiso can be installed via your package manager:\n\n"
                        f"  {install_cmd}\n\n"
                        f"Would you like ROM Converter to run this command?",
                        parent=self.master)
                    if response:
                        progress_window = Toplevel(self.master)
                        progress_window.title("Installing extract-xiso")
                        progress_window.geometry("450x120")
                        progress_window.transient(self.master)
                        progress_window.grab_set()
                        progress_window.configure(bg=COLORS["bg_dark"])
                        Label(progress_window, text=f"Running: {install_cmd}",
                              font=self.font_small, fg=COLORS["text_secondary"],
                              bg=COLORS["bg_dark"]).pack(pady=20)
                        progress_window.update()
                        try:
                            result = subprocess.run(
                                install_cmd.split(), capture_output=True, text=True, timeout=120)
                            progress_window.destroy()
                            if result.returncode == 0 or self.check_extract_xiso():
                                self._update_tool_status()
                                self.log("✅ extract-xiso installed via package manager")
                                messagebox.showinfo("Success",
                                    "extract-xiso installed successfully!", parent=self.master)
                                return
                            else:
                                self.log(f"⚠️ Package install failed: {result.stderr[:200]}")
                        except Exception as e:
                            progress_window.destroy()
                            self.log(f"⚠️ Package install error: {e}")
                    break
            else:
                messagebox.showinfo("No Package Manager Found",
                    "Could not detect a supported package manager.\n\n"
                    "Please install extract-xiso manually or place the binary\n"
                    "in the application directory.",
                    parent=self.master)
                return
        else:
            # Windows: download from a known source
            progress_window = Toplevel(self.master)
            progress_window.title("Downloading extract-xiso")
            progress_window.geometry("450x150")
            progress_window.resizable(False, False)
            progress_window.transient(self.master)
            progress_window.grab_set()
            progress_window.configure(bg=COLORS["bg_dark"])
            progress_window.lift()
            progress_window.focus_force()

            Label(progress_window, text="Downloading extract-xiso...", font=self.font_title,
                  fg=COLORS["text_primary"], bg=COLORS["bg_dark"]).pack(pady=10)
            progress_bar = ttk.Progressbar(progress_window, mode='indeterminate', length=380)
            progress_bar.pack(pady=6)
            progress_bar.start(10)
            status_label = Label(progress_window, text="Connecting...",
                                font=self.font_small, fg=COLORS["text_secondary"],
                                bg=COLORS["bg_dark"])
            status_label.pack(pady=5)
            progress_window.update()

            def do_win_download():
                try:
                    # Try GitHub releases for extract-xiso
                    api_url = "https://api.github.com/repos/XboxDev/extract-xiso/releases/latest"
                    req = urllib.request.Request(api_url, headers={
                        'User-Agent': 'ROM-Converter',
                        'Accept': 'application/vnd.github.v3+json'
                    })
                    with urllib.request.urlopen(req, timeout=30) as resp:
                        release_data = json.loads(resp.read().decode())
                    asset = None
                    for a in release_data.get("assets", []):
                        if a["name"].lower().endswith(".exe"):
                            asset = a
                            break
                    if asset:
                        dest = self.script_dir / "extract-xiso.exe"
                        self._download_file(asset["browser_download_url"], dest, status_label, progress_window)
                    else:
                        # Fallback to bundled copy
                        fallback_url = "https://raw.githubusercontent.com/WoofahRayetCode/WlfRyt-Rom-Manager/main/extract-xiso.exe"
                        self._download_file(fallback_url, self.script_dir / "extract-xiso.exe", status_label, progress_window)

                    self.check_extract_xiso()
                    if self.extract_xiso_path:
                        self.save_config()
                        self._update_tool_status()
                        self.log("✅ extract-xiso downloaded successfully")
                        progress_window.after(500, progress_window.destroy)
                        self.master.after(600, lambda: messagebox.showinfo(
                            "Success", f"extract-xiso downloaded to:\n{self.extract_xiso_path}", parent=self.master))
                    else:
                        progress_window.after(500, progress_window.destroy)
                        self.master.after(600, lambda: messagebox.showwarning(
                            "Manual Setup Needed",
                            "extract-xiso was downloaded but not detected.\nUse Settings → Tool Paths to set it manually.",
                            parent=self.master))
                except Exception as e:
                    self.log(f"❌ Failed to download extract-xiso: {e}")
                    progress_window.after(500, progress_window.destroy)
                    self.master.after(600, lambda: messagebox.showerror(
                        "Download Failed",
                        f"Error: {e}\n\nDownload manually from:\n"
                        "https://github.com/XboxDev/extract-xiso/releases",
                        parent=self.master))

            threading.Thread(target=do_win_download, daemon=True).start()

    def browse_extract_xiso(self):
        """Allow user to manually select extract-xiso executable."""
        filetypes = [("Executable files", "*.exe"), ("All files", "*.*")] if sys.platform == "win32" else [("All files", "*")]
        selected = filedialog.askopenfilename(
            title="Select extract-xiso executable", filetypes=filetypes)
        if selected:
            try:
                result = subprocess.run([selected], capture_output=True, text=True, timeout=5)
                output = (result.stdout + result.stderr).lower()
                if "extract-xiso" in output or "xiso" in output or result.returncode in [0, 1]:
                    self.extract_xiso_path = selected
                    self.save_config()
                    self.log(f"extract-xiso location set to: {selected}")
                    self._update_tool_status()
                    messagebox.showinfo("Success", f"extract-xiso location set to:\n{selected}")
                else:
                    messagebox.showerror("Error", "Selected file does not appear to be extract-xiso")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to verify extract-xiso:\n{e}")

    def show_maxcso_setup_help(self):
        """Show a platform-aware setup guide for maxcso."""
        help_window = Toplevel(self.master)
        help_window.title("maxcso Setup Guide")
        help_window.geometry("620x540")
        help_window.transient(self.master)
        help_window.grab_set()
        help_window.configure(bg=COLORS['bg_dark'])
        help_window.lift()
        help_window.focus_force()

        main_frame = Frame(help_window, bg=COLORS['bg_dark'])
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)

        Label(main_frame, text="🗜 maxcso Setup Guide", font=self.font_title,
              fg=COLORS['accent_yellow'], bg=COLORS['bg_dark']).pack(anchor="w", pady=(0, 10))

        info_text = Text(main_frame, font=self.font_body, fg=COLORS['text_primary'],
                        bg=COLORS['bg_light'], wrap="word", height=22, relief="flat", padx=10, pady=10)
        info_text.pack(fill="both", expand=True, pady=(0, 10))

        if sys.platform == "win32":
            help_content = """maxcso is required for CSO and ZSO output formats.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OPTION 1: Auto-Download (Recommended)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Click [ Download maxcso ] in Settings → Tool Paths to automatically
fetch the latest release from GitHub.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OPTION 2: Place in Program Directory
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. Click [ Open GitHub Releases ] below
2. Download the latest maxcso.exe
3. Place it in the same folder as rom_converter.py
4. Restart ROM Converter

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OPTION 3: Manual Selection with [ Set ] Button
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Use Settings → Tool Paths → maxcso → [Set] to browse and select.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Format Recommendations
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

• CHD: High compatibility, slightly larger files
• CSO: Good compression, fast decompression
• ZSO: Best compression, lower CPU usage"""
        else:
            help_content = """maxcso is required for CSO and ZSO output formats.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OPTION 1: Auto-Download (Recommended)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Click [ Download ] in Settings → Tool Paths to automatically
fetch the latest release from GitHub.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OPTION 2: Place in Program Directory
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. Click [ Open GitHub Releases ] below
2. Download the Linux binary (or compile from source)
3. Rename to 'maxcso' and place in the app directory
4. Make executable: chmod +x maxcso
5. Restart ROM Converter

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OPTION 3: Build from Source
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

git clone https://github.com/unknownbrackets/maxcso
cd maxcso && make
cp maxcso /path/to/rom-converter/

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Format Recommendations
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

• CHD: High compatibility, slightly larger files
• CSO: Good compression, fast decompression
• ZSO: Best compression, lower CPU usage"""

        info_text.config(state="normal")
        info_text.insert("1.0", help_content)
        info_text.config(state="disabled")

        button_frame = Frame(main_frame, bg=COLORS['bg_dark'])
        button_frame.pack(fill="x")

        Button(button_frame, text="[ Open GitHub Releases ]",
               command=self.open_maxcso_releases,
               font=self.font_small, bg=COLORS['button_blue'], fg="white",
               activeforeground="white", relief="flat", cursor="hand2").pack(side="left", padx=(0, 5))

        Button(button_frame, text="[ Download Now ]",
               command=lambda: [help_window.destroy(), self.download_maxcso()],
               font=self.font_small, bg=COLORS['button_green'], fg="white",
               activeforeground="white", relief="flat", cursor="hand2").pack(side="left", padx=(0, 5))

        Button(button_frame, text="[ Close ]", command=help_window.destroy,
               font=self.font_small, bg=COLORS['bg_light'],
               activeforeground="white", relief="flat", cursor="hand2").pack(side="left")
    
    def open_maxcso_releases(self):
        """Open maxcso GitHub releases page in browser"""
        import webbrowser
        webbrowser.open("https://github.com/unknownbrackets/maxcso/releases")
        self.log("Opening maxcso releases page...")
    
    def browse_chdman(self):
        """Allow user to manually select chdman executable"""
        filetypes = [("Executable files", "*.exe"), ("All files", "*.*")]
        chdman_file = filedialog.askopenfilename(
            title="Select chdman executable",
            filetypes=filetypes
        )
        if chdman_file:
            # Verify it's actually chdman by trying to run it
            try:
                result = subprocess.run(
                    [chdman_file, "--help"],
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                if "chdman" in result.stdout.lower() or "chdman" in result.stderr.lower():
                    self.chdman_path = chdman_file
                    self.save_config()
                    self.log(f"chdman location set to: {chdman_file}")
                    if hasattr(self, 'chdman_label'):
                        self.chdman_label.config(text=self.chdman_path)
                    self._update_tool_status()
                    messagebox.showinfo("Success", f"chdman location set to:\n{chdman_file}")
                else:
                    messagebox.showerror("Error", "Selected file does not appear to be chdman")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to verify chdman:\n{e}")
    
    def browse_ndecrypt(self):
        """Allow user to manually select NDecrypt executable"""
        filetypes = [("Executable files", "*.exe"), ("All files", "*.*")]
        ndecrypt_file = filedialog.askopenfilename(
            title="Select NDecrypt executable",
            filetypes=filetypes
        )
        if ndecrypt_file:
            try:
                result = subprocess.run(
                    [ndecrypt_file, "--help"],
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                output = (result.stdout + result.stderr).lower()
                if "ndecrypt" in output or "decrypt" in output:
                    self.ndecrypt_path = ndecrypt_file
                    self.save_config()
                    self.log(f"NDecrypt location set to: {ndecrypt_file}")
                    if hasattr(self, 'ndecrypt_label'):
                        self.ndecrypt_label.config(text=self.ndecrypt_path,
                                                  fg=COLORS['text_secondary'])
                    self._update_tool_status()
                    messagebox.showinfo("Success", f"NDecrypt location set to:\n{ndecrypt_file}")
                else:
                    messagebox.showerror("Error", "Selected file does not appear to be NDecrypt")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to verify NDecrypt:\n{e}")
    
    def download_ndecrypt(self):
        """Download NDecrypt from GitHub releases"""
        try:
            # Determine platform
            if sys.platform == "win32":
                asset_pattern = "win-x64"
            elif sys.platform == "darwin":
                import platform
                if platform.machine() == "arm64":
                    asset_pattern = "osx-arm64"
                else:
                    asset_pattern = "osx-x64"
            else:  # Linux
                import platform
                if platform.machine() == "aarch64":
                    asset_pattern = "linux-arm64"
                else:
                    asset_pattern = "linux-x64"
            
            self.log("Fetching NDecrypt release info from GitHub...")
            
            # Get latest release info
            req = urllib.request.Request(
                NDECRYPT_GITHUB_RELEASES_API,
                headers={'User-Agent': 'ROM-Converter'}
            )
            with urllib.request.urlopen(req, timeout=30) as response:
                release_data = json.loads(response.read().decode())
            
            # Find matching asset
            download_url = None
            asset_name = None
            for asset in release_data.get('assets', []):
                name = asset.get('name', '')
                if asset_pattern in name and name.endswith('.zip'):
                    download_url = asset.get('browser_download_url')
                    asset_name = name
                    break
            
            if not download_url:
                messagebox.showerror("Error", 
                    f"Could not find NDecrypt release for your platform ({asset_pattern}).\n\n"
                    "Please download manually from:\nhttps://github.com/SabreTools/NDecrypt/releases")
                return
            
            self.log(f"Downloading {asset_name}...")
            
            # Download the zip
            zip_path = self.script_dir / asset_name
            urllib.request.urlretrieve(download_url, zip_path)
            
            # Extract
            self.log("Extracting NDecrypt...")
            extract_dir = self.script_dir / "ndecrypt"
            extract_dir.mkdir(exist_ok=True)
            
            with zipfile.ZipFile(zip_path, 'r') as zf:
                zf.extractall(extract_dir)
            
            # Find the executable
            if sys.platform == "win32":
                ndecrypt_exe = extract_dir / "NDecrypt.exe"
            else:
                ndecrypt_exe = extract_dir / "NDecrypt"
                # Make executable on Linux/Mac
                if ndecrypt_exe.exists():
                    os.chmod(ndecrypt_exe, 0o755)
            
            if ndecrypt_exe.exists():
                self.ndecrypt_path = str(ndecrypt_exe)
                self.save_config()
                self.log(f"✅ NDecrypt installed to: {self.ndecrypt_path}")
                if hasattr(self, 'ndecrypt_label'):
                    self.ndecrypt_label.config(text=self.ndecrypt_path,
                                              fg=COLORS['text_secondary'])
                messagebox.showinfo("Success", 
                    f"NDecrypt downloaded and installed!\n\n{self.ndecrypt_path}\n\n"
                    "Note: You'll need a config.json with encryption keys for decryption to work.")
            else:
                messagebox.showerror("Error", "Could not find NDecrypt executable after extraction")
            
            # Cleanup zip
            try:
                zip_path.unlink()
            except:
                pass
                
        except Exception as e:
            self.log(f"❌ Failed to download NDecrypt: {e}")
            messagebox.showerror("Error", 
                f"Failed to download NDecrypt:\n{e}\n\n"
                "Please download manually from:\nhttps://github.com/SabreTools/NDecrypt/releases")

    def check_ps3_dumper(self):
        """Check if ps3-disc-dumper is available"""
        # Determine platform-specific binary name
        dumper_name = "ps3-disc-dumper.exe" if sys.platform == "win32" else "ps3-disc-dumper"
        
        # First check bundled resources (PyInstaller)
        bundled_dumper = self.bundle_dir / dumper_name
        if bundled_dumper.exists():
            self.ps3_dumper_path = str(bundled_dumper)
            return True
        
        # Then check for dumper directly next to the executable/script
        direct_dumper = self.script_dir / dumper_name
        if direct_dumper.exists():
            self.ps3_dumper_path = str(direct_dumper)
            return True

        # Check in a subfolder
        subfolder_dumper = self.script_dir / "ps3-disc-dumper" / dumper_name
        if subfolder_dumper.exists():
            self.ps3_dumper_path = str(subfolder_dumper)
            return True

        # Check PATH as fallback
        dumper = shutil.which("ps3-disc-dumper")
        if dumper:
            self.ps3_dumper_path = dumper
            return True

        return False

    def download_ps3_dumper(self):
        """Download ps3-disc-dumper from GitHub releases"""
        try:
            # Determine platform
            if sys.platform == "win32":
                asset_pattern = "windows"
            elif sys.platform == "linux":
                asset_pattern = "linux"
            else:
                return False # Not supported or needs manual
            
            self.log("Fetching PS3 Disc Dumper release info from GitHub...")
            
            # Get latest release info
            req = urllib.request.Request(
                PS3_DISC_DUMPER_GITHUB_RELEASES_API,
                headers={'User-Agent': 'ROM-Converter'}
            )
            with urllib.request.urlopen(req, timeout=30) as response:
                release_data = json.loads(response.read().decode())
            
            # Find matching asset
            download_url = None
            asset_name = None
            for asset in release_data.get('assets', []):
                name = asset.get('name', '').lower()
                if asset_pattern in name and name.endswith('.zip'):
                    download_url = asset.get('browser_download_url')
                    asset_name = asset.get('name')
                    break
            
            if not download_url:
                self.log(f"⚠️ Could not find PS3 Disc Dumper release for {asset_pattern}")
                return False
            
            self.log(f"Downloading {asset_name}...")
            
            # Download the zip
            zip_path = self.script_dir / asset_name
            urllib.request.urlretrieve(download_url, zip_path)
            
            # Extract
            self.log("Extracting PS3 Disc Dumper...")
            extract_dir = self.script_dir / "ps3-disc-dumper"
            extract_dir.mkdir(exist_ok=True)
            
            with zipfile.ZipFile(zip_path, 'r') as zf:
                zf.extractall(extract_dir)
            
            # Find the executable
            if sys.platform == "win32":
                dumper_exe = extract_dir / "ps3-disc-dumper.exe"
            else:
                # Might be inside the zip or have a version name
                # Look for first file that starts with ps3-disc-dumper
                dumper_exe = None
                for f in extract_dir.glob("ps3-disc-dumper*"):
                    if f.is_file() and not f.suffix == '.zip':
                        dumper_exe = f
                        break
                
                # Make executable on Linux/Mac
                if dumper_exe and dumper_exe.exists():
                    os.chmod(dumper_exe, 0o755)
            
            if dumper_exe and dumper_exe.exists():
                self.ps3_dumper_path = str(dumper_exe)
                self.save_config()
                self.log(f"✅ PS3 Disc Dumper installed to: {self.ps3_dumper_path}")
                
                # Cleanup zip
                try:
                    zip_path.unlink()
                except:
                    pass
                return True
            else:
                self.log("❌ Could not find PS3 Disc Dumper executable after extraction")
                return False
                
        except Exception as e:
            self.log(f"❌ Failed to download PS3 Disc Dumper: {e}")
            return False

    def get_installed_ps3_dumper_version(self):
        """Get the version of the currently installed ps3-disc-dumper"""
        if not self.ps3_dumper_path:
            return None
        
        try:
            # Most modern versions support --version or similar, but often it's in the filename
            filename = Path(self.ps3_dumper_path).name
            version_match = re.search(r'v(\d+\.\d+\.\d+)', filename)
            if version_match:
                return version_match.group(1)
            
            # If not in name, it might be in the folder name
            folder_name = Path(self.ps3_dumper_path).parent.name
            version_match = re.search(r'v(\d+\.\d+\.\d+)', folder_name)
            if version_match:
                return version_match.group(1)

        except Exception as e:
            print(f"Error getting PS3 dumper version: {e}")
        
        return None

    def check_for_ps3_dumper_update(self):
        """Check if a newer version of ps3-disc-dumper is available"""
        try:
            installed_version = self.get_installed_ps3_dumper_version()
            
            # Fetch latest release info
            req = urllib.request.Request(
                PS3_DISC_DUMPER_GITHUB_RELEASES_API,
                headers={'User-Agent': 'ROM-Converter'}
            )
            with urllib.request.urlopen(req, timeout=30) as response:
                release_data = json.loads(response.read().decode())
            
            latest_version = release_data.get('tag_name', '').strip('v')
            
            if not installed_version or latest_version != installed_version:
                self.log(f"ℹ️ PS3 Disc Dumper update available: {latest_version} (Current: {installed_version or 'Unknown'})")
                # Automatically download updates since the user requested update check at launch
                self.download_ps3_dumper()
        except Exception as e:
            print(f"Error checking for PS3 dumper update: {e}")

    def check_auto_decrypt_ps3(self):
        """Check if AutoDecrypt.ps1 is available"""
        script_name = "AutoDecrypt.ps1"
        
        # Check in a subfolder
        subfolder_script = self.script_dir / "ps3-auto-decrypt" / script_name
        if subfolder_script.exists():
            self.auto_decrypt_ps3_path = str(subfolder_script)
            return True

        # Check directly next to the executable/script
        direct_script = self.script_dir / script_name
        if direct_script.exists():
            self.auto_decrypt_ps3_path = str(direct_script)
            return True

        return False

    def download_auto_decrypt_ps3(self):
        """Download AutoDecrypt.ps1 from GitHub releases"""
        if sys.platform != "win32":
            return False
            
        try:
            self.log("Fetching AutoDecryptPS3 release info from GitHub...")
            
            # Get latest release info
            req = urllib.request.Request(
                AUTO_DECRYPT_PS3_GITHUB_RELEASES_API,
                headers={'User-Agent': 'ROM-Converter'}
            )
            with urllib.request.urlopen(req, timeout=30) as response:
                release_data = json.loads(response.read().decode())
            
            # Find matching asset (AutoDecrypt.ps1)
            download_url = None
            asset_name = None
            for asset in release_data.get('assets', []):
                name = asset.get('name', '')
                if name.endswith('.ps1'):
                    download_url = asset.get('browser_download_url')
                    asset_name = name
                    break
            
            if not download_url:
                self.log(f"⚠️ Could not find AutoDecrypt.ps1 release asset")
                return False
            
            self.log(f"Downloading {asset_name}...")
            
            # Create subfolder
            extract_dir = self.script_dir / "ps3-auto-decrypt"
            extract_dir.mkdir(exist_ok=True)
            
            # Download the script
            script_path = extract_dir / "AutoDecrypt.ps1"
            urllib.request.urlretrieve(download_url, script_path)
            
            self.auto_decrypt_ps3_path = str(script_path)
            # Store the version in a file since it's just a script
            version_file = extract_dir / "version.txt"
            with open(version_file, "w") as f:
                f.write(release_data.get('tag_name', ''))
                
            self.save_config()
            self.log(f"✅ AutoDecryptPS3 installed to: {self.auto_decrypt_ps3_path}")
            return True
                
        except Exception as e:
            self.log(f"❌ Failed to download AutoDecryptPS3: {e}")
            return False

    def get_installed_auto_decrypt_ps3_version(self):
        """Get the version of the currently installed AutoDecryptPS3"""
        if not self.auto_decrypt_ps3_path:
            return None
        
        try:
            version_file = Path(self.auto_decrypt_ps3_path).parent / "version.txt"
            if version_file.exists():
                return version_file.read_text().strip().strip('v')
        except Exception:
            pass
        
        return None

    def check_for_auto_decrypt_ps3_update(self):
        """Check if a newer version of AutoDecryptPS3 is available"""
        if sys.platform != "win32":
            return

        try:
            installed_version = self.get_installed_auto_decrypt_ps3_version()
            
            # Fetch latest release info
            req = urllib.request.Request(
                AUTO_DECRYPT_PS3_GITHUB_RELEASES_API,
                headers={'User-Agent': 'ROM-Converter'}
            )
            with urllib.request.urlopen(req, timeout=30) as response:
                release_data = json.loads(response.read().decode())
            
            latest_version = release_data.get('tag_name', '').strip('v')
            
            if not installed_version or latest_version != installed_version:
                self.log(f"ℹ️ AutoDecryptPS3 update available: {latest_version} (Current: {installed_version or 'Unknown'})")
                self.download_auto_decrypt_ps3()
        except Exception as e:
            print(f"Error checking for AutoDecryptPS3 update: {e}")

    def run_auto_decrypt_ps3(self):
        """Trigger the AutoDecryptPS3 script"""
        if not self.auto_decrypt_ps3_path:
            if self.download_auto_decrypt_ps3():
                pass
            else:
                messagebox.showerror("Error", "AutoDecryptPS3 script not found and could not be downloaded.")
                return
        
        try:
            # Open the script directory or the source directory?
            # User might want to run it in their ROM folder
            source = self.source_dir or str(self.script_dir)
            
            # Confirm with user
            response = messagebox.askyesno("AutoDecrypt PS3", 
                f"This will launch AutoDecrypt.ps1 in a PowerShell window.\n\n"
                f"Source: {source}\n\n"
                "Note: You may need to copy the script to your ROM folder if it doesn't work from here.\n"
                "Continue?")
            
            if not response:
                return

            # Launch PowerShell with execution policy bypass
            cmd = ["powershell.exe", "-NoExit", "-ExecutionPolicy", "Bypass", "-File", self.auto_decrypt_ps3_path]
            
            # We want it to run in the source directory if it exists
            subprocess.Popen(cmd, cwd=source if os.path.isdir(source) else None)
            self.log("✅ AutoDecryptPS3 launched in PowerShell window.")
            
        except Exception as e:
            self.log(f"❌ Failed to launch AutoDecryptPS3: {e}")
            messagebox.showerror("Error", f"Failed to launch AutoDecryptPS3:\n{e}")

    def find_aes_keys_file(self):
        """Find the bundled or local 3DS AES Keys file"""
        possible_locations = [
            self.bundle_dir / "3DS AES Keys.txt",
            self.script_dir / "3DS AES Keys.txt",
            Path("3DS AES Keys.txt"),
        ]
        for loc in possible_locations:
            if loc.exists():
                return loc
        return None
    
    def convert_aes_keys_to_config(self, aes_keys_path, output_path):
        """Convert aes_keys.txt format to NDecrypt config.json format"""
        try:
            keys = {}
            with open(aes_keys_path, 'r') as f:
                for line in f:
                    line = line.strip()
                    if '=' in line and not line.startswith('#'):
                        key, value = line.split('=', 1)
                        keys[key.strip()] = value.strip()
            
            # Map aes_keys.txt format to NDecrypt config.json format
            config = {}
            
            # Generator -> AESHardwareConstant
            if 'generator' in keys:
                config['AESHardwareConstant'] = keys['generator']
            
            # Map KeyX slots
            key_mapping = {
                'slot0x18KeyX': 'KeyX0x18',
                'slot0x1BKeyX': 'KeyX0x1B',
                'slot0x25KeyX': 'KeyX0x25',
                'slot0x2CKeyX': 'KeyX0x2C',
                'slot0x2DKeyX': 'KeyX0x2D',
                'slot0x2EKeyX': 'KeyX0x2E',
                'slot0x2FKeyX': 'KeyX0x2F',
                'slot0x30KeyX': 'KeyX0x30',
                'slot0x31KeyX': 'KeyX0x31',
                'slot0x32KeyX': 'KeyX0x32',
                'slot0x33KeyX': 'KeyX0x33',
                'slot0x34KeyX': 'KeyX0x34',
                'slot0x35KeyX': 'KeyX0x35',
                'slot0x36KeyX': 'KeyX0x36',
                'slot0x37KeyX': 'KeyX0x37',
                'slot0x38KeyX': 'KeyX0x38',
                'slot0x39KeyX': 'KeyX0x39',
                'slot0x3AKeyX': 'KeyX0x3A',
                'slot0x3BKeyX': 'KeyX0x3B',
                'slot0x3CKeyX': 'KeyX0x3C',
                'slot0x3DKeyX': 'KeyX0x3D',
                'slot0x3EKeyX': 'KeyX0x3E',
                'slot0x3FKeyX': 'KeyX0x3F',
                'slot0x03KeyX': 'KeyX0x03',
                'slot0x19KeyX': 'KeyX0x19',
                'slot0x1AKeyX': 'KeyX0x1A',
                'slot0x1CKeyX': 'KeyX0x1C',
                'slot0x1DKeyX': 'KeyX0x1D',
                'slot0x1EKeyX': 'KeyX0x1E',
                'slot0x1FKeyX': 'KeyX0x1F',
            }
            
            for aes_key, config_key in key_mapping.items():
                if aes_key in keys:
                    config[config_key] = keys[aes_key]
            
            # Map KeyY slots if present
            keyy_mapping = {
                'slot0x18KeyY': 'KeyY0x18',
                'slot0x1BKeyY': 'KeyY0x1B',
                'slot0x25KeyY': 'KeyY0x25',
                'slot0x2CKeyY': 'KeyY0x2C',
            }
            
            for aes_key, config_key in keyy_mapping.items():
                if aes_key in keys:
                    config[config_key] = keys[aes_key]
            
            # Map KeyN (normal keys) if present
            keyn_mapping = {
                'slot0x18KeyN': 'KeyN0x18',
                'slot0x1BKeyN': 'KeyN0x1B',
                'slot0x25KeyN': 'KeyN0x25',
                'slot0x2CKeyN': 'KeyN0x2C',
            }
            
            for aes_key, config_key in keyn_mapping.items():
                if aes_key in keys:
                    config[config_key] = keys[aes_key]
            
            # Write config.json
            with open(output_path, 'w') as f:
                json.dump(config, f, indent=2)
            
            return True
        except Exception as e:
            self.log(f"❌ Failed to convert AES keys: {e}")
            return False
    
    def setup_ndecrypt_keys(self):
        """Setup NDecrypt with bundled AES keys if available"""
        if not self.ndecrypt_path:
            return False
        
        ndecrypt_dir = Path(self.ndecrypt_path).parent
        config_path = ndecrypt_dir / "config.json"
        
        # Check if config.json already exists
        if config_path.exists():
            return True
        
        # Try to find and convert bundled AES keys
        aes_keys_path = self.find_aes_keys_file()
        if aes_keys_path:
            self.log(f"📋 Found AES keys at: {aes_keys_path}")
            self.log("🔄 Converting to NDecrypt config.json format...")
            if self.convert_aes_keys_to_config(aes_keys_path, config_path):
                self.log(f"✅ Created config.json at: {config_path}")
                return True
        
        return False
    
    def clean_rom_filename(self, filename):
        """Clean a ROM filename by removing unwanted tags while preserving important ones."""
        name = filename
        
        # Tags to KEEP (case-insensitive patterns that should be preserved)
        keep_patterns = [
            r'\(Disc\s*\d+\)', r'\(Disk\s*\d+\)', r'\(Bonus\s*Disc\)', r'\(Bonus\s*Disk\)',
            r'\(Demo\)', r'\(Beta\)', r'\(Proto\)', r'\(Prototype\)', r'\(Sample\)',
            r'\(Limited\s*Edition\)', r'\(Collector.?s?\s*Edition\)', r'\(Special\s*Edition\)',
            r'\(Game\s*of.*Year\)', r'\(GOTY\)', r'\(Director.?s?\s*Cut\)', r'\(Uncut\)',
            r'\(Part\s*\d+\)', r'\(Side\s*[AB]\)',
        ]
        
        # Extract tags to keep
        preserved_tags = []
        for pattern in keep_patterns:
            matches = re.findall(pattern, name, re.IGNORECASE)
            preserved_tags.extend(matches)
        
        # Tags to REMOVE
        remove_patterns = [
            r'\(USA\)', r'\(U\)', r'\(America\)', r'\(Europe\)', r'\(E\)', r'\(EU\)',
            r'\(Japan\)', r'\(J\)', r'\(JP\)', r'\(Korea\)', r'\(K\)', r'\(KR\)',
            r'\(Asia\)', r'\(A\)', r'\(World\)', r'\(W\)', r'\(Australia\)', r'\(AU\)',
            r'\(France\)', r'\(F\)', r'\(Fr\)', r'\(Germany\)', r'\(G\)', r'\(De\)',
            r'\(Spain\)', r'\(S\)', r'\(Es\)', r'\(Italy\)', r'\(I\)', r'\(It\)',
            r'\(En\)', r'\(En,.*?\)', r'\(English\)', r'\(French\)', r'\(German\)',
            r'\(Spanish\)', r'\(Italian\)', r'\(Japanese\)',
            r'\(Multi\)', r'\(Multi\d*\)', r'\(M\d+\)',
            r'\(Rev\s*[\dA-Z\.]+\)', r'\(v[\d\.]+[a-z]?\)', r'\(Ver\.?\s*[\d\.]+\)',
            r'\[!\]', r'\[a\d?\]', r'\[b\d?\]', r'\[c\]', r'\[f\d?\]',
            r'\[h\d*[A-Za-z]*\]', r'\[o\d?\]', r'\[p\d?\]', r'\[t\d?\]',
            r'\[T[+-][A-Za-z]+[^\]]*\]',
            r'\(NTSC\)', r'\(NTSC-U\)', r'\(NTSC-J\)', r'\(PAL\)', r'\(SECAM\)',
            r'\(\d{4}-\d{2}-\d{2}\)', r'\(\d{8}\)', r'\(Unl\)',
            r'\(Decrypted\)', r'\(Encrypted\)',
        ]
        
        # Multi-region patterns
        region_words = [
            'USA', 'Europe', 'Japan', 'Asia', 'World', 'Korea', 'Australia',
            'France', 'Germany', 'Spain', 'Italy', 'En', 'Fr', 'De', 'Es', 'It',
            'U', 'E', 'J', 'A', 'K', 'W', 'G', 'F', 'S', 'I', 'EU', 'JP', 'KR', 'AU'
        ]
        region_pattern = r'\(\s*(?:' + '|'.join(re.escape(r) for r in region_words) + r')(?:\s*,\s*(?:' + '|'.join(re.escape(r) for r in region_words) + r'))+\s*\)'
        name = re.sub(region_pattern, '', name, flags=re.IGNORECASE)
        
        # Remove unwanted tags
        for pattern in remove_patterns:
            name = re.sub(pattern, '', name, flags=re.IGNORECASE)
        
        # Remove 2-3 letter codes in parentheses
        name = re.sub(r'\s*\([A-Za-z]{1,3}\)(?!\s*\.)', '', name)
        
        # Clean up spaces
        name = re.sub(r'\s+', ' ', name)
        name = re.sub(r'\s+\.', '.', name)
        name = name.strip()
        
        return name

    def about_dialog(self):
        """Show About dialog with credits"""
        dialog = Toplevel(self.master)
        dialog.title("◄ ABOUT ►")
        dialog.geometry("550x520")
        dialog.resizable(False, False)
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.configure(bg=COLORS['bg_dark'])
        
        # Title
        title_frame = Frame(dialog, bg=COLORS['bg_light'], pady=12)
        title_frame.pack(fill="x", padx=10, pady=(10, 10))
        Label(title_frame, text="◄ ROM CONVERTER ►", font=self.font_title,
              fg=COLORS['text_primary'], bg=COLORS['bg_light']).pack()
        Label(title_frame, text="ROM Management & Conversion Tool", font=self.font_small,
              fg=COLORS['text_muted'], bg=COLORS['bg_light']).pack(pady=(4, 0))
        
        # Developer
        dev_frame = Frame(dialog, bg=COLORS['bg_medium'], pady=10, padx=10)
        dev_frame.pack(fill="x", padx=10, pady=5)
        Label(dev_frame, text="👨‍💻 DEVELOPED BY", font=self.font_label_bold,
              fg=COLORS['accent_yellow'], bg=COLORS['bg_medium']).pack()
        Label(dev_frame, text="WoofahRayetCode", font=self.font_heading_md,
              fg=COLORS['accent_purple'], bg=COLORS['bg_medium']).pack(pady=(4, 0))
        Label(dev_frame, text=f"Build: {self.build_timestamp}", font=self.font_small,
              fg=COLORS['text_secondary'], bg=COLORS['bg_medium']).pack(pady=(6, 0))
        
        # Credits section
        credits_frame = Frame(dialog, bg=COLORS['bg_dark'], padx=10, pady=5)
        credits_frame.pack(fill="both", expand=True, padx=10)
        
        Label(credits_frame, text="🛠️ TOOLS & CREDITS", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=COLORS['bg_dark']).pack(anchor="w", pady=(5, 8))
        
        # Scrollable credits
        credits_text = Text(credits_frame, wrap="word", height=14, font=self.font_body,
                           bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                           relief="flat", padx=10, pady=10)
        credits_text.pack(fill="both", expand=True)
        
        credits_content = """◆ CHDMAN
   Developer: MAME Team
   Purpose: CHD compression for disc images
   License: BSD/GPL

◆ MAXCSO
   Developer: Unknown W. Brackets
   Purpose: CSO/ZSO compression for PSP ISOs
   License: ISC

◆ 7-Zip
   Developer: Igor Pavlov
   Purpose: Archive extraction (7z, RAR, etc.)
   License: LGPL/BSD

◆ NDecrypt
   Developer: SabreTools Team
   Purpose: 3DS ROM decryption
   License: MIT

◆ Python & Tkinter
   Developer: Python Software Foundation
   Purpose: Application framework
   License: PSF License

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

This program is provided as-is for personal use.
Please respect copyright and only use with legally
obtained ROM files.
"""
        credits_text.insert("1.0", credits_content)
        credits_text.configure(state="disabled")
        
        # Close button
        Button(dialog, text="✕ CLOSE", command=dialog.destroy,
               font=self.font_button, bg=COLORS['bg_light'],
               fg=COLORS['text_primary'], relief="flat", cursor="hand2",
               padx=20, pady=5).pack(pady=10)

    def decrypt_3ds_dialog(self):
        """Open dialog for 3DS ROM management: extract, decrypt, and move"""
        dialog = Toplevel(self.master)
        dialog.title("◄ 3DS ROM MANAGER ►")
        dialog.geometry("850x750")
        dialog.resizable(True, True)
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.configure(bg=COLORS['bg_dark'])
        
        # Title
        title_frame = Frame(dialog, bg=COLORS['bg_light'], pady=6)
        title_frame.pack(fill="x", padx=10, pady=(10, 10))
        Label(title_frame, text="🎮 3DS ROM MANAGER", font=self.font_heading_md,
              fg=COLORS['accent_purple'], bg=COLORS['bg_light']).pack()
        Label(title_frame, text="Extract → Decrypt → Move", font=self.font_small,
              fg=COLORS['text_muted'], bg=COLORS['bg_light']).pack()
        
        # Status frame
        status_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        status_frame.pack(fill="x")
        
        # Check for keys and NDecrypt
        keys_available = self.find_aes_keys_file() is not None
        config_exists = False
        if self.ndecrypt_path:
            config_exists = (Path(self.ndecrypt_path).parent / "config.json").exists()
        
        if self.ndecrypt_path and (keys_available or config_exists):
            ndecrypt_status = "✅ NDecrypt Ready"
            status_color = COLORS['button_green']
        elif self.ndecrypt_path:
            ndecrypt_status = "⚠️ NDecrypt OK, keys missing"
            status_color = COLORS['accent_yellow']
        else:
            ndecrypt_status = "❌ NDecrypt not configured"
            status_color = COLORS['accent_red']
        
        Label(status_frame, text=f"Decryption: {ndecrypt_status}", 
              font=self.font_small, fg=status_color, bg=COLORS['bg_dark']).pack(side="left", padx=(0, 20))
        
        if not self.ndecrypt_path:
            Button(status_frame, text="[ DOWNLOAD ]", command=self.download_ndecrypt,
                   font=self.font_small, bg=COLORS['accent_purple'],
                   fg="white", relief="flat", cursor="hand2").pack(side="left")
        
        # Source directory (for archives or ROMs)
        source_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        source_frame.pack(fill="x")
        
        Label(source_frame, text="📂 Source Folder:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(side="left")
        source_entry = Entry(source_frame, font=self.font_body,
                            bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                            insertbackground=COLORS['text_primary'], relief="flat")
        source_entry.pack(side="left", fill="x", expand=True, padx=5, ipady=3)
        # Pre-fill with saved 3DS source dir, fallback to general source_dir
        if self.threeds_source_dir:
            source_entry.insert(0, self.threeds_source_dir)
        elif self.source_dir:
            source_entry.insert(0, self.source_dir)
        
        def browse_source():
            folder = filedialog.askdirectory(title="Select Source Folder")
            if folder:
                source_entry.delete(0, "end")
                source_entry.insert(0, folder)
        
        Button(source_frame, text="📁", command=browse_source,
               font=self.font_small, bg=COLORS['bg_light'],
               fg=COLORS['text_secondary'], relief="flat", cursor="hand2").pack(side="left")
        
        # Destination directory (for moving)
        dest_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        dest_frame.pack(fill="x")
        
        Label(dest_frame, text="📁 Destination:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(side="left")
        dest_entry = Entry(dest_frame, font=self.font_body,
                          bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                          insertbackground=COLORS['text_primary'], relief="flat")
        dest_entry.pack(side="left", fill="x", expand=True, padx=5, ipady=3)
        
        # Pre-fill with saved 3DS destination, fallback to system_extract_dirs
        if self.threeds_dest_dir:
            dest_entry.insert(0, self.threeds_dest_dir)
        elif 'Nintendo 3DS' in self.system_extract_dirs:
            dest_entry.insert(0, self.system_extract_dirs['Nintendo 3DS'])
        
        def browse_dest():
            folder = filedialog.askdirectory(title="Select Destination Folder for 3DS ROMs")
            if folder:
                dest_entry.delete(0, "end")
                dest_entry.insert(0, folder)
                self.system_extract_dirs['Nintendo 3DS'] = folder
                self.save_config()
        
        Button(dest_frame, text="📁", command=browse_dest,
               font=self.font_small, bg=COLORS['bg_light'],
               fg=COLORS['text_secondary'], relief="flat", cursor="hand2").pack(side="left")
        
        # Options
        options_frame = Frame(dialog, padx=10, pady=8, bg=COLORS['bg_light'])
        options_frame.pack(fill="x", padx=10, pady=5)
        
        cb_bg = COLORS['bg_light']
        
        # Row 1
        opt_row1 = Frame(options_frame, bg=cb_bg)
        opt_row1.pack(fill="x", pady=2)
        
        recursive_scan = BooleanVar(value=True)
        Checkbutton(opt_row1, text="↳ Scan subdirectories", 
                   variable=recursive_scan, font=self.font_small,
                   fg=COLORS['text_secondary'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                   activebackground=cb_bg).pack(side="left", padx=(0, 20))
        
        backup_original = BooleanVar(value=self.threeds_backup_original)
        Checkbutton(opt_row1, text="📁 Backup before decryption", 
                   variable=backup_original, font=self.font_small,
                   fg=COLORS['text_secondary'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                   activebackground=cb_bg).pack(side="left", padx=(0, 20))
        
        auto_clean_names = BooleanVar(value=self.threeds_auto_clean_names)
        Checkbutton(opt_row1, text="✨ Clean filenames", 
                   variable=auto_clean_names, font=self.font_small,
                   fg=COLORS['text_secondary'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                   activebackground=cb_bg).pack(side="left")
        
        # Row 2 - Workflow options
        opt_row2 = Frame(options_frame, bg=cb_bg)
        opt_row2.pack(fill="x", pady=2)
        
        delete_archives = BooleanVar(value=self.threeds_delete_archives)
        Checkbutton(opt_row2, text="🗑️ Delete archives after extraction", 
                   variable=delete_archives, font=self.font_small,
                   fg=COLORS['accent_red'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                   activebackground=cb_bg).pack(side="left", padx=(0, 20))
        
        delete_after_move = BooleanVar(value=self.threeds_delete_after_move)
        Checkbutton(opt_row2, text="🗑️ Delete source after move", 
                   variable=delete_after_move, font=self.font_small,
                   fg=COLORS['accent_red'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                   activebackground=cb_bg).pack(side="left")
        
        # Function to save settings when checkboxes change
        def save_3ds_settings(*args):
            self.threeds_backup_original = backup_original.get()
            self.threeds_delete_archives = delete_archives.get()
            self.threeds_delete_after_move = delete_after_move.get()
            self.threeds_auto_clean_names = auto_clean_names.get()
            self.save_config()
        
        # Trace checkbox changes
        backup_original.trace_add('write', save_3ds_settings)
        delete_archives.trace_add('write', save_3ds_settings)
        delete_after_move.trace_add('write', save_3ds_settings)
        auto_clean_names.trace_add('write', save_3ds_settings)
        
        # Save paths when focus leaves entry fields
        def save_source_path(event=None):
            self.threeds_source_dir = source_entry.get().strip()
            self.save_config()
        
        def save_dest_path(event=None):
            self.threeds_dest_dir = dest_entry.get().strip()
            self.system_extract_dirs['Nintendo 3DS'] = self.threeds_dest_dir
            self.save_config()
        
        source_entry.bind('<FocusOut>', save_source_path)
        dest_entry.bind('<FocusOut>', save_dest_path)
        
        # Results area
        results_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        results_frame.pack(fill="both", expand=True)
        
        Label(results_frame, text="◄ OPERATION LOG ►", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=COLORS['bg_dark']).pack(anchor="w", pady=(0, 4))
        
        list_frame = Frame(results_frame, bg=COLORS['bg_dark'])
        list_frame.pack(fill="both", expand=True)
        
        scrollbar = Scrollbar(list_frame, bg=COLORS['bg_light'],
                             troughcolor=COLORS['bg_dark'])
        scrollbar.pack(side="right", fill="y")
        
        results_text = Text(list_frame, wrap="word", yscrollcommand=scrollbar.set,
                           height=18, font=self.font_mono,
                           bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                           insertbackground=COLORS['text_primary'], relief="flat",
                           padx=8, pady=8)
        results_text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=results_text.yview)
        
        # Track found items
        found_archives = []
        found_roms = []
        extracted_roms = []
        
        def log_msg(msg):
            results_text.insert("end", msg + "\n")
            results_text.see("end")
            dialog.update()
        
        def clear_log():
            results_text.delete("1.0", "end")
        
        # ===== STEP 1: EXTRACT =====
        def extract_3ds_archives():
            """Extract 3DS ROMs from archives"""
            source = source_entry.get()
            if not source or not os.path.isdir(source):
                messagebox.showwarning("Warning", "Please select a valid source folder")
                return
            
            clear_log()
            found_archives.clear()
            extracted_roms.clear()
            
            log_msg("📦 STEP 1: EXTRACTING 3DS ARCHIVES")
            log_msg("=" * 50)
            log_msg(f"Scanning: {source}\n")
            
            # Find archives
            archives = self.find_compressed_files(source, recursive_scan.get())
            
            if not archives:
                log_msg("No archive files found.")
                return
            
            log_msg(f"Found {len(archives)} archive(s). Scanning for 3DS content...\n")
            
            extensions_3ds = {'.3ds', '.cia'}
            archives_with_3ds = []
            
            for archive in archives:
                try:
                    archive_path = Path(archive)
                    ext = archive_path.suffix.lower()
                    has_3ds = False
                    
                    if ext == '.zip':
                        with zipfile.ZipFile(archive_path, 'r') as zf:
                            for name in zf.namelist():
                                if Path(name).suffix.lower() in extensions_3ds:
                                    has_3ds = True
                                    break
                    elif ext in ['.7z', '.rar'] and self.seven_zip_path:
                        cmd = [self.seven_zip_path, 'l', '-ba', str(archive_path)]
                        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                        if result.returncode == 0:
                            for line in result.stdout.split('\n'):
                                if any(line.lower().endswith(e) for e in extensions_3ds):
                                    has_3ds = True
                                    break
                    
                    if has_3ds:
                        archives_with_3ds.append(archive)
                        found_archives.append(archive)
                        log_msg(f"  📦 {archive.name} - contains 3DS ROMs")
                except Exception as e:
                    log_msg(f"  ⚠️ {archive.name} - scan error: {e}")
            
            if not archives_with_3ds:
                log_msg("\nNo archives containing 3DS ROMs found.")
                return
            
            log_msg(f"\n{'─' * 50}")
            log_msg(f"Found {len(archives_with_3ds)} archive(s) with 3DS ROMs")
            log_msg(f"{'─' * 50}\n")
            
            # Extract
            for archive in archives_with_3ds:
                log_msg(f"📦 Extracting: {archive.name}")
                success, folder = self.extract_archive(archive)
                
                if success and folder:
                    # Find extracted 3DS files
                    for f in folder.rglob("*"):
                        if f.is_file() and f.suffix.lower() in extensions_3ds:
                            extracted_roms.append(f)
                            log_msg(f"   ✅ {f.name}")
                    
                    if delete_archives.get():
                        try:
                            archive.unlink()
                            log_msg(f"   🗑️ Archive deleted")
                        except Exception as e:
                            log_msg(f"   ⚠️ Could not delete archive: {e}")
                else:
                    log_msg(f"   ❌ Extraction failed")
            
            log_msg(f"\n{'━' * 50}")
            log_msg(f"✅ Extracted {len(extracted_roms)} 3DS ROM(s)")
            
            # Update found_roms with extracted ones
            found_roms.clear()
            found_roms.extend(extracted_roms)
        
        # ===== STEP 2: DECRYPT =====
        def decrypt_3ds_roms():
            """Decrypt 3DS ROM files"""
            if not self.ndecrypt_path:
                messagebox.showerror("Error", "NDecrypt not configured. Please download it first.")
                return
            
            # If no ROMs found yet, scan for them
            if not found_roms:
                source = source_entry.get()
                if not source or not os.path.isdir(source):
                    messagebox.showwarning("Warning", "Please select a valid source folder")
                    return
                
                clear_log()
                log_msg("🔓 STEP 2: DECRYPTING 3DS ROMS")
                log_msg("=" * 50)
                log_msg("Scanning for 3DS ROM files...\n")
                
                path = Path(source)
                extensions_3ds = {'.3ds', '.cia'}
                
                if recursive_scan.get():
                    all_files = list(path.rglob("*"))
                else:
                    all_files = list(path.glob("*"))
                
                for f in all_files:
                    if f.is_file() and f.suffix.lower() in extensions_3ds:
                        found_roms.append(f)
                
                if not found_roms:
                    log_msg("No .3ds or .cia files found.")
                    return
                
                log_msg(f"Found {len(found_roms)} ROM(s)")
            else:
                clear_log()
                log_msg("🔓 STEP 2: DECRYPTING 3DS ROMS")
                log_msg("=" * 50)
            
            # Setup keys
            log_msg("\nChecking encryption keys...")
            if self.setup_ndecrypt_keys():
                log_msg("✅ Keys configured\n")
            else:
                config_path = Path(self.ndecrypt_path).parent / "config.json"
                if not config_path.exists():
                    log_msg("❌ Keys not found!")
                    messagebox.showerror("Error", "Encryption keys not found!")
                    return
            
            success_count = 0
            error_count = 0
            
            for rom_file in found_roms:
                log_msg(f"🔓 {rom_file.name}")
                
                try:
                    if backup_original.get():
                        backup_dir = rom_file.parent / "encrypted_backup"
                        backup_dir.mkdir(exist_ok=True)
                        backup_path = backup_dir / rom_file.name
                        if not backup_path.exists():
                            shutil.copy2(rom_file, backup_path)
                            log_msg(f"   📁 Backup created")
                    
                    cmd = [self.ndecrypt_path, "d", str(rom_file)]
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=600,
                                          cwd=Path(self.ndecrypt_path).parent)
                    
                    if result.returncode == 0:
                        log_msg(f"   ✅ Decrypted")
                        success_count += 1
                    else:
                        error_msg = result.stderr.strip() or result.stdout.strip() or "Unknown error"
                        log_msg(f"   ❌ Failed: {error_msg[:80]}")
                        error_count += 1
                        
                except subprocess.TimeoutExpired:
                    log_msg(f"   ❌ Timeout")
                    error_count += 1
                except Exception as e:
                    log_msg(f"   ❌ Error: {e}")
                    error_count += 1
            
            log_msg(f"\n{'━' * 50}")
            log_msg(f"✅ Decrypted: {success_count} | ❌ Errors: {error_count}")
        
        # ===== STEP 3: MOVE =====
        def move_3ds_roms():
            """Move 3DS ROMs to destination folder"""
            dest = dest_entry.get().strip()
            if not dest:
                messagebox.showwarning("Warning", "Please specify a destination folder")
                return
            
            # If no ROMs tracked, scan for them
            if not found_roms:
                source = source_entry.get()
                if not source or not os.path.isdir(source):
                    messagebox.showwarning("Warning", "Please select a valid source folder")
                    return
                
                path = Path(source)
                extensions_3ds = {'.3ds', '.cia'}
                
                if recursive_scan.get():
                    all_files = list(path.rglob("*"))
                else:
                    all_files = list(path.glob("*"))
                
                for f in all_files:
                    if f.is_file() and f.suffix.lower() in extensions_3ds:
                        found_roms.append(f)
            
            if not found_roms:
                messagebox.showwarning("Warning", "No 3DS ROMs found to move")
                return
            
            clear_log()
            log_msg("📁 STEP 3: MOVING 3DS ROMS")
            log_msg("=" * 50)
            log_msg(f"Destination: {dest}\n")
            
            dest_path = Path(dest)
            dest_path.mkdir(parents=True, exist_ok=True)
            
            moved_count = 0
            error_count = 0
            
            for rom_file in found_roms:
                try:
                    # Apply filename cleaning if enabled
                    if auto_clean_names.get():
                        clean_name = self.clean_rom_filename(rom_file.name)
                        target = dest_path / clean_name
                    else:
                        target = dest_path / rom_file.name
                    
                    if target.exists():
                        log_msg(f"⚠️ {rom_file.name} - already exists, skipping")
                        continue
                    
                    if delete_after_move.get():
                        shutil.move(str(rom_file), str(target))
                        if auto_clean_names.get() and target.name != rom_file.name:
                            log_msg(f"✅ {rom_file.name} → {target.name}")
                        else:
                            log_msg(f"✅ {rom_file.name} - moved")
                    else:
                        shutil.copy2(rom_file, target)
                        if auto_clean_names.get() and target.name != rom_file.name:
                            log_msg(f"✅ {rom_file.name} → {target.name}")
                        else:
                            log_msg(f"✅ {rom_file.name} - copied")
                    moved_count += 1
                except Exception as e:
                    log_msg(f"❌ {rom_file.name} - {e}")
                    error_count += 1
            
            log_msg(f"\n{'━' * 50}")
            log_msg(f"✅ Processed: {moved_count} | ❌ Errors: {error_count}")
            
            # Save destination for future use
            self.system_extract_dirs['Nintendo 3DS'] = dest
            self.save_config()
        
        # ===== ALL-IN-ONE =====
        def run_full_workflow():
            """Run complete Extract → Decrypt → Move workflow"""
            source = source_entry.get()
            dest = dest_entry.get().strip()
            
            if not source or not os.path.isdir(source):
                messagebox.showwarning("Warning", "Please select a valid source folder")
                return
            
            if not dest:
                messagebox.showwarning("Warning", "Please specify a destination folder")
                return
            
            if not self.ndecrypt_path:
                messagebox.showerror("Error", "NDecrypt not configured. Please download it first.")
                return
            
            confirm = messagebox.askyesno(
                "Run Full Workflow",
                "This will:\n\n"
                "1. Extract 3DS ROMs from archives\n"
                "2. Decrypt all .3ds/.cia files\n"
                "3. Move decrypted ROMs to destination\n\n"
                f"Source: {source}\n"
                f"Destination: {dest}\n\n"
                "Continue?"
            )
            
            if not confirm:
                return
            
            clear_log()
            log_msg("🚀 FULL WORKFLOW: EXTRACT → DECRYPT → MOVE")
            log_msg("=" * 50)
            log_msg("")
            
            # Step 1: Extract
            found_archives.clear()
            found_roms.clear()
            extracted_roms.clear()
            
            log_msg("📦 STEP 1: EXTRACTING ARCHIVES")
            log_msg("─" * 40)
            
            archives = self.find_compressed_files(source, recursive_scan.get())
            extensions_3ds = {'.3ds', '.cia'}
            
            if archives:
                for archive in archives:
                    try:
                        archive_path = Path(archive)
                        ext = archive_path.suffix.lower()
                        has_3ds = False
                        
                        if ext == '.zip':
                            with zipfile.ZipFile(archive_path, 'r') as zf:
                                for name in zf.namelist():
                                    if Path(name).suffix.lower() in extensions_3ds:
                                        has_3ds = True
                                        break
                        
                        if has_3ds:
                            log_msg(f"📦 {archive.name}")
                            success, folder = self.extract_archive(archive)
                            if success and folder:
                                for f in folder.rglob("*"):
                                    if f.is_file() and f.suffix.lower() in extensions_3ds:
                                        found_roms.append(f)
                                        log_msg(f"   ✅ {f.name}")
                                
                                if delete_archives.get():
                                    try:
                                        archive.unlink()
                                        log_msg(f"   🗑️ Archive deleted")
                                    except:
                                        pass
                    except:
                        pass
            
            # Also scan for loose ROM files
            path = Path(source)
            if recursive_scan.get():
                all_files = list(path.rglob("*"))
            else:
                all_files = list(path.glob("*"))
            
            for f in all_files:
                if f.is_file() and f.suffix.lower() in extensions_3ds and f not in found_roms:
                    found_roms.append(f)
            
            log_msg(f"\n✅ Found {len(found_roms)} 3DS ROM(s)\n")
            
            if not found_roms:
                log_msg("No 3DS ROMs found. Workflow stopped.")
                return
            
            # Step 2: Decrypt
            log_msg("🔓 STEP 2: DECRYPTING")
            log_msg("─" * 40)
            
            if self.setup_ndecrypt_keys():
                log_msg("✅ Keys configured\n")
            
            decrypt_success = 0
            decrypt_errors = 0
            
            for rom_file in found_roms:
                log_msg(f"🔓 {rom_file.name}")
                try:
                    if backup_original.get():
                        backup_dir = rom_file.parent / "encrypted_backup"
                        backup_dir.mkdir(exist_ok=True)
                        backup_path = backup_dir / rom_file.name
                        if not backup_path.exists():
                            shutil.copy2(rom_file, backup_path)
                    
                    cmd = [self.ndecrypt_path, "d", str(rom_file)]
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=600,
                                          cwd=Path(self.ndecrypt_path).parent)
                    
                    if result.returncode == 0:
                        log_msg(f"   ✅ Decrypted")
                        decrypt_success += 1
                    else:
                        log_msg(f"   ❌ Failed")
                        decrypt_errors += 1
                except:
                    log_msg(f"   ❌ Error")
                    decrypt_errors += 1
            
            log_msg(f"\n✅ Decrypted: {decrypt_success} | ❌ Errors: {decrypt_errors}\n")
            
            # Step 3: Move
            log_msg("📁 STEP 3: MOVING TO DESTINATION")
            log_msg("─" * 40)
            log_msg(f"Destination: {dest}\n")
            
            dest_path = Path(dest)
            dest_path.mkdir(parents=True, exist_ok=True)
            
            move_success = 0
            move_errors = 0
            
            for rom_file in found_roms:
                try:
                    if not rom_file.exists():
                        continue
                    
                    # Apply filename cleaning if enabled
                    if auto_clean_names.get():
                        clean_name = self.clean_rom_filename(rom_file.name)
                        target = dest_path / clean_name
                    else:
                        target = dest_path / rom_file.name
                    
                    if target.exists():
                        log_msg(f"⚠️ {rom_file.name} - exists")
                        continue
                    
                    if delete_after_move.get():
                        shutil.move(str(rom_file), str(target))
                    else:
                        shutil.copy2(rom_file, target)
                    
                    if auto_clean_names.get() and target.name != rom_file.name:
                        log_msg(f"✅ {rom_file.name} → {target.name}")
                    else:
                        log_msg(f"✅ {rom_file.name}")
                    move_success += 1
                except Exception as e:
                    log_msg(f"❌ {rom_file.name}")
                    move_errors += 1
            
            log_msg(f"\n{'━' * 50}")
            log_msg("🎉 WORKFLOW COMPLETE!")
            log_msg(f"{'━' * 50}")
            log_msg(f"Extracted: {len(found_roms)} ROM(s)")
            log_msg(f"Decrypted: {decrypt_success} | Errors: {decrypt_errors}")
            log_msg(f"Moved: {move_success} | Errors: {move_errors}")
            
            # Save paths and settings
            self.threeds_source_dir = source
            self.threeds_dest_dir = dest
            self.system_extract_dirs['Nintendo 3DS'] = dest
            self.save_config()
            
            messagebox.showinfo("Workflow Complete",
                f"✅ Extracted: {len(found_roms)} ROM(s)\n"
                f"✅ Decrypted: {decrypt_success}\n"
                f"✅ Moved: {move_success}\n\n"
                f"Destination: {dest}")
        
        # Action buttons - Individual steps
        steps_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        steps_frame.pack(fill="x")
        
        Label(steps_frame, text="Individual Steps:", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=COLORS['bg_dark']).pack(side="left", padx=(0, 10))
        
        Button(steps_frame, text="1️⃣ EXTRACT", command=extract_3ds_archives,
               font=self.font_small,
               bg=COLORS['button_blue'], fg="white",
               activebackground=COLORS['text_secondary'],
               relief="flat", cursor="hand2", padx=10, pady=3).pack(side="left", padx=3)
        
        Button(steps_frame, text="2️⃣ DECRYPT", command=decrypt_3ds_roms,
               font=self.font_small,
               bg=COLORS['accent_purple'], fg="white",
               activebackground=COLORS['accent_pink'],
               relief="flat", cursor="hand2", padx=10, pady=3).pack(side="left", padx=3)
        
        Button(steps_frame, text="3️⃣ MOVE", command=move_3ds_roms,
               font=self.font_small,
               bg=COLORS['accent_orange'], fg="white",
               activebackground=COLORS['accent_yellow'],
               relief="flat", cursor="hand2", padx=10, pady=3).pack(side="left", padx=3)
        
        # Action buttons - All-in-one and close
        action_frame = Frame(dialog, padx=10, pady=10, bg=COLORS['bg_dark'])
        action_frame.pack(fill="x")
        
        Button(action_frame, text="✕ CLOSE", command=dialog.destroy,
               font=self.font_button,
               activebackground=COLORS['accent_red'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="right", padx=5)

    def download_bios_dialog(self):
        """Open dialog to download and extract BIOS files from this project's GitHub"""
        dialog = Toplevel(self.master)
        dialog.title("◄ BIOS DOWNLOADER ►")
        dialog.geometry("700x550")
        dialog.resizable(True, True)
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.configure(bg=COLORS['bg_dark'])

        # Title
        title_frame = Frame(dialog, bg=COLORS['bg_light'], pady=6)
        title_frame.pack(fill="x", padx=10, pady=(10, 10))
        Label(title_frame, text="🧬 BIOS DOWNLOADER", font=self.font_heading_md,
              fg=COLORS['accent_pink'], bg=COLORS['bg_light']).pack()
        Label(title_frame, text="Download official BIOS collection from WoofahRayetCode repository",
              font=self.font_small, fg=COLORS['text_muted'], bg=COLORS['bg_light']).pack()

        # Destination directory
        dest_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        dest_frame.pack(fill="x")

        Label(dest_frame, text="📁 Destination:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(side="left")
        dest_entry = Entry(dest_frame, font=self.font_body,
                          bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                          insertbackground=COLORS['text_primary'], relief="flat")
        dest_entry.pack(side="left", fill="x", expand=True, padx=5, ipady=3)
        
        # Try to find a good default: BIOS folder in app directory
        default_dest = str(self.script_dir / "BIOS")
            
        dest_entry.insert(0, default_dest)

        def browse_dest():
            folder = filedialog.askdirectory(title="Select BIOS Extraction Folder")
            if folder:
                dest_entry.delete(0, "end")
                dest_entry.insert(0, folder)

        Button(dest_frame, text="[ BROWSE ]", command=browse_dest,
               font=self.font_small, bg=COLORS['bg_light'],
               fg=COLORS['text_secondary'], relief="flat", cursor="hand2").pack(side="left")

        # Info text
        info_frame = Frame(dialog, padx=10, pady=10, bg=COLORS['bg_dark'])
        info_frame.pack(fill="x")
        
        info_label = Label(info_frame, text="This will download the BIOS collection from the WoofahRayetCode repository.\n"
                                           "Files will be organized into system-specific subfolders.",
                           font=self.font_small, fg=COLORS['accent_yellow'], bg=COLORS['bg_dark'], justify="left")
        info_label.pack(anchor="w")

        # Progress area
        results_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        results_frame.pack(fill="both", expand=True)

        Label(results_frame, text="◄ OPERATION LOG ►", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=COLORS['bg_dark']).pack(anchor="w", pady=(0, 4))

        list_frame = Frame(results_frame, bg=COLORS['bg_dark'])
        list_frame.pack(fill="both", expand=True)

        scrollbar = Scrollbar(list_frame, bg=COLORS['bg_light'],
                             troughcolor=COLORS['bg_dark'])
        scrollbar.pack(side="right", fill="y")

        results_text = Text(list_frame, wrap="word", yscrollcommand=scrollbar.set,
                           height=15, font=self.font_mono,
                           bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                           insertbackground=COLORS['text_primary'], relief="flat",
                           padx=8, pady=8)
        results_text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=results_text.yview)

        def log_msg(msg):
            results_text.insert("end", msg + "\n")
            results_text.see("end")
            dialog.update()

        def start_download():
            dest = dest_entry.get().strip()
            if not dest:
                messagebox.showwarning("Warning", "Please specify a destination folder")
                return
            
            # Disable buttons during download
            download_btn.config(state="disabled")
            
            threading.Thread(target=self.download_bios, args=(dest, log_msg, lambda: download_btn.config(state="normal")), daemon=True).start()

        # Action buttons
        action_frame = Frame(dialog, padx=10, pady=10, bg=COLORS['bg_dark'])
        action_frame.pack(fill="x")

        download_btn = Button(action_frame, text="📥 START DOWNLOAD", command=start_download,
               font=self.font_button,
               bg=COLORS['button_green'], fg=COLORS['bg_dark'],
               activebackground=COLORS['text_primary'],
               relief="flat", cursor="hand2", padx=20, pady=5)
        download_btn.pack(side="left", padx=5)

        Button(action_frame, text="✕ CLOSE", command=dialog.destroy,
               font=self.font_button,
               activebackground=COLORS['accent_red'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="right", padx=5)

    def download_bios(self, dest_dir, log_callback, on_complete):
        """Download and extract BIOS collection from project repository"""
        try:
            repo_url = "https://github.com/WoofahRayetCode/WlfRyt-Rom-Manager/archive/refs/heads/main.zip"
            dest_path = Path(dest_dir)
            dest_path.mkdir(parents=True, exist_ok=True)
            
            temp_zip = dest_path / "bios_temp.zip"
            
            log_callback("🚀 Starting BIOS download...")
            log_callback(f"🔗 Source: {repo_url}")
            log_callback(f"📂 Destination: {dest_dir}\n")
            
            # Download zip
            req = urllib.request.Request(repo_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=300) as response:
                total_size = int(response.headers.get('content-length', 0))
                downloaded = 0
                block_size = 1024 * 1024
                
                with open(temp_zip, 'wb') as f:
                    while True:
                        buffer = response.read(block_size)
                        if not buffer:
                            break
                        f.write(buffer)
                        downloaded += len(buffer)
                        if total_size > 0:
                            percent = (downloaded / total_size) * 100
                            log_callback(f"   📥 Downloading: {downloaded/(1024*1024):.1f}MB / {total_size/(1024*1024):.1f}MB ({percent:.1f}%)")
                        else:
                            log_callback(f"   📥 Downloading: {downloaded/(1024*1024):.1f}MB...")
            
            log_callback("\n✅ Download complete. Extracting files...")
            
            # Extract
            with zipfile.ZipFile(temp_zip, 'r') as zf:
                # The repo zip contains a top-level folder "WlfRyt-Rom-Manager-main"
                # and BIOS files are in "WlfRyt-Rom-Manager-main/BIOS Files/"
                prefix = "WlfRyt-Rom-Manager-main/BIOS Files/"
                
                # Filter members to only include those under the BIOS folder
                bios_members = [m for m in zf.namelist() if m.startswith(prefix)]
                
                if not bios_members:
                    # Fallback check for branch names or structure changes
                    log_callback("⚠️ Could not find 'BIOS Files' in ZIP. Checking structure...")
                    for name in zf.namelist():
                        if "BIOS Files" in name:
                            log_callback(f"Found: {name}")
                
                extracted_count = 0
                for member in bios_members:
                    # Get the path relative to the BIOS folder
                    rel_path = member[len(prefix):]
                    if not rel_path: continue
                    
                    target = dest_path / rel_path
                    
                    if member.endswith('/'):
                        target.mkdir(parents=True, exist_ok=True)
                    else:
                        target.parent.mkdir(parents=True, exist_ok=True)
                        with zf.open(member) as source, open(target, "wb") as target_file:
                            shutil.copyfileobj(source, target_file)
                        
                        # Handle nested ZIP files
                        if target.suffix.lower() == '.zip':
                            log_callback(f"   📦 Unzipping pack: {rel_path}...")
                            try:
                                with zipfile.ZipFile(target, 'r') as nested_zf:
                                    # Extract to the same directory as the zip
                                    nested_zf.extractall(target.parent)
                                # Remove the zip file after extraction
                                target.unlink()
                                log_callback(f"   ✅ Unzipped and cleaned: {rel_path}")
                            except Exception as zip_err:
                                log_callback(f"   ⚠️ Could not unzip {rel_path}: {zip_err}")
                        else:
                            log_callback(f"   📦 Extracted: {rel_path}")
                        
                        extracted_count += 1
            
            if extracted_count > 0:
                log_callback(f"\n✨ Successfully extracted {extracted_count} BIOS files!")
                log_callback(f"📍 Files are located in: {dest_dir}")
                messagebox.showinfo("Success", f"BIOS collection downloaded and extracted to:\n{dest_dir}")
            else:
                log_callback("\n❌ Error: No BIOS files found in the download.")
                messagebox.showerror("Error", "Failed to find BIOS files in the repository download.")
            
            # Cleanup
            try:
                temp_zip.unlink()
            except:
                pass
            
        except Exception as e:
            log_callback(f"\n❌ Error: {e}")
            messagebox.showerror("Error", f"Failed to download BIOS:\n{e}")
        finally:
            on_complete()

    def ia_login_dialog(self, on_success=None):
        """Open dialog to log in to Internet Archive"""
        if not REQUESTS_AVAILABLE:
            messagebox.showerror("Error", "Requests library not found. Login disabled.")
            return

        dialog = Toplevel(self.master)
        dialog.title("🔑 ARCHIVE.ORG LOGIN")
        dialog.geometry("400x300")
        dialog.resizable(False, False)
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.configure(bg=COLORS['bg_dark'])

        title_frame = Frame(dialog, bg=COLORS['bg_light'], pady=6)
        title_frame.pack(fill="x", padx=10, pady=(10, 10))
        Label(title_frame, text="🔑 ARCHIVE.ORG LOGIN", font=self.font_heading_md,
              fg=COLORS['accent_yellow'], bg=COLORS['bg_light']).pack()
        Label(title_frame, text="Required for restricted/locked collections",
              font=self.font_small, fg=COLORS['text_muted'], bg=COLORS['bg_light']).pack()

        fields_frame = Frame(dialog, padx=20, pady=10, bg=COLORS['bg_dark'])
        fields_frame.pack(fill="both", expand=True)

        Label(fields_frame, text="Email:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(anchor="w")
        email_entry = Entry(fields_frame, font=self.font_body,
                           bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                           insertbackground=COLORS['text_primary'], relief="flat")
        email_entry.pack(fill="x", pady=(2, 10), ipady=3)
        email_entry.focus_set()

        Label(fields_frame, text="Password:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(anchor="w")
        pass_entry = Entry(fields_frame, font=self.font_body, show="*",
                          bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                          insertbackground=COLORS['text_primary'], relief="flat")
        pass_entry.pack(fill="x", pady=(2, 10), ipady=3)

        status_label = Label(fields_frame, text="", font=self.font_small,
                            bg=COLORS['bg_dark'], fg=COLORS['accent_orange'])
        status_label.pack(pady=5)

        def do_login():
            email = email_entry.get().strip()
            password = pass_entry.get().strip()
            if not email or not password:
                status_label.config(text="⚠️ Please enter both email and password")
                return

            status_label.config(text="⌛ Logging in...", fg=COLORS['text_secondary'])
            dialog.update()

            success, msg = self.ia_login(email, password)
            if success:
                status_label.config(text=f"✅ {msg}", fg=COLORS['button_green'])
                dialog.after(1000, dialog.destroy)
                if on_success:
                    on_success()
            else:
                status_label.config(text=f"❌ {msg}", fg=COLORS['accent_red'])

        btn_frame = Frame(dialog, padx=10, pady=10, bg=COLORS['bg_dark'])
        btn_frame.pack(fill="x")
        Button(btn_frame, text="🔑 LOGIN", command=do_login,
               font=self.font_button, bg=COLORS['button_blue'], fg="white",
               relief="flat", cursor="hand2", padx=20, pady=5).pack(side="left", padx=5)
        Button(btn_frame, text="✕ CANCEL", command=dialog.destroy,
               font=self.font_button, bg=COLORS['bg_light'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="right", padx=5)

    def ia_s3_login_dialog(self, on_success=None):
        """Open dialog to authenticate using Archive.org S3 keys"""
        dialog = Toplevel(self.master)
        dialog.title("🔑 IA S3 KEYS AUTHENTICATION")
        dialog.geometry("450x350")
        dialog.resizable(False, False)
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.configure(bg=COLORS['bg_dark'])

        title_frame = Frame(dialog, bg=COLORS['bg_light'], pady=6)
        title_frame.pack(fill="x", padx=10, pady=(10, 10))
        Label(title_frame, text="🔑 IA S3 KEYS", font=self.font_heading_md,
              fg=COLORS['accent_yellow'], bg=COLORS['bg_light']).pack()
        Label(title_frame, text="Get keys from archive.org/account/s3.php",
              font=self.font_small, fg=COLORS['text_muted'], bg=COLORS['bg_light']).pack()

        fields_frame = Frame(dialog, padx=20, pady=10, bg=COLORS['bg_dark'])
        fields_frame.pack(fill="both", expand=True)

        Label(fields_frame, text="Access Key:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(anchor="w")
        access_entry = Entry(fields_frame, font=self.font_body,
                            bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                            insertbackground=COLORS['text_primary'], relief="flat")
        access_entry.pack(fill="x", pady=(2, 10), ipady=3)

        Label(fields_frame, text="Secret Key:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(anchor="w")
        secret_entry = Entry(fields_frame, font=self.font_body, show="*",
                            bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                            insertbackground=COLORS['text_primary'], relief="flat")
        secret_entry.pack(fill="x", pady=(2, 10), ipady=3)

        def do_auth():
            access = access_entry.get().strip()
            secret = secret_entry.get().strip()
            if not access or not secret:
                messagebox.showwarning("Warning", "Please enter both keys")
                return

            headers = {"Authorization": f"LOW {access}:{secret}"}
            try:
                res = requests.get("https://archive.org/metadata/@me", headers=headers, timeout=15)
                if res.status_code == 200:
                    self.ia_s3_keys = (access, secret)
                    self.ia_username = res.json().get('username', 'S3 User')
                    self.ia_session = None
                    messagebox.showinfo("Success", f"Authenticated as {self.ia_username}")
                    dialog.destroy()
                    if on_success: on_success()
                else:
                    messagebox.showerror("Error", "Invalid S3 keys")
            except Exception as e:
                messagebox.showerror("Error", f"Connection error: {e}")

        btn_frame = Frame(dialog, padx=10, pady=10, bg=COLORS['bg_dark'])
        btn_frame.pack(fill="x")
        Button(btn_frame, text="✅ AUTHENTICATE", command=do_auth,
               font=self.font_button, bg=COLORS['button_blue'], fg="white",
               relief="flat", cursor="hand2", padx=20, pady=5).pack(side="left", padx=5)
        Button(btn_frame, text="✕ CANCEL", command=dialog.destroy,
               font=self.font_button, bg=COLORS['bg_light'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="right", padx=5)

    def ia_login(self, email, password):
        """Perform login to Archive.org and store session"""
        try:
            session = requests.Session()
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
            }
            login_url = "https://archive.org/account/login"
            session.get(login_url, headers=headers, timeout=30)
            data = {
                "username": email,
                "password": password,
                "remember": "checked",
                "submit_by_js": "true",
                "submit-to-login": "Log in"
            }
            headers['Referer'] = login_url
            response = session.post(login_url, data=data, headers=headers, timeout=30)
            if 'logged-in-sig' in session.cookies or 'logged-in-user' in session.cookies:
                self.ia_session = session
                self.ia_username = email
                return True, "Login successful!"
            
            xauth_url = "https://archive.org/services/xauthn/"
            xauth_data = {"username": email, "password": password}
            response = session.post(xauth_url, data=xauth_data, headers=headers, timeout=30)
            if 'logged-in-sig' in session.cookies or 'logged-in-user' in session.cookies:
                self.ia_session = session
                self.ia_username = email
                return True, "Login successful via service!"
            return False, "Invalid credentials. If you have 2FA enabled, please use S3 Keys instead."
        except Exception as e:
            return False, f"Connection error: {str(e)}"

    def run_ia_downloads(self, games, dest_dir, status_label, progress_bar, on_complete):
        """Handle IA downloads in sequence"""
        total = len(games)
        for i, game in enumerate(games, 1):
            name = game['name']
            identifier = game['identifier']
            encoded_name = urllib.parse.quote(name)
            url = f"https://archive.org/download/{identifier}/{encoded_name}"
            dest_path = Path(dest_dir) / name
            def update_status(msg):
                status_label.config(text=f"[{i}/{total}] {msg}")
                self.master.update_idletasks()
            update_status(f"Starting: {name}")
            try:
                if self.ia_session:
                    with self.ia_session.get(url, stream=True, timeout=300) as response:
                        response.raise_for_status()
                        total_size = int(response.headers.get('content-length', 0))
                        downloaded = 0
                        with open(dest_path, 'wb') as f:
                            for chunk in response.iter_content(chunk_size=1024*1024):
                                if chunk:
                                    f.write(chunk)
                                    downloaded += len(chunk)
                                    if total_size > 0:
                                        percent = (downloaded / total_size) * 100
                                        progress_bar['value'] = percent
                                        update_status(f"Downloading: {downloaded/(1024*1024):.1f}MB / {total_size/(1024*1024):.1f}MB ({percent:.1f}%)")
                                    else:
                                        update_status(f"Downloading: {downloaded/(1024*1024):.1f}MB...")
                elif self.ia_s3_keys:
                    access, secret = self.ia_s3_keys
                    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0', 'Authorization': f'LOW {access}:{secret}'})
                    with urllib.request.urlopen(req, timeout=300) as response:
                        total_size = int(response.headers.get('content-length', 0))
                        downloaded = 0
                        with open(dest_path, 'wb') as f:
                            while True:
                                chunk = response.read(1024*1024)
                                if not chunk: break
                                f.write(chunk)
                                downloaded += len(chunk)
                                if total_size > 0:
                                    percent = (downloaded / total_size) * 100
                                    progress_bar['value'] = percent
                                    update_status(f"Downloading: {downloaded/(1024*1024):.1f}MB / {total_size/(1024*1024):.1f}MB ({percent:.1f}%)")
                                else:
                                    update_status(f"Downloading: {downloaded/(1024*1024):.1f}MB...")
                else:
                    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(req, timeout=300) as response:
                        total_size = int(response.headers.get('content-length', 0))
                        downloaded = 0
                        with open(dest_path, 'wb') as f:
                            while True:
                                chunk = response.read(1024*1024)
                                if not chunk: break
                                f.write(chunk)
                                downloaded += len(chunk)
                                if total_size > 0:
                                    percent = (downloaded / total_size) * 100
                                    progress_bar['value'] = percent
                                    update_status(f"Downloading: {downloaded/(1024*1024):.1f}MB / {total_size/(1024*1024):.1f}MB ({percent:.1f}%)")
                                else:
                                    update_status(f"Downloading: {downloaded/(1024*1024):.1f}MB...")
                update_status(f"✅ Finished: {name}")
                time.sleep(0.5)
            except Exception as e:
                update_status(f"❌ Error: {e}")
                if messagebox.askretrycancel("Download Error", f"Failed to download {name}:\n{e}\n\nRetry?"):
                    self.run_ia_downloads(games[i-1:], dest_dir, status_label, progress_bar, on_complete)
                    return
        status_label.config(text=f"🎉 Completed {total} download(s)!")
        progress_bar['value'] = 0
        messagebox.showinfo("Success", f"Successfully downloaded {total} game(s)")
        on_complete()

    def ia_login_dialog(self, on_success=None):
        """Open dialog to log in to Internet Archive"""
        if not REQUESTS_AVAILABLE:
            messagebox.showerror("Error", "Requests library not found. Login disabled.")
            return

        dialog = Toplevel(self.master)
        dialog.title("🔑 ARCHIVE.ORG LOGIN")
        dialog.geometry("400x300")
        dialog.resizable(False, False)
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.configure(bg=COLORS['bg_dark'])

        # Title
        title_frame = Frame(dialog, bg=COLORS['bg_light'], pady=6)
        title_frame.pack(fill="x", padx=10, pady=(10, 10))
        Label(title_frame, text="🔑 ARCHIVE.ORG LOGIN", font=self.font_heading_md,
              fg=COLORS['accent_yellow'], bg=COLORS['bg_light']).pack()
        Label(title_frame, text="Required for restricted/locked collections",
              font=self.font_small, fg=COLORS['text_muted'], bg=COLORS['bg_light']).pack()

        # Input fields
        fields_frame = Frame(dialog, padx=20, pady=10, bg=COLORS['bg_dark'])
        fields_frame.pack(fill="both", expand=True)

        Label(fields_frame, text="Email:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(anchor="w")
        email_entry = Entry(fields_frame, font=self.font_body,
                           bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                           insertbackground=COLORS['text_primary'], relief="flat")
        email_entry.pack(fill="x", pady=(2, 10), ipady=3)
        email_entry.focus_set()

        Label(fields_frame, text="Password:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(anchor="w")
        pass_entry = Entry(fields_frame, font=self.font_body, show="*",
                          bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                          insertbackground=COLORS['text_primary'], relief="flat")
        pass_entry.pack(fill="x", pady=(2, 10), ipady=3)

        status_label = Label(fields_frame, text="", font=self.font_small,
                            bg=COLORS['bg_dark'], fg=COLORS['accent_orange'])
        status_label.pack(pady=5)

        def do_login():
            email = email_entry.get().strip()
            password = pass_entry.get().strip()
            if not email or not password:
                status_label.config(text="⚠️ Please enter both email and password")
                return

            status_label.config(text="⌛ Logging in...", fg=COLORS['text_secondary'])
            dialog.update()

            success, msg = self.ia_login(email, password)
            if success:
                status_label.config(text=f"✅ {msg}", fg=COLORS['button_green'])
                dialog.after(1000, dialog.destroy)
                if on_success:
                    on_success()
            else:
                status_label.config(text=f"❌ {msg}", fg=COLORS['accent_red'])

        # Action buttons
        btn_frame = Frame(dialog, padx=10, pady=10, bg=COLORS['bg_dark'])
        btn_frame.pack(fill="x")

        Button(btn_frame, text="🔑 LOGIN", command=do_login,
               font=self.font_button, bg=COLORS['button_blue'], fg="white",
               relief="flat", cursor="hand2", padx=20, pady=5).pack(side="left", padx=5)

        Button(btn_frame, text="✕ CANCEL", command=dialog.destroy,
               font=self.font_button, bg=COLORS['bg_light'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="right", padx=5)

    def ia_browser_dialog(self):
        """Open generic Archive.org browser for multiple systems and collections"""
        dialog = Toplevel(self.master)
        dialog.title("◄ INTERNET ARCHIVE ROM BROWSER ►")
        dialog.geometry("1100x850")
        dialog.resizable(True, True)
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.configure(bg=COLORS['bg_dark'])

        # Data Mapping
        def get_collections():
            fmt = self.ia_download_format.get()
            
            # Base collections available in both or specific formats
            collections = {
                "Sony": {
                    "PlayStation Portable (Redump USA)": {
                        "pattern": "sony_psp_usa_202205",
                        "parts": ["Full Collection"]
                    }
                },
                "Nintendo": {
                    "Nintendo 64 (BigEndian Z64)": {
                        "pattern": "Nintendo_64_USA_2020_08_02",
                        "parts": ["Full Collection"]
                    },
                    "NES (No-Intro 2021)": {
                        "pattern": "nintendo-entertainment-system-no-intro-2021-08-04",
                        "parts": ["Full Collection"]
                    },
                    "SNES (No-Intro 2021)": {
                        "pattern": "super-nintendo-entertainment-system-no-intro-2021-08-04",
                        "parts": ["Full Collection"]
                    }
                },
                "Microsoft": {
                    "Xbox (Redump USA)": {
                        "pattern": "xbox_redump_usa",
                        "parts": ["Full Collection"]
                    }
                }
            }

            # Format-specific additions
            if fmt == "CHD":
                collections["Sony"]["PlayStation 1 (Redump CHD)"] = {
                    "pattern": "psx-chd-roms-{part}",
                    "parts": list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
                }
                collections["Sony"]["PlayStation 2 (Redump USA CHD)"] = {
                    "pattern": "ps2-redump-usa-chd-part-{part}",
                    "parts": ["0-9", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"],
                    "overrides": {"G": "g_202207", "0-9": "0"}
                }
                collections["Nintendo"]["GameCube (Redump USA CHD)"] = {
                    "pattern": "nintendo-gamecube_202212",
                    "parts": ["Full Collection"]
                }
                collections["Nintendo"]["Wii (Redump USA RVZ)"] = {
                    "pattern": "wii-redump-usa-rvz-part{part}",
                    "parts": ["1", "2", "3", "4", "5", "6", "7"]
                }
            else: # ISO/BIN
                collections["Sony"]["PlayStation 1 (Redump USA BIN/CUE)"] = {
                    "pattern": "Redump_PSX_2021_06_04_{part}",
                    "parts": list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
                }
                collections["Sony"]["PlayStation 2 (Redump USA ISO)"] = {
                    "pattern": "ps2usaredump{part}",
                    "parts": ["1", "2", "3"]
                }
                collections["Nintendo"]["GameCube (Redump USA ISO)"] = {
                    "pattern": "GameCube_USA_Redump_2021-08-11",
                    "parts": ["Full Collection"]
                }
                collections["Nintendo"]["Wii (Redump USA ISO)"] = {
                    "pattern": "Wii_USA_Redump_2022-07-28",
                    "parts": ["Full Collection"]
                }
            
            return collections

        # Title
        title_frame = Frame(dialog, bg=COLORS['bg_light'], pady=6)
        title_frame.pack(fill="x", padx=10, pady=(10, 10))
        Label(title_frame, text="🌐 INTERNET ARCHIVE BROWSER", font=self.font_heading_md,
              fg=COLORS['button_blue'], bg=COLORS['bg_light']).pack()

        # Login Status Frame
        login_info_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        login_info_frame.pack(fill="x")
        
        status_color = COLORS['button_green'] if (self.ia_session or self.ia_s3_keys) else COLORS['accent_red']
        auth_type = "S3 Keys" if self.ia_s3_keys else "Form Session" if self.ia_session else "None"
        status_text = f"Authenticated via {auth_type}: {self.ia_username}" if (self.ia_session or self.ia_s3_keys) else "⚠️ NOT LOGGED IN - Restricted collections will fail"
        
        login_status_label = Label(login_info_frame, text=status_text, font=self.font_label_bold,
                                 fg=status_color, bg=COLORS['bg_dark'])
        login_status_label.pack(side="left")
        
        def update_login_ui():
            s_color = COLORS['button_green'] if (self.ia_session or self.ia_s3_keys) else COLORS['accent_red']
            a_type = "S3 Keys" if self.ia_s3_keys else "Form Session" if self.ia_session else "None"
            s_text = f"Authenticated via {a_type}: {self.ia_username}" if (self.ia_session or self.ia_s3_keys) else "⚠️ NOT LOGGED IN - Restricted collections will fail"
            login_status_label.config(text=s_text, fg=s_color)

        Button(login_info_frame, text="🔑 S3 KEYS (DEFAULT)", 
              command=lambda: self.ia_s3_login_dialog(update_login_ui),
              font=self.font_small, bg=COLORS['accent_yellow'],
              fg=COLORS['bg_dark'], relief="flat", cursor="hand2", padx=10).pack(side="right", padx=5)
        
        Button(login_info_frame, text="🔑 FORM LOGIN", 
              command=lambda: self.ia_login_dialog(update_login_ui),
              font=self.font_small, bg=COLORS['bg_medium'],
              fg=COLORS['text_secondary'], relief="flat", cursor="hand2", padx=10).pack(side="right")

        # Selection Frame
        selection_frame = Frame(dialog, padx=10, pady=10, bg=COLORS['bg_light'])
        selection_frame.pack(fill="x", padx=10, pady=5)

        # Format Toggle (CHD vs ISO)
        format_frame = Frame(selection_frame, bg=COLORS['bg_light'])
        format_frame.pack(side="left", padx=(0, 20))
        Label(format_frame, text="Format:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_light']).pack(side="left")
        
        self.ia_download_format = StringVar(value="CHD")
        
        def on_format_change():
            # Refresh system collections based on format
            on_system_change()

        Radiobutton(format_frame, text="CHD", variable=self.ia_download_format, value="CHD",
                   command=on_format_change, bg=COLORS['bg_light'], fg=COLORS['text_primary'],
                   activebackground=COLORS['bg_light'], selectcolor=COLORS['bg_light'],
                   font=self.font_small).pack(side="left", padx=5)
        Radiobutton(format_frame, text="ISO/BIN", variable=self.ia_download_format, value="ISO",
                   command=on_format_change, bg=COLORS['bg_light'], fg=COLORS['text_primary'],
                   activebackground=COLORS['bg_light'], selectcolor=COLORS['bg_light'],
                   font=self.font_small).pack(side="left", padx=5)

        # System Select
        Label(selection_frame, text="System:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_light']).pack(side="left")
        system_combo = ttk.Combobox(selection_frame, values=list(get_collections().keys()), state="readonly", width=12)
        system_combo.set("Sony")
        system_combo.pack(side="left", padx=5)

        # Console Select
        Label(selection_frame, text="Collection:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_light']).pack(side="left", padx=(10, 0))
        console_combo = ttk.Combobox(selection_frame, state="readonly", width=35)
        console_combo.pack(side="left", padx=5)

        # Part Select
        Label(selection_frame, text="Part:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_light']).pack(side="left", padx=(10, 0))
        part_combo = ttk.Combobox(selection_frame, state="readonly", width=10)
        part_combo.pack(side="left", padx=5)

        # Search
        Label(selection_frame, text="🔍 Search:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_light']).pack(side="left", padx=(10, 0))
        search_entry = Entry(selection_frame, font=self.font_body,
                            bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                            insertbackground=COLORS['text_primary'], relief="flat", width=20)
        search_entry.pack(side="left", padx=5, ipady=2)

        # Folder frame
        folder_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        folder_frame.pack(fill="x")
        Label(folder_frame, text="📁 Save to:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(side="left")
        dest_entry = Entry(folder_frame, font=self.font_body,
                          bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                          insertbackground=COLORS['text_primary'], relief="flat")
        dest_entry.pack(side="left", fill="x", expand=True, padx=5, ipady=3)
        default_dest = self.source_dir or str(self.script_dir / "Downloads")
        dest_entry.insert(0, default_dest)

        def browse_dest():
            folder = filedialog.askdirectory(title="Select Download Folder")
            if folder:
                dest_entry.delete(0, "end")
                dest_entry.insert(0, folder)
        Button(folder_frame, text="[ BROWSE ]", command=browse_dest,
               font=self.font_small, bg=COLORS['bg_light'],
               fg=COLORS['text_secondary'], relief="flat", cursor="hand2").pack(side="left")

        # Game List
        list_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        list_frame.pack(fill="both", expand=True)
        columns = ("name", "size")
        tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="extended")
        tree.heading("name", text="Filename")
        tree.heading("size", text="Size")
        tree.column("name", width=700)
        tree.column("size", width=120, anchor="e")
        tree_scroll = Scrollbar(list_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=tree_scroll.set)
        tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")

        # Status
        progress_bar = ttk.Progressbar(dialog, mode='determinate', length=1000)
        progress_bar.pack(fill="x", padx=10, pady=2)
        status_var = Label(dialog, text="Ready", font=self.font_small, bg=COLORS['bg_dark'], fg=COLORS['text_muted'], anchor="w")
        status_var.pack(fill="x", padx=10)

        all_items = []

        def on_system_change(event=None):
            sys_name = system_combo.get()
            collections = get_collections()
            consoles = list(collections[sys_name].keys())
            console_combo.config(values=consoles)
            console_combo.set(consoles[0])
            on_console_change()

        def on_console_change(event=None):
            sys_name = system_combo.get()
            con_name = console_combo.get()
            collections = get_collections()
            parts = collections[sys_name][con_name]["parts"]
            part_combo.config(values=parts)
            part_combo.set(parts[0])
            load_list()

        def load_list(event=None):
            sys_name = system_combo.get()
            con_name = console_combo.get()
            part = part_combo.get()
            
            collections = get_collections()
            coll_data = collections[sys_name][con_name]
            pattern = coll_data["pattern"]
            
            # Apply overrides
            part_val = part
            if "overrides" in coll_data and part in coll_data["overrides"]:
                part_val = coll_data["overrides"][part]
            
            if "{part}" in pattern:
                identifier = pattern.replace("{part}", part_val)
            else:
                identifier = pattern
            
            url = f"https://archive.org/metadata/{identifier}"
            status_var.config(text=f"Fetching metadata for {identifier}...")
            dialog.update()

            try:
                if self.ia_session:
                    response = self.ia_session.get(url, timeout=30)
                    response.raise_for_status()
                    data = response.json()
                elif self.ia_s3_keys:
                    access, secret = self.ia_s3_keys
                    req = urllib.request.Request(url, headers={
                        'User-Agent': 'Mozilla/5.0',
                        'Authorization': f'LOW {access}:{secret}'
                    })
                    with urllib.request.urlopen(req, timeout=30) as response:
                        data = json.loads(response.read().decode())
                else:
                    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(req, timeout=30) as response:
                        data = json.loads(response.read().decode())
                
                all_items.clear()
                for f in data.get('files', []):
                    # Filter out metadata files
                    if f['name'].lower().endswith(('.chd', '.iso', '.rvz', '.z64', '.nes', '.sfc', '.smc', '.snes', '.n64', '.v64')):
                        size_bytes = int(f.get('size', 0))
                        size_str = f"{size_bytes / (1024*1024):.1f} MB" if size_bytes < 1024*1024*1024 else f"{size_bytes / (1024*1024*1024):.2f} GB"
                        all_items.append({
                            "name": f['name'],
                            "size": size_str,
                            "identifier": identifier
                        })
                
                update_view()
                status_var.config(text=f"Loaded {len(all_items)} items from {identifier}")
            except Exception as e:
                status_var.config(text=f"Error: {e}")

        def update_view(event=None):
            query = search_entry.get().lower()
            for item in tree.get_children(): tree.delete(item)
            filtered = [i for i in all_items if query in i['name'].lower()]
            for i in filtered:
                tree.insert("", "end", values=(i['name'], i['size']), tags=(i['identifier'],))

        system_combo.bind("<<ComboboxSelected>>", on_system_change)
        console_combo.bind("<<ComboboxSelected>>", on_console_change)
        part_combo.bind("<<ComboboxSelected>>", load_list)
        search_entry.bind("<KeyRelease>", update_view)

        def download_selected():
            selected = tree.selection()
            if not selected: return
            dest = dest_entry.get().strip()
            if not dest: return
            Path(dest).mkdir(parents=True, exist_ok=True)
            
            games = []
            for item in selected:
                values = tree.item(item, "values")
                identifier = tree.item(item, "tags")[0]
                games.append({"name": values[0], "identifier": identifier})
            
            download_btn.config(state="disabled")
            threading.Thread(target=self.run_ia_downloads, args=(games, dest, status_var, progress_bar, lambda: download_btn.config(state="normal")), daemon=True).start()

        # Action Buttons
        btn_frame = Frame(dialog, padx=10, pady=10, bg=COLORS['bg_dark'])
        btn_frame.pack(fill="x")
        
        load_btn = Button(btn_frame, text="🔄 REFRESH", command=load_list,
                        font=self.font_button, bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                        relief="flat", cursor="hand2", padx=15, pady=5)
        load_btn.pack(side="left", padx=5)

        download_btn = Button(btn_frame, text="📥 DOWNLOAD SELECTED", command=download_selected,
                            font=self.font_button, bg=COLORS['button_green'], fg=COLORS['bg_dark'],
                            activebackground=COLORS['text_primary'],
                            relief="flat", cursor="hand2", padx=20, pady=5)
        download_btn.pack(side="left", padx=5)

        Button(btn_frame, text="✕ CLOSE", command=dialog.destroy,
               font=self.font_button, bg=COLORS['accent_red'], fg="white",
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="right", padx=5)

        # Init
        on_system_change()

    def run_ia_downloads(self, games, dest_dir, status_label, progress_bar, on_complete):
        """Handle IA downloads in sequence"""
        total = len(games)
        for i, game in enumerate(games, 1):
            name = game['name']
            identifier = game['identifier']
            encoded_name = urllib.parse.quote(name)
            url = f"https://archive.org/download/{identifier}/{encoded_name}"
            dest_path = Path(dest_dir) / name
            
            def update_status(msg):
                status_label.config(text=f"[{i}/{total}] {msg}")
                self.master.update_idletasks()

            update_status(f"Starting: {name}")
            
            try:
                if self.ia_session:
                    with self.ia_session.get(url, stream=True, timeout=300) as response:
                        response.raise_for_status()
                        total_size = int(response.headers.get('content-length', 0))
                        downloaded = 0
                        with open(dest_path, 'wb') as f:
                            for chunk in response.iter_content(chunk_size=1024*1024):
                                if chunk:
                                    f.write(chunk)
                                    downloaded += len(chunk)
                                    if total_size > 0:
                                        percent = (downloaded / total_size) * 100
                                        progress_bar['value'] = percent
                                        update_status(f"Downloading: {downloaded/(1024*1024):.1f}MB / {total_size/(1024*1024):.1f}MB ({percent:.1f}%)")
                                    else:
                                        update_status(f"Downloading: {downloaded/(1024*1024):.1f}MB...")
                elif self.ia_s3_keys:
                    access, secret = self.ia_s3_keys
                    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0', 'Authorization': f'LOW {access}:{secret}'})
                    with urllib.request.urlopen(req, timeout=300) as response:
                        total_size = int(response.headers.get('content-length', 0))
                        downloaded = 0
                        with open(dest_path, 'wb') as f:
                            while True:
                                chunk = response.read(1024*1024)
                                if not chunk: break
                                f.write(chunk)
                                downloaded += len(chunk)
                                if total_size > 0:
                                    percent = (downloaded / total_size) * 100
                                    progress_bar['value'] = percent
                                    update_status(f"Downloading: {downloaded/(1024*1024):.1f}MB / {total_size/(1024*1024):.1f}MB ({percent:.1f}%)")
                                else:
                                    update_status(f"Downloading: {downloaded/(1024*1024):.1f}MB...")
                else:
                    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(req, timeout=300) as response:
                        total_size = int(response.headers.get('content-length', 0))
                        downloaded = 0
                        with open(dest_path, 'wb') as f:
                            while True:
                                chunk = response.read(1024*1024)
                                if not chunk: break
                                f.write(chunk)
                                downloaded += len(chunk)
                                if total_size > 0:
                                    percent = (downloaded / total_size) * 100
                                    progress_bar['value'] = percent
                                    update_status(f"Downloading: {downloaded/(1024*1024):.1f}MB / {total_size/(1024*1024):.1f}MB ({percent:.1f}%)")
                                else:
                                    update_status(f"Downloading: {downloaded/(1024*1024):.1f}MB...")
                
                update_status(f"✅ Finished: {name}")
                time.sleep(0.5)
            except Exception as e:
                update_status(f"❌ Error: {e}")
                if messagebox.askretrycancel("Download Error", f"Failed to download {name}:\n{e}\n\nRetry?"):
                    self.run_ia_downloads(games[i-1:], dest_dir, status_label, progress_bar, on_complete)
                    return
        
        status_label.config(text=f"🎉 Completed {total} download(s)!")
        progress_bar['value'] = 0
        messagebox.showinfo("Success", f"Successfully downloaded {total} game(s)")
        on_complete()

    def save_config(self):
        """Save configuration to JSON file or ConfigAdapter"""
        # Phase 4 Week 2: Sync to ConfigAdapter before saving
        if self.config_adapter:
            try:
                self.config_adapter.sync_from_rom_converter()
                if self.config_adapter.save():
                    self.log("✅ Configuration saved via ConfigAdapter")
                    return
            except Exception as e:
                self.log(f"⚠️  ConfigAdapter save failed: {e}, falling back to direct save")
        
        # Fallback: Traditional JSON saving
        try:
            config = {
                'source_dir': self._make_portable_path(self.source_dir),
                'delete_originals': self.delete_originals.get(),
                'move_to_backup': self.move_to_backup.get(),
                'recursive': self.recursive.get(),
                'process_ps1_cues': self.process_ps1_cues.get(),
                'process_ps2_cues': self.process_ps2_cues.get(),
                'process_ps2_isos': self.process_ps2_isos.get(),
                'process_psp_isos': self.process_psp_isos.get(),
                'process_ps3_isos': self.process_ps3_isos.get(),
                'process_xbox_isos': self.process_xbox_isos.get(),
                'process_nes_roms': self.process_nes_roms.get(),
                'process_snes_roms': self.process_snes_roms.get(),
                'process_n64_roms': self.process_n64_roms.get(),
                'extract_compressed': self.extract_compressed.get(),
                'delete_archives_after_extract': self.delete_archives_after_extract.get(),
                'chdman_path': self._make_portable_path(self.chdman_path),
                'seven_zip_path': self._make_portable_path(self.seven_zip_path),
                'maxcso_path': self._make_portable_path(self.maxcso_path),
                'extract_xiso_path': self._make_portable_path(self.extract_xiso_path),
                'ndecrypt_path': self._make_portable_path(self.ndecrypt_path),
                'ps3_dumper_path': self._make_portable_path(self.ps3_dumper_path),
                'auto_decrypt_ps3_path': self._make_portable_path(self.auto_decrypt_ps3_path),
                'ps1_output_format': self.ps1_output_format,
                'ps2_output_format': self.ps2_output_format,
                'psp_output_format': self.psp_output_format,
                'ps2_emulator': self.ps2_emulator,
                'max_concurrent_conversions': self.max_concurrent_conversions,
                'disk_write_throttle_mb_s': self.disk_write_throttle_mb_s,
                'theme': self.current_theme,
                'system_extract_dirs': {k: self._make_portable_path(v) for k, v in self.system_extract_dirs.items()},
                # 3DS workflow settings
                'threeds_backup_original': self.threeds_backup_original,
                'threeds_delete_archives': self.threeds_delete_archives,
                'threeds_delete_after_move': self.threeds_delete_after_move,
                'threeds_auto_clean_names': self.threeds_auto_clean_names,
                'threeds_source_dir': self._make_portable_path(self.threeds_source_dir),
                'threeds_dest_dir': self._make_portable_path(self.threeds_dest_dir),
                'window_width': self.master.winfo_width() if self.master.winfo_width() > 100 else 1100,
                'window_height': self.master.winfo_height() if self.master.winfo_height() > 100 else 800,
            }
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=2)
        except Exception as e:
            # Silently fail - don't interrupt user experience
            pass
    
    def load_config(self):
        """Load configuration from JSON file or ConfigAdapter"""
        # Phase 4 Week 2: Try loading via ConfigAdapter first if available
        if self.config_adapter:
            try:
                if self.config_adapter.load():
                    self.config_adapter.sync_to_rom_converter()
                    self.log("✅ Configuration loaded via ConfigAdapter")
                    return
            except Exception as e:
                self.log(f"⚠️  ConfigAdapter load failed: {e}, falling back to direct load")
        
        # Fallback: Traditional JSON loading
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                
                # Restore settings
                self.source_dir = self._resolve_portable_path(config.get('source_dir', ''))
                self.delete_originals.set(config.get('delete_originals', False))
                self.move_to_backup.set(config.get('move_to_backup', True))
                self.recursive.set(config.get('recursive', True))
                self.process_ps1_cues.set(config.get('process_ps1_cues', False))
                self.process_ps2_cues.set(config.get('process_ps2_cues', False))
                self.process_ps2_isos.set(config.get('process_ps2_isos', False))
                self.process_psp_isos.set(config.get('process_psp_isos', False))
                self.process_ps3_isos.set(config.get('process_ps3_isos', False))
                self.process_xbox_isos.set(config.get('process_xbox_isos', False))
                self.process_nes_roms.set(config.get('process_nes_roms', False))
                self.process_snes_roms.set(config.get('process_snes_roms', False))
                self.process_n64_roms.set(config.get('process_n64_roms', False))
                self.extract_compressed.set(config.get('extract_compressed', True))
                self.delete_archives_after_extract.set(config.get('delete_archives_after_extract', False))
                self.ps2_output_format = config.get('ps2_output_format', 'CHD')
                self.psp_output_format = config.get('psp_output_format', 'CSO')
                self.ps2_emulator = config.get('ps2_emulator', 'PCSX2')
                self.max_concurrent_conversions = config.get('max_concurrent_conversions', self._detect_optimal_workers())
                self.disk_write_throttle_mb_s = config.get('disk_write_throttle_mb_s', 100)
                self.current_theme = config.get('theme', 'PS2')
                
                # Restore chdman path if saved and still exists
                saved_chdman = self._resolve_portable_path(config.get('chdman_path'))
                if saved_chdman and os.path.exists(saved_chdman):
                    self.chdman_path = saved_chdman
                
                # Restore 7-Zip path if saved and still exists
                saved_7zip = self._resolve_portable_path(config.get('seven_zip_path'))
                if saved_7zip and os.path.exists(saved_7zip):
                    self.seven_zip_path = saved_7zip

                # Restore maxcso path if saved and still exists
                saved_maxcso = self._resolve_portable_path(config.get('maxcso_path'))
                if saved_maxcso and os.path.exists(saved_maxcso):
                    self.maxcso_path = saved_maxcso

                # Restore extract-xiso path if saved and still exists
                saved_xiso = self._resolve_portable_path(config.get('extract_xiso_path'))
                if saved_xiso and os.path.exists(saved_xiso):
                    self.extract_xiso_path = saved_xiso
                
                # Restore ndecrypt path if saved and still exists
                saved_ndecrypt = self._resolve_portable_path(config.get('ndecrypt_path'))
                if saved_ndecrypt and os.path.exists(saved_ndecrypt):
                    self.ndecrypt_path = saved_ndecrypt

                # Restore ps3-disc-dumper path if saved and still exists
                saved_ps3_dumper = self._resolve_portable_path(config.get('ps3_dumper_path'))
                if saved_ps3_dumper and os.path.exists(saved_ps3_dumper):
                    self.ps3_dumper_path = saved_ps3_dumper

                # Restore AutoDecryptPS3 path if saved and still exists
                saved_auto_decrypt = self._resolve_portable_path(config.get('auto_decrypt_ps3_path'))
                if saved_auto_decrypt and os.path.exists(saved_auto_decrypt):
                    self.auto_decrypt_ps3_path = saved_auto_decrypt
                
                # Restore system extraction directories
                saved_system_dirs = config.get('system_extract_dirs', {})
                for system, path in saved_system_dirs.items():
                    resolved_path = self._resolve_portable_path(path)
                    if resolved_path and os.path.isdir(resolved_path):
                        self.system_extract_dirs[system] = resolved_path
                
                # Restore 3DS workflow settings
                self.threeds_backup_original = config.get('threeds_backup_original', True)
                self.threeds_delete_archives = config.get('threeds_delete_archives', False)
                self.threeds_delete_after_move = config.get('threeds_delete_after_move', False)
                self.threeds_auto_clean_names = config.get('threeds_auto_clean_names', True)
                self.threeds_source_dir = self._resolve_portable_path(config.get('threeds_source_dir', '')) or ""
                self.threeds_dest_dir = self._resolve_portable_path(config.get('threeds_dest_dir', '')) or ""
        except Exception as e:
            # Silently fail - use defaults
            pass

    def load_progress(self):
        """Load progress from previous conversion sessions for crash recovery"""
        try:
            if self.progress_file.exists():
                with open(self.progress_file, 'r') as f:
                    data = json.load(f)
                    self.completed_files = set(data.get('completed_files', []))
                    self.current_batch_id = data.get('batch_id')
                    if self.completed_files:
                        self.log(f"📂 Loaded progress: {len(self.completed_files)} files previously completed")
        except Exception as e:
            self.log(f"⚠️  Could not load progress file: {e}")
            self.completed_files = set()

    def save_progress(self, source_dir):
        """Save current conversion progress for crash recovery"""
        try:
            data = {
                'batch_id': self.current_batch_id,
                'source_dir': str(source_dir),
                'completed_files': list(self.completed_files),
                'timestamp': time.time()
            }
            with open(self.progress_file, 'w') as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            self.log(f"⚠️  Could not save progress: {e}")

    def clear_progress(self):
        """Clear progress file after successful completion"""
        try:
            if self.progress_file.exists():
                self.progress_file.unlink()
                self.completed_files = set()
                self.current_batch_id = None
        except Exception:
            pass
    
    def detect_iso_system(self, name, size_bytes=None, full_path=None):
        """Determine whether an ISO is PS2, PSP, PS3, or Xbox using IDs, folder path, and size."""
        lower_name = str(name).lower()
        
        # Check for game ID patterns in filename (very reliable)
        if any(token in lower_name for token in PSP_ID_PATTERNS):
            return 'PSP'
        if any(token in lower_name for token in PS2_ID_PATTERNS):
            return 'PlayStation 2'
        if any(token in lower_name for token in PS3_ID_PATTERNS):
            return 'PlayStation 3'
        
        # Check for Xbox specific extensions/naming
        if lower_name.endswith('.xiso'):
            return 'Xbox'
            
        # Check folder path for system hints (use full path if available)
        path_to_check = str(full_path).lower() if full_path else lower_name
        path_parts = re.split(r'[\\/]', path_to_check)
        
        # Check each folder part for system names
        for part in path_parts:
            part_clean = part.strip()
            # PSP folder detection
            if part_clean == 'psp' or 'playstation portable' in part_clean:
                return 'PSP'
            # PS2 folder detection
            if part_clean == 'ps2' or part_clean == 'playstation 2' or part_clean == 'playstation2' or 'sony playstation 2' in part_clean:
                return 'PlayStation 2'
            # PS3 folder detection
            if part_clean == 'ps3' or part_clean == 'playstation 3' or part_clean == 'playstation3' or 'sony playstation 3' in part_clean:
                return 'PlayStation 3'
            # Xbox folder detection
            if part_clean == 'xbox' or 'microsoft xbox' in part_clean:
                if '360' in part_clean: return 'Xbox 360'
                return 'Xbox'
            if 'xbox 360' in part_clean or 'xbox360' in part_clean:
                return 'Xbox 360'
        
        # Size-based heuristic as last resort
        if size_bytes is not None:
            size_gb = size_bytes / (1024 * 1024 * 1024)
            # PS2 DVDs are typically 4.7GB or larger, PSP UMDs max at ~1.8GB
            if 3.0 <= size_gb <= 4.7:
                # Could be PS2 or Xbox, but PS2 is more common
                return 'PlayStation 2'
            # PS3 games are usually much larger, often 8GB to 45GB
            if size_gb > 5.0 and size_gb <= 50.0:
                return 'PlayStation 3'
            # Xbox 360 ISOs are usually 7.3GB (XGD2) or 8.1GB (XGD3)
            if 7.0 <= size_gb <= 8.5:
                return 'Xbox 360'
        
        # Check name patterns
        if 'xbox' in lower_name:
            if '360' in lower_name: return 'Xbox 360'
            return 'Xbox'
        if 'ps3' in lower_name:
            return 'PlayStation 3'
            
        return None

    def _make_portable_path(self, path_value):
        """Store paths relative to the app folder when possible for portability."""
        if not path_value:
            return ""
        try:
            path_obj = Path(path_value)
            # If already relative, keep as-is
            if not path_obj.is_absolute():
                return str(path_obj)
            # If inside the app directory, store as relative like ./subdir/file
            try:
                relative = path_obj.relative_to(self.script_dir)
                return str(Path(".") / relative)
            except ValueError:
                return str(path_obj)
        except Exception:
            return str(path_value)

    def _resolve_portable_path(self, stored_value):
        """Resolve stored paths back to absolute paths anchored at the app folder when relative."""
        if not stored_value:
            return ""
        try:
            path_obj = Path(stored_value)
            if path_obj.is_absolute():
                return str(path_obj)
            # Treat relative paths as relative to the app directory
            return str((self.script_dir / path_obj).resolve())
        except Exception:
            return stored_value

    def set_theme_colors(self, theme_name):
        """Set global COLORS to the chosen theme palette"""
        palette = THEME_PRESETS.get(theme_name, THEME_PRESETS['PS2'])
        global COLORS
        COLORS = dict(palette)
        self.current_theme = theme_name
        self.font_body_family = palette.get('font_body', 'Consolas')
        self.font_heading_family = palette.get('font_heading', self.font_body_family)
        self.font_mono_family = palette.get('font_mono', 'Consolas')

    def init_fonts(self):
        """Create reusable font objects to allow live theme switching"""
        self.font_title = tkfont.Font(family=self.font_heading_family, size=18, weight="bold")
        self.font_heading_md = tkfont.Font(family=self.font_heading_family, size=14, weight="bold")
        self.font_subtitle = tkfont.Font(family=self.font_body_family, size=9)
        self.font_label_bold = tkfont.Font(family=self.font_body_family, size=10, weight="bold")
        self.font_body = tkfont.Font(family=self.font_body_family, size=10)
        self.font_small = tkfont.Font(family=self.font_body_family, size=9)
        self.font_button = tkfont.Font(family=self.font_body_family, size=10, weight="bold")
        self.font_status = tkfont.Font(family=self.font_body_family, size=9, weight="bold")
        self.font_mono = tkfont.Font(family=self.font_mono_family, size=9)

    def update_font_families(self):
        """Update font families on theme change"""
        for f, fam in [
            (getattr(self, 'font_title', None), self.font_heading_family),
            (getattr(self, 'font_heading_md', None), self.font_heading_family),
            (getattr(self, 'font_subtitle', None), self.font_body_family),
            (getattr(self, 'font_label_bold', None), self.font_body_family),
            (getattr(self, 'font_body', None), self.font_body_family),
            (getattr(self, 'font_small', None), self.font_body_family),
            (getattr(self, 'font_button', None), self.font_body_family),
            (getattr(self, 'font_status', None), self.font_body_family),
            (getattr(self, 'font_mono', None), self.font_mono_family),
        ]:
            if f:
                f.configure(family=fam)

    def apply_theme(self):
        """Apply current theme colors across the UI."""
        self.update_font_families()
        style = ttk.Style()
        style.configure("Retro.Horizontal.TProgressbar",
                        troughcolor=COLORS['bg_light'],
                        background=COLORS['text_primary'],
                        darkcolor=COLORS['button_green'],
                        lightcolor=COLORS['text_primary'],
                        bordercolor=COLORS['text_primary'])
        style.configure("Retro.TNotebook", background=COLORS['bg_dark'], borderwidth=0)
        style.configure("Retro.TNotebook.Tab", background=COLORS['bg_medium'],
                       foreground=COLORS['text_secondary'], padding=[18, 6],
                       font=self.font_label_bold)
        style.map("Retro.TNotebook.Tab",
                  background=[("selected", COLORS['button_blue'])],
                  foreground=[("selected", "white")])

        self.master.configure(bg=COLORS['bg_dark'])
        if self.main_frame:
            self.main_frame.configure(bg=COLORS['bg_dark'])
        if self.title_frame:
            self.title_frame.configure(bg=COLORS['bg_light'])
            for child in self.title_frame.winfo_children():
                try:
                    child.configure(bg=COLORS['bg_light'])
                    if isinstance(child, Frame):
                        for sub in child.winfo_children():
                            try:
                                sub.configure(bg=COLORS['bg_light'])
                            except Exception:
                                pass
                except Exception:
                    pass

        # Tool status bar
        if hasattr(self, '_tool_status_frame'):
            self._tool_status_frame.configure(bg=COLORS['bg_medium'])
            self._tool_status_label.configure(bg=COLORS['bg_medium'], fg=COLORS['text_secondary'])

        # Status / metrics bars
        if self.status_label:
            self.status_label.configure(bg=COLORS['bg_light'], fg=COLORS['text_primary'])
        if self.metrics_label:
            self.metrics_label.configure(bg=COLORS['bg_medium'], fg=COLORS['accent_yellow'])

        # Entry
        if self.dir_entry:
            self.dir_entry.configure(bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                                     insertbackground=COLORS['text_primary'])

        # Tool path labels
        for lbl in [self.chdman_label, self.seven_zip_label, self.maxcso_label,
                    getattr(self, 'ndecrypt_label', None), getattr(self, 'ps3_dumper_label', None)]:
            if lbl:
                try:
                    lbl.configure(bg=COLORS['bg_medium'], fg=COLORS['text_secondary'])
                except Exception:
                    pass

        # Buttons
        for btn in [self.scan_button, self.convert_button, self.stop_button, self.move_chd_button,
                    self.cleanup_button, self.clean_names_button, self.extract_archives_button,
                    self.decrypt_3ds_button, self.ia_browser_button, self.bios_download_button]:
            if btn:
                try:
                    btn.configure(activebackground=COLORS['text_primary'])
                except Exception:
                    pass

        # Log
        if self.log_text:
            self.log_text.configure(bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                                    insertbackground=COLORS['text_primary'],
                                    selectbackground=COLORS['accent_purple'])
        if self.progress:
            self.progress.configure(style="Retro.Horizontal.TProgressbar")

        # Collapsible frames
        if hasattr(self, 'options_frame'):
            self.options_frame.set_colors(COLORS['bg_light'], COLORS['accent_pink'])

        # Bottom bar
        if hasattr(self, '_bottom_bar'):
            self._bottom_bar.configure(bg=COLORS['bg_dark'])
    
    def _load_window_size(self):
        """Load persisted window size from config."""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                return config.get("window_width", 1100), config.get("window_height", 800)
        except Exception:
            pass
        return 1100, 800

    def _save_window_size(self):
        """Persist current window size."""
        try:
            self.save_config()
        except Exception:
            pass

    def _on_window_resize(self, event=None):
        """Throttled window resize handler."""
        if not hasattr(self, "_resize_after_id"):
            self._resize_after_id = None
        if self._resize_after_id:
            self.master.after_cancel(self._resize_after_id)
        self._resize_after_id = self.master.after(500, self._save_window_size)

    def _on_drop(self, event):
        """Handle drag-and-drop of ROM directory."""
        data = event.data
        if data:
            # Strip braces from paths with spaces on Windows, or handle URL-encoded paths
            path = data.strip()
            if path.startswith("{") and path.endswith("}"):
                path = path[1:-1]
            if os.path.isdir(path):
                self.source_dir = path
                self.dir_entry.delete(0, "end")
                self.dir_entry.insert(0, path)
                self.save_config()
                self.log(f"📁 Dropped directory: {path}")
            elif os.path.isfile(path) and os.path.isdir(os.path.dirname(path)):
                self.source_dir = os.path.dirname(path)
                self.dir_entry.delete(0, "end")
                self.dir_entry.insert(0, self.source_dir)
                self.save_config()
                self.log(f"📁 Set directory from dropped file: {self.source_dir}")

    def _read_config_file(self):
        """Read the config file as a dict."""
        try:
            if self.config_file.exists():
                with open(self.config_file, "r") as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    def _write_config_file(self, config):
        """Write config dict to file."""
        try:
            with open(self.config_file, "w") as f:
                json.dump(config, f, indent=2)
        except Exception:
            pass

    def _is_first_run(self):
        """Check if this is the first run (no ROM dir configured)."""
        return not self.source_dir and not self.config_file.exists()

    def _show_quick_start(self):
        """Show simplified Quick-Start wizard for first-time users."""
        qs = Toplevel(self.master)
        qs.title("Quick Start Wizard")
        qs.geometry("550x420")
        qs.resizable(False, False)
        qs.transient(self.master)
        qs.grab_set()
        qs.configure(bg=COLORS["bg_dark"])
        qs.lift()
        qs.focus_force()

        # Header
        header = Frame(qs, bg=COLORS["bg_light"], pady=12)
        header.pack(fill="x")
        Label(header, text="⚡ Welcome to ROM Converter!", font=self.font_title,
              fg=COLORS["text_primary"], bg=COLORS["bg_light"]).pack()
        Label(header, text="Let's get you set up in 3 steps.",
              font=self.font_small, fg=COLORS["text_secondary"], bg=COLORS["bg_light"]).pack(pady=(4, 0))

        content = Frame(qs, bg=COLORS["bg_dark"], padx=20, pady=15)
        content.pack(fill="both", expand=True)

        # Step 1: ROM Directory
        Label(content, text="Step 1: Select your ROM directory",
              font=self.font_label_bold, fg=COLORS["accent_yellow"], bg=COLORS["bg_dark"]).pack(anchor="w")
        dir_row = Frame(content, bg=COLORS["bg_dark"])
        dir_row.pack(fill="x", pady=(4, 12))
        dir_entry_qs = Entry(dir_row, font=self.font_body,
                             bg=COLORS["bg_input"], fg=COLORS["text_primary"],
                             insertbackground=COLORS["text_primary"], relief="flat")
        dir_entry_qs.pack(side="left", fill="x", expand=True, ipady=4)
        if self.source_dir:
            dir_entry_qs.insert(0, self.source_dir)
        Button(dir_row, text="Browse", command=lambda: self._qs_browse(dir_entry_qs),
               font=self.font_small, bg=COLORS["bg_light"], fg=COLORS["text_secondary"],
               relief="flat", cursor="hand2").pack(side="left", padx=(8, 0))

        # Step 2: Systems
        Label(content, text="Step 2: Select systems to convert",
              font=self.font_label_bold, fg=COLORS["accent_yellow"], bg=COLORS["bg_dark"]).pack(anchor="w")
        sys_frame = Frame(content, bg=COLORS["bg_medium"])
        sys_frame.pack(fill="x", pady=(4, 12))

        qs_system_vars = {}
        systems_row1 = ["PlayStation (PS1)", "PlayStation 2", "PSP"]
        systems_row2 = ["Xbox", "PlayStation 3", "Dreamcast"]
        for i, systems in enumerate([systems_row1, systems_row2]):
            row = Frame(sys_frame, bg=COLORS["bg_medium"])
            row.pack(fill="x", padx=8, pady=4)
            for sys_name in systems:
                var = BooleanVar(value=(sys_name in ["PlayStation (PS1)", "PlayStation 2"]))
                qs_system_vars[sys_name] = var
                cb = Checkbutton(row, text=sys_name, variable=var,
                                 font=self.font_small, fg=COLORS["text_primary"],
                                 bg=COLORS["bg_medium"], selectcolor=COLORS["bg_dark"],
                                 activebackground=COLORS["bg_medium"],
                                 activeforeground=COLORS["text_primary"])
                cb.pack(side="left", padx=4)

        # Step 3: Format
        Label(content, text="Step 3: Choose output format",
              font=self.font_label_bold, fg=COLORS["accent_yellow"], bg=COLORS["bg_dark"]).pack(anchor="w")
        fmt_frame = Frame(content, bg=COLORS["bg_dark"])
        fmt_frame.pack(fill="x", pady=(4, 16))
        fmt_var = StringVar(value="CHD (Recommended)")
        for fmt in ["CHD (Recommended)", "ISO", "CSO", "ZSO"]:
            rb = Radiobutton(fmt_frame, text=fmt, variable=fmt_var, value=fmt,
                             font=self.font_small, fg=COLORS["text_primary"],
                             bg=COLORS["bg_dark"], selectcolor=COLORS["bg_dark"],
                             activebackground=COLORS["bg_dark"],
                             activeforeground=COLORS["text_primary"])
            rb.pack(side="left", padx=(0, 16))

        # Action buttons
        btn_row = Frame(content, bg=COLORS["bg_dark"])
        btn_row.pack(fill="x", pady=(0, 8))

        def apply_quick_start():
            dir_path = dir_entry_qs.get().strip()
            if dir_path and os.path.isdir(dir_path):
                self.source_dir = dir_path
                self.dir_entry.delete(0, "end")
                self.dir_entry.insert(0, dir_path)
            selected_systems = [s for s, v in qs_system_vars.items() if v.get()]
            # Map system names to checkboxes
            system_map = {
                "PlayStation (PS1)": "ps1", "PlayStation 2": "ps2",
                "PSP": "psp", "Xbox": "xbox", "PlayStation 3": "ps3",
            }
            for sys_name, key in system_map.items():
                if sys_name in selected_systems:
                    if key == "ps1":
                        self.process_ps1_cues.set(True)
                    elif key == "ps2":
                        self.process_ps2_isos.set(True)
                    elif key == "psp":
                        self.process_psp_isos.set(True)
                    elif key == "xbox":
                        self.process_xbox_isos.set(True)
                    elif key == "ps3":
                        self.process_ps3_isos.set(True)
            chosen_fmt = fmt_var.get()
            if "CHD" in chosen_fmt:
                self.ps1_output_format = "CHD"
                self.ps2_output_format = "CHD"
            elif "ISO" in chosen_fmt:
                self.ps1_output_format = "ISO"
                self.ps2_output_format = "ISO"
            elif chosen_fmt == "CSO":
                self.ps1_output_format = "CHD"
                self.ps2_output_format = "CSO"
            elif chosen_fmt == "ZSO":
                self.ps1_output_format = "CHD"
                self.ps2_output_format = "ZSO"
            self.ps1_format_combo.set(self.ps1_output_format)
            self.ps2_format_combo.set(self.ps2_output_format)
            if self.psp_output_format:
                self.psp_format_combo.set(self.psp_output_format)
            self.save_config()
            self.log("✅ Quick Start configuration applied. Ready to convert!")
            qs.destroy()

        Button(btn_row, text="⚡ Get Started!", command=apply_quick_start,
               font=self.font_button, bg=COLORS["button_green"], fg=COLORS["bg_dark"],
               activebackground=COLORS["text_primary"], relief="flat", cursor="hand2",
               padx=24, pady=6).pack(side="left", padx=(0, 12))
        Button(btn_row, text="Skip", command=qs.destroy,
               font=self.font_small, bg=COLORS["bg_medium"], fg=COLORS["text_muted"],
               relief="flat", cursor="hand2", padx=16, pady=4).pack(side="left")

        Label(content, text="Tip: Drag & drop a ROM folder onto the main window anytime!",
              font=self.font_small, fg=COLORS["text_muted"], bg=COLORS["bg_dark"]).pack()

    def _qs_browse(self, entry_widget):
        """Browse for a directory in the Quick-Start wizard."""
        directory = filedialog.askdirectory(title="Select ROM Directory")
        if directory:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, directory)

    def _active_tab(self):
        """Return the name of the currently active notebook tab."""
        try:
            return self._notebook.tab(self._notebook.select(), "text")
        except Exception:
            return "Convert"

    def _toggle_log(self):
        """Toggle the log panel visibility."""
        if self.log_visible:
            self._log_frame.pack_forget()
            self._log_toggle_btn.config(text="▲ Show Log")
            self.log_visible = False
        else:
            self._log_frame.pack(fill="both", expand=True, before=self._bottom_bar)
            self._log_toggle_btn.config(text="▼ Hide Log")
            self.log_visible = True

    def setup_ui(self):
        """Setup the user interface with tabbed retro gaming aesthetic."""
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Retro.Horizontal.TProgressbar",
                       troughcolor=COLORS['bg_light'],
                       background=COLORS['text_primary'],
                       darkcolor=COLORS['button_green'],
                       lightcolor=COLORS['text_primary'],
                       bordercolor=COLORS['text_primary'])
        # Notebook tab style
        style.configure("Retro.TNotebook", background=COLORS['bg_dark'], borderwidth=0)
        style.configure("Retro.TNotebook.Tab", background=COLORS['bg_medium'],
                       foreground=COLORS['text_secondary'], padding=[18, 6],
                       font=self.font_label_bold)
        style.map("Retro.TNotebook.Tab",
                  background=[("selected", COLORS['button_blue'])],
                  foreground=[("selected", "white")])

        # --- Main container ---
        self.main_frame = Frame(self.master, padx=8, pady=6, bg=COLORS['bg_dark'])
        self.main_frame.pack(fill="both", expand=True)

        # --- Title banner ---
        self.title_frame = Frame(self.main_frame, bg=COLORS['bg_light'], pady=6)
        self.title_frame.pack(fill="x", pady=(0, 8))
        Label(self.title_frame, text="◄ ROM CONVERTER ►", font=self.font_title,
              fg=COLORS['text_primary'], bg=COLORS['bg_light']).pack()

        # Theme selector + About in title bar
        theme_row = Frame(self.title_frame, bg=COLORS['bg_light'])
        theme_row.pack(pady=(4, 0))
        Label(theme_row, text="Theme:", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=COLORS['bg_light']).pack(side="left", padx=(0, 6))
        self.theme_combo = ttk.Combobox(theme_row, values=list(THEME_PRESETS.keys()),
                                        state="readonly", width=8)
        if self.current_theme not in THEME_PRESETS:
            self.current_theme = 'PS2'
        self.theme_combo.set(self.current_theme)
        self.theme_combo.pack(side="left")
        self.theme_combo.bind("<<ComboboxSelected>>", lambda e: (
            self.set_theme_colors(self.theme_combo.get()),
            self.update_font_families(),
            self.save_config(),
            self.apply_theme(),
            self.log(f"🖌 Theme set to {self.theme_combo.get()}.")
        ))
        Button(theme_row, text="ℹ️ About", command=self.about_dialog,
               font=self.font_small, bg=COLORS['bg_medium'],
               fg=COLORS['text_secondary'], relief="flat", cursor="hand2",
               padx=8).pack(side="left", padx=(15, 0))

        # --- Tool status indicator ---
        self._tool_status_frame = Frame(self.main_frame, bg=COLORS['bg_medium'], pady=3)
        self._tool_status_frame.pack(fill="x", pady=(0, 6))
        self._tool_status_label = Label(self._tool_status_frame,
            text=self._build_tool_status_text(),
            font=self.font_small, fg=COLORS['text_muted'], bg=COLORS['bg_medium'], anchor="w")
        self._tool_status_label.pack(fill="x", padx=8)

        # --- Tabbed notebook ---
        self._notebook = ttk.Notebook(self.main_frame, style="Retro.TNotebook")
        self._notebook.pack(fill="both", expand=True)

        # === CONVERT TAB ===
        convert_tab = Frame(self._notebook, bg=COLORS['bg_dark'])
        self._notebook.add(convert_tab, text="  Convert  ")

        # ROM directory row
        dir_row = Frame(convert_tab, bg=COLORS['bg_dark'])
        dir_row.pack(fill="x", pady=(8, 4))
        Label(dir_row, text="📁 ROM Directory:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(side="left", padx=(0, 8))
        self.dir_entry = Entry(dir_row, font=self.font_body,
                              bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                              insertbackground=COLORS['text_primary'],
                              relief="flat", highlightthickness=1,
                              highlightcolor=COLORS['text_secondary'],
                              highlightbackground=COLORS['text_muted'])
        self.dir_entry.pack(side="left", fill="x", expand=True, padx=(0, 6), ipady=4)
        if self.source_dir:
            self.dir_entry.insert(0, self.source_dir)
        Button(dir_row, text="Browse", command=self.browse_directory,
               font=self.font_small, bg=COLORS['bg_light'],
               fg=COLORS['text_secondary'], activebackground=COLORS['accent_purple'],
               activeforeground="white", relief="flat", cursor="hand2",
               padx=10).pack(side="left")

        # --- Collapsible: General Options ---
        self.options_frame = CollapsibleFrame(convert_tab, title="⚙ General Options",
            font=self.font_label_bold, fg=COLORS['accent_pink'],
            bg=COLORS['bg_light'], collapsed=True)
        self.options_frame.pack(fill="x", pady=(6, 2))
        cb_bg = COLORS['bg_light']
        cb_font = self.font_small
        for txt, var, fg_color in [
            ("Scan subdirectories recursively", self.recursive, COLORS['text_primary']),
            ("Move originals to backup folder after conversion", self.move_to_backup, COLORS['text_primary']),
            ("⚠ Delete original files after successful conversion", self.delete_originals, COLORS['accent_red']),
        ]:
            Checkbutton(self.options_frame.content, text=txt, variable=var,
                       font=cb_font, fg=fg_color, bg=cb_bg, selectcolor=COLORS['bg_dark'],
                       activebackground=cb_bg, activeforeground=fg_color).pack(anchor="w")

        # --- Collapsible: System Filters ---
        sys_filter = CollapsibleFrame(convert_tab, title="🎮 System Filters",
            font=self.font_label_bold, fg=COLORS['accent_yellow'],
            bg=COLORS['bg_medium'], collapsed=True)
        sys_filter.pack(fill="x", pady=(0, 2))
        sys_content = sys_filter.content
        # Select All / None row
        sel_row = Frame(sys_content, bg=COLORS['bg_medium'])
        sel_row.pack(fill="x", pady=(2, 4))
        system_vars = [
            ("PS1 CUE", self.process_ps1_cues), ("PS2 BIN/CUE", self.process_ps2_cues),
            ("PS2 ISO", self.process_ps2_isos), ("Xbox ISO", self.process_xbox_isos),
            ("PSP ISO", self.process_psp_isos), ("PS3 ISO", self.process_ps3_isos),
            ("NES ROMs", self.process_nes_roms), ("SNES ROMs", self.process_snes_roms),
            ("N64 ROMs", self.process_n64_roms),
        ]
        def _sel_all():
            for _, v in system_vars: v.set(True)
        def _sel_none():
            for _, v in system_vars: v.set(False)
        Button(sel_row, text="All", command=_sel_all,
               font=self.font_small, bg=COLORS['bg_light'], fg=COLORS['text_secondary'],
               relief="flat", cursor="hand2", padx=6).pack(side="left", padx=(0, 4))
        Button(sel_row, text="None", command=_sel_none,
               font=self.font_small, bg=COLORS['bg_light'], fg=COLORS['text_secondary'],
               relief="flat", cursor="hand2", padx=6).pack(side="left")
        # Grid of system checkboxes
        grid = Frame(sys_content, bg=COLORS['bg_medium'])
        grid.pack(fill="x")
        for i, (label, var) in enumerate(system_vars):
            row, col = divmod(i, 3)
            Checkbutton(grid, text=label, variable=var, font=cb_font,
                       fg=COLORS['text_primary'], bg=COLORS['bg_medium'],
                       selectcolor=COLORS['bg_dark'],
                       activebackground=COLORS['bg_medium'],
                       activeforeground=COLORS['text_primary']).grid(row=row, column=col, sticky="w", padx=4, pady=1)

        # --- Collapsible: Output Formats ---
        fmt_section = CollapsibleFrame(convert_tab, title="📀 Output Formats",
            font=self.font_label_bold, fg=COLORS['button_green'],
            bg=COLORS['bg_light'], collapsed=False)
        fmt_section.pack(fill="x", pady=(0, 2))
        c = fmt_section.content
        # PS1
        r1 = Frame(c, bg=COLORS['bg_light'])
        r1.pack(fill="x", pady=2)
        Label(r1, text="PS1:", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=COLORS['bg_light']).pack(side="left", padx=(0, 6))
        self.ps1_format_combo = ttk.Combobox(r1, values=PS1_OUTPUT_FORMATS, state="readonly", width=6)
        self.ps1_format_combo.set(self.ps1_output_format)
        self.ps1_format_combo.pack(side="left")
        self.ps1_format_combo.bind("<<ComboboxSelected>>", lambda e: setattr(self, 'ps1_output_format', self.ps1_format_combo.get()) or self.save_config())
        # PS2
        r2 = Frame(c, bg=COLORS['bg_light'])
        r2.pack(fill="x", pady=2)
        Label(r2, text="PS2:", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=COLORS['bg_light']).pack(side="left", padx=(0, 6))
        self.ps2_emulator_combo = ttk.Combobox(r2, values=PS2_EMULATORS, state="readonly", width=10)
        self.ps2_emulator_combo.set(self.ps2_emulator)
        self.ps2_emulator_combo.pack(side="left", padx=(0, 6))
        self.ps2_format_combo = ttk.Combobox(r2, values=PS2_OUTPUT_FORMATS, state="readonly", width=6)
        self.ps2_format_combo.set(self.ps2_output_format)
        self.ps2_format_combo.pack(side="left", padx=6)
        self.ps2_format_combo.bind("<<ComboboxSelected>>", lambda e: (
            setattr(self, 'ps2_output_format', self.ps2_format_combo.get()),
            self.save_config(),
            self.log("⚠ maxcso required for CSO/ZSO") if self.ps2_output_format in ['CSO','ZSO'] and not self.maxcso_path else None
        ))
        self.ps2_emulator_combo.bind("<<ComboboxSelected>>", lambda e: self._on_ps2_emu_change())
        # PSP
        r3 = Frame(c, bg=COLORS['bg_light'])
        r3.pack(fill="x", pady=2)
        Label(r3, text="PSP:", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=COLORS['bg_light']).pack(side="left", padx=(0, 6))
        self.psp_format_combo = ttk.Combobox(r3, values=['CSO', 'ZSO'], state="readonly", width=6)
        self.psp_format_combo.set(self.psp_output_format)
        self.psp_format_combo.pack(side="left")
        self.psp_format_combo.bind("<<ComboboxSelected>>", lambda e: (
            setattr(self, 'psp_output_format', self.psp_format_combo.get()),
            self.save_config(),
            self.log("⚠ maxcso required for PSP CSO/ZSO") if not self.maxcso_path else None
        ))

        # --- Collapsible: Archive Handling ---
        arch_section = CollapsibleFrame(convert_tab, title="📦 Archive Handling",
            font=self.font_label_bold, fg=COLORS['accent_orange'],
            bg=COLORS['bg_medium'], collapsed=True)
        arch_section.pack(fill="x", pady=(0, 2))
        Checkbutton(arch_section.content, text="Extract compressed files before conversion",
                   variable=self.extract_compressed, font=cb_font,
                   fg=COLORS['accent_orange'], bg=COLORS['bg_medium'],
                   selectcolor=COLORS['bg_dark'],
                   activebackground=COLORS['bg_medium'],
                   activeforeground=COLORS['accent_orange']).pack(anchor="w")
        Checkbutton(arch_section.content, text="⚠ Delete archive files after extraction",
                   variable=self.delete_archives_after_extract, font=cb_font,
                   fg=COLORS['accent_red'], bg=COLORS['bg_medium'],
                   selectcolor=COLORS['bg_dark'],
                   activebackground=COLORS['bg_medium'],
                   activeforeground=COLORS['accent_red']).pack(anchor="w")

        # --- Action buttons ---
        btn_frame = Frame(convert_tab, bg=COLORS['bg_dark'])
        btn_frame.pack(fill="x", pady=(8, 4))
        self.scan_button = Button(btn_frame, text="🔍 SCAN", command=self.scan_directory,
            font=self.font_button, bg=COLORS['button_green'], fg=COLORS['bg_dark'],
            activebackground=COLORS['text_primary'], activeforeground=COLORS['bg_dark'],
            relief="flat", cursor="hand2", padx=14, pady=5)
        self.scan_button.pack(side="left", padx=(0, 6))
        self.convert_button = Button(btn_frame, text="🔥 CONVERT", command=self.start_conversion,
            font=self.font_button, bg=COLORS['button_blue'], fg="white",
            activebackground=COLORS['text_secondary'], activeforeground=COLORS['bg_dark'],
            disabledforeground=COLORS['text_muted'], relief="flat", cursor="hand2",
            padx=14, pady=5, state="disabled")
        self.convert_button.pack(side="left", padx=(0, 6))
        self.stop_button = Button(btn_frame, text="■ STOP", command=self.stop_conversion,
            font=self.font_button, bg=COLORS['accent_red'], fg="white",
            activebackground=COLORS['accent_orange'], disabledforeground="white",
            relief="flat", cursor="hand2", padx=14, pady=5, state="disabled")
        self.stop_button.pack(side="left", padx=(0, 6))
        self.move_chd_button = Button(btn_frame, text="📁 MOVE CHD", command=self.move_chd_files_dialog,
            font=self.font_button, bg=COLORS['accent_purple'], fg="white",
            activebackground=COLORS['accent_pink'], relief="flat", cursor="hand2", padx=14, pady=5)
        self.move_chd_button.pack(side="left")
        # Progress bar
        self.progress = ttk.Progressbar(convert_tab, mode='determinate',
                                        style="Retro.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(8, 4), ipady=3)

        # Phase 5: BatchProgressPanel — per-file queue with ETA and cancel
        if BATCH_PROGRESS_AVAILABLE:
            try:
                self._batch_progress_panel = create_batch_progress_panel(
                    convert_tab,
                    log_callback=self.log,
                    on_pause_resume=self._on_batch_pause_resume,
                    on_cancel_file=self._on_cancel_file,
                )
                self._batch_progress_panel.pack(fill="both", expand=True, pady=(0, 4))
            except Exception as e:
                print(f"⚠️  BatchProgressPanel init failed: {e}")
                self._batch_progress_panel = None

        # --- Scan preview (hidden by default) ---
        self._scan_preview_frame = Frame(convert_tab, bg=COLORS['bg_dark'])
        self._scan_preview_frame.pack(fill="both", expand=True)
        self._scan_preview_frame.pack_forget()

        # === TOOLS TAB ===
        tools_tab = Frame(self._notebook, bg=COLORS['bg_dark'])
        self._notebook.add(tools_tab, text="  Tools  ")
        tools_inner = Frame(tools_tab, bg=COLORS['bg_dark'], padx=10, pady=10)
        tools_inner.pack(fill="both", expand=True)
        Label(tools_inner, text="🛠 Utility Tools", font=self.font_heading_md,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(anchor="w", pady=(0, 10))
        tool_buttons = [
            ("🗑️  CLEANUP – Remove leftover files", self.cleanup_compressed_dialog, COLORS['accent_orange']),
            ("✨  CLEAN NAMES – Normalize ROM filenames", self.clean_names_dialog, COLORS['accent_pink']),
            ("📦  EXTRACT – Bulk extract archives", self.extract_archives_dialog, COLORS['accent_yellow']),
            ("🔓  3DS DECRYPT – Decrypt 3DS ROMs", self.decrypt_3ds_dialog, COLORS['accent_purple']),
            ("🌐  ROM BROWSER – Browse Archive.org", self.ia_browser_dialog, COLORS['button_blue']),
            ("🧬  BIOS – Download BIOS files", self.download_bios_dialog, COLORS['accent_pink']),
            ("📁  MOVE CHD – Move converted files", self.move_chd_files_dialog, COLORS['accent_purple']),
        ]
        for label, cmd, color in tool_buttons:
            Button(tools_inner, text=label, command=cmd,
                   font=self.font_button, bg=color, fg="white",
                   activebackground=COLORS['text_primary'], activeforeground=COLORS['bg_dark'],
                   relief="flat", cursor="hand2", padx=16, pady=6).pack(fill="x", pady=3)

        # Store references for tools
        self.cleanup_button = tools_inner.winfo_children()[1]
        self.clean_names_button = tools_inner.winfo_children()[2]
        self.extract_archives_button = tools_inner.winfo_children()[3]
        self.decrypt_3ds_button = tools_inner.winfo_children()[4]
        self.ia_browser_button = tools_inner.winfo_children()[5]
        self.bios_download_button = tools_inner.winfo_children()[6]

        # === SETTINGS TAB ===
        settings_tab = Frame(self._notebook, bg=COLORS['bg_dark'])
        self._notebook.add(settings_tab, text="  Settings  ")
        settings_inner = Frame(settings_tab, bg=COLORS['bg_dark'], padx=10, pady=10)
        settings_inner.pack(fill="both", expand=True)

        # Tool paths section
        path_section = CollapsibleFrame(settings_inner, title="🔧 Tool Paths",
            font=self.font_label_bold, fg=COLORS['accent_yellow'],
            bg=COLORS['bg_medium'], collapsed=False)
        path_section.pack(fill="x", pady=(0, 8))
        pc = path_section.content
        self._build_tool_path_row(pc, "⚙ chdman:", self.chdman_path, self.browse_chdman, "Change")
        self._build_tool_path_row(pc, "📦 7-Zip:", self.seven_zip_path, self.browse_7zip, "Set")
        self._build_tool_path_row(pc, "🗜 maxcso:", self.maxcso_path, self.browse_maxcso, "Set")
        self._build_tool_path_row_with_dl(pc, "💚 extract-xiso:", self.extract_xiso_path, self.browse_extract_xiso, self.download_extract_xiso)
        self._build_tool_path_row_with_dl(pc, "🔓 NDecrypt:", self.ndecrypt_path, self.browse_ndecrypt, self.download_ndecrypt)
        self._build_tool_path_row_with_dl(pc, "🎮 PS3 Dumper:", self.ps3_dumper_path, self._browse_ps3_dumper, self.download_ps3_dumper)

        # Concurrency
        conc_section = CollapsibleFrame(settings_inner, title="⚡ Performance",
            font=self.font_label_bold, fg=COLORS['button_green'],
            bg=COLORS['bg_light'], collapsed=False)
        conc_section.pack(fill="x", pady=(0, 8))
        cc = conc_section.content
        cr = Frame(cc, bg=COLORS['bg_light'])
        cr.pack(fill="x", pady=4)
        Label(cr, text="Max concurrent conversions:", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=COLORS['bg_light']).pack(side="left")
        self.concurrent_label = Label(cr, text=str(self.max_concurrent_conversions),
            font=self.font_label_bold, fg=COLORS['accent_yellow'], bg=COLORS['bg_light'], width=3)
        self.concurrent_label.pack(side="left", padx=(8, 0))
        max_cores = multiprocessing.cpu_count()
        self.concurrent_slider = ttk.Scale(cr, from_=1, to=max_cores,
            orient="horizontal", length=150, command=self.on_concurrent_change)
        self.concurrent_slider.set(self.max_concurrent_conversions)
        self.concurrent_slider.pack(side="left", padx=(8, 0))
        Label(cr, text=f"(1-{max_cores} cores)", font=self.font_small,
              fg=COLORS['text_muted'], bg=COLORS['bg_light']).pack(side="left", padx=(8, 0))

        # Phase 4 Week 1: Tool Status Display
        if CONFIG_MANAGER_UI_AVAILABLE:
            try:
                self._tool_status_display = create_tool_status_display(
                    settings_inner,
                    tool_registry=getattr(self, "tool_registry", None),
                )
                self._tool_status_display.pack(fill="x", pady=(0, 8))
            except Exception as exc:
                print(f"⚠️  ToolStatusDisplay failed: {exc}")

        # Phase 4 Week 1: Config Validation Display
        if CONFIG_MANAGER_UI_AVAILABLE:
            try:
                self._config_validation_display = create_config_validation_display(
                    settings_inner,
                    config_adapter=getattr(self, "config_adapter", None),
                )
                self._config_validation_display.pack(fill="x", pady=(0, 8))
            except Exception as exc:
                print(f"⚠️  ConfigValidationDisplay failed: {exc}")

        # Phase 4 Week 1: Config Export/Import
        if CONFIG_MANAGER_UI_AVAILABLE:
            try:
                self._config_export_import = create_config_export_import(
                    settings_inner,
                    config_adapter=getattr(self, "config_adapter", None),
                )
                self._config_export_import.pack(fill="x", pady=(0, 8))
            except Exception as exc:
                print(f"⚠️  ConfigExportImport failed: {exc}")

        # --- Bottom bar (status + metrics) ---
        self._bottom_bar = Frame(self.main_frame, bg=COLORS['bg_dark'])
        self._bottom_bar.pack(fill="x", side="bottom")
        total_cores = multiprocessing.cpu_count()
        self.status_label = Label(self._bottom_bar,
            text=f"▶ READY | {self.cpu_cores}/{total_cores} CPU CORES | 1 CORE RESERVED",
            font=self.font_status, fg=COLORS['text_primary'], bg=COLORS['bg_light'],
            anchor="w", padx=8, pady=4)
        self.status_label.pack(fill="x")
        self.metrics_label = Label(self._bottom_bar, text="◆ METRICS: IDLE ◆", anchor="w",
            bg=COLORS['bg_medium'], fg=COLORS['accent_yellow'],
            font=self.font_status, padx=8, pady=4)
        self.metrics_label.pack(fill="x")

        # --- Collapsible log panel ---
        self._log_toggle_btn = Button(self.main_frame, text="▲ Show Log", command=self._toggle_log,
            font=self.font_small, bg=COLORS['bg_medium'], fg=COLORS['text_muted'],
            relief="flat", cursor="hand2")
        self._log_toggle_btn.pack(fill="x", side="bottom", before=self._bottom_bar)

        self._log_frame = Frame(self.main_frame, bg=COLORS['bg_dark'])
        log_scroll = Scrollbar(self._log_frame, bg=COLORS['bg_light'],
                               troughcolor=COLORS['bg_dark'],
                               activebackground=COLORS['text_primary'])
        log_scroll.pack(side="right", fill="y")
        self.log_text = Text(self._log_frame, wrap="word", yscrollcommand=log_scroll.set,
                            height=10, font=self.font_mono,
                            bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                            insertbackground=COLORS['text_primary'],
                            selectbackground=COLORS['accent_purple'],
                            selectforeground="white", relief="flat", padx=8, pady=8)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.config(command=self.log_text.yview)
        # Log starts hidden

        # --- Trace config save on option changes ---
        for var in [self.delete_originals, self.move_to_backup, self.recursive,
                    self.process_ps1_cues, self.process_ps2_cues, self.process_ps2_isos,
                    self.process_psp_isos, self.process_ps3_isos, self.process_xbox_isos,
                    self.process_nes_roms, self.process_snes_roms, self.process_n64_roms,
                    self.extract_compressed, self.delete_archives_after_extract]:
            var.trace_add('write', lambda *args: self.save_config())

        if not PSUTIL_AVAILABLE:
            self.log("ℹ Resource metrics disabled (psutil not installed - this is optional)")

        # === PERFORMANCE TAB (Phase 4 Week 6) ===
        self._setup_performance_tab()

        self.process_log_queue()
        self.apply_theme()

        # Show quick-start wizard on first run
        self.master.after(500, lambda: self._show_quick_start() if self._is_first_run() else None)

    def _on_perf_settings_apply(self, settings: dict) -> None:
        """Called when the PerformanceSettingsPanel Apply button is clicked."""
        self.log(f"⚙ Performance settings applied: {settings}")
        # Update optimizer if available
        optimizer = getattr(self, "performance_optimizer", None)
        if optimizer:
            try:
                if "max_workers" in settings:
                    optimizer.set_max_workers(settings["max_workers"])
            except Exception:
                pass

    def _on_batch_pause_resume(self, paused: bool) -> None:
        """Called by BatchProgressPanel when user clicks Pause/Resume."""
        self.log("⏸ Batch paused by user" if paused else "▶ Batch resumed by user")
        # Underlying pause flag checked in conversion loop
        self._batch_paused = paused

    def _on_cancel_file(self, file_id: str) -> None:
        """Called by BatchProgressPanel when user cancels a specific file."""
        self.log(f"🚫 Cancel requested for: {file_id}")
        if not hasattr(self, "_cancelled_files"):
            self._cancelled_files = set()
        self._cancelled_files.add(file_id)

    def _setup_performance_tab(self):
        """Add the Performance tab to the notebook (Phase 4 Week 6)."""
        try:
            from performance_status_panel import create_performance_panel
        except ImportError:
            return  # Module not available — skip silently

        perf_tab = Frame(self._notebook, bg=COLORS['bg_dark'])
        self._notebook.add(perf_tab, text="  ⚡ Performance  ")

        inner = Frame(perf_tab, bg=COLORS['bg_dark'], padx=10, pady=10)
        inner.pack(fill="both", expand=True)

        Label(
            inner,
            text="⚡ Performance Monitor",
            font=self.font_heading_md,
            fg=COLORS['text_primary'],
            bg=COLORS['bg_dark'],
        ).pack(anchor="w", pady=(0, 8))

        # Live status panel
        self._perf_panel = create_performance_panel(
            inner,
            optimizer=getattr(self, "performance_optimizer", None),
            error_handler=getattr(self, "error_handler", None),
            analyzer=getattr(self, "performance_analyzer", None),
            log_callback=self.log,
        )
        self._perf_panel.pack(fill="x", pady=(0, 8))
        self._perf_panel.start_updates()

        # Session summary section
        Label(
            inner,
            text="Session Metrics",
            font=self.font_label_bold,
            fg=COLORS['accent_yellow'],
            bg=COLORS['bg_dark'],
        ).pack(anchor="w", pady=(8, 4))

        self._session_summary_text = Text(
            inner,
            height=8,
            bg=COLORS['bg_medium'],
            fg=COLORS['text_secondary'],
            font=("Courier", 9),
            relief="flat",
            state="disabled",
        )
        self._session_summary_text.pack(fill="x")

        # Refresh button
        Frame(inner, bg=COLORS['bg_dark'], height=4).pack()
        Button(
            inner,
            text="🔄 Refresh Session Summary",
            command=self._refresh_session_summary,
            font=self.font_button,
            bg=COLORS['button_blue'],
            fg="white",
            relief="flat",
            cursor="hand2",
            padx=12,
            pady=4,
        ).pack(anchor="w")

        # Performance settings panel (Phase 5 Week 2)
        if PERF_SETTINGS_PANEL_AVAILABLE:
            try:
                self._perf_settings_panel = create_performance_settings_panel(
                    inner,
                    config_adapter=getattr(self, "config_adapter", None),
                    on_apply=self._on_perf_settings_apply,
                )
                self._perf_settings_panel.pack(fill="x", pady=(8, 0))
            except Exception as exc:
                print(f"⚠️  PerformanceSettingsPanel failed: {exc}")

        # History tab (Phase 5 Week 2)
        if SESSION_HISTORY_PANEL_AVAILABLE:
            try:
                history_tab = Frame(self._notebook, bg=COLORS['bg_dark'])
                self._notebook.add(history_tab, text="  📜 History  ")
                self._session_history_panel = create_session_history_panel(
                    history_tab,
                    metrics_store=getattr(self, "metrics_store", None),
                )
                self._session_history_panel.pack(fill="both", expand=True, padx=6, pady=6)
            except Exception as exc:
                print(f"⚠️  SessionHistoryPanel failed: {exc}")

    def _refresh_session_summary(self):
        """Populate the session summary text box from MetricsStore."""
        store = getattr(self, "metrics_store", None)
        if not store:
            self._set_session_summary("MetricsStore not available.")
            return

        try:
            stats = store.get_stats()
            total = stats.get("total") or 0
            if total == 0:
                self._set_session_summary("No conversions recorded yet.")
                return

            successful = stats.get("successful") or 0
            failed     = stats.get("failed") or 0
            rate       = f"{successful / total * 100:.1f}%" if total else "N/A"
            avg_tp     = stats.get("avg_throughput_mbps") or 0
            max_dur    = stats.get("max_duration") or 0
            avg_dur    = stats.get("avg_duration") or 0
            total_in   = (stats.get("total_input_bytes") or 0) / 1024 / 1024 / 1024

            lines = [
                f"All-time stats ({total} conversions)",
                f"  Success rate : {rate}  ({successful} ok / {failed} failed)",
                f"  Avg throughput : {avg_tp:.2f} MB/s",
                f"  Avg duration   : {avg_dur:.1f}s",
                f"  Max duration   : {max_dur:.1f}s",
                f"  Total input    : {total_in:.2f} GB processed",
            ]
            tool_stats = store.get_tool_stats()
            if tool_stats:
                lines.append("\nPer-tool averages:")
                for t in tool_stats[:5]:
                    lines.append(
                        f"  {t['tool_name'] or '(unknown)':12s}  "
                        f"{t['total']:4d} conversions  "
                        f"{t['avg_throughput_mbps']:.2f} MB/s"
                    )
            self._set_session_summary("\n".join(lines))
        except Exception as exc:
            self._set_session_summary(f"Error loading metrics: {exc}")

    def _set_session_summary(self, text: str):
        """Helper to safely update the session summary text widget."""
        try:
            self._session_summary_text.config(state="normal")
            self._session_summary_text.delete("1.0", "end")
            self._session_summary_text.insert("end", text)
            self._session_summary_text.config(state="disabled")
        except Exception:
            pass

    def _build_tool_path_row(self, parent, label, path, browse_cmd, btn_text):
        """Build a tool path row for the settings tab."""
        row = Frame(parent, bg=COLORS['bg_medium'], pady=2)
        row.pack(fill="x", pady=1)
        Label(row, text=label, font=self.font_label_bold,
              fg=COLORS['accent_yellow'], bg=COLORS['bg_medium'], width=14, anchor="w").pack(side="left")
        label_widget = Label(row, text=path or "Not set", font=self.font_small,
            fg=COLORS['text_secondary'] if path else COLORS['text_muted'],
            bg=COLORS['bg_medium'], anchor="w")
        label_widget.pack(side="left", fill="x", expand=True, padx=(4, 8))
        btn = Button(row, text=f"[ {btn_text} ]", command=browse_cmd,
            font=self.font_small, bg=COLORS['button_blue'], fg="white",
            activebackground=COLORS['accent_purple'], relief="flat", cursor="hand2")
        btn.pack(side="left")
        # Store label references
        if "chdman" in label:
            self.chdman_label = label_widget
        elif "7-Zip" in label:
            self.seven_zip_label = label_widget
        elif "maxcso" in label:
            self.maxcso_label = label_widget

    def _build_tool_path_row_with_dl(self, parent, label, path, browse_cmd, download_cmd):
        """Build a tool path row with Set + Download buttons."""
        row = Frame(parent, bg=COLORS['bg_medium'], pady=2)
        row.pack(fill="x", pady=1)
        Label(row, text=label, font=self.font_label_bold,
              fg=COLORS['accent_yellow'], bg=COLORS['bg_medium'], width=14, anchor="w").pack(side="left")
        label_widget = Label(row, text=path or "Not set", font=self.font_small,
            fg=COLORS['text_secondary'] if path else COLORS['text_muted'],
            bg=COLORS['bg_medium'], anchor="w")
        label_widget.pack(side="left", fill="x", expand=True, padx=(4, 8))
        Button(row, text="[ Set ]", command=browse_cmd,
            font=self.font_small, bg=COLORS['bg_light'], fg=COLORS['text_secondary'],
            relief="flat", cursor="hand2").pack(side="left", padx=2)
        Button(row, text="[ Download ]", command=download_cmd,
            font=self.font_small, bg=COLORS['accent_purple'], fg="white",
            relief="flat", cursor="hand2").pack(side="left", padx=2)
        # Store label references
        if "NDecrypt" in label:
            self.ndecrypt_label = label_widget
        elif "PS3" in label:
            self.ps3_dumper_label = label_widget

    def _browse_ps3_dumper(self):
        filetypes = [("Executable files", "*.exe"), ("All files", "*.*")]
        selected = filedialog.askopenfilename(title="Select ps3-disc-dumper executable", filetypes=filetypes)
        if selected:
            self.ps3_dumper_path = selected
            self.ps3_dumper_label.config(text=self.ps3_dumper_path, fg=COLORS['text_secondary'])
            self._update_tool_status()
            self.save_config()

    def _on_ps2_emu_change(self):
        self.ps2_emulator = self.ps2_emulator_combo.get()
        recommended = PS2_EMULATOR_RECOMMENDATIONS.get(self.ps2_emulator, 'CHD')
        if recommended in PS2_OUTPUT_FORMATS:
            self.ps2_output_format = recommended
            self.ps2_format_combo.set(recommended)
            self.log(f"ℹ Using recommended format for {self.ps2_emulator}: {recommended}")
            self.save_config()

    def _build_tool_status_text(self):
        """Build compact tool status string."""
        parts = []
        parts.append(f"⚙ chdman: {'✓' if self.chdman_path else '✗'}")
        parts.append(f"7z: {'✓' if self.seven_zip_path else '✗'}")
        parts.append(f"maxcso: {'✓' if self.maxcso_path else '✗'}")
        parts.append(f"xiso: {'✓' if self.extract_xiso_path else '✗'}")
        parts.append(f"NDecrypt: {'✓' if self.ndecrypt_path else '✗'}")
        parts.append(f"PS3: {'✓' if self.ps3_dumper_path else '✗'}")
        return " | ".join(parts)

    def _update_tool_status(self):
        """Update the compact tool status bar."""
        if hasattr(self, '_tool_status_label'):
            self._tool_status_label.config(text=self._build_tool_status_text())
    
    def browse_directory(self):
        """Open directory browser"""
        directory = filedialog.askdirectory(title="Select ROM Directory")
        if directory:
            self.source_dir = directory
            self.dir_entry.delete(0, "end")
            self.dir_entry.insert(0, directory)
            self.save_config()
            self.log(f"Selected directory: {directory}")
    
    def log(self, message):
        """Add message to log (thread-safe)"""
        self.log_queue.put(message)
    
    def process_log_queue(self):
        """Process queued log messages from threads - batched for performance."""
        try:
            messages = []
            for _ in range(20):
                try:
                    message = self.log_queue.get_nowait()
                    messages.append(message)
                except:
                    break

            if messages:
                # Auto-expand log panel during conversion or when new messages arrive
                if not self.log_visible and (self.is_converting or len(messages) >= 3):
                    self._toggle_log()
                self.log_text.insert("end", "\n".join(messages) + "\n")
                self.log_text.see("end")
                self.log_text.update_idletasks()
                # Periodically refresh tool status
                if hasattr(self, '_tool_status_label'):
                    self._tool_status_label.config(text=self._build_tool_status_text())
        except:
            pass
        finally:
            interval = 50 if self.is_converting else 200
            self.master.after(interval, self.process_log_queue)
    
    def keep_ui_responsive(self):
        """Call periodically during long operations to keep UI responsive.
        
        Throttled to avoid excessive updates.
        """
        current_time = time.time()
        if current_time - self.last_ui_update >= 0.1:  # Max 10 updates per second
            try:
                self.master.update_idletasks()
                self.last_ui_update = current_time
            except:
                pass
    
    def find_cue_files(self, directory, recursive=True):
        """Find all .cue files in directory"""
        cue_files = []
        path = Path(directory)
        
        if recursive:
            cue_files = list(path.rglob("*.cue"))
        else:
            cue_files = list(path.glob("*.cue"))
        
        return sorted(cue_files)

    def find_compressed_files(self, directory, recursive=True):
        """Find all compressed files in directory"""
        path = Path(directory)
        compressed_files = []
        
        for ext in COMPRESSED_EXTENSIONS:
            if recursive:
                compressed_files.extend(path.rglob(f"*{ext}"))
            else:
                compressed_files.extend(path.glob(f"*{ext}"))
        
        return sorted(compressed_files)
    
    def extract_archive_with_registry(self, archive_path):
        """Extract a compressed archive using ExtractionRegistry (Phase 3 Week 4 refactored version)
        
        Returns tuple of (success: bool, output_dir: Path or None)
        """
        if not self.extraction_registry:
            # Fallback to original implementation
            return self._extract_archive_original(archive_path)
        
        archive_path = Path(archive_path)
        
        # Wait for memory pressure before extraction
        self._wait_for_memory_pressure()
        
        # Determine output directory
        extract_folder = archive_path.parent / archive_path.stem
        
        # Handle multi-extension like .tar.gz
        if archive_path.name.lower().endswith(('.tar.gz', '.tgz')):
            extract_folder = archive_path.parent / archive_path.name.replace('.tar.gz', '').replace('.tgz', '')
        elif archive_path.name.lower().endswith(('.tar.bz2', '.tbz2')):
            extract_folder = archive_path.parent / archive_path.name.replace('.tar.bz2', '').replace('.tbz2', '')
        elif archive_path.name.lower().endswith(('.tar.xz', '.txz')):
            extract_folder = archive_path.parent / archive_path.name.replace('.tar.xz', '').replace('.txz', '')
        
        try:
            success, output_dir = self.extraction_registry.extract(archive_path, extract_folder)
            if success and output_dir:
                gc.collect()  # Free memory after extraction
                return True, output_dir
            else:
                return False, None
        except Exception as e:
            self.log(f"  ❌ Exception in registry extraction: {e}")
            return False, None
    
    def extract_archive(self, archive_path):
        """Extract a compressed archive to a folder with the same name
        
        Dispatches to ExtractionRegistry if available (Phase 3 Week 4),
        otherwise uses original implementation for compatibility.
        
        Returns tuple of (success: bool, output_dir: Path or None)
        """
        # Try registry-based extraction first if available
        if self.extraction_registry and EXTRACTION_REGISTRY_AVAILABLE:
            return self.extract_archive_with_registry(archive_path)
        
        # Fallback to original implementation
        return self._extract_archive_original(archive_path)
    
    def _extract_archive_original(self, archive_path):
        """Original extract_archive implementation (kept for compatibility)"""
        archive_path = Path(archive_path)
        
        # Wait for memory pressure before extraction
        self._wait_for_memory_pressure()
        
        # Create extraction folder (same name as archive without extension)
        extract_folder = archive_path.parent / archive_path.stem
        
        # Handle multi-extension like .tar.gz
        if archive_path.name.endswith('.tar.gz') or archive_path.name.endswith('.tgz'):
            extract_folder = archive_path.parent / archive_path.name.replace('.tar.gz', '').replace('.tgz', '')
        
        try:
            extract_folder.mkdir(exist_ok=True)
            ext = archive_path.suffix.lower()
            
            # Handle .zip files using Python's built-in zipfile
            if ext == '.zip':
                self.log(f"  📦 Extracting ZIP: {archive_path.name}")
                with zipfile.ZipFile(archive_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_folder)
                self.log(f"  ✅ Extracted to: {extract_folder.name}/")
                gc.collect()  # Free memory after extraction
                return True, extract_folder
            
            # Handle .tar, .tar.gz, .tgz files using Python's tarfile
            elif ext in ['.tar', '.gz', '.tgz'] or archive_path.name.endswith('.tar.gz'):
                self.log(f"  📦 Extracting TAR: {archive_path.name}")
                mode = 'r:gz' if ext in ['.gz', '.tgz'] or archive_path.name.endswith('.tar.gz') else 'r'
                with tarfile.open(archive_path, mode) as tar_ref:
                    tar_ref.extractall(extract_folder)
                self.log(f"  ✅ Extracted to: {extract_folder.name}/")
                gc.collect()  # Free memory after extraction
                return True, extract_folder
            
            # Handle .7z and .rar files using 7-Zip
            elif ext in ['.7z', '.rar']:
                if not self.seven_zip_path:
                    self.log(f"  ⚠️  Cannot extract {ext} file: 7-Zip not configured")
                    return False, None
                
                self.log(f"  📦 Extracting with 7-Zip: {archive_path.name}")
                cmd = [self.seven_zip_path, 'x', str(archive_path), f'-o{extract_folder}', '-y']
                
                # Use lower process priority to reduce system impact
                creationflags = 0
                if sys.platform == 'win32':
                    creationflags = 0x00004000  # BELOW_NORMAL_PRIORITY_CLASS
                
                result = subprocess.run(
                    cmd, 
                    capture_output=True, 
                    text=True, 
                    timeout=3600,
                    creationflags=creationflags if sys.platform == 'win32' else 0
                )
                
                if result.returncode == 0:
                    self.log(f"  ✅ Extracted to: {extract_folder.name}/")
                    gc.collect()  # Free memory after extraction
                    return True, extract_folder
                else:
                    self.log(f"  ❌ 7-Zip extraction failed: {result.stderr.strip()}")
                    return False, None
            
            else:
                self.log(f"  ⚠️  Unsupported archive format: {ext}")
                return False, None
                
        except zipfile.BadZipFile:
            self.log(f"  ❌ Invalid or corrupted ZIP file: {archive_path.name}")
            return False, None
        except tarfile.TarError as e:
            self.log(f"  ❌ TAR extraction error: {e}")
            return False, None
        except subprocess.TimeoutExpired:
            self.log(f"  ❌ Extraction timeout: {archive_path.name}")
            return False, None
        except Exception as e:
            self.log(f"  ❌ Extraction error: {e}")
            return False, None
    
    def extract_all_archives(self, directory, recursive=True):
        """Find and extract all compressed files in the directory"""
        compressed_files = self.find_compressed_files(directory, recursive)
        
        if not compressed_files:
            self.log("No compressed files found to extract.")
            return []
        
        self.log(f"\n📦 Found {len(compressed_files)} compressed file(s) to extract:")
        for cf in compressed_files:
            self.log(f"   - {cf.name}")
        self.log("")
        
        extracted_folders = []
        for archive in compressed_files:
            success, folder = self.extract_archive(archive)
            if success and folder:
                extracted_folders.append(folder)
                
                # Delete archive if option is enabled
                if self.delete_archives_after_extract.get():
                    try:
                        archive.unlink()
                        self.log(f"  🗑️  Deleted archive: {archive.name}")
                    except Exception as e:
                        self.log(f"  ⚠️  Could not delete archive: {e}")
        
        return extracted_folders

    def find_game_files(self, directory, recursive=True):
        """Find all supported game descriptor files (.cue and optionally .iso, .nes, .sfc, .smc, .snes, .n64, .z64, .v64)"""
        path = Path(directory)
        files = set()
        if recursive:
            if self.process_ps1_cues.get():
                files.update(path.rglob("*.cue"))
                if self.ps1_output_format == 'ISO':
                    files.update(path.rglob("*.chd"))
            if self.process_ps2_cues.get():
                files.update(path.rglob("*.cue"))
                if self.ps2_output_format == 'ISO':
                    files.update(path.rglob("*.chd"))
            if self.process_ps2_isos.get():
                files.update(path.rglob("*.iso"))
                if self.ps2_output_format == 'CHD':
                    # Already handled by .iso glob, but if we wanted CHD->ISO we'd need more
                    pass
            if self.process_psp_isos.get():
                files.update(path.rglob("*.iso"))
            if self.process_ps3_isos.get():
                files.update(path.rglob("*.iso"))
            if self.process_xbox_isos.get():
                files.update(path.rglob("*.iso"))
                files.update(path.rglob("*.xiso"))
            if self.process_nes_roms.get():
                files.update(path.rglob("*.nes"))
            if self.process_snes_roms.get():
                files.update(path.rglob("*.sfc"))
                files.update(path.rglob("*.smc"))
                files.update(path.rglob("*.snes"))
            if self.process_n64_roms.get():
                files.update(path.rglob("*.n64"))
                files.update(path.rglob("*.z64"))
                files.update(path.rglob("*.v64"))
        else:
            if self.process_ps1_cues.get():
                files.update(path.glob("*.cue"))
            if self.process_ps2_cues.get():
                files.update(path.glob("*.cue"))
            if self.process_ps2_isos.get():
                files.update(path.glob("*.iso"))
            if self.process_psp_isos.get():
                files.update(path.glob("*.iso"))
            if self.process_ps3_isos.get():
                files.update(path.glob("*.iso"))
            if self.process_xbox_isos.get():
                files.update(path.glob("*.iso"))
                files.update(path.glob("*.xiso"))
            if self.process_nes_roms.get():
                files.update(path.glob("*.nes"))
            if self.process_snes_roms.get():
                files.update(path.glob("*.sfc"))
                files.update(path.glob("*.smc"))
                files.update(path.glob("*.snes"))
            if self.process_n64_roms.get():
                files.update(path.glob("*.n64"))
                files.update(path.glob("*.z64"))
                files.update(path.glob("*.v64"))
        # Sort for stable processing order
        return sorted(files)
    
    def parse_cue_file(self, cue_path, auto_repair=True):
        """Parse CUE file to find associated BIN files.
        
        If auto_repair is True, will attempt to fix CUE references to BIN files
        that have been renamed (e.g., locale tags removed).
        """
        bin_files = []
        cue_dir = cue_path.parent
        cue_needs_repair = False
        repairs = {}  # old_name -> new_name
        
        try:
            with open(cue_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
                
            # Find FILE entries in CUE
            file_pattern = re.compile(r'FILE\s+"([^"]+)"\s+BINARY', re.IGNORECASE)
            matches = file_pattern.findall(content)
            
            for match in matches:
                bin_path = cue_dir / match
                if bin_path.exists():
                    bin_files.append(bin_path)
                else:
                    # Try to find BIN file with cleaned name
                    found_bin = None
                    
                    # Get base name without extension
                    bin_stem = Path(match).stem
                    
                    # Try finding by partial match - look for BINs that might be the cleaned version
                    # Extract track info if present
                    track_match = re.search(r'(Track\s*\d+|\(Track\s*\d+\))', match, re.IGNORECASE)
                    track_info = track_match.group(1) if track_match else None
                    
                    # Get the CUE file's base name (likely already cleaned)
                    cue_base = cue_path.stem
                    
                    # Look for BIN files that match the CUE base name
                    for existing_bin in cue_dir.glob("*.bin"):
                        existing_stem = existing_bin.stem
                        
                        # Check if this BIN matches the CUE name pattern
                        if track_info:
                            # Multi-track: look for "CueName (Track N).bin" or "CueName Track N.bin"
                            expected_patterns = [
                                f"{cue_base} ({track_info})",
                                f"{cue_base} {track_info}",
                                f"{cue_base}({track_info})",
                            ]
                            for pattern in expected_patterns:
                                if existing_stem.lower() == pattern.lower():
                                    found_bin = existing_bin
                                    break
                        else:
                            # Single track: look for "CueName.bin"
                            if existing_stem.lower() == cue_base.lower():
                                found_bin = existing_bin
                                break
                    
                    if found_bin:
                        bin_files.append(found_bin)
                        repairs[match] = found_bin.name
                        cue_needs_repair = True
                        self.log(f"  INFO: Found renamed BIN: {match} → {found_bin.name}")
                    else:
                        self.log(f"  WARNING: Referenced BIN file not found: {match}")
            
            # Auto-repair CUE file if we found renamed BINs
            if auto_repair and cue_needs_repair and repairs:
                try:
                    new_content = content
                    for old_name, new_name in repairs.items():
                        new_content = new_content.replace(f'"{old_name}"', f'"{new_name}"')
                    
                    with open(cue_path, 'w', encoding='utf-8') as f:
                        f.write(new_content)
                    self.log(f"  ✅ Auto-repaired CUE file references")
                except Exception as e:
                    self.log(f"  WARNING: Could not auto-repair CUE file: {e}")
        
        except Exception as e:
            self.log(f"  ERROR parsing CUE file: {e}")
        
        return bin_files
    
    def parse_nes_header(self, nes_path):
        """Parse NES ROM header (iNES/NES 2.0 format) to extract game information.
        
        Returns:
            dict with ROM info: {
                'format': 'iNES' or 'NES 2.0' or 'Unknown',
                'prg_rom_size': PRG ROM size in KB,
                'chr_rom_size': CHR ROM size in KB,
                'mapper': Mapper number,
                'mirroring': 'Horizontal', 'Vertical', or 'Four-screen',
                'battery': True/False (has battery-backed RAM),
                'trainer': True/False (has 512-byte trainer),
                'region': 'NTSC' or 'PAL' or 'Dual' (NES 2.0 only),
                'valid': True/False
            }
        """
        rom_info = {
            'format': 'Unknown',
            'prg_rom_size': 0,
            'chr_rom_size': 0,
            'mapper': 0,
            'mirroring': 'Unknown',
            'battery': False,
            'trainer': False,
            'region': 'Unknown',
            'valid': False
        }
        
        try:
            with open(nes_path, 'rb') as f:
                header = f.read(16)
                
            # Check for iNES header signature "NES\x1A"
            if len(header) < 16 or header[0:4] != b'NES\x1a':
                self.log(f"  ⚠️  Invalid NES ROM header (missing NES signature): {nes_path.name}")
                return rom_info
            
            rom_info['valid'] = True
            
            # Parse header bytes
            prg_rom_units = header[4]  # 16 KB units
            chr_rom_units = header[5]  # 8 KB units
            flags6 = header[6]
            flags7 = header[7]
            flags9 = header[9] if len(header) > 9 else 0
            
            # Determine format (iNES vs NES 2.0)
            if (flags7 & 0x0C) == 0x08:
                rom_info['format'] = 'NES 2.0'
                # NES 2.0 extended fields
                # Mapper number is composed of:
                # - Lower nibble: bits 4-7 of flags6
                # - Middle nibble: bits 4-7 of flags7  
                # - Upper nibble: bits 0-3 of byte 8 (if available)
                mapper_low = (flags6 & 0xF0) >> 4
                mapper_mid = flags7 & 0xF0
                mapper_high = 0
                if len(header) > 8:
                    mapper_high = (header[8] & 0x0F) << 8
                rom_info['mapper'] = mapper_high | mapper_mid | mapper_low
                
                # NES 2.0 region
                if len(header) > 12:
                    flags12 = header[12]
                    region_bits = flags12 & 0x03
                    if region_bits == 0:
                        rom_info['region'] = 'NTSC'
                    elif region_bits == 1:
                        rom_info['region'] = 'PAL'
                    else:
                        rom_info['region'] = 'Dual'
            else:
                rom_info['format'] = 'iNES'
                # Standard iNES mapper
                mapper_low = (flags6 & 0xF0) >> 4
                mapper_high = flags7 & 0xF0
                rom_info['mapper'] = mapper_high | mapper_low
                
                # Try to guess region from ROM size/mapper (not reliable)
                rom_info['region'] = 'NTSC'  # Default to NTSC for iNES 1.0
            
            # PRG ROM size (16 KB units)
            rom_info['prg_rom_size'] = prg_rom_units * 16
            
            # CHR ROM size (8 KB units)
            rom_info['chr_rom_size'] = chr_rom_units * 8
            
            # Mirroring
            if flags6 & 0x08:
                rom_info['mirroring'] = 'Four-screen'
            elif flags6 & 0x01:
                rom_info['mirroring'] = 'Vertical'
            else:
                rom_info['mirroring'] = 'Horizontal'
            
            # Battery-backed RAM
            rom_info['battery'] = bool(flags6 & 0x02)
            
            # Trainer
            rom_info['trainer'] = bool(flags6 & 0x04)
            
        except Exception as e:
            self.log(f"  ERROR parsing NES ROM: {e}")
            rom_info['valid'] = False
        
        return rom_info
    
    def parse_snes_header(self, snes_path):
        """Parse SNES ROM header to extract game information.
        
        SNES ROMs can have different formats:
        - With or without 512-byte SMC copier header
        - LoROM or HiROM memory mapping
        
        Returns:
            dict with ROM info: {
                'title': Game title string,
                'rom_size': ROM size in KB,
                'rom_makeup': 'LoROM', 'HiROM', 'ExLoROM', 'ExHiROM', etc.,
                'rom_type': Cart type description,
                'region': 'Japan', 'USA', 'Europe', 'Other',
                'version': Version number,
                'has_smc_header': True/False (512-byte copier header),
                'checksum_valid': True/False,
                'valid': True/False
            }
        """
        rom_info = {
            'title': 'Unknown',
            'rom_size': 0,
            'rom_makeup': 'Unknown',
            'rom_type': 'Unknown',
            'region': 'Unknown',
            'version': 0,
            'has_smc_header': False,
            'checksum_valid': False,
            'valid': False
        }
        
        rom_type_map = {
            0x00: 'ROM',
            0x01: 'ROM + RAM',
            0x02: 'ROM + SRAM',
            0x03: 'ROM + DSP1',
            0x04: 'ROM + DSP1 + RAM',
            0x05: 'ROM + DSP1 + SRAM',
            0x13: 'ROM + SuperFX',
            0x14: 'ROM + SuperFX + RAM',
            0x15: 'ROM + SuperFX + SRAM',
            0x1A: 'ROM + SuperFX + SRAM (alt)',
            0x33: 'ROM + SA-1',
            0x34: 'ROM + SA-1 + RAM',
            0x35: 'ROM + SA-1 + SRAM',
        }
        
        region_map = {
            0x00: 'Japan',
            0x01: 'USA',
            0x02: 'Europe',
            0x03: 'Sweden',
            0x04: 'Finland',
            0x05: 'Denmark',
            0x06: 'France',
            0x07: 'Netherlands',
            0x08: 'Spain',
            0x09: 'Germany',
            0x0A: 'Italy',
            0x0B: 'China',
            0x0C: 'Indonesia',
            0x0D: 'Korea',
            0x0E: 'Global',
            0x0F: 'Canada',
            0x10: 'Brazil',
            0x11: 'Australia',
        }
        
        try:
            with open(snes_path, 'rb') as f:
                rom_data = f.read()
            
            rom_size_actual = len(rom_data)
            rom_info['rom_size'] = rom_size_actual // 1024
            
            # Check for 512-byte SMC header
            has_header = (rom_size_actual % 1024) == 512
            header_offset = 512 if has_header else 0
            rom_info['has_smc_header'] = has_header
            
            # Try to find valid header at LoROM and HiROM positions
            lorom_offset = 0x7FB0 + header_offset
            hirom_offset = 0xFFB0 + header_offset
            
            def read_header_at(offset):
                """Read and validate header at given offset"""
                if offset + 0x30 > len(rom_data):
                    return None
                
                header = rom_data[offset:offset + 0x30]
                
                # Extract fields
                title_bytes = header[0x00:0x15]
                rom_makeup = header[0x15]
                rom_type_byte = header[0x16]
                rom_size_byte = header[0x17]
                region_byte = header[0x19]
                version = header[0x1B]
                checksum_comp = (header[0x1D] << 8) | header[0x1C]
                checksum = (header[0x1F] << 8) | header[0x1E]
                
                # Validate: checksum + complement should equal 0xFFFF
                checksum_valid = (checksum + checksum_comp) == 0xFFFF
                
                # Validate: title should be printable ASCII/extended ASCII
                try:
                    title = title_bytes.decode('ascii', errors='ignore').strip('\x00 ')
                    title_valid = all(32 <= ord(c) < 127 or c == '\x00' for c in title_bytes.decode('ascii', errors='replace'))
                except:
                    title_valid = False
                    title = ''
                
                # Validate: ROM size makes sense
                if rom_size_byte < 20:
                    calculated_size = 1 << rom_size_byte  # 2^rom_size_byte KB
                    size_valid = abs(calculated_size - rom_info['rom_size']) < calculated_size * 0.1
                else:
                    size_valid = False
                
                # Score the header validity
                score = 0
                if checksum_valid:
                    score += 3
                if title_valid and len(title) > 0:
                    score += 2
                if size_valid:
                    score += 2
                if rom_makeup in [0x20, 0x21, 0x22, 0x23, 0x25, 0x30, 0x31, 0x32, 0x35]:
                    score += 1
                
                return {
                    'score': score,
                    'title': title,
                    'rom_makeup': rom_makeup,
                    'rom_type': rom_type_byte,
                    'rom_size_byte': rom_size_byte,
                    'region': region_byte,
                    'version': version,
                    'checksum_valid': checksum_valid
                }
            
            # Try both positions and pick the best match
            lorom_header = read_header_at(lorom_offset)
            hirom_header = read_header_at(hirom_offset)
            
            best_header = None
            rom_mapping = 'Unknown'
            
            if lorom_header and hirom_header:
                # Both found, use highest scoring one
                if lorom_header['score'] >= hirom_header['score']:
                    best_header = lorom_header
                    rom_mapping = 'LoROM'
                else:
                    best_header = hirom_header
                    rom_mapping = 'HiROM'
            elif lorom_header:
                best_header = lorom_header
                rom_mapping = 'LoROM'
            elif hirom_header:
                best_header = hirom_header
                rom_mapping = 'HiROM'
            
            if best_header and best_header['score'] >= 3:
                rom_info['valid'] = True
                rom_info['title'] = best_header['title']
                rom_info['checksum_valid'] = best_header['checksum_valid']
                rom_info['version'] = best_header['version']
                
                # Decode ROM makeup
                makeup = best_header['rom_makeup']
                if makeup == 0x20:
                    rom_info['rom_makeup'] = 'LoROM'
                elif makeup == 0x21:
                    rom_info['rom_makeup'] = 'HiROM'
                elif makeup == 0x22:
                    rom_info['rom_makeup'] = 'LoROM + S-DD1'
                elif makeup == 0x23:
                    rom_info['rom_makeup'] = 'LoROM + SA-1'
                elif makeup == 0x25:
                    rom_info['rom_makeup'] = 'ExHiROM'
                elif makeup == 0x30:
                    rom_info['rom_makeup'] = 'LoROM + FastROM'
                elif makeup == 0x31:
                    rom_info['rom_makeup'] = 'HiROM + FastROM'
                elif makeup == 0x32:
                    rom_info['rom_makeup'] = 'ExLoROM'
                elif makeup == 0x35:
                    rom_info['rom_makeup'] = 'ExHiROM + FastROM'
                else:
                    rom_info['rom_makeup'] = f'{rom_mapping} (0x{makeup:02X})'
                
                # Decode ROM type
                rom_info['rom_type'] = rom_type_map.get(best_header['rom_type'], f'Type 0x{best_header["rom_type"]:02X}')
                
                # Decode region
                rom_info['region'] = region_map.get(best_header['region'], f'Unknown (0x{best_header["region"]:02X})')
            
        except Exception as e:
            self.log(f"  ERROR parsing SNES ROM: {e}")
            rom_info['valid'] = False
        
        return rom_info
    
    def parse_n64_header(self, n64_path):
        """Parse N64 ROM header to extract game information.
        
        N64 ROMs can have different byte orderings:
        - Big-endian (.z64): Native format, 0x80371240
        - Little-endian (.n64): Byte-swapped, 0x40123780
        - Byte-swapped (.v64): Word-swapped, 0x37804012
        
        Returns:
            dict with ROM info: {
                'title': Game title string,
                'game_code': Game ID code (4 chars),
                'region': 'NTSC', 'PAL', 'Gateway 64', etc.,
                'rom_size': ROM size in MB,
                'byte_order': 'Big-endian', 'Little-endian', 'Byte-swapped',
                'crc1': CRC1 checksum,
                'crc2': CRC2 checksum,
                'valid': True/False
            }
        """
        rom_info = {
            'title': 'Unknown',
            'game_code': 'Unknown',
            'region': 'Unknown',
            'rom_size': 0,
            'byte_order': 'Unknown',
            'crc1': 0,
            'crc2': 0,
            'valid': False
        }
        
        region_codes = {
            0x41: 'Japan/USA (NTSC)',  # 'A'
            0x42: 'Brazil (NTSC)',  # 'B'
            0x43: 'China (PAL)',  # 'C'
            0x44: 'Germany (PAL)',  # 'D'
            0x45: 'USA (NTSC)',  # 'E'
            0x46: 'France (PAL)',  # 'F'
            0x47: 'Gateway 64 (NTSC)',  # 'G'
            0x48: 'Netherlands (PAL)',  # 'H'
            0x49: 'Italy (PAL)',  # 'I'
            0x4A: 'Japan (NTSC)',  # 'J'
            0x4B: 'Korea (NTSC)',  # 'K'
            0x4C: 'Gateway 64 (PAL)',  # 'L'
            0x4E: 'Canada (NTSC)',  # 'N'
            0x50: 'Europe (PAL)',  # 'P'
            0x53: 'Spain (PAL)',  # 'S'
            0x55: 'Australia (PAL)',  # 'U'
            0x57: 'Scandinavia (PAL)',  # 'W'
            0x58: 'Europe (PAL)',  # 'X'
            0x59: 'Europe (PAL)',  # 'Y'
        }
        
        try:
            with open(n64_path, 'rb') as f:
                header_data = f.read(4096)  # Read first 4KB
            
            if len(header_data) < 64:
                return rom_info
            
            rom_size_bytes = n64_path.stat().st_size
            rom_info['rom_size'] = rom_size_bytes / (1024 * 1024)  # Convert to MB
            
            # Detect byte order from first 4 bytes (PI_BSB_DOM1 register)
            first_dword = header_data[0:4]
            
            if first_dword == b'\x80\x37\x12\x40':
                # Big-endian (.z64) - native N64 format
                rom_info['byte_order'] = 'Big-endian (.z64)'
                data = header_data
            elif first_dword == b'\x37\x80\x40\x12':
                # Byte-swapped (.v64)
                rom_info['byte_order'] = 'Byte-swapped (.v64)'
                # Convert to big-endian for easier parsing
                data = bytearray(len(header_data))
                for i in range(0, len(header_data), 2):
                    data[i] = header_data[i + 1]
                    data[i + 1] = header_data[i]
            elif first_dword == b'\x40\x12\x37\x80':
                # Little-endian (.n64)
                rom_info['byte_order'] = 'Little-endian (.n64)'
                # Convert to big-endian for easier parsing
                data = bytearray(len(header_data))
                for i in range(0, len(header_data), 4):
                    data[i] = header_data[i + 3]
                    data[i + 1] = header_data[i + 2]
                    data[i + 2] = header_data[i + 1]
                    data[i + 3] = header_data[i]
            else:
                # Unknown format
                return rom_info
            
            rom_info['valid'] = True
            
            # Extract CRC values (bytes 0x10-0x17, big-endian)
            rom_info['crc1'] = (data[0x10] << 24) | (data[0x11] << 16) | (data[0x12] << 8) | data[0x13]
            rom_info['crc2'] = (data[0x14] << 24) | (data[0x15] << 16) | (data[0x16] << 8) | data[0x17]
            
            # Extract game title (bytes 0x20-0x33, 20 bytes, ASCII)
            title_bytes = data[0x20:0x34]
            try:
                title = title_bytes.decode('ascii', errors='ignore').strip('\x00 ')
                if title:
                    rom_info['title'] = title
            except:
                pass
            
            # Extract game code (bytes 0x3B-0x3E, 4 bytes)
            try:
                game_code = data[0x3B:0x3F].decode('ascii', errors='ignore')
                if game_code.isprintable():
                    rom_info['game_code'] = game_code
            except:
                pass
            
            # Extract region code (byte 0x3E)
            region_byte = data[0x3E]
            rom_info['region'] = region_codes.get(region_byte, f'Unknown (0x{region_byte:02X})')
            
        except Exception as e:
            self.log(f"  ERROR parsing N64 ROM: {e}")
            rom_info['valid'] = False
        
        return rom_info
    
    def scan_directory(self):
        """Scan directory for CUE files"""
        if not self.source_dir or not os.path.isdir(self.source_dir):
            messagebox.showwarning("Warning", "Please select a valid directory")
            return
        
        self.log("\n" + "="*60)
        self.log("SCANNING FOR GAME FILES...")
        self.log("="*60)
        
        # Keep UI responsive during scan
        self.keep_ui_responsive()
        
        # First, check for compressed files
        compressed_count = 0
        compressed_size = 0
        if self.extract_compressed.get():
            compressed_files = self.find_compressed_files(self.source_dir, self.recursive.get())
            if compressed_files:
                compressed_count = len(compressed_files)
                self.log(f"\n📦 Found {compressed_count} compressed file(s):")
                for cf in compressed_files:
                    size_mb = cf.stat().st_size / (1024 * 1024)
                    compressed_size += cf.stat().st_size
                    self.log(f"   - {cf.name} ({size_mb:.1f} MB)")
                self.log("\n⚠️  Compressed files will be extracted when you click 'Start Conversion'.")
        
        # Keep UI responsive
        self.keep_ui_responsive()

        game_files = self.find_game_files(self.source_dir, self.recursive.get())

        if not game_files and compressed_count == 0:
            self.log("No game descriptor files found (.cue/.iso) and no compressed files!")
            self.status_label.config(text="No game files found")
            self.convert_button.config(state="disabled")
            return
        
        # If we have compressed files but no game files, still allow conversion
        # (extraction will reveal game files)
        if not game_files and compressed_count > 0:
            self.log(f"\nNo extracted game files found yet, but {compressed_count} compressed file(s) will be extracted.")
            compressed_size_mb = compressed_size / (1024 * 1024)
            self.status_label.config(text=f"Found {compressed_count} compressed file(s) ({compressed_size_mb:.1f} MB) - Ready to extract & convert")
            self.convert_button.config(state="normal")
            return

        ps1_count = 0
        ps2_count = 0
        psp_count = 0
        nes_count = 0
        snes_count = 0
        n64_count = 0
        total_size = 0

        self.log(f"\nFound {len(game_files)} game descriptor file(s):\n")

        for game_file in game_files:
            game_size = 0
            if game_file.suffix.lower() == '.cue':
                # Decide label based on toggles
                if self.process_ps1_cues.get() and not self.process_ps2_cues.get():
                    cue_label = 'PS1'
                elif self.process_ps2_cues.get() and not self.process_ps1_cues.get():
                    cue_label = 'PS2'
                else:
                    cue_label = 'CUE'
                ps1_count += 1  # keep legacy counters but treat as cue/CD
                self.log(f"📀 [{cue_label}] {game_file.name}")
                self.log(f"   Path: {game_file}")
                bin_files = self.parse_cue_file(game_file)
                if bin_files:
                    for bin_file in bin_files:
                        size_mb = bin_file.stat().st_size / (1024 * 1024)
                        game_size += bin_file.stat().st_size
                        self.log(f"   └─ {bin_file.name} ({size_mb:.1f} MB)")
                game_size += game_file.stat().st_size  # CUE size (small)
            elif game_file.suffix.lower() == '.iso':
                iso_size = game_file.stat().st_size
                size_gb = iso_size / (1024 * 1024 * 1024)
                size_mb = iso_size / (1024 * 1024)
                system_guess = self.detect_iso_system(game_file.name, iso_size)
                iso_label = 'PS2'
                if system_guess == 'PSP':
                    iso_label = 'PSP'
                    psp_count += 1
                else:
                    ps2_count += 1
                self.log(f"💿 [{iso_label}] {game_file.name}")
                self.log(f"   Path: {game_file}")
                if size_gb >= 1:
                    self.log(f"   └─ ISO size: {size_gb:.2f} GB")
                else:
                    self.log(f"   └─ ISO size: {size_mb:.1f} MB")
                game_size += iso_size
            elif game_file.suffix.lower() == '.nes':
                nes_count += 1
                rom_size = game_file.stat().st_size
                size_kb = rom_size / 1024
                self.log(f"🎮 [NES] {game_file.name}")
                self.log(f"   Path: {game_file}")
                self.log(f"   └─ ROM size: {size_kb:.1f} KB")
                
                # Parse NES header for additional info
                nes_info = self.parse_nes_header(game_file)
                if nes_info['valid']:
                    self.log(f"   └─ Format: {nes_info['format']}")
                    self.log(f"   └─ Mapper: {nes_info['mapper']}")
                    self.log(f"   └─ PRG ROM: {nes_info['prg_rom_size']} KB | CHR ROM: {nes_info['chr_rom_size']} KB")
                    self.log(f"   └─ Mirroring: {nes_info['mirroring']}")
                    if nes_info['battery']:
                        self.log(f"   └─ Battery: Yes (save data)")
                    if nes_info['trainer']:
                        self.log(f"   └─ Trainer: Yes")
                    if nes_info['region'] != 'Unknown':
                        self.log(f"   └─ Region: {nes_info['region']}")
                
                game_size += rom_size
            elif game_file.suffix.lower() in ['.sfc', '.smc', '.snes']:
                snes_count += 1
                rom_size = game_file.stat().st_size
                size_kb = rom_size / 1024
                self.log(f"🎮 [SNES] {game_file.name}")
                self.log(f"   Path: {game_file}")
                self.log(f"   └─ ROM size: {size_kb:.1f} KB")
                
                # Parse SNES header for additional info
                snes_info = self.parse_snes_header(game_file)
                if snes_info['valid']:
                    self.log(f"   └─ Title: {snes_info['title']}")
                    self.log(f"   └─ Mapping: {snes_info['rom_makeup']}")
                    self.log(f"   └─ Cartridge: {snes_info['rom_type']}")
                    self.log(f"   └─ Region: {snes_info['region']}")
                    if snes_info['has_smc_header']:
                        self.log(f"   └─ SMC Header: Yes (512 bytes)")
                    self.log(f"   └─ Checksum: {'Valid' if snes_info['checksum_valid'] else 'Invalid'}")
                    if snes_info['version'] > 0:
                        self.log(f"   └─ Version: 1.{snes_info['version']}")
                
                game_size += rom_size
            elif game_file.suffix.lower() in ['.n64', '.z64', '.v64']:
                n64_count += 1
                rom_size = game_file.stat().st_size
                size_mb = rom_size / (1024 * 1024)
                self.log(f"🎮 [N64] {game_file.name}")
                self.log(f"   Path: {game_file}")
                self.log(f"   └─ ROM size: {size_mb:.1f} MB")
                
                # Parse N64 header for additional info
                n64_info = self.parse_n64_header(game_file)
                if n64_info['valid']:
                    self.log(f"   └─ Title: {n64_info['title']}")
                    self.log(f"   └─ Game Code: {n64_info['game_code']}")
                    self.log(f"   └─ Region: {n64_info['region']}")
                    self.log(f"   └─ Format: {n64_info['byte_order']}")
                    self.log(f"   └─ CRC1: 0x{n64_info['crc1']:08X}")
                    self.log(f"   └─ CRC2: 0x{n64_info['crc2']:08X}")
                
                game_size += rom_size
            total_size += game_size
            self.log("")

        total_size_mb = total_size / (1024 * 1024)
        total_size_gb = total_size / (1024 * 1024 * 1024)

        # Build totals string
        totals_parts = []
        if ps1_count > 0:
            totals_parts.append(f"CUE: {ps1_count}")
        if ps2_count > 0:
            totals_parts.append(f"PS2: {ps2_count}")
        if psp_count > 0:
            totals_parts.append(f"PSP: {psp_count}")
        if nes_count > 0:
            totals_parts.append(f"NES: {nes_count}")
        if snes_count > 0:
            totals_parts.append(f"SNES: {snes_count}")
        if n64_count > 0:
            totals_parts.append(f"N64: {n64_count}")
        self.log(f"Totals: {' | '.join(totals_parts) if totals_parts else 'None'}  Combined: {len(game_files)}")
        if compressed_count > 0:
            self.log(f"📦 Plus {compressed_count} compressed file(s) to extract")
        if total_size_gb >= 1:
            self.log(f"💾 Current total size: {total_size_gb:.2f} GB ({total_size_mb:.1f} MB)")
        else:
            self.log(f"💾 Current total size: {total_size_mb:.1f} MB")

        # Build status text
        status_parts = []
        if ps1_count > 0:
            status_parts.append(f"CUE:{ps1_count}")
        if ps2_count > 0:
            status_parts.append(f"PS2:{ps2_count}")
        if psp_count > 0:
            status_parts.append(f"PSP:{psp_count}")
        if nes_count > 0:
            status_parts.append(f"NES:{nes_count}")
        if snes_count > 0:
            status_parts.append(f"SNES:{snes_count}")
        if n64_count > 0:
            status_parts.append(f"N64:{n64_count}")
        status_text = f"Found {' '.join(status_parts) if status_parts else 'none'}"
        if compressed_count > 0:
            status_text += f" + {compressed_count} archives"
        status_text += f" | Size: {total_size_gb:.2f} GB"
        self.status_label.config(text=status_text)
        self.convert_button.config(state="normal")
    
    def convert_game_with_factory(self, path):
        """Convert a game file using ConverterFactory (Phase 3 Week 4 refactored version)"""
        if not self.converter_factory:
            # Fallback to original implementation
            return self._convert_game_original(path)
        
        path = Path(path)
        ext = path.suffix.lower()
        
        # Determine output format based on file type
        if ext == '.cue':
            output_format = self.ps1_output_format.upper()
        elif ext == '.chd':
            # For CHD extraction, use the target format from preferences
            is_ps1 = self.process_ps1_cues.get()
            output_format = self.ps1_output_format.upper() if is_ps1 else self.ps2_output_format.upper()
        elif ext == '.iso':
            # Determine system and format
            iso_size = path.stat().st_size
            system_guess = self.detect_iso_system(path.name, iso_size, full_path=path)
            
            if system_guess == 'PlayStation 3':
                output_format = 'DECRYPT'
            elif system_guess == 'Xbox':
                output_format = 'EXTRACT'
            elif iso_size < 1.4e9:  # PSP detection
                output_format = self.psp_output_format.upper()
            else:
                output_format = self.ps2_output_format.upper()
        else:
            # Other formats use factory's detection
            output_format = 'AUTO'
        
        # Prepare kwargs for factory
        kwargs = {
            'max_processors': self.chdman_max_processors,
            'max_threads': self.maxcso_threads,
            'treat_as_psp': self.process_psp_isos.get() and not self.process_ps2_isos.get(),
            'system_guess': self.detect_iso_system(path.name, path.stat().st_size if path.exists() else 0, full_path=path) if ext == '.iso' else None,
        }
        
        # Use factory to convert
        try:
            result = self.converter_factory.convert(path, output_format, **kwargs)
            
            # Track statistics
            if result and path.exists():
                original_size = path.stat().st_size
                self.total_original_size += original_size
                self.completed_files.add(str(path))
                self.save_progress(self.source_dir)
            
            return result
        except Exception as e:
            self.log(f"  ❌ Exception in factory conversion: {e}")
            return False
    
    def convert_game(self, path):
        """Convert a game file to the selected output format
        
        Dispatches to ConverterFactory if available (Phase 3 Week 4),
        otherwise uses original implementation for compatibility.
        """
        # Try factory-based conversion first if available
        if self.converter_factory and CONVERTER_FACTORY_AVAILABLE:
            return self.convert_game_with_factory(path)
        
        # Fallback to original implementation
        return self._convert_game_original(path)
    
    def _convert_game_original(self, path):
        """Original convert_game implementation (kept for compatibility)"""
        ext = path.suffix.lower()
        path = Path(path)

        if ext == '.cue':
            # Check system-specific output format
            is_ps1 = self.process_ps1_cues.get()
            is_ps2 = self.process_ps2_cues.get()
            
            # Use PS1 setting as default if both are enabled or neither (generic CUE)
            fmt = self.ps1_output_format.upper() if is_ps1 else self.ps2_output_format.upper()
            
            if fmt == 'CHD':
                output_path = path.with_suffix('.chd')
                if output_path.exists():
                    self.log(f"  ⚠️  CHD already exists, skipping: {output_path.name}")
                    return True
                cmd = [self.chdman_path, 'createcd', '-np', str(self.chdman_max_processors), '-i', str(path), '-o', str(output_path)]
                original_size = sum(f.stat().st_size for f in self.parse_cue_file(path)) + path.stat().st_size
                label = 'CD (CUE)'
                format_label = 'CHD'
            else:
                # ISO/BIN output - CUE/BIN is already the "uncompressed" format for CDs
                self.log(f"  ℹ️  ISO/BIN selected for {path.name} - keeping as-is.")
                return True
        elif ext == '.chd':
            # Handle CHD to ISO/BIN extraction (reverse conversion)
            is_ps1 = self.process_ps1_cues.get()
            is_ps2 = self.process_ps2_cues.get()
            
            # Determine target format
            fmt = self.ps1_output_format.upper() if is_ps1 else self.ps2_output_format.upper()
            
            if fmt != 'ISO':
                # Only process CHDs if output is set to ISO (extraction mode)
                return True
            
            # Detect if this is a CD or DVD CHD
            try:
                info_cmd = [self.chdman_path, 'info', '-i', str(path)]
                info_res = subprocess.run(info_cmd, capture_output=True, text=True, timeout=10)
                is_cd = "CDROM" in info_res.stdout or "createcd" in info_res.stdout
            except:
                is_cd = True # Default to CD if check fails
            
            if is_cd:
                # Extract to BIN/CUE
                output_path = path.with_suffix('.cue')
                if output_path.exists():
                    self.log(f"  ⚠️  CUE already exists, skipping extraction: {output_path.name}")
                    return True
                cmd = [self.chdman_path, 'extractcd', '-i', str(path), '-o', str(output_path)]
                label = 'CHD (CD)'
                format_label = 'BIN/CUE'
            else:
                # Extract to ISO
                output_path = path.with_suffix('.iso')
                if output_path.exists():
                    self.log(f"  ⚠️  ISO already exists, skipping extraction: {output_path.name}")
                    return True
                cmd = [self.chdman_path, 'extractdvd', '-i', str(path), '-o', str(output_path)]
                label = 'CHD (DVD)'
                format_label = 'ISO'
            
            original_size = path.stat().st_size
        elif ext == '.iso':
            # Determine if this is a PSP, PS2, PS3, or Xbox ISO based on detection and settings
            iso_size = path.stat().st_size
            system_guess = self.detect_iso_system(path.name, iso_size, full_path=path)
            
            # Respect user settings
            psp_enabled = self.process_psp_isos.get()
            ps2_enabled = self.process_ps2_isos.get()
            ps3_enabled = self.process_ps3_isos.get()
            xbox_enabled = self.process_xbox_isos.get()
            
            # PS3 Decryption
            if ps3_enabled and system_guess == 'PlayStation 3':
                if not self.ps3_dumper_path:
                    self.log(f"  ⚠️  PS3 Dumper not configured. Cannot process: {path.name}")
                    return False
                
                self.log(f"  🎮 PS3 Decrypt: Launching dumper for {path.name}...")
                try:
                    # Launch ps3-disc-dumper with the ISO path as argument
                    # It's a GUI tool, so it will pop up a window.
                    # We use Popen and don't wait if it's meant to be interactive, 
                    # but the user said "support for PS3 ISOs", and usually this implies
                    # being able to trigger the process.
                    cmd = [self.ps3_dumper_path, str(path)]
                    subprocess.Popen(cmd)
                    self.log(f"  ✅ PS3 Dumper launched. Please complete decryption in the dumper window.")
                    return True # Treat as success in terms of launching
                except Exception as e:
                    self.log(f"  ❌ Failed to launch PS3 Dumper: {e}")
                    return False

            # Determine system path
            if xbox_enabled and (system_guess == 'Xbox' or system_guess == 'Xbox 360'):
                # Xbox extraction
                return self.extract_xbox_iso(path, path.parent)
            
            # Determine which system to treat this as (PS2/PSP)
            if psp_enabled and not ps2_enabled:
                treat_psp = True
            elif ps2_enabled and not psp_enabled:
                treat_psp = False
            elif psp_enabled and ps2_enabled:
                if system_guess == 'PSP':
                    treat_psp = True
                else:
                    treat_psp = False
            else:
                # Fallback to Xbox if specifically selected or detected and others not available
                if xbox_enabled:
                    return self.extract_xbox_iso(path, path.parent)
                self.log(f"  ❌ No ISO processing enabled for: {path.name}")
                return False
            
            if treat_psp:
                # PSP ISO conversion (CSO/ZSO only)
                fmt = self.psp_output_format.upper()
                if fmt == 'CSO':
                    output_path = path.with_suffix('.cso')
                    if output_path.exists():
                        self.log(f"  ⚠️  CSO already exists, skipping: {output_path.name}")
                        return True
                    # Use dynamic thread count based on system specs
                    cmd = [self.maxcso_path, '--threads', str(self.maxcso_threads), str(path), '-o', str(output_path)]
                    format_label = 'CSO'
                elif fmt == 'ZSO':
                    output_path = path.with_suffix('.zso')
                    if output_path.exists():
                        self.log(f"  ⚠️  ZSO already exists, skipping: {output_path.name}")
                        return True
                    # Use dynamic thread count based on system specs
                    cmd = [self.maxcso_path, '--zso', '--threads', str(self.maxcso_threads), str(path), '-o', str(output_path)]
                    format_label = 'ZSO'
                else:
                    self.log(f"  ❌ Unsupported PSP format: {fmt}")
                    return False
                label = 'PSP'
            else:
                # PS2 ISO conversion (CHD/CSO/ZSO)
                fmt = self.ps2_output_format.upper()
                if fmt == 'CHD':
                    output_path = path.with_suffix('.chd')
                    if output_path.exists():
                        self.log(f"  ⚠️  CHD already exists, skipping: {output_path.name}")
                        return True
                    # Use --numprocessors to limit chdman RAM usage (each processor uses ~500MB-1GB)
                    cmd = [self.chdman_path, 'createdvd', '-np', str(self.chdman_max_processors), '-i', str(path), '-o', str(output_path)]
                    format_label = 'CHD'
                elif fmt == 'CSO':
                    output_path = path.with_suffix('.cso')
                    if output_path.exists():
                        self.log(f"  ⚠️  CSO already exists, skipping: {output_path.name}")
                        return True
                    # Use dynamic thread count based on system specs
                    cmd = [self.maxcso_path, '--threads', str(self.maxcso_threads), str(path), '-o', str(output_path)]
                    format_label = 'CSO'
                elif fmt == 'ZSO':
                    output_path = path.with_suffix('.zso')
                    if output_path.exists():
                        self.log(f"  ⚠️  ZSO already exists, skipping: {output_path.name}")
                        return True
                    # Use dynamic thread count based on system specs
                    cmd = [self.maxcso_path, '--ziso', '--threads', str(self.maxcso_threads), str(path), '-o', str(output_path)]
                    format_label = 'ZSO'
                else:
                    self.log(f"  ❌ Unsupported PS2 format: {fmt}")
                    return False
                label = 'PS2'

            original_size = iso_size
        else:
            self.log(f"  ❌ Unsupported file type: {path.name}")
            return False

        try:
            self.log(f"  Converting ({label} → {format_label}): {path.name} -> {output_path.name}")
            
            # Use idle process priority to minimize system impact
            creationflags = 0
            if sys.platform == 'win32':
                # IDLE_PRIORITY_CLASS = 0x00000040 (lowest priority)
                creationflags = 0x00000040
            
            # Use Popen for real-time progress monitoring
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=creationflags if sys.platform == 'win32' else 0
            )
            
            # Track progress for real-time display
            last_progress_update = time.time()
            progress_update_interval = 2.0  # Update every 2 seconds for stability
            start_time = time.time()
            last_written_size = 0
            stall_start_time = None  # Track when write speed dropped to 0
            last_shown_phase = None  # Avoid duplicate phase messages
            tool_progress = [None]  # Progress reported by chdman/maxcso (list for thread access)
            stderr_lines = []  # Collect stderr for error reporting
            
            # Background thread to read stderr and parse tool progress
            def read_stderr():
                try:
                    for line in process.stderr:
                        stderr_lines.append(line)
                        # Parse chdman progress (e.g., "Compressing, 45.3% complete...")
                        # or maxcso progress
                        match = re.search(r'(\d+\.?\d*)%', line)
                        if match:
                            tool_progress[0] = float(match.group(1))
                except Exception:
                    pass
            
            stderr_thread = threading.Thread(target=read_stderr, daemon=True)
            stderr_thread.start()
            
            # Monitor output file size for progress
            while process.poll() is None:
                time.sleep(0.5)  # Check every 0.5 seconds
                
                current_time = time.time()
                elapsed = current_time - start_time
                
                # Update progress display periodically
                if current_time - last_progress_update >= progress_update_interval:
                    try:
                        if output_path.exists():
                            current_size = output_path.stat().st_size
                            size_delta = current_size - last_written_size
                            time_delta = current_time - last_progress_update
                            write_speed_mb = (size_delta / time_delta) / (1024 * 1024) if time_delta > 0 else 0
                            last_written_size = current_size
                            
                            # Calculate progress percentage based on expected output size
                            # CHD typically compresses to 40-60% of original, estimate 50%
                            estimated_final = original_size * 0.5
                            progress_pct = min(99, (current_size / estimated_final) * 100) if estimated_final > 0 else 0
                            
                            # Use tool's reported progress if available (more accurate)
                            if tool_progress[0] is not None:
                                progress_pct = tool_progress[0]
                            
                            # Format sizes for display
                            if current_size >= 1024*1024*1024:
                                size_str = f"{current_size / (1024*1024*1024):.2f} GB"
                            else:
                                size_str = f"{current_size / (1024*1024):.1f} MB"
                            
                            # Detect stall phases and provide informative feedback
                            if write_speed_mb < 0.1:  # Essentially stalled
                                if stall_start_time is None:
                                    stall_start_time = current_time
                                stall_duration = current_time - stall_start_time
                                
                                # Determine likely phase based on progress and stall duration
                                if progress_pct >= 95:
                                    phase = "Finalizing/Verifying"
                                elif progress_pct >= 80:
                                    phase = "Final compression pass"
                                elif stall_duration > 30:
                                    phase = "Processing (CPU-intensive)"
                                else:
                                    phase = "Buffering"
                                
                                # Only show phase message if it changed or every 10 seconds
                                if phase != last_shown_phase or stall_duration % 10 < progress_update_interval:
                                    self.log(f"  ⏳ {phase}: {progress_pct:.0f}% | Written: {size_str} | Elapsed: {elapsed:.0f}s")
                                    last_shown_phase = phase
                            else:
                                # Normal progress - reset stall tracking
                                stall_start_time = None
                                last_shown_phase = None
                                self.log(f"  📊 Progress: {progress_pct:.0f}% | Written: {size_str} | Speed: {write_speed_mb:.1f} MB/s | Time: {elapsed:.0f}s")
                    except (OSError, FileNotFoundError):
                        pass  # File might not exist yet or be locked
                    
                    last_progress_update = current_time
            
            # Wait for stderr thread to finish and get final output
            stderr_thread.join(timeout=5.0)
            stdout, _ = process.communicate(timeout=60)
            stderr = ''.join(stderr_lines)
            returncode = process.returncode
            
            elapsed_total = time.time() - start_time

            if returncode == 0 and output_path.exists():
                new_size = output_path.stat().st_size
                savings = ((original_size - new_size) / original_size) * 100 if original_size > 0 else 0
                avg_speed = (new_size / elapsed_total) / (1024 * 1024) if elapsed_total > 0 else 0

                # Update totals
                self.total_original_size += original_size
                self.total_chd_size += new_size

                # Track completion for crash recovery
                self.completed_files.add(str(path))
                self.save_progress(self.source_dir)

                self.log(f"  ✅ Complete! Saved {savings:.1f}% space in {elapsed_total:.1f}s (avg: {avg_speed:.1f} MB/s)")
                if original_size >= 1024*1024*1024:
                    self.log(f"     Original: {original_size / (1024*1024*1024):.2f} GB → {format_label}: {new_size / (1024*1024*1024):.2f} GB")
                else:
                    self.log(f"     Original: {original_size / (1024*1024):.1f} MB → {format_label}: {new_size / (1024*1024):.1f} MB")
                return True
            else:
                error_text = stderr.strip() or stdout.strip()
                self.log(f"  ❌ Conversion failed: {error_text}")
                # Clean up partial output file if it exists
                if output_path.exists():
                    try:
                        output_path.unlink()
                    except Exception:
                        pass
                return False
        except subprocess.TimeoutExpired:
            self.log(f"  ❌ Timeout: Conversion took too long")
            # Try to kill the process
            try:
                process.kill()
            except Exception:
                pass
            return False
        except Exception as e:
            self.log(f"  ❌ Exception: {e}")
            return False
    
    def move_to_backup_folder(self, cue_path):
        """Move original CUE and BIN files to backup folder"""
        try:
            # Create backup folder in the same directory as the CUE file
            backup_dir = cue_path.parent / "original_backup"
            backup_dir.mkdir(exist_ok=True)
            
            bin_files = self.parse_cue_file(cue_path)
            
            # Move BIN files
            for bin_file in bin_files:
                if bin_file.exists():
                    dest = backup_dir / bin_file.name
                    # Handle duplicate names
                    counter = 1
                    while dest.exists():
                        dest = backup_dir / f"{bin_file.stem}_{counter}{bin_file.suffix}"
                        counter += 1
                    shutil.move(str(bin_file), str(dest))
                    self.log(f"  📦 Moved to backup: {bin_file.name}")
            
            # Move CUE file
            if cue_path.exists():
                dest = backup_dir / cue_path.name
                counter = 1
                while dest.exists():
                    dest = backup_dir / f"{cue_path.stem}_{counter}{cue_path.suffix}"
                    counter += 1
                shutil.move(str(cue_path), str(dest))
                self.log(f"  📦 Moved to backup: {cue_path.name}")
            
            return True
        except Exception as e:
            self.log(f"  ❌ Error moving files to backup: {e}")
            return False
    
    def delete_original_files(self, cue_path):
        """Delete original CUE and BIN files"""
        try:
            bin_files = self.parse_cue_file(cue_path)
            
            # Delete BIN files
            for bin_file in bin_files:
                if bin_file.exists():
                    bin_file.unlink()
                    self.log(f"  🗑️  Deleted: {bin_file.name}")
            
            # Delete CUE file
            if cue_path.exists():
                cue_path.unlink()
                self.log(f"  🗑️  Deleted: {cue_path.name}")
            
            return True
        except Exception as e:
            self.log(f"  ❌ Error deleting files: {e}")
            return False
    
    def process_single_file(self, cue_file, file_num, total):
        """Process a single CUE file (for parallel execution)"""
        if not self.is_converting:
            return None
        
        # Wait if memory pressure is too high (prevents system freeze)
        self._wait_for_memory_pressure()
        
        # Check available RAM before starting - wait if too low
        if PSUTIL_AVAILABLE:
            try:
                mem = psutil.virtual_memory()
                available_gb = mem.available / (1024 ** 3)
                # Need at least 2GB available RAM to start a new conversion
                while available_gb < 2.0 and self.is_converting:
                    self.log(f"  ⏳ Waiting for RAM (available: {available_gb:.1f}GB, need: 2GB)...")
                    gc.collect()
                    time.sleep(3)
                    mem = psutil.virtual_memory()
                    available_gb = mem.available / (1024 ** 3)
            except Exception:
                pass
        
        # Record start time for metrics
        with self.metrics_lock:
            self.file_start_times[cue_file] = time.time()
        self.log(f"\n[{file_num}/{total}] Processing: {cue_file.name}")
        
        success = self.convert_game(cue_file)
        
        # Force garbage collection after each conversion to free memory
        gc.collect()
        
        if success:
            if self.delete_originals.get():
                self.delete_original_files(cue_file)
            elif self.move_to_backup.get():
                self.move_to_backup_folder(cue_file)
        
        return success
    
    def _wait_for_memory_pressure(self, max_wait=120):
        """Wait if RAM usage is critically high to prevent system freeze.
        
        Uses a two-tier system:
        - ram_critical_percent (85%): Start waiting and warn user
        - ram_hard_limit_percent (90%): Absolute max, wait indefinitely until below this
        
        Args:
            max_wait: Maximum seconds to wait before proceeding (only for critical, not hard limit)
        """
        if not PSUTIL_AVAILABLE:
            return
        
        waited = 0
        wait_interval = 2.0  # Check every 2 seconds
        logged_warning = False
        logged_hard_limit = False
        
        while self.is_converting:
            try:
                mem = psutil.virtual_memory()
                
                # Hard limit - never exceed this, wait indefinitely
                if mem.percent >= self.ram_hard_limit_percent:
                    if not logged_hard_limit:
                        self.log(f"  🛑 HARD LIMIT - RAM at {mem.percent:.1f}% (limit: {self.ram_hard_limit_percent}%), must wait...")
                        logged_hard_limit = True
                    gc.collect()
                    time.sleep(wait_interval)
                    continue  # Keep waiting, don't increment waited counter for hard limit
                
                # Critical threshold - wait with timeout
                if mem.percent >= self.ram_critical_percent:
                    if waited >= max_wait:
                        self.log(f"  ⚠️  Max wait time reached at {mem.percent:.1f}%, proceeding carefully...")
                        break
                    
                    if not logged_warning:
                        self.log(f"  ⏸️  Pausing - RAM usage high ({mem.percent:.1f}% >= {self.ram_critical_percent}%), waiting...")
                        logged_warning = True
                    
                    gc.collect()
                    time.sleep(wait_interval)
                    waited += wait_interval
                    continue
                
                # Memory is acceptable
                if logged_warning or logged_hard_limit:
                    self.log(f"  ✅ Memory pressure relieved ({mem.percent:.1f}%), resuming...")
                break
                
            except Exception:
                break
        
        # Also check disk I/O throttling
        self._wait_for_disk_io_throttle()
    
    def _wait_for_disk_io_throttle(self, max_wait=30):
        """Wait if disk write rate is too high to prevent I/O saturation.
        
        Args:
            max_wait: Maximum seconds to wait before proceeding anyway
        """
        if not PSUTIL_AVAILABLE:
            return
        
        waited = 0
        wait_interval = 0.5  # Check every 0.5 seconds
        logged_warning = False
        
        try:
            # Get initial disk I/O reading
            io_start = psutil.disk_io_counters()
            if io_start is None:
                return
            
            while waited < max_wait and self.is_converting:
                time.sleep(wait_interval)
                waited += wait_interval
                
                io_now = psutil.disk_io_counters()
                if io_now is None:
                    return
                
                # Calculate write rate in MB/s
                write_bytes_delta = io_now.write_bytes - io_start.write_bytes
                write_rate_mb_s = (write_bytes_delta / wait_interval) / (1024 * 1024)
                
                # Update reference point for next check
                io_start = io_now
                waited = 0  # Reset wait counter since we got a new reading
                
                if write_rate_mb_s < self.disk_write_throttle_mb_s:
                    if logged_warning:
                        self.log(f"  ✅ Disk I/O normalized ({write_rate_mb_s:.1f} MB/s), resuming...")
                    return
                
                if not logged_warning:
                    self.log(f"  ⏸️  Throttling - Disk write rate high ({write_rate_mb_s:.1f} MB/s > {self.disk_write_throttle_mb_s} MB/s limit)...")
                    logged_warning = True
                    
        except Exception:
            return
        
        if logged_warning:
            self.log(f"  ⚠️  Disk throttle timeout, proceeding anyway...")
    
    def check_system_resources(self):
        """Check system resources and return recommended worker count"""
        if not PSUTIL_AVAILABLE:
            return self.cpu_cores
        
        try:
            cpu_percent = psutil.cpu_percent(interval=0.1)
            mem = psutil.virtual_memory()
            
            # If RAM is critically high, reduce workers significantly
            if mem.percent >= self.ram_threshold_percent:
                return max(1, self.cpu_cores // 2)
            
            # If CPU is critically high, reduce workers
            if cpu_percent >= self.cpu_threshold_percent:
                return max(1, self.cpu_cores // 2)
            
            # Check disk I/O - if write rate is high, reduce workers
            try:
                io = psutil.disk_io_counters()
                if io and hasattr(self, 'last_disk_write_bytes') and self.last_disk_write_bytes > 0:
                    write_delta = io.write_bytes - self.last_disk_write_bytes
                    # Estimate write rate over a short period
                    write_rate_mb_s = (write_delta / self.disk_io_check_interval) / (1024 * 1024)
                    if write_rate_mb_s > self.disk_write_throttle_mb_s:
                        return max(1, self.cpu_cores // 2)
            except Exception:
                pass
            
            # If moderate load, use 3/4 of available cores (use threshold - 5% for moderate)
            moderate_threshold = max(50, self.ram_threshold_percent - 10)
            if mem.percent >= moderate_threshold or cpu_percent >= 80:
                return max(1, int(self.cpu_cores * 0.75))
            
            # Normal operation - use all allocated cores
            return self.cpu_cores
            
        except Exception:
            return self.cpu_cores
    
    def _detect_optimal_workers(self):
        """Detect optimal number of concurrent conversions based on system specs.
        
        Heuristics:
        - 1 worker per 6GB of RAM (conversions are very memory-intensive)
        - Cap at (CPU cores - 2) to keep system responsive
        - Minimum of 1, maximum of 4 (diminishing returns beyond this)
        """
        total_cores = multiprocessing.cpu_count()
        max_by_cpu = max(1, total_cores - 2)  # Reserve 2 cores for system
        
        if PSUTIL_AVAILABLE:
            try:
                # Get available RAM (not total) in GB
                mem = psutil.virtual_memory()
                available_ram_gb = mem.available / (1024 ** 3)
                
                # 1 worker per 6GB available RAM (conservative for heavy conversions)
                max_by_ram = max(1, int(available_ram_gb / 6))
                
                # Use the more conservative of CPU or RAM limits
                optimal = min(max_by_cpu, max_by_ram)
                
                # Cap at 4 (more aggressive cap to prevent RAM issues)
                return min(optimal, 4)
            except Exception:
                pass
        
        # Fallback: very conservative defaults
        if total_cores >= 8:
            return 2
        else:
            return 1
    
    def _detect_chdman_processors(self):
        """Detect optimal number of processors for chdman based on available RAM.
        
        Each chdman processor uses approximately 500MB-1GB RAM.
        We want to balance speed vs memory usage.
        """
        total_cores = multiprocessing.cpu_count()
        
        if PSUTIL_AVAILABLE:
            try:
                mem = psutil.virtual_memory()
                total_ram_gb = mem.total / (1024 ** 3)
                available_ram_gb = mem.available / (1024 ** 3)
                
                # Calculate based on available RAM (1 processor per 1.5GB available)
                # This is more aggressive than before but still safe
                max_by_ram = max(1, int(available_ram_gb / 1.5))
                
                # Also cap by CPU cores (no benefit beyond core count)
                max_by_cpu = max(1, total_cores - 1)
                
                # Use minimum of RAM and CPU limits
                optimal = min(max_by_ram, max_by_cpu)
                
                # For systems with lots of RAM (16GB+), allow more processors
                if total_ram_gb >= 24:
                    # 24GB+ RAM: allow up to 8 processors
                    return min(optimal, 8)
                elif total_ram_gb >= 16:
                    # 16GB+ RAM: allow up to 6 processors
                    return min(optimal, 6)
                elif total_ram_gb >= 8:
                    # 8GB+ RAM: allow up to 4 processors
                    return min(optimal, 4)
                else:
                    # Less than 8GB: limit to 2
                    return min(optimal, 2)
            except Exception:
                pass
        
        # Fallback based on CPU count
        if total_cores >= 8:
            return 4
        elif total_cores >= 4:
            return 2
        else:
            return 1
    
    def _detect_maxcso_threads(self):
        """Detect optimal number of threads for maxcso based on system specs.
        
        maxcso is less memory-intensive than chdman, so we can be more aggressive.
        """
        total_cores = multiprocessing.cpu_count()
        
        if PSUTIL_AVAILABLE:
            try:
                mem = psutil.virtual_memory()
                total_ram_gb = mem.total / (1024 ** 3)
                
                # maxcso uses less RAM per thread (~200-300MB)
                # Allow more threads, cap at (cores - 1)
                max_threads = max(1, total_cores - 1)
                
                # Scale based on RAM
                if total_ram_gb >= 16:
                    return max_threads
                elif total_ram_gb >= 8:
                    return min(max_threads, 4)
                else:
                    return min(max_threads, 2)
            except Exception:
                pass
        
        # Fallback
        if total_cores >= 8:
            return 4
        elif total_cores >= 4:
            return 2
        else:
            return 1
    
    def conversion_thread(self):
        """Run conversion in separate thread with parallel processing"""
        
        extracted_folders = []
        extracted_archives = []
        
        # First, extract any compressed files if enabled
        if self.extract_compressed.get():
            self.log("\n" + "="*60)
            self.log("EXTRACTING COMPRESSED FILES...")
            self.log("="*60)
            
            extracted_folders = self.extract_all_archives(self.source_dir, self.recursive.get())
            # Track which archives were extracted
            if extracted_folders:
                extracted_archives = self.find_compressed_files(self.source_dir, self.recursive.get())
            
            if extracted_folders:
                self.log(f"\n✅ Extracted {len(extracted_folders)} archive(s)")
                self.log("Now scanning for game files in extracted folders...\n")
        
        game_files = self.find_game_files(self.source_dir, self.recursive.get())
        
        # Filter out already completed files (crash recovery)
        original_count = len(game_files)
        game_files = [f for f in game_files if str(f) not in self.completed_files]
        skipped_count = original_count - len(game_files)
        
        total = len(game_files)
        self.total_jobs = total
        self.completed_jobs = 0
        
        if skipped_count > 0:
            self.log(f"\n📂 RESUME MODE: Skipping {skipped_count} already completed file(s)")
        
        if total == 0:
            if skipped_count > 0:
                self.log("✅ All files already converted!")
                self.clear_progress()
            else:
                self.log("No game files to convert!")
            self.is_converting = False
            self.master.after(0, self.conversion_complete)
            return
        
        # Reset size tracking
        self.total_original_size = 0
        self.total_chd_size = 0
        
        self.log("\n" + "="*60)
        self.log("STARTING CONVERSION...")
        total_cores = multiprocessing.cpu_count()
        if PSUTIL_AVAILABLE:
            try:
                mem = psutil.virtual_memory()
                total_ram_gb = mem.total / (1024 ** 3)
                avail_ram_gb = mem.available / (1024 ** 3)
                self.log(f"System: {total_cores} CPU cores | {total_ram_gb:.1f} GB RAM ({avail_ram_gb:.1f} GB available)")
            except Exception:
                self.log(f"System: {total_cores} CPU cores")
        else:
            self.log(f"System: {total_cores} CPU cores")
        self.log(f"Parallelism: chdman={self.chdman_max_processors} proc | maxcso={self.maxcso_threads} threads | workers={self.max_concurrent_conversions}")
        if self.process_ps2_isos.get():
            self.log(f"PS2 emulator: {self.ps2_emulator} | Output format: {self.ps2_output_format}")
        if PSUTIL_AVAILABLE:
            self.log(f"Throttling: RAM soft={self.ram_critical_percent}% hard={self.ram_hard_limit_percent}% | Disk={self.disk_write_throttle_mb_s} MB/s")
        self.log("="*60 + "\n")
        
        successful = 0
        failed = 0
        completed = 0
        
        # Dynamic resource management - adjust workers based on system load and user limit
        resource_limit = self.check_system_resources()
        self.max_workers = min(self.max_concurrent_conversions, resource_limit)
        if self.max_workers < self.cpu_cores:
            if resource_limit < self.max_concurrent_conversions:
                self.log(f"⚠️  System load detected - using {self.max_workers} workers (throttled)")
            else:
                self.log(f"🔧 Using {self.max_workers} concurrent conversion(s) (user limit)")
        
        # Use ThreadPoolExecutor for parallel processing with dynamic worker adjustment
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Submit all conversion jobs
            futures = {executor.submit(self.process_single_file, f, i, total): f for i, f in enumerate(game_files, 1)}
            
            # Track last resource check time
            last_resource_check = time.time()
            resource_check_interval = 5.0  # Check every 5 seconds
            
            # Process results as they complete
            for future in as_completed(futures):
                if not self.is_converting:
                    self.log("\n⛔ Conversion stopped by user")
                    executor.shutdown(wait=False, cancel_futures=True)
                    break
                
                # Periodically check system resources and warn if needed
                current_time = time.time()
                if current_time - last_resource_check > resource_check_interval:
                    if PSUTIL_AVAILABLE:
                        try:
                            mem = psutil.virtual_memory()
                            if mem.percent >= self.ram_threshold_percent:
                                self.log(f"⚠️  WARNING: RAM usage high ({mem.percent:.1f}%) - conversions may slow down")
                        except Exception:
                            pass
                    last_resource_check = current_time
                
                try:
                    result = future.result()
                    if result is not None:
                        if result:
                            successful += 1
                        else:
                            failed += 1
                        
                        completed += 1
                        # Metrics update
                        with self.metrics_lock:
                            self.completed_jobs = completed
                            started_at = self.file_start_times.get(futures[future])
                            if started_at:
                                self.file_durations.append(time.time() - started_at)
                        
                        # Throttle progress bar updates (every 1% or every file if < 100 files)
                        progress_value = (completed / total) * 100
                        should_update = (total < 100) or (int(progress_value) > int((completed - 1) / total * 100))
                        if should_update:
                            self.master.after(0, lambda v=progress_value: self.progress.config(value=v))
                
                except Exception as e:
                    failed += 1
                    cue_file = futures[future]
                    self.log(f"❌ Exception processing {cue_file.name}: {e}")
        
        self.log("\n" + "="*60)
        self.log("CONVERSION COMPLETE!")
        self.log("="*60)
        self.log(f"✅ Successful: {successful}")
        self.log(f"❌ Failed: {failed}")
        self.log(f"📊 Total: {total}")
        
        # Display space savings
        if self.total_original_size > 0:
            original_gb = self.total_original_size / (1024 * 1024 * 1024)
            chd_gb = self.total_chd_size / (1024 * 1024 * 1024)
            saved_gb = (self.total_original_size - self.total_chd_size) / (1024 * 1024 * 1024)
            savings_percent = ((self.total_original_size - self.total_chd_size) / self.total_original_size) * 100
            
            self.log("\n" + "-"*60)
            self.log("💾 SPACE SAVINGS:")
            self.log(f"   Original size: {original_gb:.2f} GB")
            self.log(f"   Output size:   {chd_gb:.2f} GB")
            self.log(f"   Space saved:   {saved_gb:.2f} GB ({savings_percent:.1f}%)")
            self.log("-"*60)
        
        # Clean up extracted folders and archive files after conversion is complete
        if extracted_folders or extracted_archives:
            self.log("\n" + "="*60)
            self.log("CLEANUP: DELETING EXTRACTED FILES...")
            self.log("="*60)
            
            # Delete extracted folders
            for folder in extracted_folders:
                try:
                    if folder.exists():
                        shutil.rmtree(folder)
                        self.log(f"  🗑️  Deleted extracted folder: {folder.name}")
                except Exception as e:
                    self.log(f"  ⚠️  Error deleting folder {folder.name}: {e}")
            
            # Delete original archive files (only if delete_archives_after_extract is enabled or after extraction)
            for archive in extracted_archives:
                try:
                    if archive.exists():
                        archive.unlink()
                        self.log(f"  🗑️  Deleted archive file: {archive.name}")
                except Exception as e:
                    self.log(f"  ⚠️  Error deleting archive {archive.name}: {e}")
            
            self.log("="*60)
        
        # Clear progress file after successful completion
        if failed == 0 or total == successful:
            self.clear_progress()
        
        self.is_converting = False
        self.master.after(0, self.conversion_complete)
    
    def on_concurrent_change(self, value):
        """Handle slider change for max concurrent conversions"""
        new_value = int(float(value))
        self.max_concurrent_conversions = new_value
        if hasattr(self, 'concurrent_label'):
            self.concurrent_label.config(text=str(new_value))
        self.save_config()
    
    def start_conversion(self):
        """Start the conversion process"""
        if not self.source_dir or not os.path.isdir(self.source_dir):
            messagebox.showwarning("Warning", "Please select a valid directory")
            return
        
        if self.delete_originals.get() and self.move_to_backup.get():
            messagebox.showwarning(
                "Conflicting Options",
                "Please choose only one option: either move to backup OR delete originals."
            )
            return
        
        if self.delete_originals.get():
            response = messagebox.askyesno(
                "Confirm Deletion",
                "Are you sure you want to delete original files after conversion?\n\n"
                "This action cannot be undone!"
            )
            if not response:
                return

        # Validate maxcso availability when CSO/ZSO is selected
        if self.process_ps2_isos.get() and self.ps2_output_format in ['CSO', 'ZSO']:
            if not self.maxcso_path:
                messagebox.showwarning(
                    "maxcso Required",
                    "CSO/ZSO output requires maxcso. Set its location in the header section."
                )
                return
        
        # Initialize new batch session
        import uuid
        self.current_batch_id = str(uuid.uuid4())
        
        self.is_converting = True
        # Initialize metrics tracking
        self.metrics_running = True
        self.conversion_start_time = time.time()
        self.file_start_times.clear()
        self.file_durations.clear()
        self.total_jobs = len(self.find_game_files(self.source_dir, self.recursive.get()))
        self.completed_jobs = 0
        if PSUTIL_AVAILABLE:
            try:
                io = psutil.disk_io_counters()
                self.initial_disk_write_bytes = io.write_bytes
                self.last_disk_write_bytes = io.write_bytes
            except Exception:
                self.initial_disk_write_bytes = 0
                self.last_disk_write_bytes = 0
        self.convert_button.config(state="disabled")
        self.scan_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.status_label.config(text="⚡ CONVERTING...", fg=COLORS['accent_yellow'])
        
        # Run conversion in separate thread
        thread = threading.Thread(target=self.conversion_thread, daemon=True)
        thread.start()
        # Start metrics update loop
        self.master.after(500, self.update_metrics)
    
    def stop_conversion(self):
        """Stop the conversion process"""
        self.is_converting = False
        self.stop_button.config(state="disabled")
        self.status_label.config(text="■ STOPPING...", fg=COLORS['accent_orange'])
    
    def conversion_complete(self):
        """Called when conversion is complete"""
        self.convert_button.config(state="normal")
        self.scan_button.config(state="normal")
        self.stop_button.config(state="disabled")
        self.progress.config(value=0)
        total_cores = multiprocessing.cpu_count()
        self.status_label.config(text=f"▶ READY | {self.cpu_cores}/{total_cores} CPU CORES | 1 CORE RESERVED", 
                                fg=COLORS['text_primary'])
        self.metrics_running = False
        self.metrics_label.config(text="◆ METRICS: IDLE ◆")

    def format_seconds(self, seconds):
        if seconds is None or seconds < 0:
            return "--"
        m, s = divmod(int(seconds), 60)
        h, m = divmod(m, 60)
        if h > 0:
            return f"{h}h {m}m {s}s"
        if m > 0:
            return f"{m}m {s}s"
        return f"{s}s"

    def update_metrics(self):
        if not self.metrics_running:
            return
        with self.metrics_lock:
            completed = self.completed_jobs
            total = self.total_jobs
            durations = list(self.file_durations)
        avg_time = (sum(durations)/len(durations)) if durations else 0
        remaining = max(total - completed, 0)
        overall_eta = avg_time * (remaining / max(self.cpu_cores, 1)) if avg_time else None
        elapsed = time.time() - self.conversion_start_time if self.conversion_start_time else 0
        if PSUTIL_AVAILABLE:
            try:
                cpu = psutil.cpu_percent(interval=None)
                mem = psutil.virtual_memory()
                io = psutil.disk_io_counters()
                written_total = io.write_bytes - self.initial_disk_write_bytes
                rate_write = (io.write_bytes - self.last_disk_write_bytes) / 0.5
                self.last_disk_write_bytes = io.write_bytes
                metrics_text = (
                    f"◆ CPU {cpu:.0f}% │ MEM {mem.percent:.0f}% │ DISK {written_total/1024/1024:.1f}MB (+{rate_write/1024/1024:.1f}MB/s) │ "
                    f"JOBS {completed}/{total} │ AVG {avg_time:.1f}s │ ELAPSED {self.format_seconds(elapsed)} │ ETA {self.format_seconds(overall_eta)} ◆"
                )
            except Exception:
                metrics_text = f"◆ JOBS {completed}/{total} │ AVG {avg_time:.1f}s │ ETA {self.format_seconds(overall_eta)} ◆"
        else:
            metrics_text = f"◆ JOBS {completed}/{total} │ AVG {avg_time:.1f}s │ ETA {self.format_seconds(overall_eta)} ◆"
        self.metrics_label.config(text=metrics_text)
        self.status_label.config(text=f"⚡ CONVERTING {completed}/{total} │ ETA {self.format_seconds(overall_eta)}")
        self.master.after(500, self.update_metrics)

    def clean_game_name(self, filename):
        """Remove all parenthetical tags except disc numbers from filename"""
        name = filename
        
        # First, extract disc number if present (to preserve it)
        disc_match = re.search(r'\(Disc\s*\d+\)', name, flags=re.IGNORECASE)
        disc_tag = disc_match.group(0) if disc_match else ""
        
        # Remove version numbers (V1.0, V2.00, v1, etc.)
        name = re.sub(r'\s*\(V[\d.]+\)', '', name, flags=re.IGNORECASE)
        
        # Remove ALL parenthetical content (USA, Europe, Rev 1, v1.0, etc.)
        name = re.sub(r'\s*\([^)]+\)', '', name)
        
        # Remove [!] verified dump markers and similar brackets
        name = re.sub(r'\s*\[[^\]]+\]', '', name)
        
        # Clean up any double spaces
        name = re.sub(r'\s+', ' ', name).strip()
        
        # Re-add disc number if it was present
        if disc_tag:
            name = f"{name} {disc_tag}"
        
        return name
    
    def find_chd_files(self, directory, recursive=True):
        """Find all CHD files in directory"""
        path = Path(directory)
        if recursive:
            return sorted(path.rglob("*.chd"))
        else:
            return sorted(path.glob("*.chd"))
    
    def move_chd_files_dialog(self):
        """Open dialog to move CHD files"""
        # Create dialog window with retro styling
        dialog = Toplevel(self.master)
        dialog.title("◄ MOVE CHD FILES ►")
        dialog.geometry("650x550")
        dialog.resizable(True, True)
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.configure(bg=COLORS['bg_dark'])
        
        # Title
        title_frame = Frame(dialog, bg=COLORS['bg_light'], pady=6)
        title_frame.pack(fill="x", padx=10, pady=(10, 10))
        Label(title_frame, text="📁 CHD FILE MANAGER", font=self.font_heading_md,
              fg=COLORS['accent_purple'], bg=COLORS['bg_light']).pack()
        
        # Source directory
        source_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        source_frame.pack(fill="x")
        
        Label(source_frame, text="📂 Source:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(side="left")
        source_entry = Entry(source_frame, font=self.font_body,
                            bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                            insertbackground=COLORS['text_primary'], relief="flat")
        source_entry.pack(side="left", fill="x", expand=True, padx=5, ipady=3)
        if self.source_dir:
            source_entry.insert(0, self.source_dir)
        
        def browse_source():
            folder = filedialog.askdirectory(title="Select Source Folder")
            if folder:
                source_entry.delete(0, "end")
                source_entry.insert(0, folder)
        
        Button(source_frame, text="[ BROWSE ]", command=browse_source,
               font=self.font_small, bg=COLORS['bg_light'],
               fg=COLORS['text_secondary'], relief="flat", cursor="hand2").pack(side="left")
        
        # Destination directory
        dest_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        dest_frame.pack(fill="x")
        
        Label(dest_frame, text="📁 Destination:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(side="left")
        dest_entry = Entry(dest_frame, font=self.font_body,
                          bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                          insertbackground=COLORS['text_primary'], relief="flat")
        dest_entry.pack(side="left", fill="x", expand=True, padx=5, ipady=3)
        
        def browse_dest():
            folder = filedialog.askdirectory(title="Select Destination Folder")
            if folder:
                dest_entry.delete(0, "end")
                dest_entry.insert(0, folder)
        
        Button(dest_frame, text="[ BROWSE ]", command=browse_dest,
               font=self.font_small, bg=COLORS['bg_light'],
               fg=COLORS['text_secondary'], relief="flat", cursor="hand2").pack(side="left")
        
        # Options
        options_frame = Frame(dialog, padx=10, pady=8, bg=COLORS['bg_light'])
        options_frame.pack(fill="x", padx=10, pady=5)
        
        cb_font = ("Consolas", 9)
        cb_bg = COLORS['bg_light']
        
        remove_locale = BooleanVar(value=True)
        Checkbutton(options_frame, text="↳ Remove locale descriptors (USA, Europe, Japan, etc.)", 
                   variable=remove_locale, font=self.font_small,
                   fg=COLORS['text_secondary'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                   activebackground=cb_bg).pack(anchor="w")
        
        recursive_scan = BooleanVar(value=True)
        Checkbutton(options_frame, text="↳ Scan subdirectories", 
                   variable=recursive_scan, font=self.font_small,
                   fg=COLORS['text_secondary'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                   activebackground=cb_bg).pack(anchor="w")
        
        copy_instead = BooleanVar(value=False)
        Checkbutton(options_frame, text="↳ Copy files instead of moving", 
                   variable=copy_instead, font=self.font_small,
                   fg=COLORS['accent_orange'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                   activebackground=cb_bg).pack(anchor="w")
        
        # Results area
        results_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        results_frame.pack(fill="both", expand=True)
        
        Label(results_frame, text="◄ SCAN RESULTS ►", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=COLORS['bg_dark']).pack(anchor="w", pady=(0, 4))
        
        list_frame = Frame(results_frame, bg=COLORS['bg_dark'])
        list_frame.pack(fill="both", expand=True)
        
        scrollbar = Scrollbar(list_frame, bg=COLORS['bg_light'],
                             troughcolor=COLORS['bg_dark'])
        scrollbar.pack(side="right", fill="y")
        
        results_text = Text(list_frame, wrap="word", yscrollcommand=scrollbar.set,
                           height=10, font=self.font_mono,
                           bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                           insertbackground=COLORS['text_primary'], relief="flat",
                           padx=8, pady=8)
        results_text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=results_text.yview)
        
        # Store found files
        found_files = []
        
        def scan_for_chd():
            source = source_entry.get()
            if not source or not os.path.isdir(source):
                messagebox.showwarning("Warning", "Please select a valid source folder")
                return
            
            results_text.delete("1.0", "end")
            found_files.clear()
            
            chd_files = self.find_chd_files(source, recursive_scan.get())
            
            if not chd_files:
                results_text.insert("end", "No CHD files found in the selected folder.\n")
                return
            
            results_text.insert("end", f"Found {len(chd_files)} CHD file(s):\n\n")
            
            total_size = 0
            for chd in chd_files:
                found_files.append(chd)
                size_mb = chd.stat().st_size / (1024 * 1024)
                total_size += chd.stat().st_size
                
                original_name = chd.stem
                if remove_locale.get():
                    clean_name = self.clean_game_name(original_name)
                    if clean_name != original_name:
                        results_text.insert("end", f"📀 {original_name}.chd\n")
                        results_text.insert("end", f"   → {clean_name}.chd ({size_mb:.1f} MB)\n\n")
                    else:
                        results_text.insert("end", f"📀 {original_name}.chd ({size_mb:.1f} MB)\n\n")
                else:
                    results_text.insert("end", f"📀 {original_name}.chd ({size_mb:.1f} MB)\n\n")
            
            total_gb = total_size / (1024 * 1024 * 1024)
            results_text.insert("end", f"\n{'='*50}\n")
            results_text.insert("end", f"Total: {len(chd_files)} files, {total_gb:.2f} GB\n")
        
        def execute_move():
            source = source_entry.get()
            dest = dest_entry.get()
            
            if not source or not os.path.isdir(source):
                messagebox.showwarning("Warning", "Please select a valid source folder")
                return
            
            if not dest:
                messagebox.showwarning("Warning", "Please select a destination folder")
                return
            
            if not found_files:
                messagebox.showwarning("Warning", "Please scan for CHD files first")
                return
            
            # Create destination if it doesn't exist
            dest_path = Path(dest)
            dest_path.mkdir(parents=True, exist_ok=True)
            
            action = "Copying" if copy_instead.get() else "Moving"
            confirm = messagebox.askyesno(
                "Confirm",
                f"{action} {len(found_files)} CHD file(s) to:\n{dest}\n\n"
                f"{'Names will be cleaned (locale removed)' if remove_locale.get() else 'Names unchanged'}\n\n"
                "Continue?"
            )
            
            if not confirm:
                return
            
            results_text.delete("1.0", "end")
            results_text.insert("end", f"{action} files...\n\n")
            
            success_count = 0
            error_count = 0
            
            for chd in found_files:
                try:
                    original_name = chd.stem
                    if remove_locale.get():
                        new_name = self.clean_game_name(original_name) + ".chd"
                    else:
                        new_name = chd.name
                    
                    dest_file = dest_path / new_name
                    
                    # Handle duplicates
                    counter = 1
                    while dest_file.exists():
                        base_name = new_name.rsplit('.', 1)[0]
                        dest_file = dest_path / f"{base_name} ({counter}).chd"
                        counter += 1
                    
                    if copy_instead.get():
                        shutil.copy2(chd, dest_file)
                    else:
                        shutil.move(str(chd), str(dest_file))
                    
                    results_text.insert("end", f"✅ {original_name}.chd → {dest_file.name}\n")
                    success_count += 1
                    
                except Exception as e:
                    results_text.insert("end", f"❌ {chd.name}: {e}\n")
                    error_count += 1
                
                dialog.update()
            
            results_text.insert("end", f"\n{'='*50}\n")
            results_text.insert("end", f"Complete! ✅ {success_count} succeeded, ❌ {error_count} failed\n")
            
            if success_count > 0:
                messagebox.showinfo("Complete", f"Successfully {'copied' if copy_instead.get() else 'moved'} {success_count} file(s)")
        
        # Action buttons with retro styling
        action_frame = Frame(dialog, padx=10, pady=10, bg=COLORS['bg_dark'])
        action_frame.pack(fill="x")
        
        Button(action_frame, text="▶ SCAN", command=scan_for_chd,
               font=self.font_button,
               bg=COLORS['button_green'], fg=COLORS['bg_dark'],
               activebackground=COLORS['text_primary'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="left", padx=5)
 
        Button(action_frame, text="📁 MOVE/COPY", command=execute_move,
               font=self.font_button,
               bg=COLORS['button_blue'], fg="white",
               activebackground=COLORS['text_secondary'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="left", padx=5)
 
        Button(action_frame, text="✕ CLOSE", command=dialog.destroy,
               font=self.font_button,
               activebackground=COLORS['accent_red'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="right", padx=5)

    def cleanup_compressed_dialog(self):
        """Open dialog to clean up compressed files and extracted folders"""
        dialog = Toplevel(self.master)
        dialog.title("◄ CLEANUP COMPRESSED FILES ►")
        dialog.geometry("650x550")
        dialog.resizable(True, True)
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.configure(bg=COLORS['bg_dark'])
        
        # Title
        title_frame = Frame(dialog, bg=COLORS['bg_light'], pady=6)
        title_frame.pack(fill="x", padx=10, pady=(10, 10))
        Label(title_frame, text="🗑️ CLEANUP MANAGER", font=self.font_heading_md,
              fg=COLORS['accent_orange'], bg=COLORS['bg_light']).pack()
        
        # Source directory
        source_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        source_frame.pack(fill="x")
        
        Label(source_frame, text="📂 Source:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(side="left")
        source_entry = Entry(source_frame, font=self.font_body,
                            bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                            insertbackground=COLORS['text_primary'], relief="flat")
        source_entry.pack(side="left", fill="x", expand=True, padx=5, ipady=3)
        if self.source_dir:
            source_entry.insert(0, self.source_dir)
        
        def browse_source():
            folder = filedialog.askdirectory(title="Select Source Folder")
            if folder:
                source_entry.delete(0, "end")
                source_entry.insert(0, folder)
        
        Button(source_frame, text="[ BROWSE ]", command=browse_source,
               font=self.font_small, bg=COLORS['bg_light'],
               fg=COLORS['text_secondary'], relief="flat", cursor="hand2").pack(side="left")
        
        # Options
        options_frame = Frame(dialog, padx=10, pady=8, bg=COLORS['bg_light'])
        options_frame.pack(fill="x", padx=10, pady=5)
        
        cb_bg = COLORS['bg_light']
        
        recursive_scan = BooleanVar(value=True)
        Checkbutton(options_frame, text="↳ Scan subdirectories", 
                   variable=recursive_scan, font=self.font_small,
                   fg=COLORS['text_secondary'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                   activebackground=cb_bg).pack(anchor="w")
        
        # Results area
        results_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        results_frame.pack(fill="both", expand=True)
        
        Label(results_frame, text="◄ SCAN RESULTS ►", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=COLORS['bg_dark']).pack(anchor="w", pady=(0, 4))
        
        list_frame = Frame(results_frame, bg=COLORS['bg_dark'])
        list_frame.pack(fill="both", expand=True)
        
        scrollbar = Scrollbar(list_frame, bg=COLORS['bg_light'],
                             troughcolor=COLORS['bg_dark'])
        scrollbar.pack(side="right", fill="y")
        
        results_text = Text(list_frame, wrap="word", yscrollcommand=scrollbar.set,
                           height=10, font=self.font_mono,
                           bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                           insertbackground=COLORS['text_primary'], relief="flat",
                           padx=8, pady=8)
        results_text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=results_text.yview)
        
        # Store found files and folders
        found_archives = []
        found_folders = []
        
        def scan_for_cleanup():
            source = source_entry.get()
            if not source or not os.path.isdir(source):
                messagebox.showwarning("Warning", "Please select a valid source folder")
                return
            
            results_text.delete("1.0", "end")
            found_archives.clear()
            found_folders.clear()
            
            # Find compressed files
            archives = self.find_compressed_files(source, recursive_scan.get())
            
            # Find extracted folders (folders matching archive names without extensions)
            path = Path(source)
            all_dirs = []
            if recursive_scan.get():
                all_dirs = [d for d in path.rglob("*") if d.is_dir()]
            else:
                all_dirs = [d for d in path.glob("*") if d.is_dir()]
            
            total_archive_size = 0
            total_folder_size = 0
            
            if archives:
                results_text.insert("end", "COMPRESSED FILES:\n\n")
                for archive in archives:
                    found_archives.append(archive)
                    size_mb = archive.stat().st_size / (1024 * 1024)
                    total_archive_size += archive.stat().st_size
                    results_text.insert("end", f"📦 {archive.name} ({size_mb:.1f} MB)\n")
            
            if found_archives:
                results_text.insert("end", "\n")
            
            # Look for extracted folders
            if found_archives:
                results_text.insert("end", "EXTRACTED FOLDERS (matching archive names):\n\n")
                for archive in found_archives:
                    # Look for folder with same name as archive (without extension)
                    folder_name = archive.stem
                    potential_folder = archive.parent / folder_name
                    
                    if potential_folder.exists() and potential_folder.is_dir():
                        found_folders.append(potential_folder)
                        # Calculate folder size
                        folder_size = 0
                        for item in potential_folder.rglob("*"):
                            if item.is_file():
                                folder_size += item.stat().st_size
                        folder_mb = folder_size / (1024 * 1024)
                        total_folder_size += folder_size
                        results_text.insert("end", f"📁 {potential_folder.name}/ ({folder_mb:.1f} MB)\n")
            
            if not found_archives and not found_folders:
                results_text.insert("end", "No compressed files or extracted folders found.\n")
                return
            
            total_archive_mb = total_archive_size / (1024 * 1024)
            total_folder_mb = total_folder_size / (1024 * 1024)
            combined_mb = (total_archive_size + total_folder_size) / (1024 * 1024)
            
            results_text.insert("end", f"\n{'='*50}\n")
            results_text.insert("end", f"Archives: {len(found_archives)} files ({total_archive_mb:.1f} MB)\n")
            results_text.insert("end", f"Folders: {len(found_folders)} folders ({total_folder_mb:.1f} MB)\n")
            results_text.insert("end", f"TOTAL: {combined_mb:.1f} MB\n")
        
        def execute_cleanup():
            if not found_archives and not found_folders:
                messagebox.showwarning("Warning", "Please scan for files first")
                return
            
            confirm = messagebox.askyesno(
                "Confirm Cleanup",
                f"Delete {len(found_archives)} archive file(s) and {len(found_folders)} extracted folder(s)?\n\n"
                "This action cannot be undone!"
            )
            
            if not confirm:
                return
            
            results_text.delete("1.0", "end")
            results_text.insert("end", "Cleaning up...\n\n")
            
            success_count = 0
            error_count = 0
            
            # Delete folders first
            for folder in found_folders:
                try:
                    if folder.exists():
                        shutil.rmtree(folder)
                        results_text.insert("end", f"✅ Deleted folder: {folder.name}/\n")
                        success_count += 1
                except Exception as e:
                    results_text.insert("end", f"❌ Error deleting {folder.name}: {e}\n")
                    error_count += 1
                dialog.update()
            
            # Delete archive files
            for archive in found_archives:
                try:
                    if archive.exists():
                        archive.unlink()
                        results_text.insert("end", f"✅ Deleted archive: {archive.name}\n")
                        success_count += 1
                except Exception as e:
                    results_text.insert("end", f"❌ Error deleting {archive.name}: {e}\n")
                    error_count += 1
                dialog.update()
            
            results_text.insert("end", f"\n{'='*50}\n")
            results_text.insert("end", f"Complete! ✅ {success_count} deleted, ❌ {error_count} errors\n")
            
            if success_count > 0:
                messagebox.showinfo("Complete", f"Successfully deleted {success_count} item(s)")
            
            found_archives.clear()
            found_folders.clear()
        
        # Action buttons with retro styling
        action_frame = Frame(dialog, padx=10, pady=10, bg=COLORS['bg_dark'])
        action_frame.pack(fill="x")
        
        Button(action_frame, text="▶ SCAN", command=scan_for_cleanup,
               font=self.font_button,
               bg=COLORS['button_green'], fg=COLORS['bg_dark'],
               activebackground=COLORS['text_primary'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="left", padx=5)
 
        Button(action_frame, text="🗑️ DELETE", command=execute_cleanup,
               font=self.font_button,
               bg=COLORS['accent_red'], fg="white",
               activebackground=COLORS['accent_orange'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="left", padx=5)
 
        Button(action_frame, text="✕ CLOSE", command=dialog.destroy,
               font=self.font_button,
               activebackground=COLORS['accent_red'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="right", padx=5)

    def clean_names_dialog(self):
        """Open dialog to clean ROM file names by removing region/revision tags"""
        dialog = Toplevel(self.master)
        dialog.title("◄ CLEAN ROM NAMES ►")
        dialog.geometry("800x600")
        dialog.resizable(True, True)
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.configure(bg=COLORS['bg_dark'])
        
        # Title
        title_frame = Frame(dialog, bg=COLORS['bg_light'], pady=6)
        title_frame.pack(fill="x", padx=10, pady=(10, 10))
        Label(title_frame, text="✨ CLEAN ROM NAMES", font=self.font_heading_md,
              fg=COLORS['accent_pink'], bg=COLORS['bg_light']).pack()
        Label(title_frame, text="Remove region codes, revision tags, and other metadata from filenames",
              font=self.font_small, fg=COLORS['text_muted'], bg=COLORS['bg_light']).pack()
        
        # Source directory
        source_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        source_frame.pack(fill="x")
        
        Label(source_frame, text="📂 Source:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(side="left")
        source_entry = Entry(source_frame, font=self.font_body,
                            bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                            insertbackground=COLORS['text_primary'], relief="flat")
        source_entry.pack(side="left", fill="x", expand=True, padx=5, ipady=3)
        if self.source_dir:
            source_entry.insert(0, self.source_dir)
        
        def browse_source():
            folder = filedialog.askdirectory(title="Select Source Folder")
            if folder:
                source_entry.delete(0, "end")
                source_entry.insert(0, folder)
        
        Button(source_frame, text="[ BROWSE ]", command=browse_source,
               font=self.font_small, bg=COLORS['bg_light'],
               fg=COLORS['text_secondary'], relief="flat", cursor="hand2").pack(side="left")
        
        # Options
        options_frame = Frame(dialog, padx=10, pady=8, bg=COLORS['bg_light'])
        options_frame.pack(fill="x", padx=10, pady=5)
        
        cb_bg = COLORS['bg_light']
        
        recursive_scan = BooleanVar(value=True)
        Checkbutton(options_frame, text="↳ Scan subdirectories", 
                   variable=recursive_scan, font=self.font_small,
                   fg=COLORS['text_secondary'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                   activebackground=cb_bg).pack(anchor="w")
        
        # Info about what will be removed
        info_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        info_frame.pack(fill="x")
        
        Label(info_frame, text="Will REMOVE: (USA), (Europe), (Japan), (En), (Rev 1), (v1.0), [!], etc.",
              font=self.font_small, fg=COLORS['accent_red'], bg=COLORS['bg_dark']).pack(anchor="w")
        Label(info_frame, text="Will KEEP: (Disc 1), (Disc 2), (Bonus Disc), (Demo), (Beta), (Proto), etc.",
              font=self.font_small, fg=COLORS['button_green'], bg=COLORS['bg_dark']).pack(anchor="w")
        
        # Results area
        results_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        results_frame.pack(fill="both", expand=True)
        
        Label(results_frame, text="◄ PREVIEW CHANGES ►", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=COLORS['bg_dark']).pack(anchor="w", pady=(0, 4))
        
        list_frame = Frame(results_frame, bg=COLORS['bg_dark'])
        list_frame.pack(fill="both", expand=True)
        
        scrollbar = Scrollbar(list_frame, bg=COLORS['bg_light'],
                             troughcolor=COLORS['bg_dark'])
        scrollbar.pack(side="right", fill="y")
        
        results_text = Text(list_frame, wrap="word", yscrollcommand=scrollbar.set,
                           height=15, font=self.font_mono,
                           bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                           insertbackground=COLORS['text_primary'], relief="flat",
                           padx=8, pady=8)
        results_text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=results_text.yview)
        
        # Store files to rename
        files_to_rename = []
        # Store rename history for undo (list of (new_path, original_path) tuples)
        rename_history = []
        
        def get_clean_name(filename):
            """Clean a ROM filename by removing unwanted tags while preserving important ones."""
            name = filename
            
            # Tags to KEEP (case-insensitive patterns that should be preserved)
            keep_patterns = [
                r'\(Disc\s*\d+\)',           # (Disc 1), (Disc 2), etc.
                r'\(Disk\s*\d+\)',           # (Disk 1), (Disk 2), etc.
                r'\(Bonus\s*Disc\)',         # (Bonus Disc)
                r'\(Bonus\s*Disk\)',         # (Bonus Disk)
                r'\(Custom\s*Install\s*Disc\)', # (Custom Install Disc)
                r'\(Install\s*Disc\)',       # (Install Disc)
                r'\(Demo\)',                 # (Demo)
                r'\(Beta\)',                 # (Beta)
                r'\(Proto\)',                # (Proto)
                r'\(Prototype\)',            # (Prototype)
                r'\(Sample\)',               # (Sample)
                r'\(Promo\)',                # (Promo)
                r'\(Kiosk\)',                # (Kiosk)
                r'\(Limited\s*Edition\)',    # (Limited Edition)
                r'\(Collector.?s?\s*Edition\)', # (Collector's Edition)
                r'\(Special\s*Edition\)',    # (Special Edition)
                r'\(Game\s*of.*Year\)',      # (Game of the Year)
                r'\(GOTY\)',                 # (GOTY)
                r'\(Director.?s?\s*Cut\)',   # (Director's Cut)
                r'\(Uncut\)',                # (Uncut)
                r'\(Black\s*Label\)',        # (Black Label)
                r'\(Greatest\s*Hits\)',      # (Greatest Hits)
                r'\(Platinum\)',             # (Platinum)
                r'\(Player.?s?\s*Choice\)',  # (Player's Choice)
                r'\(Nintendo\s*Selects\)',   # (Nintendo Selects)
                r'\(Budget\)',               # (Budget)
                r'\(Reprint\)',              # (Reprint)
                r'\(Alt\)',                  # (Alt) - alternate version
                r'\(Part\s*\d+\)',           # (Part 1), (Part 2)
                r'\(Side\s*[AB]\)',          # (Side A), (Side B)
            ]
            
            # Extract tags to keep
            preserved_tags = []
            for pattern in keep_patterns:
                matches = re.findall(pattern, name, re.IGNORECASE)
                preserved_tags.extend(matches)
            
            # Tags to REMOVE (region codes, languages, revisions, etc.)
            remove_patterns = [
                r'\(USA\)',
                r'\(U\)',
                r'\(America\)',
                r'\(Europe\)',
                r'\(E\)',
                r'\(EU\)',
                r'\(Japan\)',
                r'\(J\)',
                r'\(JP\)',
                r'\(Korea\)',
                r'\(K\)',
                r'\(KR\)',
                r'\(Asia\)',
                r'\(A\)',
                r'\(World\)',
                r'\(W\)',
                r'\(Australia\)',
                r'\(AU\)',
                r'\(France\)',
                r'\(F\)',
                r'\(Fr\)',
                r'\(Germany\)',
                r'\(G\)',
                r'\(De\)',
                r'\(Spain\)',
                r'\(S\)',
                r'\(Es\)',
                r'\(Italy\)',
                r'\(I\)',
                r'\(It\)',
                r'\(Netherlands\)',
                r'\(Nl\)',
                r'\(Sweden\)',
                r'\(Sw\)',
                r'\(Sv\)',
                r'\(Norway\)',
                r'\(No\)',
                r'\(Denmark\)',
                r'\(Dk\)',
                r'\(Da\)',
                r'\(Finland\)',
                r'\(Fi\)',
                r'\(Portugal\)',
                r'\(Pt\)',
                r'\(Brazil\)',
                r'\(Br\)',
                r'\(Russia\)',
                r'\(Ru\)',
                r'\(China\)',
                r'\(Cn\)',
                r'\(Zh\)',
                r'\(Taiwan\)',
                r'\(Tw\)',
                r'\(Hong\s*Kong\)',
                r'\(HK\)',
                r'\(En\)',                   # Language: English
                r'\(En,.*?\)',               # (En,Fr), (En,De,Es), etc.
                r'\(English\)',
                r'\(French\)',
                r'\(German\)',
                r'\(Spanish\)',
                r'\(Italian\)',
                r'\(Japanese\)',
                r'\(Multi\)',                # Multi-language
                r'\(Multi\d*\)',             # (Multi5), (Multi6), etc.
                r'\(M\d+\)',                 # (M3), (M5), etc.
                r'\(Rev\s*[\dA-Z\.]+\)',    # (Rev 1), (Rev A), (Rev 1.1)
                r'\(v[\d\.]+[a-z]?\)',      # (v1.0), (v1.1), (v2.0a)
                r'\(Ver\.?\s*[\d\.]+\)',   # (Ver 1.0), (Ver. 2.0)
                r'\(Version\s*[\d\.]+\)',  # (Version 1.0)
                r'\[!\]',                    # Good dump indicator
                r'\[a\d?\]',                 # Alternate version [a], [a1]
                r'\[b\d?\]',                 # Bad dump [b], [b1]
                r'\[c\]',                    # Cracked
                r'\[f\d?\]',                 # Fixed [f], [f1]
                r'\[h\d*[A-Za-z]*\]',        # Hack indicators
                r'\[o\d?\]',                 # Overdump
                r'\[p\d?\]',                 # Pirate
                r'\[t\d?\]',                 # Trained/Trainer
                r'\[T[+-][A-Za-z]+[^\]]*\]', # Translation [T+Eng], [T-Spa]
                r'\(NTSC\)',
                r'\(NTSC-U\)',
                r'\(NTSC-J\)',
                r'\(PAL\)',
                r'\(SECAM\)',
                r'\(\d{4}-\d{2}-\d{2}\)',   # Date stamps (2001-12-25)
                r'\(\d{8}\)',                # Date stamps (20011225)
                r'\(Unl\)',                  # Unlicensed
            ]
            
            # Multi-region/multi-language combined patterns (e.g., "(USA, Europe, Asia)")
            # These need to be handled separately with a more flexible regex
            region_words = [
                'USA', 'Europe', 'Japan', 'Asia', 'World', 'Korea', 'Australia',
                'France', 'Germany', 'Spain', 'Italy', 'Netherlands', 'Sweden',
                'Norway', 'Denmark', 'Finland', 'Portugal', 'Brazil', 'Russia',
                'China', 'Taiwan', 'Hong Kong', 'Canada', 'UK', 'America',
                'En', 'Fr', 'De', 'Es', 'It', 'Ja', 'Ko', 'Zh', 'Pt', 'Ru', 'Nl',
                'English', 'French', 'German', 'Spanish', 'Italian', 'Japanese',
                'U', 'E', 'J', 'A', 'K', 'W', 'G', 'F', 'S', 'I',
                'EU', 'JP', 'KR', 'AU', 'Br', 'Cn', 'Tw', 'HK', 'Dk', 'Fi', 'No', 'Sv', 'Sw'
            ]
            # Build a pattern that matches "(Region1, Region2, ...)" with 2+ regions
            region_pattern = r'\(\s*(?:' + '|'.join(re.escape(r) for r in region_words) + r')(?:\s*,\s*(?:' + '|'.join(re.escape(r) for r in region_words) + r'))+\s*\)'
            name = re.sub(region_pattern, '', name, flags=re.IGNORECASE)
            
            # Remove the unwanted tags
            for pattern in remove_patterns:
                name = re.sub(pattern, '', name, flags=re.IGNORECASE)
            
            # Also remove any parentheses containing just 2-3 letter codes that weren't caught
            # but avoid removing preserved tags
            name = re.sub(r'\s*\([A-Za-z]{1,3}\)(?!\s*\.)', '', name)
            
            # Clean up multiple spaces
            name = re.sub(r'\s+', ' ', name)
            
            # Clean up spaces before file extension
            name = re.sub(r'\s+\.', '.', name)
            
            # Clean up leading/trailing spaces
            name = name.strip()
            
            return name
        
        def scan_for_cleaning():
            source = source_entry.get()
            if not source or not os.path.isdir(source):
                messagebox.showerror("Error", "Please select a valid source directory")
                return
            
            files_to_rename.clear()
            results_text.delete("1.0", "end")
            
            # Scan for ROM files
            rom_extensions = {'.chd', '.cue', '.bin', '.iso', '.img', '.cso', '.zso', 
                             '.gba', '.gbc', '.gb', '.sgb', '.nes', '.snes', '.sfc', '.smc',
                             '.n64', '.z64', '.v64', '.nds', '.3ds', '.cia',
                             '.psx', '.pbp', '.gcm', '.gcz', '.rvz', '.wbfs', '.wad',
                             '.xci', '.nsp', '.xiso'}
            
            path = Path(source)
            if recursive_scan.get():
                all_files = list(path.rglob("*"))
            else:
                all_files = list(path.glob("*"))
            
            # Filter to ROM files only
            rom_files = [f for f in all_files if f.is_file() and f.suffix.lower() in rom_extensions]
            
            if not rom_files:
                results_text.insert("end", "No ROM files found in the selected directory.\n")
                return
            
            results_text.insert("end", f"Found {len(rom_files)} ROM file(s)...\n\n")
            
            # Group CUE files with their BIN files for coordinated renaming
            cue_bin_groups = {}  # cue_path -> [bin_paths]
            standalone_files = []
            
            for rom_file in rom_files:
                if rom_file.suffix.lower() == '.cue':
                    # Parse CUE to find its BIN files
                    bin_files = []
                    try:
                        with open(rom_file, 'r', encoding='utf-8', errors='ignore') as f:
                            content = f.read()
                        file_pattern = re.compile(r'FILE\s+"([^"]+)"\s+BINARY', re.IGNORECASE)
                        matches = file_pattern.findall(content)
                        for match in matches:
                            bin_path = rom_file.parent / match
                            if bin_path.exists():
                                bin_files.append(bin_path)
                    except:
                        pass
                    cue_bin_groups[rom_file] = bin_files
                elif rom_file.suffix.lower() == '.bin':
                    # Check if this BIN is already associated with a CUE
                    # Will be handled via cue_bin_groups
                    pass
                else:
                    standalone_files.append(rom_file)
            
            # Find orphan BIN files (not referenced by any CUE)
            all_grouped_bins = set()
            for bins in cue_bin_groups.values():
                all_grouped_bins.update(bins)
            
            for rom_file in rom_files:
                if rom_file.suffix.lower() == '.bin' and rom_file not in all_grouped_bins:
                    standalone_files.append(rom_file)
            
            changes_found = 0
            
            # Process CUE/BIN groups - rename CUE and its BINs together
            for cue_file, bin_files in sorted(cue_bin_groups.items()):
                original_cue_name = cue_file.name
                clean_cue_name = get_clean_name(original_cue_name)
                
                if clean_cue_name != original_cue_name:
                    changes_found += 1
                    files_to_rename.append((cue_file, clean_cue_name))
                    relative_path = cue_file.parent.relative_to(path) if cue_file.parent != path else Path('.')
                    results_text.insert("end", f"📁 {relative_path}\n")
                    results_text.insert("end", f"  ❌ {original_cue_name}\n")
                    results_text.insert("end", f"  ✅ {clean_cue_name}\n")
                    
                    # Also rename associated BIN files to match
                    clean_base = Path(clean_cue_name).stem
                    for i, bin_file in enumerate(bin_files):
                        original_bin_name = bin_file.name
                        # Construct new BIN name based on clean CUE name
                        if len(bin_files) == 1:
                            new_bin_name = f"{clean_base}.bin"
                        else:
                            # Multi-track: preserve track numbering if present
                            track_match = re.search(r'[\s\(\[]*(Track\s*\d+|T\d+|\d{2})[\s\)\]]*\.bin$', original_bin_name, re.IGNORECASE)
                            if track_match:
                                new_bin_name = f"{clean_base} ({track_match.group(1).strip()}).bin"
                            else:
                                new_bin_name = f"{clean_base} (Track {i+1}).bin"
                        
                        if new_bin_name != original_bin_name:
                            files_to_rename.append((bin_file, new_bin_name))
                            results_text.insert("end", f"    ❌ {original_bin_name}\n")
                            results_text.insert("end", f"    ✅ {new_bin_name}\n")
                    results_text.insert("end", "\n")
            
            # Process standalone files
            for rom_file in sorted(standalone_files):
                original_name = rom_file.name
                clean_name = get_clean_name(original_name)
                
                if clean_name != original_name:
                    changes_found += 1
                    files_to_rename.append((rom_file, clean_name))
                    relative_path = rom_file.parent.relative_to(path) if rom_file.parent != path else Path('.')
                    results_text.insert("end", f"📁 {relative_path}\n")
                    results_text.insert("end", f"  ❌ {original_name}\n")
                    results_text.insert("end", f"  ✅ {clean_name}\n\n")
            
            if changes_found == 0:
                results_text.insert("end", "✨ All file names are already clean! No changes needed.\n")
            else:
                results_text.insert("end", f"\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n")
                results_text.insert("end", f"Total: {changes_found} file(s) will be renamed.\n")
            
            results_text.see("1.0")
        
        def execute_rename():
            if not files_to_rename:
                messagebox.showwarning("No Changes", "No files to rename. Run SCAN first.")
                return
            
            confirm = messagebox.askyesno(
                "Confirm Rename",
                f"This will rename {len(files_to_rename)} file(s).\n\n"
                "This action cannot be undone!\n\n"
                "Continue?"
            )
            
            if not confirm:
                return
            
            renamed_count = 0
            error_count = 0
            error_files = []
            
            # Clear previous history and start fresh for this batch
            rename_history.clear()
            
            results_text.delete("1.0", "end")
            results_text.insert("end", "Renaming files...\n\n")
            
            # Separate CUE files from other files - process CUE files specially
            cue_renames = []  # (cue_file, new_cue_name)
            bin_renames = []  # (bin_file, new_bin_name)
            other_renames = []
            
            for rom_file, new_name in files_to_rename:
                if rom_file.suffix.lower() == '.cue':
                    cue_renames.append((rom_file, new_name))
                elif rom_file.suffix.lower() == '.bin':
                    bin_renames.append((rom_file, new_name))
                else:
                    other_renames.append((rom_file, new_name))
            
            # Build a mapping of old BIN names to new BIN names for CUE updates
            bin_name_map = {bf.name: new_name for bf, new_name in bin_renames}
            
            # Process CUE files first - update contents BEFORE renaming BINs
            for cue_file, new_cue_name in cue_renames:
                try:
                    new_cue_path = cue_file.parent / new_cue_name
                    
                    if new_cue_path.exists():
                        results_text.insert("end", f"⚠️ SKIP (exists): {cue_file.name} → {new_cue_name}\n")
                        error_files.append((cue_file.name, "Target file already exists"))
                        error_count += 1
                        continue
                    
                    # Read CUE file contents and update BIN references
                    try:
                        with open(cue_file, 'r', encoding='utf-8', errors='ignore') as f:
                            content = f.read()
                        
                        # Update all BIN references
                        for old_bin, new_bin in bin_name_map.items():
                            content = content.replace(f'"{old_bin}"', f'"{new_bin}"')
                            content = content.replace(f"'{old_bin}'", f"'{new_bin}'")
                        
                        # Write updated content
                        with open(cue_file, 'w', encoding='utf-8') as f:
                            f.write(content)
                    except Exception as e:
                        results_text.insert("end", f"  ⚠️ Could not update CUE contents: {e}\n")
                    
                    # Now rename the CUE file
                    cue_file.rename(new_cue_path)
                    rename_history.append((new_cue_path, cue_file))
                    results_text.insert("end", f"✅ {cue_file.name} → {new_cue_name}\n")
                    renamed_count += 1
                    
                except Exception as e:
                    results_text.insert("end", f"❌ ERROR: {cue_file.name} → {new_cue_name}\n   Reason: {e}\n")
                    error_files.append((cue_file.name, str(e)))
                    error_count += 1
            
            # Process BIN files
            for bin_file, new_bin_name in bin_renames:
                try:
                    new_path = bin_file.parent / new_bin_name
                    
                    if new_path.exists():
                        results_text.insert("end", f"⚠️ SKIP (exists): {bin_file.name} → {new_bin_name}\n")
                        error_files.append((bin_file.name, "Target file already exists"))
                        error_count += 1
                        continue
                    
                    bin_file.rename(new_path)
                    rename_history.append((new_path, bin_file))
                    results_text.insert("end", f"✅ {bin_file.name} → {new_bin_name}\n")
                    renamed_count += 1
                    
                except Exception as e:
                    results_text.insert("end", f"❌ ERROR: {bin_file.name} → {new_bin_name}\n   Reason: {e}\n")
                    error_files.append((bin_file.name, str(e)))
                    error_count += 1
            
            # Process other files
            for rom_file, new_name in other_renames:
                try:
                    new_path = rom_file.parent / new_name
                    
                    if new_path.exists():
                        results_text.insert("end", f"⚠️ SKIP (exists): {rom_file.name} → {new_name}\n")
                        error_files.append((rom_file.name, "Target file already exists"))
                        error_count += 1
                        continue
                    
                    rom_file.rename(new_path)
                    rename_history.append((new_path, rom_file))
                    results_text.insert("end", f"✅ {rom_file.name} → {new_name}\n")
                    renamed_count += 1
                    
                except Exception as e:
                    results_text.insert("end", f"❌ ERROR: {rom_file.name} → {new_name}\n   Reason: {e}\n")
                    error_files.append((rom_file.name, str(e)))
                    error_count += 1
            
            results_text.insert("end", f"\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n")
            results_text.insert("end", f"✅ Renamed: {renamed_count}\n")
            if error_count > 0:
                results_text.insert("end", f"❌ Errors: {error_count}\n\n")
                results_text.insert("end", "Files with errors:\n")
                for filename, reason in error_files:
                    results_text.insert("end", f"  • {filename}\n    → {reason}\n")
            if renamed_count > 0:
                results_text.insert("end", f"\n💡 Use UNDO to revert these changes.\n")
            
            files_to_rename.clear()
            messagebox.showinfo("Complete", f"Renamed {renamed_count} file(s)." + 
                              (f"\n\n{error_count} file(s) had errors." if error_count > 0 else ""))
        
        def undo_rename():
            """Undo the last batch of renames"""
            if not rename_history:
                messagebox.showwarning("Nothing to Undo", "No rename operations to undo.")
                return
            
            confirm = messagebox.askyesno(
                "Confirm Undo",
                f"This will revert {len(rename_history)} file(s) to their original names.\n\n"
                "Continue?"
            )
            
            if not confirm:
                return
            
            reverted_count = 0
            error_count = 0
            
            results_text.delete("1.0", "end")
            results_text.insert("end", "Undoing renames...\n\n")
            
            # Process in reverse order
            for new_path, original_path in reversed(rename_history):
                try:
                    # Check if renamed file still exists
                    if not new_path.exists():
                        results_text.insert("end", f"⚠️ SKIP (not found): {new_path.name}\n")
                        error_count += 1
                        continue
                    
                    # Check if original name is now taken
                    if original_path.exists():
                        results_text.insert("end", f"⚠️ SKIP (conflict): {original_path.name} already exists\n")
                        error_count += 1
                        continue
                    
                    new_path.rename(original_path)
                    results_text.insert("end", f"↩️ {new_path.name} → {original_path.name}\n")
                    reverted_count += 1
                    
                except Exception as e:
                    results_text.insert("end", f"❌ ERROR: {new_path.name} - {e}\n")
                    error_count += 1
            
            results_text.insert("end", f"\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n")
            results_text.insert("end", f"↩️ Reverted: {reverted_count}\n")
            if error_count > 0:
                results_text.insert("end", f"❌ Errors: {error_count}\n")
            
            rename_history.clear()
            messagebox.showinfo("Undo Complete", f"Reverted {reverted_count} file(s) to original names.")
        
        # Action buttons
        action_frame = Frame(dialog, padx=10, pady=10, bg=COLORS['bg_dark'])
        action_frame.pack(fill="x")
        
        Button(action_frame, text="▶ SCAN", command=scan_for_cleaning,
               font=self.font_button,
               bg=COLORS['button_green'], fg=COLORS['bg_dark'],
               activebackground=COLORS['text_primary'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="left", padx=5)
        
        Button(action_frame, text="✨ RENAME", command=execute_rename,
               font=self.font_button,
               bg=COLORS['accent_pink'], fg="white",
               activebackground=COLORS['accent_purple'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="left", padx=5)
        
        Button(action_frame, text="↩️ UNDO", command=undo_rename,
               font=self.font_button,
               bg=COLORS['accent_orange'], fg="white",
               activebackground=COLORS['accent_yellow'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="left", padx=5)
        
        Button(action_frame, text="✕ CLOSE", command=dialog.destroy,
               font=self.font_button,
               activebackground=COLORS['accent_red'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="right", padx=5)

    def extract_archives_dialog(self):
        """Open dialog to scan archives, detect ROM systems, and extract to configured folders"""
        dialog = Toplevel(self.master)
        dialog.title("◄ EXTRACT ARCHIVES BY SYSTEM ►")
        dialog.geometry("900x700")
        dialog.resizable(True, True)
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.configure(bg=COLORS['bg_dark'])
        
        # Title
        title_frame = Frame(dialog, bg=COLORS['bg_light'], pady=6)
        title_frame.pack(fill="x", padx=10, pady=(10, 10))
        Label(title_frame, text="📦 ARCHIVE EXTRACTOR - SORT BY SYSTEM", font=self.font_heading_md,
              fg=COLORS['accent_yellow'], bg=COLORS['bg_light']).pack()
        
        # Source directory
        source_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        source_frame.pack(fill="x")
        
        Label(source_frame, text="📂 Source:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(side="left")
        source_entry = Entry(source_frame, font=self.font_body,
                            bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                            insertbackground=COLORS['text_primary'], relief="flat")
        source_entry.pack(side="left", fill="x", expand=True, padx=5, ipady=3)
        if self.source_dir:
            source_entry.insert(0, self.source_dir)
        
        def browse_source():
            folder = filedialog.askdirectory(title="Select Source Folder with Archives")
            if folder:
                source_entry.delete(0, "end")
                source_entry.insert(0, folder)
        
        Button(source_frame, text="[ BROWSE ]", command=browse_source,
               font=self.font_small, bg=COLORS['bg_light'],
               fg=COLORS['text_secondary'], relief="flat", cursor="hand2").pack(side="left")
        
        # Options
        options_frame = Frame(dialog, padx=10, pady=8, bg=COLORS['bg_light'])
        options_frame.pack(fill="x", padx=10, pady=5)
        
        cb_bg = COLORS['bg_light']
        
        recursive_scan = BooleanVar(value=True)
        Checkbutton(options_frame, text="↳ Scan subdirectories", 
                   variable=recursive_scan, font=self.font_small,
                   fg=COLORS['text_secondary'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                   activebackground=cb_bg).pack(side="left", padx=(0, 20))
        
        delete_after_extract = BooleanVar(value=False)
        Checkbutton(options_frame, text="⚠ Delete archives after extraction", 
                   variable=delete_after_extract, font=self.font_small,
                   fg=COLORS['accent_red'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                   activebackground=cb_bg).pack(side="left")

        copy_archives_var = BooleanVar(value=False)
        Checkbutton(options_frame, text="📁 Copy archives instead of move", 
               variable=copy_archives_var, font=self.font_small,
               fg=COLORS['text_secondary'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
               activebackground=cb_bg).pack(side="left", padx=(20, 0))
        
        # System folder configuration frame
        system_config_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        system_config_frame.pack(fill="x")
        
        # Header with label and base folder button
        system_header_frame = Frame(system_config_frame, bg=COLORS['bg_dark'])
        system_header_frame.pack(fill="x", pady=(0, 4))
        
        Label(system_header_frame, text="◄ SYSTEM EXTRACTION FOLDERS ►", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=COLORS['bg_dark']).pack(side="left")
        
        # Base folder for quick setup
        base_folder_var = {"path": ""}
        
        def set_base_folder():
            """Set a base folder and auto-create system subfolders"""
            folder = filedialog.askdirectory(title="Select Base ROM Folder (subfolders will be created per system)")
            if folder:
                base_folder_var["path"] = folder
                # Update all detected system entries with base_folder/SystemName
                for system_name, entry in system_folder_entries.items():
                    system_subfolder = os.path.join(folder, system_name.replace("/", "-").replace(" ", "_"))
                    entry.delete(0, "end")
                    entry.insert(0, system_subfolder)
                    self.system_extract_dirs[system_name] = system_subfolder
                self.save_config()
                messagebox.showinfo("Base Folder Set", 
                    f"All systems configured to extract to:\n{folder}/[SystemName]\n\n"
                    "You can still edit individual paths if needed.")
        
        Button(system_header_frame, text="📁 SET BASE FOLDER (Auto-create subfolders)", 
               command=set_base_folder,
               font=self.font_small, bg=COLORS['accent_yellow'], fg=COLORS['bg_dark'],
               relief="flat", cursor="hand2", padx=10).pack(side="right")
        
        # Scrollable frame for system folder configuration
        system_canvas_frame = Frame(system_config_frame, bg=COLORS['bg_medium'], height=150)
        system_canvas_frame.pack(fill="x", pady=5)
        system_canvas_frame.pack_propagate(False)
        
        system_scrollbar = Scrollbar(system_canvas_frame, bg=COLORS['bg_light'],
                                     troughcolor=COLORS['bg_dark'])
        system_scrollbar.pack(side="right", fill="y")
        
        from tkinter import Canvas
        system_canvas = Canvas(system_canvas_frame, bg=COLORS['bg_medium'],
                              yscrollcommand=system_scrollbar.set, highlightthickness=0)
        system_canvas.pack(side="left", fill="both", expand=True)
        system_scrollbar.config(command=system_canvas.yview)
        
        system_list_frame = Frame(system_canvas, bg=COLORS['bg_medium'])
        system_canvas.create_window((0, 0), window=system_list_frame, anchor="nw")
        
        def update_scroll_region(event=None):
            system_canvas.configure(scrollregion=system_canvas.bbox("all"))
        system_list_frame.bind("<Configure>", update_scroll_region)
        
        # Store system folder entries for easy access
        system_folder_entries = {}
        system_folder_labels = {}
        system_checkboxes = {}  # Track which systems to extract
        
        def create_system_folder_row(parent, system_name, rom_count=0):
            """Create a row for configuring a system's extraction folder"""
            row_frame = Frame(parent, bg=COLORS['bg_medium'], pady=3)
            row_frame.pack(fill="x", padx=5, pady=2)
            
            # Checkbox to include/exclude this system
            include_var = BooleanVar(value=True)
            system_checkboxes[system_name] = include_var
            Checkbutton(row_frame, variable=include_var,
                       bg=COLORS['bg_medium'], selectcolor=COLORS['bg_dark'],
                       activebackground=COLORS['bg_medium']).pack(side="left")
            
            # System name with ROM count
            Label(row_frame, text=f"{system_name} ({rom_count} ROMs):", width=25, anchor="w",
                  font=self.font_small, fg=COLORS['text_primary'], 
                  bg=COLORS['bg_medium']).pack(side="left")
            
            # Editable entry field for folder path
            folder_entry = Entry(row_frame, font=self.font_small,
                                bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                                insertbackground=COLORS['text_primary'], relief="flat", width=40)
            folder_entry.pack(side="left", fill="x", expand=True, padx=5, ipady=2)
            
            # Pre-fill with saved path if available
            if system_name in self.system_extract_dirs:
                folder_entry.insert(0, self.system_extract_dirs[system_name])
            
            system_folder_entries[system_name] = folder_entry
            
            def browse_system_folder():
                folder = filedialog.askdirectory(title=f"Select folder for {system_name} ROMs")
                if folder:
                    folder_entry.delete(0, "end")
                    folder_entry.insert(0, folder)
                    self.system_extract_dirs[system_name] = folder
                    self.save_config()
            
            Button(row_frame, text="📂", command=browse_system_folder,
                   font=self.font_small, bg=COLORS['bg_light'],
                   fg=COLORS['text_secondary'], relief="flat", cursor="hand2",
                   width=3).pack(side="left", padx=2)
            
            def save_entry_path(event=None):
                path = folder_entry.get().strip()
                if path:
                    self.system_extract_dirs[system_name] = path
                elif system_name in self.system_extract_dirs:
                    del self.system_extract_dirs[system_name]
                self.save_config()
            
            folder_entry.bind("<FocusOut>", save_entry_path)
            folder_entry.bind("<Return>", save_entry_path)
        
        # Placeholder text when no systems detected
        placeholder_label = Label(system_list_frame, 
                                 text="Scan archives to detect systems and configure extraction folders",
                                 font=self.font_small, fg=COLORS['text_muted'],
                                 bg=COLORS['bg_medium'], pady=20)
        placeholder_label.pack()
        
        # Results area
        results_frame = Frame(dialog, padx=10, pady=5, bg=COLORS['bg_dark'])
        results_frame.pack(fill="both", expand=True)
        
        Label(results_frame, text="◄ SCAN RESULTS ►", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=COLORS['bg_dark']).pack(anchor="w", pady=(0, 4))
        
        list_frame = Frame(results_frame, bg=COLORS['bg_dark'])
        list_frame.pack(fill="both", expand=True)
        
        scrollbar = Scrollbar(list_frame, bg=COLORS['bg_light'],
                             troughcolor=COLORS['bg_dark'])
        scrollbar.pack(side="right", fill="y")
        
        results_text = Text(list_frame, wrap="word", yscrollcommand=scrollbar.set,
                           height=15, font=self.font_mono,
                           bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                           insertbackground=COLORS['text_primary'], relief="flat",
                           padx=8, pady=8)
        results_text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=results_text.yview)
        
        # Store scan results
        archives_by_system = {}  # {system_name: [(archive_path, [rom_files])]}
        detected_systems = set()
        
        def detect_system_from_name(name, size_bytes=None):
            """Detect system from filename, with metadata/size/ID heuristics for ISO (PS2 vs PSP)."""
            file_ext = Path(name).suffix.lower()
            if file_ext == '.iso':
                system = self.detect_iso_system(name, size_bytes)
                if system:
                    return system
            return SYSTEM_EXTENSIONS.get(file_ext)
        
        def scan_archive_contents(archive_path):
            """Scan archive to detect ROM systems without extracting"""
            archive_path = Path(archive_path)
            ext = archive_path.suffix.lower()
            systems_found = {}
            
            try:
                if ext == '.zip':
                    with zipfile.ZipFile(archive_path, 'r') as zf:
                        for info in zf.infolist():
                            system = detect_system_from_name(info.filename, info.file_size)
                            if system:
                                if system not in systems_found:
                                    systems_found[system] = []
                                systems_found[system].append(info.filename)
                
                elif ext in ['.tar', '.gz', '.tgz'] or archive_path.name.endswith('.tar.gz'):
                    mode = 'r:gz' if ext in ['.gz', '.tgz'] or archive_path.name.endswith('.tar.gz') else 'r'
                    with tarfile.open(archive_path, mode) as tf:
                        for member in tf.getmembers():
                            if member.isfile():
                                system = detect_system_from_name(member.name, getattr(member, 'size', None))
                                if system:
                                    if system not in systems_found:
                                        systems_found[system] = []
                                    systems_found[system].append(member.name)
                
                elif ext in ['.7z', '.rar']:
                    if self.seven_zip_path:
                        # Use 7z to list archive contents
                        cmd = [self.seven_zip_path, 'l', '-ba', str(archive_path)]
                        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                        if result.returncode == 0:
                            for line in result.stdout.split('\n'):
                                # 7z list output format varies, try to extract filename
                                parts = line.strip().split()
                                if parts:
                                    name = parts[-1]  # Filename is usually last
                                    size_val = None
                                    if len(parts) >= 3 and parts[-3].isdigit():
                                        size_val = int(parts[-3])
                                    system = detect_system_from_name(name, size_val)
                                    if system:
                                        if system not in systems_found:
                                            systems_found[system] = []
                                        systems_found[system].append(name)
            except Exception as e:
                pass  # Silently fail for individual archives
            
            return systems_found
        
        def scan_archives():
            """Scan all archives in source directory and detect ROM systems"""
            source = source_entry.get()
            if not source or not os.path.isdir(source):
                messagebox.showwarning("Warning", "Please select a valid source folder")
                return
            
            results_text.delete("1.0", "end")
            archives_by_system.clear()
            detected_systems.clear()
            
            # Clear previous system folder rows and checkboxes
            for widget in system_list_frame.winfo_children():
                widget.destroy()
            system_folder_entries.clear()
            system_checkboxes.clear()
            
            results_text.insert("end", "Scanning archives...\n\n")
            dialog.update()
            
            # Find all compressed files
            archives = self.find_compressed_files(source, recursive_scan.get())
            
            if not archives:
                results_text.insert("end", "No archive files found in the selected directory.\n")
                # Re-add placeholder
                Label(system_list_frame, 
                      text="No archives found. Select a different source folder.",
                      font=self.font_small, fg=COLORS['text_muted'],
                      bg=COLORS['bg_medium'], pady=20).pack()
                return
            
            results_text.insert("end", f"Found {len(archives)} archive(s). Analyzing contents...\n\n")
            dialog.update()
            
            total_roms = 0
            for archive in archives:
                systems_in_archive = scan_archive_contents(archive)
                
                for system, roms in systems_in_archive.items():
                    detected_systems.add(system)
                    if system not in archives_by_system:
                        archives_by_system[system] = []
                    archives_by_system[system].append((archive, roms))
                    total_roms += len(roms)
            
            if not detected_systems:
                results_text.insert("end", "No ROM files detected in archives.\n")
                results_text.insert("end", "Archives may be empty or contain unsupported file types.\n")
                # Re-add placeholder
                Label(system_list_frame, 
                      text="No ROMs detected. Archives may contain unsupported formats.",
                      font=self.font_small, fg=COLORS['text_muted'],
                      bg=COLORS['bg_medium'], pady=20).pack()
                return
            
            # Create system folder configuration rows for detected systems with ROM counts
            for system in sorted(detected_systems):
                rom_count = sum(len(roms) for _, roms in archives_by_system.get(system, []))
                create_system_folder_row(system_list_frame, system, rom_count)
            
            # Update scroll region
            system_list_frame.update_idletasks()
            system_canvas.configure(scrollregion=system_canvas.bbox("all"))
            
            # Display results organized by system
            results_text.insert("end", f"{'━' * 60}\n")
            results_text.insert("end", f"📊 SCAN SUMMARY\n")
            results_text.insert("end", f"{'━' * 60}\n")
            results_text.insert("end", f"   Systems Detected: {len(detected_systems)}\n")
            results_text.insert("end", f"   Total ROM Files:  {total_roms}\n")
            results_text.insert("end", f"   Archives Scanned: {len(archives)}\n")
            results_text.insert("end", f"{'━' * 60}\n\n")
            results_text.insert("end", "💡 Configure extraction folders above, then click EXTRACT BY SYSTEM\n\n")
            
            for system in sorted(archives_by_system.keys()):
                archive_list = archives_by_system[system]
                rom_count = sum(len(roms) for _, roms in archive_list)
                config_status = "✅ CONFIGURED" if system in self.system_extract_dirs else "⚠️ SET FOLDER ABOVE"
                dest_path = self.system_extract_dirs.get(system, "")
                
                results_text.insert("end", f"{'═' * 60}\n")
                results_text.insert("end", f"🎮 {system.upper()}\n")
                results_text.insert("end", f"{'═' * 60}\n")
                results_text.insert("end", f"   Status: {config_status}\n")
                if dest_path:
                    results_text.insert("end", f"   Destination: {dest_path}\n")
                results_text.insert("end", f"   ROMs: {rom_count} file(s) in {len(archive_list)} archive(s)\n")
                results_text.insert("end", f"{'─' * 60}\n")
                
                for archive, roms in archive_list:
                    results_text.insert("end", f"\n   📦 {archive.name}\n")
                    results_text.insert("end", f"   {'─' * 40}\n")
                    for rom in roms:
                        results_text.insert("end", f"      • {Path(rom).name}\n")
                
                results_text.insert("end", "\n")
        
        def reassign_rom_dialog():
            """Allow user to move detected ROMs to another system before extraction."""
            if not archives_by_system:
                messagebox.showinfo("No Data", "Please scan archives first.")
                return
            dialog.lift()
            dialog.attributes('-topmost', True)
            rom_query = simpledialog.askstring("Reassign ROM", "Enter part of the ROM filename to move:", parent=dialog)
            if not rom_query:
                dialog.attributes('-topmost', False)
                return
            dest_system = simpledialog.askstring(
                "Target System",
                "Enter target system name (existing or new):\n" +
                ", ".join(sorted(detected_systems)) if detected_systems else "e.g., PlayStation 2 or PSP",
                parent=dialog
            )
            dialog.attributes('-topmost', False)
            if not dest_system:
                return
            dest_system = dest_system.strip()
            moved_entries = []
            for system, archive_list in list(archives_by_system.items()):
                new_archive_list = []
                for archive, roms in archive_list:
                    moving = [r for r in roms if rom_query.lower() in r.lower()]
                    staying = [r for r in roms if r not in moving]
                    if moving:
                        moved_entries.append((archive, moving))
                    if staying:
                        new_archive_list.append((archive, staying))
                archives_by_system[system] = new_archive_list
            if not moved_entries:
                messagebox.showinfo("Not Found", f"No ROM matching '{rom_query}' was found.")
                return
            dest_list = archives_by_system.get(dest_system, [])
            dest_list.extend(moved_entries)
            archives_by_system[dest_system] = dest_list
            detected_systems.add(dest_system)
            moved_count = sum(len(m) for _, m in moved_entries)
            results_text.insert("end", f"\n🔀 Moved {moved_count} ROM(s) to {dest_system} (query: {rom_query})\n")
            results_text.see("end")

        def organize_archives_to_folders():
            """Move or copy archives into system-specific folders without extracting."""
            if not archives_by_system:
                messagebox.showwarning("Warning", "Please scan archives first")
                return

            # Persist any path edits from the UI
            for system, entry in system_folder_entries.items():
                path = entry.get().strip()
                if path:
                    self.system_extract_dirs[system] = path
                elif system in self.system_extract_dirs:
                    del self.system_extract_dirs[system]
            self.save_config()

            systems_to_process = []
            systems_unchecked = []
            systems_unconfigured = []

            for system in detected_systems:
                is_checked = system_checkboxes.get(system, BooleanVar(value=True)).get()
                entry = system_folder_entries.get(system)
                folder_path = entry.get().strip() if entry else self.system_extract_dirs.get(system, "")
                has_folder = bool(folder_path)
                if not is_checked:
                    systems_unchecked.append(system)
                elif not has_folder:
                    systems_unconfigured.append(system)
                else:
                    systems_to_process.append((system, folder_path))

            if not systems_to_process:
                messagebox.showwarning(
                    "Warning",
                    "No systems selected for organization.\n\nCheck at least one system and configure its folder."
                )
                return

            results_text.delete("1.0", "end")
            action_word = "Copying" if copy_archives_var.get() else "Moving"
            results_text.insert("end", f"{action_word} archives into system folders...\n\n")
            dialog.update()

            moved_count = 0
            copied_count = 0
            error_count = 0
            skipped_archives = set()
            routed_archives = {}

            for system, folder_path in systems_to_process:
                dest_folder = Path(folder_path)
                archive_list = archives_by_system.get(system, [])

                if not dest_folder.exists():
                    try:
                        dest_folder.mkdir(parents=True, exist_ok=True)
                    except Exception as e:
                        error_count += len(archive_list)
                        results_text.insert("end", f"❌ Cannot create folder for {system}: {e}\n")
                        continue

                results_text.insert("end", f"\n🎮 {system} → {dest_folder}\n")
                dialog.update()

                for archive, roms in archive_list:
                    archive_path = Path(archive)

                    if archive_path in routed_archives:
                        results_text.insert("end", f"   ↪️ {archive_path.name} already placed in {routed_archives[archive_path]}\n")
                        continue

                    target_path = dest_folder / archive_path.name
                    final_target = target_path
                    suffix = 1
                    while final_target.exists():
                        final_target = target_path.with_name(f"{target_path.stem} ({suffix}){target_path.suffix}")
                        suffix += 1

                    try:
                        if copy_archives_var.get():
                            shutil.copy2(archive_path, final_target)
                            copied_count += 1
                        else:
                            shutil.move(str(archive_path), str(final_target))
                            moved_count += 1
                        routed_archives[archive_path] = final_target
                        results_text.insert("end", f"   ✅ {archive_path.name} → {final_target}\n")
                    except Exception as e:
                        error_count += 1
                        results_text.insert("end", f"   ❌ {archive_path.name}: {e}\n")

                results_text.insert("end", "\n")
                dialog.update()

            if systems_unchecked:
                skipped_archives.update({Path(a) for sys_name in systems_unchecked for a, _ in archives_by_system.get(sys_name, [])})
            if systems_unconfigured:
                skipped_archives.update({Path(a) for sys_name in systems_unconfigured for a, _ in archives_by_system.get(sys_name, [])})

            results_text.insert("end", f"\n{'━' * 50}\n")
            results_text.insert("end", f"✅ Moved: {moved_count}\n")
            results_text.insert("end", f"✅ Copied: {copied_count}\n")
            if skipped_archives:
                results_text.insert("end", f"⏭️ Skipped: {len(skipped_archives)} archive(s) (unchecked or unconfigured)\n")
            if error_count:
                results_text.insert("end", f"❌ Errors: {error_count}\n")

            messagebox.showinfo(
                "Organization Complete",
                f"Moved: {moved_count}\n" +
                f"Copied: {copied_count}\n" +
                (f"Skipped: {len(skipped_archives)}\n" if skipped_archives else "") +
                (f"Errors: {error_count}" if error_count else "")
            )
        
        def extract_to_system_folders():
            """Extract archives to their configured system folders"""
            if not archives_by_system:
                messagebox.showwarning("Warning", "Please scan archives first")
                return
            
            # Save any paths entered in the entry fields and build extraction map
            extraction_paths = {}
            for system, entry in system_folder_entries.items():
                path = entry.get().strip()
                if path:
                    self.system_extract_dirs[system] = path
                    extraction_paths[system] = path
                elif system in self.system_extract_dirs:
                    del self.system_extract_dirs[system]
            self.save_config()
            
            # Get systems to extract (checked and configured)
            systems_to_extract = []
            systems_unchecked = []
            systems_unconfigured = []
            
            for system in detected_systems:
                is_checked = system_checkboxes.get(system, BooleanVar(value=True)).get()
                # Check both the entry field directly and the saved config
                entry_path = system_folder_entries.get(system)
                has_folder = (entry_path and entry_path.get().strip()) or \
                            (system in self.system_extract_dirs and self.system_extract_dirs[system].strip())
                
                if not is_checked:
                    systems_unchecked.append(system)
                elif not has_folder:
                    systems_unconfigured.append(system)
                else:
                    systems_to_extract.append(system)
            
            if not systems_to_extract:
                messagebox.showwarning("Warning", 
                    "No systems selected for extraction.\n\n" +
                    "Please check the systems you want to extract and ensure folders are configured.")
                return
            
            # Warn about unconfigured systems
            if systems_unconfigured:
                proceed = messagebox.askyesno(
                    "Unconfigured Systems",
                    f"{len(systems_unconfigured)} checked system(s) don't have folders configured:\n" +
                    "\n".join(f"  • {s}" for s in systems_unconfigured) +
                    "\n\nThese will be skipped.\nContinue anyway?"
                )
                if not proceed:
                    return
            
            results_text.delete("1.0", "end")
            results_text.insert("end", "Extracting archives by system...\n\n")
            dialog.update()
            
            extracted_count = 0
            skipped_count = 0
            error_count = 0
            
            for system, archive_list in archives_by_system.items():
                # Check if system is selected and configured
                is_checked = system_checkboxes.get(system, BooleanVar(value=True)).get()
                
                # Get folder path from entry field or saved config
                entry = system_folder_entries.get(system)
                folder_path = entry.get().strip() if entry else self.system_extract_dirs.get(system, "")
                has_folder = bool(folder_path)
                
                if not is_checked:
                    for archive, roms in archive_list:
                        skipped_count += len(roms)
                    results_text.insert("end", f"⏭️ Skipping {system} (unchecked)\n")
                    continue
                
                if not has_folder:
                    for archive, roms in archive_list:
                        skipped_count += len(roms)
                    results_text.insert("end", f"⏭️ Skipping {system} (no folder configured)\n")
                    continue
                
                dest_folder = Path(folder_path)
                if not dest_folder.exists():
                    try:
                        dest_folder.mkdir(parents=True, exist_ok=True)
                    except Exception as e:
                        results_text.insert("end", f"❌ Cannot create folder for {system}: {e}\n")
                        error_count += 1
                        continue
                
                results_text.insert("end", f"\n🎮 Extracting {system} ROMs to: {dest_folder}\n")
                dialog.update()
                
                for archive, roms in archive_list:
                    try:
                        archive_path = Path(archive)
                        ext = archive_path.suffix.lower()
                        
                        # Extract specific files based on archive type
                        if ext == '.zip':
                            with zipfile.ZipFile(archive_path, 'r') as zf:
                                for rom in roms:
                                    try:
                                        zf.extract(rom, dest_folder)
                                        extracted_count += 1
                                        results_text.insert("end", f"   ✅ {Path(rom).name}\n")
                                    except Exception as e:
                                        results_text.insert("end", f"   ❌ {Path(rom).name}: {e}\n")
                                        error_count += 1
                        
                        elif ext in ['.tar', '.gz', '.tgz'] or archive_path.name.endswith('.tar.gz'):
                            mode = 'r:gz' if ext in ['.gz', '.tgz'] or archive_path.name.endswith('.tar.gz') else 'r'
                            with tarfile.open(archive_path, mode) as tf:
                                for rom in roms:
                                    try:
                                        member = tf.getmember(rom)
                                        tf.extract(member, dest_folder)
                                        extracted_count += 1
                                        results_text.insert("end", f"   ✅ {Path(rom).name}\n")
                                    except Exception as e:
                                        results_text.insert("end", f"   ❌ {Path(rom).name}: {e}\n")
                                        error_count += 1
                        
                        elif ext in ['.7z', '.rar'] and self.seven_zip_path:
                            for rom in roms:
                                try:
                                    cmd = [self.seven_zip_path, 'e', str(archive_path), 
                                           f'-o{dest_folder}', rom, '-y']
                                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
                                    if result.returncode == 0:
                                        extracted_count += 1
                                        results_text.insert("end", f"   ✅ {Path(rom).name}\n")
                                    else:
                                        results_text.insert("end", f"   ❌ {Path(rom).name}: 7z error\n")
                                        error_count += 1
                                except Exception as e:
                                    results_text.insert("end", f"   ❌ {Path(rom).name}: {e}\n")
                                    error_count += 1
                        
                        dialog.update()
                        
                    except Exception as e:
                        results_text.insert("end", f"   ❌ Archive error: {archive.name}: {e}\n")
                        error_count += 1
            
            # Delete archives if option is enabled
            if delete_after_extract.get() and extracted_count > 0:
                results_text.insert("end", f"\n{'━' * 50}\n")
                results_text.insert("end", "Deleting extracted archives...\n")
                deleted_archives = set()
                for system, archive_list in archives_by_system.items():
                    if system in self.system_extract_dirs:
                        for archive, roms in archive_list:
                            if archive not in deleted_archives:
                                try:
                                    archive.unlink()
                                    results_text.insert("end", f"   🗑️ {archive.name}\n")
                                    deleted_archives.add(archive)
                                except Exception as e:
                                    results_text.insert("end", f"   ⚠️ Could not delete {archive.name}: {e}\n")
            
            results_text.insert("end", f"\n{'━' * 50}\n")
            results_text.insert("end", f"✅ Extracted: {extracted_count} ROM(s)\n")
            if skipped_count > 0:
                results_text.insert("end", f"⏭️ Skipped: {skipped_count} ROM(s) (no folder configured)\n")
            if error_count > 0:
                results_text.insert("end", f"❌ Errors: {error_count}\n")
            
            messagebox.showinfo("Extraction Complete", 
                              f"Extracted {extracted_count} ROM(s)\n" +
                              (f"Skipped {skipped_count} ROM(s)\n" if skipped_count > 0 else "") +
                              (f"Errors: {error_count}" if error_count > 0 else ""))
        
        # Action buttons
        action_frame = Frame(dialog, padx=10, pady=10, bg=COLORS['bg_dark'])
        action_frame.pack(fill="x")
        
        Button(action_frame, text="▶ SCAN ARCHIVES", command=scan_archives,
               font=self.font_button,
               bg=COLORS['button_green'], fg=COLORS['bg_dark'],
               activebackground=COLORS['text_primary'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="left", padx=5)
        
        Button(action_frame, text="🔀 REASSIGN ROM", command=reassign_rom_dialog,
               font=self.font_button,
               bg=COLORS['button_blue'], fg="white",
               activebackground=COLORS['text_secondary'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="left", padx=5)
        
        Button(action_frame, text="📁 ORGANIZE ARCHIVES", command=organize_archives_to_folders,
               font=self.font_button,
               bg=COLORS['accent_purple'], fg="white",
               activebackground=COLORS['accent_pink'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="left", padx=5)
        
        Button(action_frame, text="📦 EXTRACT BY SYSTEM", command=extract_to_system_folders,
               font=self.font_button,
               bg=COLORS['accent_yellow'], fg=COLORS['bg_dark'],
               activebackground=COLORS['accent_orange'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="left", padx=5)
        
        Button(action_frame, text="✕ CLOSE", command=dialog.destroy,
               font=self.font_button,
               activebackground=COLORS['accent_red'],
               relief="flat", cursor="hand2", padx=15, pady=5).pack(side="right", padx=5)


def main():
    # Set multiprocessing start method for Windows to ensure all cores are utilized
    # This prevents issues with the default 'spawn' method on Windows
    if os.name == 'nt':  # Windows
        try:
            multiprocessing.set_start_method('spawn', force=True)
        except RuntimeError:
            # Already set, ignore
            pass
    
    root = Tk()
    app = ROMConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
