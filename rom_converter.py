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
from pathlib import Path
from tkinter import Tk, Frame, Label, Button, Entry, Text, Scrollbar, Checkbutton, BooleanVar, filedialog, messagebox, Toplevel
from tkinter import ttk
import tkinter.font as tkfont
import threading
import re
import multiprocessing
from concurrent.futures import ThreadPoolExecutor, as_completed
from queue import Queue
import time
import threading as pythread
try:
    import psutil  # For CPU, memory, disk usage
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

# MAME download configuration
MAME_RELEASE_URL = "https://www.mamedev.org/release.html"
MAME_GITHUB_RELEASES_API = "https://api.github.com/repos/mamedev/mame/releases/latest"

# Supported compressed file extensions
COMPRESSED_EXTENSIONS = {'.zip', '.7z', '.rar', '.gz', '.tar', '.tar.gz', '.tgz'}

# Supported PS2 output formats
PS2_OUTPUT_FORMATS = ['CHD', 'CSO', 'ZSO']

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
        'bg_dark': '#0d0d11', 'bg_medium': '#151520', 'bg_light': '#1f1f2b', 'bg_input': '#191927',
        'text_primary': '#f5f5f5', 'text_secondary': '#9ad0ff', 'text_muted': '#707788',
        'accent_pink': '#ff65a3', 'accent_purple': '#c792ea', 'accent_yellow': '#ffd166',
        'accent_orange': '#ff9f1c', 'accent_red': '#ff5c8d', 'button_green': '#6ce37e',
        'button_blue': '#6fa8ff', 'scanline': '#ffffff08',
        'font_body': 'Courier New', 'font_heading': 'Courier New', 'font_mono': 'Courier New'
    },
    'PS2': {
        'bg_dark': '#0a0a12', 'bg_medium': '#12121a', 'bg_light': '#1a1a2e', 'bg_input': '#16213e',
        'text_primary': '#00ff88', 'text_secondary': '#00ccff', 'text_muted': '#5c6b7a',
        'accent_pink': '#ff00aa', 'accent_purple': '#9945ff', 'accent_yellow': '#ffdd00',
        'accent_orange': '#ff6b35', 'accent_red': '#ff3366', 'button_green': '#00cc66',
        'button_blue': '#0088ff', 'scanline': '#ffffff08',
        'font_body': 'Consolas', 'font_heading': 'Consolas', 'font_mono': 'Consolas'
    },
    'PS3': {
        'bg_dark': '#0b0b0f', 'bg_medium': '#14141c', 'bg_light': '#1d1d26', 'bg_input': '#181823',
        'text_primary': '#e8ecf1', 'text_secondary': '#6fc3ff', 'text_muted': '#7d8799',
        'accent_pink': '#ff6fa1', 'accent_purple': '#8c7bff', 'accent_yellow': '#ffd479',
        'accent_orange': '#ff8f5a', 'accent_red': '#ff5f6d', 'button_green': '#4ade80',
        'button_blue': '#3ea7ff', 'scanline': '#ffffff08',
        'font_body': 'Segoe UI', 'font_heading': 'Segoe UI Semibold', 'font_mono': 'Consolas'
    },
    'PS4': {
        'bg_dark': '#0b132b', 'bg_medium': '#1c2541', 'bg_light': '#243b55', 'bg_input': '#1a233b',
        'text_primary': '#f2f7ff', 'text_secondary': '#89c2ff', 'text_muted': '#7a8699',
        'accent_pink': '#ff7bba', 'accent_purple': '#8d7bff', 'accent_yellow': '#ffd166',
        'accent_orange': '#f8961e', 'accent_red': '#ef476f', 'button_green': '#5be7a9',
        'button_blue': '#3a86ff', 'scanline': '#ffffff08',
        'font_body': 'Segoe UI', 'font_heading': 'Segoe UI Semibold', 'font_mono': 'Consolas'
    },
    'PS5': {
        'bg_dark': '#0f1419', 'bg_medium': '#1a1f2e', 'bg_light': '#252d3d', 'bg_input': '#1f2735',
        'text_primary': '#e8ecf1', 'text_secondary': '#60a5fa', 'text_muted': '#6b7280',
        'accent_pink': '#ec4899', 'accent_purple': '#a855f7', 'accent_yellow': '#fbbf24',
        'accent_orange': '#f97316', 'accent_red': '#ef4444', 'button_green': '#10b981',
        'button_blue': '#3b82f6', 'scanline': '#ffffff08',
        'font_body': 'Segoe UI', 'font_heading': 'Segoe UI Semibold', 'font_mono': 'Consolas'
    },
    'PSP': {
        'bg_dark': '#0d0f14', 'bg_medium': '#161922', 'bg_light': '#1f2430', 'bg_input': '#191e29',
        'text_primary': '#f0f4ff', 'text_secondary': '#5dd4ff', 'text_muted': '#7e8899',
        'accent_pink': '#ff6fb7', 'accent_purple': '#9d7bff', 'accent_yellow': '#ffd166',
        'accent_orange': '#ff9f1c', 'accent_red': '#ff5f6d', 'button_green': '#4ade80',
        'button_blue': '#4aa3ff', 'scanline': '#ffffff08',
        'font_body': 'Tahoma', 'font_heading': 'Tahoma', 'font_mono': 'Consolas'
    },
    'PSVita': {
        'bg_dark': '#0b1024', 'bg_medium': '#131b33', 'bg_light': '#1c2540', 'bg_input': '#16203a',
        'text_primary': '#eef3ff', 'text_secondary': '#7cd2ff', 'text_muted': '#7c89a3',
        'accent_pink': '#ff7ba5', 'accent_purple': '#8b7bff', 'accent_yellow': '#ffd479',
        'accent_orange': '#ff9e5a', 'accent_red': '#ff6b81', 'button_green': '#5de0a4',
        'button_blue': '#4d9dff', 'scanline': '#ffffff08',
        'font_body': 'Segoe UI', 'font_heading': 'Segoe UI Semibold', 'font_mono': 'Consolas'
    },
}

# Active colors (will be set from selected theme)
COLORS = dict(THEME_PRESETS['PS2'])


class ROMConverter:
    def __init__(self, master):
        self.master = master
        master.title("⚡ ROM CONVERTER ⚡")
        master.geometry("900x1200")
        master.resizable(True, True)
        master.configure(bg=COLORS['bg_dark'])
        
        # Maximize window on startup
        master.state('zoomed')
        
        # Config file location (portable: lives beside the app)
        # Handle PyInstaller's temporary folder for bundled resources
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            self.script_dir = Path(sys.executable).parent.resolve()
            # PyInstaller extracts bundled files to sys._MEIPASS
            self.bundle_dir = Path(getattr(sys, '_MEIPASS', self.script_dir))
        else:
            # Running as script
            self.script_dir = Path(__file__).parent.resolve()
            self.bundle_dir = self.script_dir
        
        self.config_file = self.script_dir / ".rom_converter_config.json"
        
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
        self.ram_threshold_percent = 90  # Throttle if RAM usage exceeds this
        self.cpu_threshold_percent = 95  # Throttle if CPU usage exceeds this
        self.log_queue = Queue()
        self.total_original_size = 0
        self.total_chd_size = 0
        self.process_ps1_cues = BooleanVar(value=False)  # Toggle for PS1 CUE processing
        self.process_ps2_cues = BooleanVar(value=False)  # Toggle for PS2 BIN/CUE processing (CD-based games)
        self.process_ps2_isos = BooleanVar(value=False)  # Toggle for PS2 ISO processing
        self.extract_compressed = BooleanVar(value=True)  # Toggle for extracting compressed files
        self.delete_archives_after_extract = BooleanVar(value=False)  # Delete archives after extraction
        self.seven_zip_path = None  # Path to 7z executable for .7z and .rar files
        self.maxcso_path = None  # Path to maxcso executable for CSO/ZSO
        self.ps2_output_format = 'CHD'  # Default PS2 output format
        self.ps2_emulator = 'PCSX2'  # Default emulator preference
        self.current_theme = 'PS2'  # Default UI theme
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
        self.chdman_path = None  # Will store path to chdman executable
        
        # Progress tracking for crash recovery
        self.progress_file = self.script_dir / ".rom_converter_progress.json"
        self.completed_files = set()  # Track completed conversions
        self.current_batch_id = None
        
        # Load saved configuration
        self.load_config()
        self.load_progress()

        # Apply theme colors before UI construction
        self.set_theme_colors(self.current_theme)
        self.init_fonts()
        
        # Check for 7-Zip (for .7z and .rar support)
        self.check_7zip()

        # Check for maxcso (for CSO/ZSO output)
        self.check_maxcso()
        
        # Check for chdman
        if not self.check_chdman():
            # Offer to download MAME tools or manual selection
            response = messagebox.askyesnocancel(
                "chdman Not Found",
                "chdman not found!\n\n"
                "Would you like to download MAME tools automatically?\n\n"
                "Yes = Download from mamedev.org\n"
                "No = Manually select chdman.exe\n"
            )
            if response is True:  # Yes - download
                if self.download_mame_tools():
                    if not self.check_chdman():
                        messagebox.showerror("Error", "Failed to find chdman after download")
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
        
        self.setup_ui()
    
    def check_chdman(self):
        """Check if chdman is available"""
        # First check bundled resources (PyInstaller)
        bundled_chdman = self.bundle_dir / "chdman.exe"
        if bundled_chdman.exists():
            self.chdman_path = str(bundled_chdman)
            return True
        
        # Then check for chdman.exe directly next to the executable/script
        direct_chdman = self.script_dir / "chdman.exe"
        if direct_chdman.exists():
            self.chdman_path = str(direct_chdman)
            return True
        
        # Check PATH as fallback
        chdman = shutil.which("chdman")
        if chdman:
            self.chdman_path = chdman
            return True
        
        return False
    
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
                response = messagebox.askyesno(
            direct_maxcso = self.script_dir / "maxcso.exe"
                    f"Installed: MAME {installed_version[0]}.{installed_version[1:]}\n"
                    f"Latest: MAME {latest_version[0]}.{latest_version[1:]}\n\n"
                    "Would you like to update now?"
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
        """Download and extract MAME tools from mamedev.org"""
        try:
            # Get latest version
            version = self.get_latest_mame_version()
            if not version:
                messagebox.showerror("Error", "Could not determine latest MAME version")
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
            
            # Show progress dialog
            progress_window = Tk()
            progress_window.title("Downloading MAME Tools")
            progress_window.geometry("400x150")
            progress_window.resizable(False, False)
            
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
                messagebox.showerror("Error", "Could not download MAME tools from any mirror")
                return False
            
            # Extract using 7-Zip or the self-extracting exe
            status_label.config(text="Extracting chdman.exe...")
            progress_window.update()
            
            # The MAME exe is a self-extracting 7z archive
            # We need 7-Zip to extract just chdman.exe, or run with specific args
            extracted = False
            
            if self.seven_zip_path:
                # Use 7-Zip to extract only chdman.exe directly to script directory
                try:
                    cmd = [
                        self.seven_zip_path, 'e', str(download_path),
                        '-o' + str(script_dir),
                        'chdman.exe',
                        '-y'
                    ]
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
                    if result.returncode == 0 and (script_dir / "chdman.exe").exists():
                        extracted = True
                except Exception as e:
                    print(f"7-Zip extraction error: {e}")
            
            if not extracted:
                # Try running as self-extracting archive with output directory
                try:
                    cmd = [str(download_path), '-o' + str(temp_dir), '-y']
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
                    # Move chdman.exe to script directory
                    temp_chdman = temp_dir / "chdman.exe"
                    if temp_chdman.exists():
                        shutil.move(str(temp_chdman), str(script_dir / "chdman.exe"))
                        extracted = True
                except Exception as e:
                    print(f"Self-extraction error: {e}")
            
            progress_window.destroy()
            
            # Clean up the temp directory and download file
            if temp_dir.exists():
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass
            
            if extracted and (script_dir / "chdman.exe").exists():
                self.chdman_path = str(script_dir / "chdman.exe")
                self.save_config()
                messagebox.showinfo("Success", f"chdman.exe downloaded successfully!\n\nLocation:\n{self.chdman_path}")
                return True
            else:
                messagebox.showerror(
                    "Extraction Failed",
                    "Could not extract chdman.exe from MAME package.\n\n"
                    "Please install 7-Zip and try again, or manually download MAME from:\n"
                    "https://www.mamedev.org/release.html"
                )
                return False
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to download MAME tools:\n{e}")
            return False
    
    def check_7zip(self):
        """Check if 7-Zip is available"""
        # Check common locations on Windows
        common_paths = [
            r"C:\Program Files\7-Zip\7z.exe",
            r"C:\Program Files (x86)\7-Zip\7z.exe",
        ]
        
        # Check PATH first
        seven_zip = shutil.which("7z")
        if seven_zip:
            self.seven_zip_path = seven_zip
            return True
        
        # Check common install locations
        for path in common_paths:
            if os.path.exists(path):
                self.seven_zip_path = path
                return True
        
        return False

    def check_maxcso(self):
        """Check if maxcso is available for CSO/ZSO output"""
        # First check bundled resources (PyInstaller)
        bundled_maxcso = self.bundle_dir / "maxcso.exe"
        if bundled_maxcso.exists():
            self.maxcso_path = str(bundled_maxcso)
            return True
        
        # Then check for maxcso.exe directly next to the executable/script
        direct_maxcso = self.script_dir / "maxcso.exe"
        if direct_maxcso.exists():
            self.maxcso_path = str(direct_maxcso)
            return True

        # Check PATH as fallback
        maxcso = shutil.which("maxcso")
        if maxcso:
            self.maxcso_path = maxcso
            return True

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
                    messagebox.showinfo("Success", f"maxcso location set to:\n{maxcso_file}")
                else:
                    messagebox.showerror("Error", "Selected file does not appear to be maxcso")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to verify maxcso:\n{e}")
    
    def show_maxcso_setup_help(self):
        """Show a detailed setup guide for maxcso"""
        help_window = Toplevel(self.root)
        help_window.title("maxcso Setup Guide")
        help_window.geometry("600x500")
        
        # Apply theme to window
        help_window.configure(bg=COLORS['bg_dark'])
        
        # Create main frame
        main_frame = Frame(help_window, bg=COLORS['bg_dark'])
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        # Title
        title = Label(main_frame, text="maxcso Setup Guide", font=self.font_title,
                      fg=COLORS['accent_yellow'], bg=COLORS['bg_dark'])
        title.pack(anchor="w", pady=(0, 10))
        
        # Info text
        info_text = Text(main_frame, font=self.font_body, fg=COLORS['text_primary'],
                        bg=COLORS['bg_light'], wrap="word", height=20, relief="flat", padx=10, pady=10)
        info_text.pack(fill="both", expand=True, pady=(0, 10))
        
        # Disable editing
        info_text.config(state="disabled")
        
        # Add help content
        help_content = """maxcso is required for CSO and ZSO output formats.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OPTION 1: Place in Program Directory (Recommended)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. Click [ Open GitHub Releases ] below
2. Download the latest "maxcso.exe" file
3. Place maxcso.exe in the same folder as rom_converter.py
4. Restart ROM Converter - it will auto-detect it


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OPTION 2: Manual Selection with [ SET ] Button
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. Download maxcso.exe from GitHub releases
2. Save to any location on your computer
3. Click [ SET ] and browse to select maxcso.exe


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OPTION 3: Add to System PATH (Advanced)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. Download maxcso.exe and place in a folder
2. Add that folder to Windows PATH
3. Restart ROM Converter


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Format Recommendations
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

• CHD: High compatibility, slightly larger files
• CSO: Good compression, fast decompression
• ZSO: Best compression, lower CPU usage
"""
        
        info_text.config(state="normal")
        info_text.insert("1.0", help_content)
        info_text.config(state="disabled")
        
        # Button frame
        button_frame = Frame(main_frame, bg=COLORS['bg_dark'])
        button_frame.pack(fill="x", pady=(0, 0))
        
        Button(button_frame, text="[ Open GitHub Releases ]", 
               command=lambda: self.open_maxcso_releases(),
               font=self.font_small, bg=COLORS['button_blue'], fg="white",
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
                    # Update UI label if it exists
                    if hasattr(self, 'chdman_label'):
                        self.chdman_label.config(text=self.chdman_path)
                    messagebox.showinfo("Success", f"chdman location set to:\n{chdman_file}")
                else:
                    messagebox.showerror("Error", "Selected file does not appear to be chdman")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to verify chdman:\n{e}")
    
    def save_config(self):
        """Save configuration to JSON file"""
        try:
            config = {
                'source_dir': self._make_portable_path(self.source_dir),
                'delete_originals': self.delete_originals.get(),
                'move_to_backup': self.move_to_backup.get(),
                'recursive': self.recursive.get(),
                'process_ps1_cues': self.process_ps1_cues.get(),
                'process_ps2_cues': self.process_ps2_cues.get(),
                'process_ps2_isos': self.process_ps2_isos.get(),
                'extract_compressed': self.extract_compressed.get(),
                'delete_archives_after_extract': self.delete_archives_after_extract.get(),
                'chdman_path': self._make_portable_path(self.chdman_path),
                'seven_zip_path': self._make_portable_path(self.seven_zip_path),
                'maxcso_path': self._make_portable_path(self.maxcso_path),
                'ps2_output_format': self.ps2_output_format,
                'ps2_emulator': self.ps2_emulator,
                'theme': self.current_theme
            }
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=2)
        except Exception as e:
            # Silently fail - don't interrupt user experience
            pass
    
    def load_config(self):
        """Load configuration from JSON file"""
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
                self.extract_compressed.set(config.get('extract_compressed', True))
                self.delete_archives_after_extract.set(config.get('delete_archives_after_extract', False))
                self.ps2_output_format = config.get('ps2_output_format', 'CHD')
                self.ps2_emulator = config.get('ps2_emulator', 'PCSX2')
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
        self.font_button = tkfont.Font(family=self.font_body_family, size=11, weight="bold")
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
        """Apply current theme colors across the UI"""
        self.update_font_families()
        # Update ttk progress style
        style = ttk.Style()
        style.configure("Retro.Horizontal.TProgressbar",
                        troughcolor=COLORS['bg_light'],
                        background=COLORS['text_primary'],
                        darkcolor=COLORS['button_green'],
                        lightcolor=COLORS['text_primary'],
                        bordercolor=COLORS['text_primary'])

        # Window background
        self.master.configure(bg=COLORS['bg_dark'])
        for widget in [self.main_frame, getattr(self, 'title_frame', None), getattr(self, 'options_frame', None)]:
            if widget:
                widget.configure(bg=COLORS['bg_dark'] if widget is self.main_frame else COLORS['bg_light'])

        # Header labels
        for lbl in [getattr(self, 'status_label', None), getattr(self, 'metrics_label', None)]:
            if lbl:
                lbl.configure(bg=COLORS['bg_light'] if lbl is self.status_label else COLORS['bg_medium'],
                              fg=COLORS['text_primary'] if lbl is self.status_label else COLORS['accent_yellow'])

        # Inputs and labels
        for inp in [getattr(self, 'dir_entry', None)]:
            if inp:
                inp.configure(bg=COLORS['bg_input'], fg=COLORS['text_primary'], insertbackground=COLORS['text_primary'])
        for lbl in [getattr(self, 'chdman_label', None), getattr(self, 'seven_zip_label', None), getattr(self, 'maxcso_label', None)]:
            if lbl:
                lbl.configure(bg=COLORS['bg_dark'], fg=COLORS['text_secondary'])

        # Buttons
        buttons = [getattr(self, 'scan_button', None), getattr(self, 'convert_button', None),
                   getattr(self, 'stop_button', None), getattr(self, 'move_chd_button', None)]
        for btn in buttons:
            if btn:
                btn.configure(activebackground=COLORS['text_primary'])
        # Log area
        if getattr(self, 'log_text', None):
            self.log_text.configure(bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                                    insertbackground=COLORS['text_primary'],
                                    selectbackground=COLORS['accent_purple'])
        # Progress bar
        if getattr(self, 'progress', None):
            self.progress.configure(style="Retro.Horizontal.TProgressbar")

        # Title frame background
        if getattr(self, 'title_frame', None):
            self.title_frame.configure(bg=COLORS['bg_light'])
            for child in self.title_frame.winfo_children():
                try:
                    child.configure(bg=COLORS['bg_light'], fg=COLORS['text_secondary'])
                except Exception:
                    pass
    
    def setup_ui(self):
        """Setup the user interface with retro gaming aesthetic"""
        # Configure ttk styles for retro look
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Retro.Horizontal.TProgressbar",
                       troughcolor=COLORS['bg_light'],
                       background=COLORS['text_primary'],
                       darkcolor=COLORS['button_green'],
                       lightcolor=COLORS['text_primary'],
                       bordercolor=COLORS['text_primary'])
        
        # Main container with dark background
        self.main_frame = Frame(self.master, padx=15, pady=15, bg=COLORS['bg_dark'])
        self.main_frame.pack(fill="both", expand=True)
        
        # Title banner
        title_frame = Frame(self.main_frame, bg=COLORS['bg_light'], pady=8)
        title_frame.pack(fill="x", pady=(0, 15))
        self.title_frame = title_frame
        
        title_label = Label(title_frame, text="◄ ROM CONVERTER ►", 
                   font=self.font_title,
                           fg=COLORS['text_primary'], bg=COLORS['bg_light'])
        title_label.pack()

        # Theme selector
        theme_frame = Frame(title_frame, bg=COLORS['bg_light'])
        theme_frame.pack(pady=(6, 0))
        Label(theme_frame, text="Theme:", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=COLORS['bg_light']).pack(side="left", padx=(0, 6))
        self.theme_combo = ttk.Combobox(theme_frame, values=list(THEME_PRESETS.keys()),
                                        state="readonly", width=8)
        if self.current_theme not in THEME_PRESETS:
            self.current_theme = 'PS2'
        self.theme_combo.set(self.current_theme)
        self.theme_combo.pack(side="left")

        def on_theme_change(event=None):
            chosen = self.theme_combo.get()
            self.set_theme_colors(chosen)
            self.update_font_families()
            self.save_config()
            self.apply_theme()
            self.log(f"🖌 Theme set to {chosen}.")

        self.theme_combo.bind("<<ComboboxSelected>>", on_theme_change)
        
        # Directory selection
        dir_frame = Frame(self.main_frame, bg=COLORS['bg_dark'])
        dir_frame.pack(fill="x", pady=(0, 8))
        
        Label(dir_frame, text="📁 ROM Directory:", font=self.font_label_bold,
              fg=COLORS['text_primary'], bg=COLORS['bg_dark']).pack(side="left", padx=(0, 10))
        
        self.dir_entry = Entry(dir_frame, font=self.font_body,
                              bg=COLORS['bg_input'], fg=COLORS['text_primary'],
                              insertbackground=COLORS['text_primary'],
                              relief="flat", highlightthickness=1,
                              highlightcolor=COLORS['text_secondary'],
                              highlightbackground=COLORS['text_muted'])
        self.dir_entry.pack(side="left", fill="x", expand=True, padx=(0, 10), ipady=4)
        if self.source_dir:
            self.dir_entry.insert(0, self.source_dir)
        
        Button(dir_frame, text="[ BROWSE ]", command=self.browse_directory,
               font=self.font_small, bg=COLORS['bg_light'],
               fg=COLORS['text_secondary'], activebackground=COLORS['accent_purple'],
               activeforeground="white", relief="flat", cursor="hand2").pack(side="left")
        
        # chdman location
        chdman_frame = Frame(self.main_frame, bg=COLORS['bg_dark'])
        chdman_frame.pack(fill="x", pady=(0, 8))
        
        Label(chdman_frame, text="⚙ chdman:", font=self.font_label_bold,
              fg=COLORS['accent_yellow'], bg=COLORS['bg_dark']).pack(side="left", padx=(0, 10))
        self.chdman_label = Label(chdman_frame, text=self.chdman_path or "Not set",
                                  font=self.font_small,
                                  fg=COLORS['text_secondary'], bg=COLORS['bg_dark'], anchor="w")
        self.chdman_label.pack(side="left", fill="x", expand=True, padx=(0, 10))
        Button(chdman_frame, text="[ CHANGE ]", command=self.browse_chdman,
               font=self.font_small, bg=COLORS['button_blue'],
               fg="white", activebackground=COLORS['accent_purple'],
               activeforeground="white", relief="flat", cursor="hand2").pack(side="left")
        
        # 7-Zip location
        seven_zip_frame = Frame(self.main_frame, bg=COLORS['bg_dark'])
        seven_zip_frame.pack(fill="x", pady=(0, 12))
        
        Label(seven_zip_frame, text="📦 7-Zip:", font=self.font_label_bold,
              fg=COLORS['accent_yellow'], bg=COLORS['bg_dark']).pack(side="left", padx=(0, 10))
        self.seven_zip_label = Label(seven_zip_frame, 
                                     text=self.seven_zip_path or "Not set (optional for .7z/.rar)",
                                     font=self.font_small,
                                     fg=COLORS['text_secondary'] if self.seven_zip_path else COLORS['text_muted'],
                                     bg=COLORS['bg_dark'], anchor="w")
        self.seven_zip_label.pack(side="left", fill="x", expand=True, padx=(0, 10))
        Button(seven_zip_frame, text="[ SET ]", command=self.browse_7zip,
               font=self.font_small, bg=COLORS['button_blue'],
               fg="white", activebackground=COLORS['accent_purple'],
               activeforeground="white", relief="flat", cursor="hand2").pack(side="left")

        # maxcso location (for CSO/ZSO output)
        maxcso_frame = Frame(self.main_frame, bg=COLORS['bg_dark'])
        maxcso_frame.pack(fill="x", pady=(0, 12))

        Label(maxcso_frame, text="🗜  maxcso:", font=self.font_label_bold,
              fg=COLORS['accent_yellow'], bg=COLORS['bg_dark']).pack(side="left", padx=(0, 10))
        self.maxcso_label = Label(maxcso_frame, 
                        text=self.maxcso_path or "Not set (required for CSO/ZSO)",
                        font=self.font_small,
                                   fg=COLORS['text_secondary'] if self.maxcso_path else COLORS['text_muted'],
                                   bg=COLORS['bg_dark'], anchor="w")
        self.maxcso_label.pack(side="left", fill="x", expand=True)
        
        # Options panel
        options_frame = Frame(self.main_frame, bg=COLORS['bg_light'], padx=10, pady=8)
        options_frame.pack(fill="x", pady=(0, 12))
        self.options_frame = options_frame
        
        options_title = Label(options_frame, text="▼ OPTIONS ▼", font=self.font_label_bold,
                             fg=COLORS['accent_pink'], bg=COLORS['bg_light'])
        options_title.pack(anchor="w", pady=(0, 5))
        
        # Custom checkbox style
        cb_font = self.font_small
        cb_bg = COLORS['bg_light']
        
        Checkbutton(options_frame, text="↳ Scan subdirectories recursively",
                   variable=self.recursive, font=cb_font,
                   fg=COLORS['text_primary'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                   activebackground=cb_bg, activeforeground=COLORS['text_primary']).pack(anchor="w")
        
        Checkbutton(options_frame, text="↳ Move originals to backup folder after conversion",
                   variable=self.move_to_backup, font=cb_font,
                   fg=COLORS['text_secondary'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                   activebackground=cb_bg, activeforeground=COLORS['text_secondary']).pack(anchor="w")
        
        Checkbutton(options_frame, text="⚠ Delete original files after successful conversion",
                   variable=self.delete_originals, font=cb_font,
                   fg=COLORS['accent_red'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                   activebackground=cb_bg, activeforeground=COLORS['accent_red']).pack(anchor="w")

        Checkbutton(options_frame, text="🎮 Process PS1 CUE files (.cue)",
                variable=self.process_ps1_cues, font=cb_font,
                fg=COLORS['button_green'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                activebackground=cb_bg, activeforeground=COLORS['button_green']).pack(anchor="w")

        Checkbutton(options_frame, text="🎮 Process PS2 BIN/CUE files (.cue)",
            variable=self.process_ps2_cues, font=cb_font,
            fg=COLORS['button_green'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
            activebackground=cb_bg, activeforeground=COLORS['button_green']).pack(anchor="w")

        Checkbutton(options_frame, text="🎮 Process PS2 ISO files (.iso)",
                variable=self.process_ps2_isos, font=cb_font,
                fg=COLORS['accent_purple'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                activebackground=cb_bg, activeforeground=COLORS['accent_purple']).pack(anchor="w")

        # Emulator preset selection
        emulator_frame = Frame(options_frame, bg=cb_bg)
        emulator_frame.pack(fill="x", pady=(4, 2))
        Label(emulator_frame, text="↳ PS2 emulator:", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=cb_bg).pack(side="left")
        self.ps2_emulator_combo = ttk.Combobox(emulator_frame, values=PS2_EMULATORS,
                                               state="readonly", width=10)
        if self.ps2_emulator not in PS2_EMULATORS:
            self.ps2_emulator = 'PCSX2'
        self.ps2_emulator_combo.set(self.ps2_emulator)
        self.ps2_emulator_combo.pack(side="left", padx=8)

        # PS2 output format selector
        format_frame = Frame(options_frame, bg=cb_bg)
        format_frame.pack(fill="x", pady=(4, 4))
        Label(format_frame, text="↳ PS2 output format:", font=self.font_label_bold,
              fg=COLORS['text_secondary'], bg=cb_bg).pack(side="left")
        self.ps2_format_combo = ttk.Combobox(format_frame, values=PS2_OUTPUT_FORMATS,
                            state="readonly", width=6)
        if self.ps2_output_format not in PS2_OUTPUT_FORMATS:
            self.ps2_output_format = 'CHD'
        self.ps2_format_combo.set(self.ps2_output_format)
        self.ps2_format_combo.pack(side="left", padx=8)

        def on_ps2_format_change(event=None):
            self.ps2_output_format = self.ps2_format_combo.get()
            self.save_config()
            if self.ps2_output_format in ['CSO', 'ZSO'] and not self.maxcso_path:
                self.log("⚠ maxcso is required for CSO/ZSO output. Set the path above.")

        self.ps2_format_combo.bind("<<ComboboxSelected>>", on_ps2_format_change)

        def on_ps2_emulator_change(event=None):
            self.ps2_emulator = self.ps2_emulator_combo.get()
            # Apply recommended format for selected emulator
            recommended = PS2_EMULATOR_RECOMMENDATIONS.get(self.ps2_emulator, 'CHD')
            if recommended in PS2_OUTPUT_FORMATS:
                self.ps2_output_format = recommended
                self.ps2_format_combo.set(recommended)
                self.log(f"ℹ Using recommended format for {self.ps2_emulator}: {recommended}")
                self.save_config()
                if recommended in ['CSO', 'ZSO'] and not self.maxcso_path:
                    self.log("⚠ maxcso is required for CSO/ZSO output. Set the path above.")

        self.ps2_emulator_combo.bind("<<ComboboxSelected>>", on_ps2_emulator_change)

        Checkbutton(options_frame, text="📦 Extract compressed files before conversion",
                variable=self.extract_compressed, font=cb_font,
                fg=COLORS['accent_orange'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                activebackground=cb_bg, activeforeground=COLORS['accent_orange']).pack(anchor="w")

        Checkbutton(options_frame, text="⚠ Delete archive files after extraction",
                variable=self.delete_archives_after_extract, font=cb_font,
                fg=COLORS['accent_red'], bg=cb_bg, selectcolor=COLORS['bg_dark'],
                activebackground=cb_bg, activeforeground=COLORS['accent_red']).pack(anchor="w")
        
        # Action buttons
        button_frame = Frame(self.main_frame, bg=COLORS['bg_dark'])
        button_frame.pack(fill="x", pady=(0, 8))
        
        self.scan_button = Button(button_frame, text="▶ SCAN", 
                                 command=self.scan_directory,
                                 font=self.font_button,
                                 bg=COLORS['button_green'], fg=COLORS['bg_dark'],
                                 activebackground=COLORS['text_primary'],
                                 activeforeground=COLORS['bg_dark'],
                                 relief="flat", cursor="hand2", padx=15, pady=5)
        self.scan_button.pack(side="left", padx=(0, 8))
        
        self.convert_button = Button(button_frame, text="⚡ CONVERT", 
                                    command=self.start_conversion,
                                    font=self.font_button,
                                    bg=COLORS['button_blue'], fg="white",
                                    activebackground=COLORS['text_secondary'],
                                    activeforeground=COLORS['bg_dark'],
                                    disabledforeground=COLORS['text_muted'],
                                    relief="flat", cursor="hand2", padx=15, pady=5,
                                    state="disabled")
        self.convert_button.pack(side="left", padx=(0, 8))
        
        self.stop_button = Button(button_frame, text="■ STOP", 
                                 command=self.stop_conversion,
                                 font=self.font_button,
                                 bg=COLORS['accent_red'], fg="white",
                                 activebackground=COLORS['accent_orange'],
                                 disabledforeground="white",
                                 relief="flat", cursor="hand2", padx=15, pady=5,
                                 state="disabled")
        self.stop_button.pack(side="left", padx=(0, 8))
        
        self.move_chd_button = Button(button_frame, text="📁 MOVE CHD", 
                                     command=self.move_chd_files_dialog,
                                     font=self.font_button,
                                     bg=COLORS['accent_purple'], fg="white",
                                     activebackground=COLORS['accent_pink'],
                                     relief="flat", cursor="hand2", padx=15, pady=5)
        self.move_chd_button.pack(side="left", padx=(0, 8))
        
        self.cleanup_button = Button(button_frame, text="🗑️ CLEANUP", 
                                    command=self.cleanup_compressed_dialog,
                                    font=self.font_button,
                                    bg=COLORS['accent_orange'], fg="white",
                                    activebackground=COLORS['accent_red'],
                                    relief="flat", cursor="hand2", padx=15, pady=5)
        self.cleanup_button.pack(side="left", padx=(0, 8))
        
        self.clean_names_button = Button(button_frame, text="✨ CLEAN NAMES", 
                                        command=self.clean_names_dialog,
                                        font=self.font_button,
                                        bg=COLORS['accent_pink'], fg="white",
                                        activebackground=COLORS['accent_purple'],
                                        relief="flat", cursor="hand2", padx=15, pady=5)
        self.clean_names_button.pack(side="left")
        
        # Progress bar with retro style
        progress_frame = Frame(self.main_frame, bg=COLORS['bg_dark'])
        progress_frame.pack(fill="x", pady=(0, 8))
        
        self.progress = ttk.Progressbar(progress_frame, mode='determinate',
                                        style="Retro.Horizontal.TProgressbar")
        self.progress.pack(fill="x", ipady=3)
        
        # Log area with terminal aesthetic
        log_label = Label(self.main_frame, text="◄ TERMINAL OUTPUT ►", anchor="w",
                         font=self.font_label_bold,
                         fg=COLORS['text_secondary'], bg=COLORS['bg_dark'])
        log_label.pack(fill="x", pady=(0, 4))
        
        log_frame = Frame(self.main_frame, bg=COLORS['bg_dark'])
        log_frame.pack(fill="both", expand=True)
        
        scrollbar = Scrollbar(log_frame, bg=COLORS['bg_light'],
                             troughcolor=COLORS['bg_dark'],
                             activebackground=COLORS['text_primary'])
        scrollbar.pack(side="right", fill="y")
        
        self.log_text = Text(log_frame, wrap="word", yscrollcommand=scrollbar.set,
                            height=1000, font=self.font_mono,
                            bg=COLORS['bg_medium'], fg=COLORS['text_primary'],
                            insertbackground=COLORS['text_primary'],
                            selectbackground=COLORS['accent_purple'],
                            selectforeground="white",
                            relief="flat", padx=8, pady=8)
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.log_text.yview)
        
        # Status bar
        total_cores = multiprocessing.cpu_count()
        self.status_label = Label(self.main_frame, 
                                 text=f"▶ READY | {self.cpu_cores}/{total_cores} CPU CORES | 1 CORE RESERVED",
                                 font=self.font_status,
                                 fg=COLORS['text_primary'], bg=COLORS['bg_light'],
                                 anchor="w", padx=8, pady=4)
        self.status_label.pack(fill="x", pady=(8, 4))

        # Metrics label with retro styling
        self.metrics_label = Label(self.main_frame, text="◆ METRICS: IDLE ◆", anchor="w", 
                                   bg=COLORS['bg_medium'], fg=COLORS['accent_yellow'],
                                   font=self.font_status, padx=8, pady=4)
        self.metrics_label.pack(fill="x")

        if not PSUTIL_AVAILABLE:
            self.log("ℹ Resource metrics disabled (psutil not installed - this is optional)")
        
        # Add trace callbacks to save config when options change
        self.delete_originals.trace_add('write', lambda *args: self.save_config())
        self.move_to_backup.trace_add('write', lambda *args: self.save_config())
        self.recursive.trace_add('write', lambda *args: self.save_config())
        self.process_ps1_cues.trace_add('write', lambda *args: self.save_config())
        self.process_ps2_cues.trace_add('write', lambda *args: self.save_config())
        self.process_ps2_isos.trace_add('write', lambda *args: self.save_config())
        self.extract_compressed.trace_add('write', lambda *args: self.save_config())
        self.delete_archives_after_extract.trace_add('write', lambda *args: self.save_config())
        
        # Start log queue processor
        self.process_log_queue()
        # Apply theme after UI construction
        self.apply_theme()
    
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
        """Process queued log messages from threads"""
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.insert("end", message + "\n")
                self.log_text.see("end")
        except:
            pass
        finally:
            self.master.after(100, self.process_log_queue)
    
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
    
    def extract_archive(self, archive_path):
        """Extract a compressed archive to a folder with the same name"""
        archive_path = Path(archive_path)
        
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
                return True, extract_folder
            
            # Handle .tar, .tar.gz, .tgz files using Python's tarfile
            elif ext in ['.tar', '.gz', '.tgz'] or archive_path.name.endswith('.tar.gz'):
                self.log(f"  📦 Extracting TAR: {archive_path.name}")
                mode = 'r:gz' if ext in ['.gz', '.tgz'] or archive_path.name.endswith('.tar.gz') else 'r'
                with tarfile.open(archive_path, mode) as tar_ref:
                    tar_ref.extractall(extract_folder)
                self.log(f"  ✅ Extracted to: {extract_folder.name}/")
                return True, extract_folder
            
            # Handle .7z and .rar files using 7-Zip
            elif ext in ['.7z', '.rar']:
                if not self.seven_zip_path:
                    self.log(f"  ⚠️  Cannot extract {ext} file: 7-Zip not configured")
                    return False, None
                
                self.log(f"  📦 Extracting with 7-Zip: {archive_path.name}")
                cmd = [self.seven_zip_path, 'x', str(archive_path), f'-o{extract_folder}', '-y']
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=3600)
                
                if result.returncode == 0:
                    self.log(f"  ✅ Extracted to: {extract_folder.name}/")
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
        """Find all supported game descriptor files (.cue and optionally .iso)"""
        path = Path(directory)
        files = set()
        if recursive:
            if self.process_ps1_cues.get():
                files.update(path.rglob("*.cue"))
            if self.process_ps2_cues.get():
                files.update(path.rglob("*.cue"))
            if self.process_ps2_isos.get():
                files.update(path.rglob("*.iso"))
        else:
            if self.process_ps1_cues.get():
                files.update(path.glob("*.cue"))
            if self.process_ps2_cues.get():
                files.update(path.glob("*.cue"))
            if self.process_ps2_isos.get():
                files.update(path.glob("*.iso"))
        # Sort for stable processing order
        return sorted(files)
    
    def parse_cue_file(self, cue_path):
        """Parse CUE file to find associated BIN files"""
        bin_files = []
        cue_dir = cue_path.parent
        
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
                    self.log(f"  WARNING: Referenced BIN file not found: {match}")
        
        except Exception as e:
            self.log(f"  ERROR parsing CUE file: {e}")
        
        return bin_files
    
    def scan_directory(self):
        """Scan directory for CUE files"""
        if not self.source_dir or not os.path.isdir(self.source_dir):
            messagebox.showwarning("Warning", "Please select a valid directory")
            return
        
        self.log("\n" + "="*60)
        self.log("SCANNING FOR GAME FILES...")
        self.log("="*60)
        
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
                ps2_count += 1
                iso_size = game_file.stat().st_size
                size_gb = iso_size / (1024 * 1024 * 1024)
                size_mb = iso_size / (1024 * 1024)
                self.log(f"💿 [PS2] {game_file.name}")
                self.log(f"   Path: {game_file}")
                if size_gb >= 1:
                    self.log(f"   └─ ISO size: {size_gb:.2f} GB")
                else:
                    self.log(f"   └─ ISO size: {size_mb:.1f} MB")
                game_size += iso_size
            total_size += game_size
            self.log("")

        total_size_mb = total_size / (1024 * 1024)
        total_size_gb = total_size / (1024 * 1024 * 1024)

        self.log(f"Totals: PS1: {ps1_count}  PS2: {ps2_count}  Combined: {ps1_count + ps2_count}")
        if compressed_count > 0:
            self.log(f"📦 Plus {compressed_count} compressed file(s) to extract")
        if total_size_gb >= 1:
            self.log(f"💾 Current total size: {total_size_gb:.2f} GB ({total_size_mb:.1f} MB)")
        else:
            self.log(f"💾 Current total size: {total_size_mb:.1f} MB")

        status_text = f"Found PS1:{ps1_count} PS2:{ps2_count}"
        if compressed_count > 0:
            status_text += f" + {compressed_count} archives"
        status_text += f" | Size: {total_size_gb:.2f} GB"
        self.status_label.config(text=status_text)
        self.convert_button.config(state="normal")
    
    def convert_game(self, path):
        """Convert a game file to the selected output format"""
        ext = path.suffix.lower()

        if ext == '.cue':
            output_path = path.with_suffix('.chd')
            if output_path.exists():
                self.log(f"  ⚠️  CHD already exists, skipping: {output_path.name}")
                return True
            cmd = [self.chdman_path, 'createcd', '-i', str(path), '-o', str(output_path)]
            original_size = sum(f.stat().st_size for f in self.parse_cue_file(path)) + path.stat().st_size
            # Label cues generically since PS1/PS2 CD games both use createcd
            label = 'CD (CUE)'
            format_label = 'CHD'
        elif ext == '.iso':
            # Determine output format for PS2
            fmt = self.ps2_output_format.upper()
            if fmt == 'CHD':
                output_path = path.with_suffix('.chd')
                if output_path.exists():
                    self.log(f"  ⚠️  CHD already exists, skipping: {output_path.name}")
                    return True
                cmd = [self.chdman_path, 'createdvd', '-i', str(path), '-o', str(output_path)]
                format_label = 'CHD'
            elif fmt == 'CSO':
                output_path = path.with_suffix('.cso')
                if output_path.exists():
                    self.log(f"  ⚠️  CSO already exists, skipping: {output_path.name}")
                    return True
                # Use dynamic worker count for maxcso
                workers = self.check_system_resources()
                cmd = [self.maxcso_path, '--threads', str(workers), str(path), str(output_path)]
                format_label = 'CSO'
            elif fmt == 'ZSO':
                output_path = path.with_suffix('.zso')
                if output_path.exists():
                    self.log(f"  ⚠️  ZSO already exists, skipping: {output_path.name}")
                    return True
                # Use dynamic worker count for maxcso
                workers = self.check_system_resources()
                cmd = [self.maxcso_path, '--ziso', '--threads', str(workers), str(path), str(output_path)]
                format_label = 'ZSO'
            else:
                self.log(f"  ❌ Unsupported PS2 format: {fmt}")
                return False

            original_size = path.stat().st_size
            label = 'PS2'
        else:
            self.log(f"  ❌ Unsupported file type: {path.name}")
            return False

        try:
            self.log(f"  Converting ({label} → {format_label}): {path.name} -> {output_path.name}")
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=1800)

            if result.returncode == 0 and output_path.exists():
                new_size = output_path.stat().st_size
                savings = ((original_size - new_size) / original_size) * 100 if original_size > 0 else 0

                # Update totals
                self.total_original_size += original_size
                self.total_chd_size += new_size

                # Track completion for crash recovery
                self.completed_files.add(str(path))
                self.save_progress(self.source_dir)

                self.log(f"  ✅ Success! Saved {savings:.1f}% space")
                if original_size >= 1024*1024*1024:
                    self.log(f"     Original: {original_size / (1024*1024*1024):.2f} GB -> {format_label}: {new_size / (1024*1024*1024):.2f} GB")
                else:
                    self.log(f"     Original: {original_size / (1024*1024):.1f} MB -> {format_label}: {new_size / (1024*1024):.1f} MB")
                return True
            else:
                error_text = result.stderr.strip() or result.stdout.strip()
                self.log(f"  ❌ Conversion failed: {error_text}")
                return False
        except subprocess.TimeoutExpired:
            self.log(f"  ❌ Timeout: Conversion took too long")
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
        # Record start time for metrics
        with self.metrics_lock:
            self.file_start_times[cue_file] = time.time()
        self.log(f"\n[{file_num}/{total}] Processing: {cue_file.name}")
        
        success = self.convert_game(cue_file)
        
        if success:
            if self.delete_originals.get():
                self.delete_original_files(cue_file)
            elif self.move_to_backup.get():
                self.move_to_backup_folder(cue_file)
        
        return success
    
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
            
            # If moderate load, use 3/4 of available cores
            if mem.percent >= 75 or cpu_percent >= 80:
                return max(1, int(self.cpu_cores * 0.75))
            
            # Normal operation - use all allocated cores
            return self.cpu_cores
            
        except Exception:
            return self.cpu_cores
    
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
        self.log(f"Using {self.cpu_cores} of {total_cores} CPU cores (1 core reserved for system)")
        if self.process_ps2_isos.get():
            self.log(f"PS2 emulator: {self.ps2_emulator} | Output format: {self.ps2_output_format}")
        if PSUTIL_AVAILABLE:
            self.log(f"Resource monitoring: Enabled (RAM threshold: {self.ram_threshold_percent}%)")
        self.log("="*60 + "\n")
        
        successful = 0
        failed = 0
        completed = 0
        
        # Dynamic resource management - adjust workers based on system load
        self.max_workers = self.check_system_resources()
        if self.max_workers < self.cpu_cores:
            self.log(f"⚠️  System load detected - using {self.max_workers} workers (throttled)")
        
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
                        progress_value = (completed / total) * 100
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
                             '.gba', '.gbc', '.gb', '.nes', '.snes', '.sfc', '.smc',
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
            
            changes_found = 0
            for rom_file in sorted(rom_files):
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
            
            results_text.delete("1.0", "end")
            results_text.insert("end", "Renaming files...\n\n")
            
            for rom_file, new_name in files_to_rename:
                try:
                    new_path = rom_file.parent / new_name
                    
                    # Check if target already exists
                    if new_path.exists():
                        results_text.insert("end", f"⚠️ SKIP (exists): {new_name}\n")
                        error_count += 1
                        continue
                    
                    rom_file.rename(new_path)
                    results_text.insert("end", f"✅ {rom_file.name} → {new_name}\n")
                    renamed_count += 1
                    
                except Exception as e:
                    results_text.insert("end", f"❌ ERROR: {rom_file.name} - {e}\n")
                    error_count += 1
            
            results_text.insert("end", f"\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n")
            results_text.insert("end", f"✅ Renamed: {renamed_count}\n")
            if error_count > 0:
                results_text.insert("end", f"❌ Errors: {error_count}\n")
            
            files_to_rename.clear()
            messagebox.showinfo("Complete", f"Renamed {renamed_count} file(s).")
        
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
