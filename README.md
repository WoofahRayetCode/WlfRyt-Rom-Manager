# ⚡ ROM Converter

A powerful, cross-platform GUI tool for bulk converting disc game images to optimized formats (CHD, CSO, ZSO). Perfect for preparing ROM collections for emulators.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20Linux-lightgrey)

## ✨ Features

### 🎮 Multi-System Support
- **PlayStation**: PS1, PS2, PS3, PSP, PS Vita
- **Nintendo**: NES, SNES, N64, GameCube, Wii, Switch, 3DS, DS
- **Sega**: Genesis, 32X, Master System, Game Gear, Dreamcast
- **Xbox**: Xbox, Xbox 360
- **Other**: Atari 2600/7800/Lynx, PC Engine, Neo Geo, WonderSwan, Virtual Boy

### 📦 Output Formats
- **PS1/PS2**: CHD, ISO
- **PS2 (extra)**: CSO, ZSO
- **PSP**: CSO, ZSO
- **Nintendo Switch**: XCI, NSP
- *...and many more*

### 🤖 Smart Features
- **Auto-Dependency Management**: Automatically downloads and installs required tools (chdman, maxcso, 7-Zip)
- **Intelligent Resource Throttling**: Auto-detects CPU cores, RAM, and disk I/O limits to prevent system overload
- **Parallel Processing**: Convert multiple ROMs simultaneously with optimized worker threads
- **Crash Recovery**: Tracks progress and automatically resumes interrupted conversions
- **Archive Extraction**: Supports ZIP, 7Z, RAR, TAR, GZ, TAR.GZ
- **3DS Workflow**: Extract → Decrypt → Organize pipeline with automatic key detection
- **Archive.org Integration**: Auto-detects and logs into Archive.org S3 accounts

### 🎨 User Experience
- **Beautiful Dark Themes**: System-themed UI presets (PS1, PS2, PS3, PS4, PS5, PSP, PSVita)
- **Drag-and-Drop Support**: Drop ROM folders directly into the app
- **Real-Time Progress**: Live conversion status with ETA calculations
- **Emulator-Specific Recommendations**: Smart format suggestions (e.g., PCSX2 → CHD, OPL → ZSO)
- **Persistent Configuration**: Settings saved across sessions (portable or home-directory compatible)

### 🖥️ Cross-Platform
- **Windows**: Native EXE with PyInstaller bundling
- **Linux**: AppImage, Flatpak, or direct Python execution
- **Immutable Linux Support**: Auto-detects Fedora Atomic, Bazzite, SteamOS, Kinoite

## 🚀 Quick Start

### Windows

1. **Download** the latest `ROM_Converter.exe` from [Releases](../../releases)
2. **Optional**: Place `chdman.exe` in the same folder (the app will auto-download if missing)
3. **Run** `ROM_Converter.exe`
4. **Select** your ROM folder and output format
5. **Click** "Start Conversion"

> **Note**: On first run, the app will automatically download required tools (chdman, maxcso, 3DS keys)

### Linux (Flatpak)

```bash
# Install MAME (provides chdman)
flatpak install flathub org.mamedev.MAME

# Run ROM Converter (when available)
flatpak run com.rom_converter.ROMConverter
```

### Linux (AppImage)

```bash
# Make executable
chmod +x ROM_Converter-x86_64.AppImage

# Run
./ROM_Converter-x86_64.AppImage
```

### From Source (Any Platform)

```bash
# Clone repository
git clone https://github.com/WoofahRayetCode/WlfRyt-Rom-Manager.git
cd WlfRyt-Rom-Manager

# Install dependencies
pip install -r requirements.txt

# Run
python3 rom_converter.py
```

## 📊 Supported Formats

### PlayStation Family
| System | Input Formats | Output Formats | Recommended |
|--------|---------------|----------------|------------|
| PS1 | ISO, CUE+BIN, CHD | CHD, ISO | CHD (universal) |
| PS2 | ISO, CUE+BIN, CHD | CHD, ISO, CSO, ZSO | CHD (PCSX2), ZSO (OPL) |
| PSP | ISO, CSO, ZSO | CSO, ZSO | CSO (best compatibility) |
| PS3 | ISO, JB ISOs | ISO | ISO (PS3 discs) |

### Nintendo Systems
| System | Input Formats | Output Formats |
|--------|---------------|----------------|
| NES/SNES | .nes, .smc, .sfc | No conversion needed |
| Nintendo 64 | .z64, .n64, .v64 | No conversion needed |
| GameCube | .gcm, .iso | .gcz (recommended) |
| Wii | .iso, .wbfs | No conversion needed |
| 3DS | .3ds, .cia | .cia (decrypted) |
| Switch | .xci, .nsp | No conversion needed |

### Other Systems
| System | Input Formats |
|--------|---------------|
| Sega Genesis | .md, .gen, .smd |
| Dreamcast | .gdi, .cdi |
| Atari Lynx | .lnx |
| Virtual Boy | .vb |
| WonderSwan | .ws, .wsc |

## ⚙️ Configuration

ROM Converter saves configuration in one of these locations (checked in order):
1. `.rom_converter_config.json` - In the app directory (portable)
2. `~/.rom_converter_config.json` - In your home directory

### Configuration Options

```json
{
  "_version": 2,
  "ps1_output_format": "CHD",
  "ps2_output_format": "CHD",
  "ps2_emulator": "PCSX2",
  "current_theme": "PS2",
  "delete_originals": false,
  "move_to_backup": true,
  "extract_compressed": true,
  "max_concurrent_conversions": null,
  "chdman_max_processors": null,
  "maxcso_threads": null
}
```

## 🔧 System Requirements

### Minimum
- **OS**: Windows 7+ / Linux (any modern distro)
- **Python**: 3.8+
- **RAM**: 2 GB
- **Disk Space**: 10 GB free (for conversions)

### Recommended
- **OS**: Windows 10+ / Ubuntu 20.04+ / Fedora 36+
- **Python**: 3.10+
- **RAM**: 8 GB+
- **CPU**: 4+ cores
- **Disk**: Fast SSD (NVMe recommended for large collections)

### External Tools (Auto-Downloaded)
- **chdman** (from MAME) - CHD conversion
- **maxcso** - CSO/ZSO compression
- **7-Zip** - Archive extraction
- **extract-xiso** - Xbox ISO extraction
- **NDecrypt** - 3DS decryption (for 3DS workflow)

## 💾 Building from Source

### Windows

```batch
cd WlfRyt-Rom-Manager
build.bat
```

This creates `dist\ROM_Converter.exe` (bundled with PyInstaller).

### Linux

```bash
cd WlfRyt-Rom-Manager
chmod +x build.sh
./build.sh
```

This creates `ROM_Converter-x86_64.AppImage`.

For Flatpak builds, see `flatpakbuild.sh`.

## 🎮 3DS Workflow

ROM Converter includes a complete 3DS ROM preparation pipeline:

1. **Extract**: Decompress 3DS ROMs from ZIP/7Z/RAR archives
2. **Decrypt**: Use NDecrypt to decrypt encrypted .3ds/.cia files
3. **Organize**: Auto-clean filenames and move to output folder

**Steps**:
1. Open ROM Converter → Select "3DS Workflow" tab
2. Choose source folder (archives with encrypted ROMs)
3. Choose destination folder for decrypted ROMs
4. Select which steps to run (extract, decrypt, move)
5. Click "Run Full Workflow"

> **Note**: Requires `3DS AES Keys.txt` in the app directory. The app auto-downloads this file.

## 🔑 Archive.org Integration

ROM Converter can auto-login to Archive.org using S3 credentials:

1. **Create** `Archive.org_keys.txt` or `keys.docx` in app directory with:
   ```
   access_key: YOUR_ACCESS_KEY_HERE
   secret_key: YOUR_SECRET_KEY_HERE
   ```

2. The app detects and auto-logs in on startup

3. Use for uploading converted ROMs to Archive.org directly

## 📈 Performance Tips

### For Large Collections (1000+ ROMs)

1. **Increase Worker Threads**:
   - Set `max_concurrent_conversions` in config
   - Recommended: `CPU_cores - 1`

2. **Use Fast Disk**:
   - SSD/NVMe much faster than HDD
   - Separate source/destination disks recommended

3. **Optimize Format**:
   - CHD: Slowest but best compression (PS2 emulators love it)
   - CSO: Faster, good compression (PS2 handhelds)
   - ZSO: Fastest, lower compression (OPL in PS2)

4. **Monitor Resources**:
   - Watch RAM usage (app auto-throttles at 85%)
   - CPU usage shown in log
   - Disk I/O throttled automatically

### Emulator-Specific Recommendations

| Emulator | System | Recommended Format | Rationale |
|----------|--------|-------------------|-----------|
| PCSX2 | PS2 | CHD | Best speed, native support |
| AetherSX2 | PS2 | CHD | Mobile-friendly, optimal |
| OPL (PS2) | PS2 | ZSO | Native support, efficient |
| Citra | 3DS | CIA | Decrypted, direct load |
| Yuzu | Switch | XCI/NSP | Native formats |

## 🐛 Troubleshooting

### "chdman not found"
- **Solution**: App will offer to download automatically
- **Alternative**: Manually place `chdman.exe` in app directory
- **Linux**: Install MAME via package manager: `sudo apt install mame`

### "Out of Memory" during conversion
- **Cause**: Too many parallel workers
- **Solution**: Reduce `max_concurrent_conversions` in config (default auto-detects)
- **Alternative**: Close other applications to free RAM

### "Archive extraction failed"
- **Cause**: Missing 7-Zip or corrupt archive
- **Solution**: App auto-downloads 7-Zip on first run
- **Alternative**: Pre-extract archives and run converter on extracted files

### "Conversion timeout"
- **Cause**: Very large ROM or slow disk
- **Solution**: Increase timeout in advanced settings (if added)
- **Alternative**: Convert one file at a time (set workers to 1)

### "3DS decryption failed"
- **Cause**: Missing or incorrect AES keys
- **Solution**: App auto-downloads `3DS AES Keys.txt`
- **Alternative**: Manually download from [key database](https://github.com/BernardoGiordano/KeySavRedirector/wiki/Extracting-Keys)

## 📚 Documentation

- **[ARCHITECTURE.md](ARCHITECTURE.md)** - System design and component overview
- **[CONTRIBUTING.md](CONTRIBUTING.md)** - How to contribute improvements and new features
- **[BUILD.md](docs/build.md)** - Detailed build instructions for all platforms
- **[API.md](docs/api.md)** - Developer API reference

## 🤝 Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for:
- Development setup
- How to add new converters
- Testing guidelines
- Pull request process

### Common Contributions

**Adding a new system converter**:
```python
# converters/new_system.py
from .base import BaseConverter

class MySystemConverter(BaseConverter):
    def convert(self, source_file, output_dir, **kwargs):
        # Your conversion logic here
        pass
```

**Adding a new extraction type**:
```python
# extractors/myformat.py
from .base import BaseExtractor

class MyFormatExtractor(BaseExtractor):
    def extract(self, archive_path, output_dir, **kwargs):
        # Your extraction logic here
        pass
```

## 📄 License

This project is licensed under the MIT License - see [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **MAME Team** - For chdman tool
- **maxcso** - CSO/ZSO compression tool
- **NDecrypt** - 3DS decryption tool
- **7-Zip** - Archive extraction
- All emulator developers who inspired this project

## 📞 Support

**Issues & Bug Reports**: [GitHub Issues](../../issues)

**Feature Requests**: [GitHub Discussions](../../discussions)

**Email**: (maintainer contact info here)

---

**Latest Version**: 1.0.0  
**Last Updated**: 2026-08-02  
**Status**: ✅ Actively Maintained
