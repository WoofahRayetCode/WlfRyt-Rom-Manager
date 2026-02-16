#!/bin/bash
# Build script for ROM Converter using PyInstaller (Arch Linux)
#
# Usage: ./build.sh [--clean] [--download-tools] [--onedir] [--appimage] [--flatpak]
#
# Options:
#   --clean           Remove previous build artifacts before building
#   --download-tools  Download chdman (from MAME) and maxcso before building
#   --onedir          Build loose files in dist/ROM_Converter/ instead of a single binary
#   --appimage        Build an AppImage package
#   --flatpak         Build a Flatpak package
#
# Notes:
#   - If chdman and maxcso are in the project directory, they will be bundled into the executable
#   - The resulting binary will be fully self-contained and portable

set -e

# Colors for output
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

echo -e "${YELLOW}ROM Converter Build Script (Arch Linux)${NC}"
echo "======================================"
echo

# Parse arguments
CLEAN=false
DOWNLOAD_TOOLS=false
BUILD_APPIMAGE=false
BUILD_FLATPAK=false
BUILD_ONEDIR=false
for arg in "$@"; do
    case $arg in
        --clean)
            CLEAN=true
            shift
            ;;
        --download-tools)
            DOWNLOAD_TOOLS=true
            shift
            ;;
        --onedir)
            BUILD_ONEDIR=true
            shift
            ;;
        --appimage)
            BUILD_APPIMAGE=true
            shift
            ;;
        --flatpak)
            BUILD_FLATPAK=true
            shift
            ;;
    esac
done

# Check if we're in the right directory
if [[ ! -f "rom_converter.py" ]]; then
    echo -e "${RED}Error: rom_converter.py not found in root folder${NC}"
    echo "Please run this script from the project root directory"
    exit 1
fi

# Clean previous builds if requested
if [[ "$CLEAN" == true ]]; then
    echo "Cleaning previous builds..."
    rm -rf dist build rom_converter.spec 2>/dev/null || true
    echo -e "${GREEN}Clean complete${NC}"
fi

# Capture build timestamp for embedding into the binary
BUILD_TIMESTAMP=$(date +%s)
RUNTIME_HOOK=".build_timestamp_hook.py"
cat > "$RUNTIME_HOOK" << EOF
import os
os.environ.setdefault("BUILD_TIMESTAMP", "${BUILD_TIMESTAMP}")
EOF

# Function to download chdman from MAME
download_chdman() {
    echo -e "${CYAN}Downloading chdman from MAME...${NC}"
    
    # Get latest MAME version from GitHub API
    echo "  Fetching latest MAME version..."
    MAME_VERSION=$(curl -s "https://api.github.com/repos/mamedev/mame/releases/latest" | grep -oP '"tag_name": "mame\K[0-9]+' | head -1)
    
    if [[ -z "$MAME_VERSION" ]]; then
        echo -e "${RED}  Failed to get MAME version from GitHub API${NC}"
        return 1
    fi
    
    echo -e "  ${GREEN}Latest MAME version: $MAME_VERSION${NC}"
    
    # Download URL for Linux binary
    MAME_URL="https://github.com/mamedev/mame/releases/download/mame${MAME_VERSION}/mame${MAME_VERSION}b_64bit.exe"
    
    # Create temp directory
    TEMP_DIR=$(mktemp -d)
    trap "rm -rf $TEMP_DIR" EXIT
    
    echo "  Downloading MAME package (this may take a while, ~96MB)..."
    if ! curl -L -# -o "$TEMP_DIR/mame.exe" "$MAME_URL"; then
        echo -e "${RED}  Failed to download MAME${NC}"
        return 1
    fi
    
    # Extract chdman using 7z
    echo "  Extracting chdman..."
    if command -v 7z &> /dev/null; then
        if 7z e "$TEMP_DIR/mame.exe" -o"$TEMP_DIR" chdman.exe -y > /dev/null 2>&1; then
            if [[ -f "$TEMP_DIR/chdman.exe" ]]; then
                # On Linux, we need the Linux binary, not Windows .exe
                # MAME doesn't provide standalone Linux binaries, so we need to build or use system package
                echo -e "${YELLOW}  Note: MAME releases only contain Windows binaries${NC}"
                echo -e "${YELLOW}  Installing chdman from Arch repositories instead...${NC}"
                rm -rf "$TEMP_DIR"
                trap - EXIT
                
                if sudo pacman -S --noconfirm --needed mame-tools 2>/dev/null; then
                    # Create symlink to system chdman
                    CHDMAN_PATH=$(command -v chdman)
                    if [[ -n "$CHDMAN_PATH" ]]; then
                        cp "$CHDMAN_PATH" ./chdman
                        chmod +x ./chdman
                        echo -e "${GREEN}  chdman installed successfully!${NC}"
                        return 0
                    fi
                fi
                
                echo -e "${RED}  Failed to install mame-tools${NC}"
                return 1
            fi
        fi
    fi
    
    # Fallback: try to install from pacman
    echo -e "${YELLOW}  7z extraction failed, trying pacman...${NC}"
    if sudo pacman -S --noconfirm --needed mame-tools 2>/dev/null; then
        CHDMAN_PATH=$(command -v chdman)
        if [[ -n "$CHDMAN_PATH" ]]; then
            cp "$CHDMAN_PATH" ./chdman
            chmod +x ./chdman
            echo -e "${GREEN}  chdman installed successfully!${NC}"
            return 0
        fi
    fi
    
    echo -e "${RED}  Failed to get chdman${NC}"
    echo -e "${YELLOW}  Try: sudo pacman -S mame-tools${NC}"
    return 1
}

# Function to download maxcso
download_maxcso() {
    echo -e "${CYAN}Downloading maxcso...${NC}"
    
    # maxcso doesn't have prebuilt Linux binaries, need to build from source
    echo "  maxcso needs to be built from source on Linux..."
    
    # Check if make and g++ are available
    if ! command -v make &> /dev/null || ! command -v g++ &> /dev/null; then
        echo -e "${YELLOW}  Installing build dependencies...${NC}"
        sudo pacman -S --noconfirm --needed base-devel libuv lz4 zlib
    fi
    
    # Save current directory
    ORIG_DIR=$(pwd)
    
    # Create temp directory
    TEMP_DIR=$(mktemp -d)
    cd "$TEMP_DIR"
    
    echo "  Cloning maxcso repository..."
    if ! git clone --depth 1 https://github.com/unknownbrackets/maxcso.git maxcso_src > /dev/null 2>&1; then
        echo -e "${RED}  Failed to clone maxcso repository${NC}"
        cd "$ORIG_DIR"
        rm -rf "$TEMP_DIR"
        return 1
    fi
    
    cd maxcso_src
    
    echo "  Building maxcso..."
    if make -j$(nproc) > /dev/null 2>&1; then
        if [[ -f "maxcso" ]]; then
            cp maxcso "$ORIG_DIR/"
            cd "$ORIG_DIR"
            chmod +x ./maxcso
            rm -rf "$TEMP_DIR"
            echo -e "${GREEN}  maxcso built successfully!${NC}"
            return 0
        fi
    fi
    
    cd "$ORIG_DIR"
    rm -rf "$TEMP_DIR"
    
    echo -e "${RED}  Failed to build maxcso${NC}"
    echo -e "${YELLOW}  You can try building manually:${NC}"
    echo -e "${YELLOW}    git clone https://github.com/unknownbrackets/maxcso.git${NC}"
    echo -e "${YELLOW}    cd maxcso && make${NC}"
    return 1
}

# Download tools if requested
if [[ "$DOWNLOAD_TOOLS" == true ]]; then
    echo "Downloading required tools..."
    echo
    
    if [[ ! -f "chdman" ]]; then
        download_chdman || echo -e "${YELLOW}Continuing without chdman...${NC}"
        echo
    else
        echo -e "${GREEN}chdman already exists, skipping download${NC}"
    fi
    
    if [[ ! -f "maxcso" ]]; then
        download_maxcso || echo -e "${YELLOW}Continuing without maxcso...${NC}"
        echo
    else
        echo -e "${GREEN}maxcso already exists, skipping download${NC}"
    fi
fi

# Check if Python is available
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}Error: Python3 not found. Please install Python first.${NC}"
    echo "On Arch Linux: sudo pacman -S python"
    exit 1
fi

# Check if PyInstaller is installed
echo "Checking dependencies..."
if ! python3 -m pip list 2>/dev/null | grep -qi pyinstaller; then
    echo "Installing PyInstaller..."
    python3 -m pip install pyinstaller --quiet --break-system-packages
    if [[ $? -ne 0 ]]; then
        echo -e "${RED}Failed to install PyInstaller${NC}"
        exit 1
    fi
fi
echo -e "${GREEN}PyInstaller available${NC}"

# Check if psutil is installed (for resource monitoring)
if ! python3 -m pip list 2>/dev/null | grep -qi psutil; then
    echo "Installing psutil..."
    python3 -m pip install psutil --quiet --break-system-packages
    if [[ $? -ne 0 ]]; then
        echo -e "${YELLOW}Warning: Failed to install psutil - resource monitoring may be limited${NC}"
    fi
fi
echo -e "${GREEN}psutil available${NC}"

# Check if tkinter is installed (for GUI)
# Detect the Python version for package naming (used by install steps and error messages)
PY_VERSION=$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
PY_MAJOR=$(python3 -c 'import sys; print(sys.version_info.major)')

# Detect if running on an immutable/atomic OS (Bazzite, Bluefin, Fedora Atomic, etc.)
IS_IMMUTABLE=false
if [[ -f /run/ostree-booted ]] || command -v rpm-ostree &> /dev/null; then
    IS_IMMUTABLE=true
fi

# Helper: check if system Python (outside any venv) has tkinter
system_has_tkinter() {
    # Try common system Python paths directly
    for sys_python in /usr/bin/python${PY_VERSION} /usr/bin/python${PY_MAJOR} /usr/bin/python3; do
        if [[ -x "$sys_python" ]]; then
            if "$sys_python" -c "import tkinter" &> /dev/null; then
                return 0
            fi
            break
        fi
    done
    return 1
}

# Helper: find system tkinter directory and _tkinter shared object
find_system_tkinter() {
    # Search standard system library paths for the tkinter package directory
    for base in /usr/lib64/python${PY_VERSION} /usr/lib/python${PY_VERSION} \
                /usr/lib64/python${PY_MAJOR} /usr/lib/python${PY_MAJOR}; do
        if [[ -d "${base}/tkinter" ]]; then
            echo "${base}/tkinter"
            return 0
        fi
    done
    # Fallback: ask system Python directly
    for sys_python in /usr/bin/python${PY_VERSION} /usr/bin/python${PY_MAJOR} /usr/bin/python3; do
        if [[ -x "$sys_python" ]]; then
            local tk_dir
            tk_dir=$("$sys_python" -c "import tkinter, os; print(os.path.dirname(tkinter.__file__))" 2>/dev/null)
            if [[ -n "$tk_dir" && -d "$tk_dir" ]]; then
                echo "$tk_dir"
                return 0
            fi
            break
        fi
    done
    return 1
}

# Helper: link system tkinter into the active virtual environment
link_tkinter_to_venv() {
    local sys_tk_dir="$1"
    local venv_site
    venv_site=$(python3 -c 'import site; print(site.getsitepackages()[0])' 2>/dev/null)
    if [[ -z "$venv_site" ]]; then
        echo -e "${RED}  Could not determine venv site-packages path${NC}"
        return 1
    fi

    echo -e "${CYAN}  Linking system tkinter into virtual environment...${NC}"

    # Link the tkinter package directory
    ln -sf "$sys_tk_dir" "$venv_site/tkinter"
    echo "    -> $venv_site/tkinter"

    # Find and link the _tkinter C extension (.so)
    local parent
    parent=$(dirname "$sys_tk_dir")
    local tkinter_so
    tkinter_so=$(find "$parent" /usr/lib64/python${PY_VERSION} /usr/lib/python${PY_VERSION} \
                      -maxdepth 2 -name '_tkinter*.so' -print -quit 2>/dev/null)
    if [[ -n "$tkinter_so" ]]; then
        # Determine correct lib-dynload directory inside the venv
        local venv_dynload
        venv_dynload=$(python3 -c '
import importlib.machinery, os, sysconfig
dynload = sysconfig.get_path("platlib")
# lib-dynload is typically a sibling of the site-packages dir
candidate = os.path.join(os.path.dirname(dynload), "lib-dynload")
if not os.path.isdir(candidate):
    candidate = os.path.join(os.path.dirname(os.path.dirname(dynload)), "lib-dynload")
print(candidate)' 2>/dev/null)
        if [[ -z "$venv_dynload" ]]; then
            venv_dynload="$(dirname "$venv_site")/lib-dynload"
        fi
        mkdir -p "$venv_dynload" 2>/dev/null
        ln -sf "$tkinter_so" "$venv_dynload/"
        echo "    -> $venv_dynload/$(basename "$tkinter_so")"
    else
        echo -e "${YELLOW}  Warning: _tkinter.so not found - tkinter may not work${NC}"
        return 1
    fi

    echo -e "${GREEN}  System tkinter linked into venv successfully${NC}"
    return 0
}

if ! python3 -c "import tkinter" &> /dev/null; then
    echo -e "${YELLOW}tkinter not found - attempting to set up...${NC}"

    # --- Step 1: If in a venv and system already has tkinter, just link it ---
    if [[ -n "$VIRTUAL_ENV" ]]; then
        echo "  Virtual environment detected: $VIRTUAL_ENV"
        if system_has_tkinter; then
            echo -e "${GREEN}  System Python already has tkinter${NC}"
            SYS_TK_DIR=$(find_system_tkinter)
            if [[ -n "$SYS_TK_DIR" ]]; then
                link_tkinter_to_venv "$SYS_TK_DIR"
            fi
        fi
    fi

    # --- Step 2: If still not available, install the system package ---
    if ! python3 -c "import tkinter" &> /dev/null; then
        echo "  Installing tkinter system package..."

        TKINTER_INSTALLED=false

        if [[ "$IS_IMMUTABLE" == true ]]; then
            # Immutable Fedora-based OS (Bazzite, Bluefin, Fedora Atomic/Silverblue)
            echo -e "  ${CYAN}Detected immutable OS (ostree/rpm-ostree)${NC}"
            # Check if python3-tkinter is already layered or in the base image
            if rpm -q python${PY_MAJOR}-tkinter &> /dev/null; then
                echo -e "  ${GREEN}python${PY_MAJOR}-tkinter is already installed in the image${NC}"
                TKINTER_INSTALLED=true
            else
                echo "  Layering python${PY_MAJOR}-tkinter via rpm-ostree..."
                echo -e "  ${YELLOW}Note: This requires a reboot to take effect.${NC}"
                if sudo rpm-ostree install -y --idempotent python${PY_MAJOR}-tkinter; then
                    echo -e "  ${YELLOW}Package layered. A reboot is required for it to become available.${NC}"
                    echo -e "  ${YELLOW}After rebooting, re-run this build script.${NC}"
                    # Check if it's usable without reboot (rpm-ostree live apply)
                    if sudo rpm-ostree ex apply-live --allow-replacement 2>/dev/null; then
                        echo -e "  ${GREEN}Live update applied - no reboot needed!${NC}"
                        TKINTER_INSTALLED=true
                    else
                        echo -e "  ${YELLOW}Could not live-apply. Please reboot and re-run this script.${NC}"
                        exit 1
                    fi
                else
                    echo -e "${RED}  Failed to layer python${PY_MAJOR}-tkinter${NC}"
                fi
            fi
        elif command -v pacman &> /dev/null; then
            echo "  Detected pacman (Arch/Manjaro)..."
            sudo pacman -S --noconfirm --needed tk && TKINTER_INSTALLED=true
        elif command -v apt-get &> /dev/null; then
            echo "  Detected apt (Debian/Ubuntu)..."
            sudo apt-get update -qq && sudo apt-get install -y python${PY_MAJOR}-tk && TKINTER_INSTALLED=true
        elif command -v dnf &> /dev/null; then
            echo "  Detected dnf (Fedora/RHEL)..."
            sudo dnf install -y python${PY_MAJOR}-tkinter && TKINTER_INSTALLED=true
        elif command -v zypper &> /dev/null; then
            echo "  Detected zypper (openSUSE)..."
            sudo zypper install -y python${PY_MAJOR}-tk && TKINTER_INSTALLED=true
        elif command -v brew &> /dev/null; then
            echo "  Detected Homebrew (macOS)..."
            brew install python-tk@${PY_VERSION} && TKINTER_INSTALLED=true
        fi

        # --- Step 3: After installing system package, link into venv if needed ---
        if [[ "$TKINTER_INSTALLED" == true && -n "$VIRTUAL_ENV" ]]; then
            if ! python3 -c "import tkinter" &> /dev/null; then
                SYS_TK_DIR=$(find_system_tkinter)
                if [[ -n "$SYS_TK_DIR" ]]; then
                    link_tkinter_to_venv "$SYS_TK_DIR"
                fi
            fi
        fi
    fi

    # Final verification
    if ! python3 -c "import tkinter" &> /dev/null; then
        echo -e "${RED}Error: Failed to set up tkinter${NC}"
        echo "Please install tkinter manually for your distribution:"
        if [[ "$IS_IMMUTABLE" == true ]]; then
            echo "  Bazzite/Bluefin/Fedora Atomic:"
            echo "    sudo rpm-ostree install python${PY_MAJOR}-tkinter"
            echo "    systemctl reboot"
        fi
        echo "  Arch/Manjaro:   sudo pacman -S tk"
        echo "  Debian/Ubuntu:  sudo apt-get install python3-tk"
        echo "  Fedora/RHEL:    sudo dnf install python3-tkinter"
        echo "  openSUSE:       sudo zypper install python3-tk"
        echo "  macOS (Brew):   brew install python-tk@${PY_VERSION}"
        if [[ -n "$VIRTUAL_ENV" ]]; then
            echo
            echo "If your system Python already has tkinter, try recreating the venv with:"
            echo "  python3 -m venv --system-site-packages $VIRTUAL_ENV"
        fi
        exit 1
    fi
    echo -e "${GREEN}tkinter set up successfully${NC}"
else
    echo -e "${GREEN}tkinter available${NC}"
fi

# Add ~/.local/bin to PATH if not already there (for user-installed pip packages)
export PATH="$HOME/.local/bin:$PATH"

# Build command arguments
BUILD_ARGS=(
    "--name=ROM_Converter"
    "--distpath=dist"
    "--hidden-import=psutil"
    "--runtime-hook=${RUNTIME_HOOK}"
    "rom_converter.py"
)

# Default to onefile unless onedir explicitly requested
if [[ "$BUILD_ONEDIR" == false ]]; then
    BUILD_ARGS=("--onefile" "${BUILD_ARGS[@]}")
fi

# Check for and include binaries
echo "Checking for bundled binaries..."
BINARIES_FOUND=false

if [[ -f "chdman" ]]; then
    echo -e "  ${GREEN}Found chdman - bundling into executable${NC}"
    BUILD_ARGS+=("--add-data=chdman:.")
    BINARIES_FOUND=true
else
    echo -e "  ${YELLOW}chdman not found (optional - can be downloaded at runtime)${NC}"
fi

if [[ -f "maxcso" ]]; then
    echo -e "  ${GREEN}Found maxcso - bundling into executable${NC}"
    BUILD_ARGS+=("--add-data=maxcso:.")
    BINARIES_FOUND=true
else
    echo -e "  ${YELLOW}maxcso not found (optional for CSO/ZSO support)${NC}"
fi

if [[ -f "3DS AES Keys.txt" ]]; then
    echo -e "  ${GREEN}Found 3DS AES Keys - bundling into executable${NC}"
    BUILD_ARGS+=("--add-data=3DS AES Keys.txt:.")
else
    echo -e "  ${YELLOW}3DS AES Keys.txt not found (optional for 3DS decryption)${NC}"
fi

if [[ "$BINARIES_FOUND" == false ]]; then
    echo
    echo -e "${YELLOW}Note: No binaries found to bundle. The executable will work but will need to download chdman at first run.${NC}"
fi

# Build the executable
echo
if [[ "$BUILD_ONEDIR" == true ]]; then
    echo "Building loose onedir bundle (dist/ROM_Converter/)..."
else
    echo "Building single-file executable (dist/ROM_Converter)..."
fi
python3 -m PyInstaller "${BUILD_ARGS[@]}"

if [[ $? -ne 0 ]]; then
    echo -e "${RED}Build failed${NC}"
    exit 1
fi

APPIMAGE_SOURCE=""

# Verify build
if [[ "$BUILD_ONEDIR" == true ]]; then
    if [[ -d "dist/ROM_Converter" ]]; then
        BIN_PATH="dist/ROM_Converter/ROM_Converter"
        [[ -f "$BIN_PATH" ]] && chmod +x "$BIN_PATH"
        APPIMAGE_SOURCE="$BIN_PATH"
        # Write a simple launcher script for onedir builds
        LAUNCHER="dist/ROM_Converter/run_rom_converter.sh"
        printf '#!/bin/bash\nDIR="$(cd "$(dirname "$0")" && pwd)"\nexec "$DIR/ROM_Converter" "$@"\n' > "$LAUNCHER"
        chmod +x "$LAUNCHER"
        SIZE=$(du -sh "dist/ROM_Converter" | cut -f1)
        echo
        echo -e "${GREEN}Build successful!${NC}"
        echo "======================================"
        echo
        echo "Output folder: dist/ROM_Converter/"
        echo "Size: $SIZE"
        echo
        echo "Next Steps:"
        echo "  1. Run dist/ROM_Converter/ROM_Converter"
        echo "     or dist/ROM_Converter/run_rom_converter.sh"
        echo "  2. Distribute the entire dist/ROM_Converter/ folder"
    else
        echo -e "${RED}Build verification failed - dist/ROM_Converter folder not found${NC}"
        exit 1
    fi
else
    if [[ -f "dist/ROM_Converter" ]]; then
        chmod +x "dist/ROM_Converter"
        APPIMAGE_SOURCE="dist/ROM_Converter"
        SIZE=$(du -h "dist/ROM_Converter" | cut -f1)
        echo
        echo -e "${GREEN}Build successful!${NC}"
        echo "======================================"
        echo
        echo "Output: dist/ROM_Converter"
        echo "Size: $SIZE"
        
        if [[ "$BINARIES_FOUND" == true ]]; then
            echo
            echo -e "${GREEN}Bundled binaries included in executable:${NC}"
            [[ -f "chdman" ]] && echo -e "  ${GREEN}- chdman${NC}"
            [[ -f "maxcso" ]] && echo -e "  ${GREEN}- maxcso${NC}"
            echo
            echo -e "${GREEN}The executable is fully self-contained!${NC}"
        else
            echo
            echo -e "${YELLOW}No binaries bundled - chdman will be downloaded on first run if needed${NC}"
        fi
        
        echo
        echo "Next Steps:"
        echo "  1. Run dist/ROM_Converter directly"
        if [[ "$BINARIES_FOUND" == false ]]; then
            echo "  2. (Optional) Place chdman and maxcso next to ROM_Converter to avoid downloading"
        fi
        echo
        echo "The executable is portable and can be moved anywhere!"
    else
        echo -e "${RED}Build verification failed - ROM_Converter not found${NC}"
        exit 1
    fi
fi

# Clean up build-time hook
rm -f "$RUNTIME_HOOK"

# Build AppImage if requested
if [[ "$BUILD_APPIMAGE" == true ]]; then
    echo
    echo -e "${YELLOW}Building AppImage...${NC}"
    echo "======================================"
    
    # Check for FUSE - can be on host, in DistroBox, or we use extract-and-run fallback
    FUSE_AVAILABLE=false
    DISTROBOX_CMD=""
    
    if command -v fusermount &> /dev/null || command -v fusermount3 &> /dev/null; then
        FUSE_AVAILABLE=true
        echo -e "${GREEN}FUSE available on host${NC}"
    elif command -v distrobox &> /dev/null; then
        # Check if there's a DistroBox with FUSE available
        echo "Checking DistroBox for FUSE support..."
        DISTROBOX_CONTAINERS=$(distrobox list --no-color 2>/dev/null | tail -n +2 | awk '{print $3}' | head -5)
        for container in $DISTROBOX_CONTAINERS; do
            if [[ -n "$container" ]]; then
                if distrobox enter "$container" -- command -v fusermount &> /dev/null 2>&1 || \
                   distrobox enter "$container" -- command -v fusermount3 &> /dev/null 2>&1; then
                    FUSE_AVAILABLE=true
                    DISTROBOX_CMD="distrobox enter $container --"
                    echo -e "${GREEN}FUSE available in DistroBox: $container${NC}"
                    break
                fi
            fi
        done
    fi
    
    if [[ "$FUSE_AVAILABLE" == false ]]; then
        echo -e "${YELLOW}FUSE not found on host or in DistroBox - using extract-and-run fallback${NC}"
    fi
    
    # Check for appimagetool
    APPIMAGETOOL=""
    APPIMAGETOOL_ARGS=""
    if command -v appimagetool &> /dev/null; then
        APPIMAGETOOL="appimagetool"
    elif [[ -f "./appimagetool-x86_64.AppImage" ]]; then
        APPIMAGETOOL="./appimagetool-x86_64.AppImage"
        # Use extract-and-run if no FUSE available
        if [[ "$FUSE_AVAILABLE" == false ]] || [[ -z "$DISTROBOX_CMD" ]]; then
            APPIMAGETOOL_ARGS="--appimage-extract-and-run"
        fi
    else
        echo "Downloading appimagetool..."
        curl -L -o appimagetool-x86_64.AppImage "https://github.com/AppImage/AppImageKit/releases/download/continuous/appimagetool-x86_64.AppImage"
        chmod +x appimagetool-x86_64.AppImage
        APPIMAGETOOL="./appimagetool-x86_64.AppImage"
        # Use extract-and-run if no FUSE available
        if [[ "$FUSE_AVAILABLE" == false ]] || [[ -z "$DISTROBOX_CMD" ]]; then
            APPIMAGETOOL_ARGS="--appimage-extract-and-run"
        fi
    fi
    
    # Require a built binary path
    if [[ -z "$APPIMAGE_SOURCE" || ! -f "$APPIMAGE_SOURCE" ]]; then
        echo -e "${RED}Cannot build AppImage: built binary not found${NC}"
        echo "Make sure the build step succeeded before requesting --appimage"
        exit 1
    fi

    # Create AppDir structure
    APPDIR="dist/ROM_Converter.AppDir"
    rm -rf "$APPDIR"
    mkdir -p "$APPDIR/usr/bin"
    mkdir -p "$APPDIR/usr/share/applications"
    mkdir -p "$APPDIR/usr/share/icons/hicolor/256x256/apps"
    
    # Copy executable
    cp "$APPIMAGE_SOURCE" "$APPDIR/usr/bin/"
    
    # Create desktop file
    cat > "$APPDIR/usr/share/applications/rom-converter.desktop" << EOF
[Desktop Entry]
Type=Application
Name=ROM Converter
Comment=Convert disc images to CHD format
Exec=ROM_Converter
Icon=rom-converter
Categories=Utility;Game;
Terminal=false
EOF
    
    # Create symlinks required by AppImage
    ln -sf usr/share/applications/rom-converter.desktop "$APPDIR/rom-converter.desktop"
    
    # Create a simple icon (or use existing if available)
    if [[ -f "icon.png" ]]; then
        cp icon.png "$APPDIR/usr/share/icons/hicolor/256x256/apps/rom-converter.png"
    else
        # Create a placeholder icon using ImageMagick if available
        if command -v convert &> /dev/null; then
            convert -size 256x256 xc:#1a1a2e -fill '#00ff88' -gravity center -pointsize 48 -annotate 0 'ROM\nConverter' "$APPDIR/usr/share/icons/hicolor/256x256/apps/rom-converter.png"
        else
            echo -e "${YELLOW}Warning: No icon found and ImageMagick not available${NC}"
            # Create minimal 1x1 PNG as fallback
            printf '\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc\xf8\x0f\x00\x00\x01\x01\x00\x05\x18\xd8N\x00\x00\x00\x00IEND\xaeB`\x82' > "$APPDIR/usr/share/icons/hicolor/256x256/apps/rom-converter.png"
        fi
    fi
    ln -sf usr/share/icons/hicolor/256x256/apps/rom-converter.png "$APPDIR/rom-converter.png"
    
    # Create AppRun script
    cat > "$APPDIR/AppRun" << 'EOF'
#!/bin/bash
SELF=$(readlink -f "$0")
HERE=${SELF%/*}
export PATH="${HERE}/usr/bin:${PATH}"
exec "${HERE}/usr/bin/ROM_Converter" "$@"
EOF
    chmod +x "$APPDIR/AppRun"
    
    # Build AppImage
    # Use DistroBox if available and has FUSE, otherwise use extract-and-run fallback
    APPDIR_ABS=$(realpath "$APPDIR")
    DIST_ABS=$(realpath "dist")
    
    if [[ -n "$DISTROBOX_CMD" ]]; then
        echo "Building AppImage via DistroBox..."
        $DISTROBOX_CMD bash -c "cd '$(pwd)' && ARCH=x86_64 '$APPIMAGETOOL' '$APPDIR_ABS' '$DIST_ABS/ROM_Converter-x86_64.AppImage'"
    elif [[ -n "$APPIMAGETOOL_ARGS" ]]; then
        ARCH=x86_64 $APPIMAGETOOL $APPIMAGETOOL_ARGS "$APPDIR" "dist/ROM_Converter-x86_64.AppImage"
    else
        ARCH=x86_64 $APPIMAGETOOL "$APPDIR" "dist/ROM_Converter-x86_64.AppImage"
    fi
    
    if [[ -f "dist/ROM_Converter-x86_64.AppImage" ]]; then
        chmod +x "dist/ROM_Converter-x86_64.AppImage"
        APPIMAGE_SIZE=$(du -h "dist/ROM_Converter-x86_64.AppImage" | cut -f1)
        echo
        echo -e "${GREEN}AppImage built successfully!${NC}"
        echo "Output: dist/ROM_Converter-x86_64.AppImage"
        echo "Size: $APPIMAGE_SIZE"
    else
        echo -e "${RED}AppImage build failed${NC}"
    fi
fi

# Build Flatpak if requested
if [[ "$BUILD_FLATPAK" == true ]]; then
    echo
    echo -e "${YELLOW}Building Flatpak...${NC}"
    echo "======================================"
    
    # Check for flatpak-builder
    if ! command -v flatpak-builder &> /dev/null; then
        echo -e "${YELLOW}Installing flatpak-builder...${NC}"
        sudo pacman -S --noconfirm --needed flatpak-builder || {
            echo -e "${RED}Failed to install flatpak-builder${NC}"
            echo "Install manually: sudo pacman -S flatpak-builder"
            exit 1
        }
    fi
    
    # Ensure Flathub repo is added
    flatpak remote-add --if-not-exists --user flathub https://flathub.org/repo/flathub.flatpakrepo 2>/dev/null || true
    
    # Install required runtime and SDK if not present
    echo "Checking Flatpak runtime..."
    flatpak install --user -y flathub org.freedesktop.Platform//23.08 org.freedesktop.Sdk//23.08 2>/dev/null || true
    
    # Create flatpak directory
    FLATPAK_DIR="flatpak-build"
    mkdir -p "$FLATPAK_DIR"
    
    # Create Flatpak manifest
    cat > "$FLATPAK_DIR/io.github.wlfryt.RomConverter.yml" << EOF
app-id: io.github.wlfryt.RomConverter
runtime: org.freedesktop.Platform
runtime-version: '23.08'
sdk: org.freedesktop.Sdk
command: ROM_Converter
finish-args:
  - --share=ipc
  - --socket=x11
  - --socket=wayland
  - --filesystem=home
  - --filesystem=/run/media
  - --filesystem=/media
  - --device=dri

modules:
  - name: rom-converter
    buildsystem: simple
    build-commands:
      - install -Dm755 ROM_Converter /app/bin/ROM_Converter
      - install -Dm644 rom-converter.desktop /app/share/applications/io.github.wlfryt.RomConverter.desktop
      - install -Dm644 rom-converter.png /app/share/icons/hicolor/256x256/apps/io.github.wlfryt.RomConverter.png
    sources:
      - type: file
        path: ../dist/ROM_Converter
      - type: file
        path: rom-converter.desktop
      - type: file
        path: rom-converter.png
EOF
    
    # Create desktop file for Flatpak
    cat > "$FLATPAK_DIR/rom-converter.desktop" << EOF
[Desktop Entry]
Type=Application
Name=ROM Converter
Comment=Convert disc images to CHD format
Exec=ROM_Converter
Icon=io.github.wlfryt.RomConverter
Categories=Utility;Game;
Terminal=false
EOF
    
    # Create or copy icon
    if [[ -f "icon.png" ]]; then
        cp icon.png "$FLATPAK_DIR/rom-converter.png"
    else
        if command -v convert &> /dev/null; then
            convert -size 256x256 xc:#1a1a2e -fill '#00ff88' -gravity center -pointsize 48 -annotate 0 'ROM\nConverter' "$FLATPAK_DIR/rom-converter.png"
        else
            # Create minimal placeholder
            printf '\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc\xf8\x0f\x00\x00\x01\x01\x00\x05\x18\xd8N\x00\x00\x00\x00IEND\xaeB`\x82' > "$FLATPAK_DIR/rom-converter.png"
        fi
    fi
    
    # Build Flatpak
    cd "$FLATPAK_DIR"
    flatpak-builder --force-clean --user --install-deps-from=flathub build-dir io.github.wlfryt.RomConverter.yml
    
    if [[ $? -eq 0 ]]; then
        # Export to repo and create bundle
        flatpak-builder --repo=repo --force-clean build-dir io.github.wlfryt.RomConverter.yml
        flatpak build-bundle repo ../dist/ROM_Converter.flatpak io.github.wlfryt.RomConverter
        cd ..
        
        if [[ -f "dist/ROM_Converter.flatpak" ]]; then
            FLATPAK_SIZE=$(du -h "dist/ROM_Converter.flatpak" | cut -f1)
            echo
            echo -e "${GREEN}Flatpak built successfully!${NC}"
            echo "Output: dist/ROM_Converter.flatpak"
            echo "Size: $FLATPAK_SIZE"
            echo
            echo "To install: flatpak install --user dist/ROM_Converter.flatpak"
            echo "To run: flatpak run io.github.wlfryt.RomConverter"
        else
            echo -e "${RED}Flatpak bundle creation failed${NC}"
        fi
    else
        cd ..
        echo -e "${RED}Flatpak build failed${NC}"
    fi
fi
