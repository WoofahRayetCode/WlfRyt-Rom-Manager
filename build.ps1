# Build script for ROM Converter using PyInstaller
# 
# Usage: .\build.ps1 [-Clean]
#
# Options:
#   -Clean   Remove previous build artifacts before building
#
# Notes:
#   - If chdman.exe and maxcso.exe are in the project directory, they will be bundled into the executable
#   - The resulting .exe will be fully self-contained and portable

param(
    [switch]$Clean
)

# Colors for output
$Green = @{ ForegroundColor = 'Green' }
$Yellow = @{ ForegroundColor = 'Yellow' }
$Red = @{ ForegroundColor = 'Red' }

Write-Host "ROM Converter Build Script" @Yellow
Write-Host "======================================`n"

# Check if we're in the right directory
if (-not (Test-Path "rom_converter.py")) {
    Write-Host "Error: rom_converter.py not found in root folder" @Red
    Write-Host "Please run this script from the project root directory"
    exit 1
}

# Clean previous builds if requested
if ($Clean) {
    Write-Host "Cleaning previous builds..."
    Remove-Item -Path "dist" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "build" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "rom_converter.spec" -Force -ErrorAction SilentlyContinue
    Write-Host "Clean complete" @Green
}

# Check if PyInstaller is installed
Write-Host "Checking dependencies..."
try {
    $PyInstaller = python -m pip list | Select-String -Pattern "pyinstaller"
    if (-not $PyInstaller) {
        Write-Host "Installing PyInstaller..."
        python -m pip install pyinstaller -q
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Failed to install PyInstaller" @Red
            exit 1
        }
    }
    Write-Host "PyInstaller available" @Green
} catch {
    Write-Host "Error checking dependencies: $_" @Red
    exit 1
}

# Build command
$BuildArgs = @(
    "--onefile",
    "--windowed",
    "--name=ROM_Converter",
    "--distpath=dist",
    "rom_converter.py"
)

# Always check for and include binaries
Write-Host "Checking for bundled binaries..."
$binariesFound = $false
if (Test-Path "chdman.exe") {
    Write-Host "  Found chdman.exe - bundling into executable" @Green
    $BuildArgs += "--add-data=chdman.exe;."
    $binariesFound = $true
} else {
    Write-Host "  chdman.exe not found (optional - can be downloaded at runtime)" @Yellow
}

if (Test-Path "maxcso.exe") {
    Write-Host "  Found maxcso.exe - bundling into executable" @Green
    $BuildArgs += "--add-data=maxcso.exe;."
    $binariesFound = $true
} else {
    Write-Host "  maxcso.exe not found (optional for CSO/ZSO support)" @Yellow
}

if (-not $binariesFound) {
    Write-Host "`nNote: No binaries found to bundle. The executable will work but will need to download chdman at first run." @Yellow
}

# Build the executable
Write-Host "Building executable..."
try {
    python -m PyInstaller $BuildArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Build failed" @Red
        exit 1
    }
} catch {
    Write-Host "Build error: $_" @Red
    exit 1
}

# Verify build
if (Test-Path "dist\ROM_Converter.exe") {
    Write-Host "`nBuild successful!" @Green
    Write-Host "======================================`n"
    Write-Host "Output: dist\ROM_Converter.exe"
    Write-Host "Size: $(Get-Item dist\ROM_Converter.exe | ForEach-Object { "{0:N2} MB" -f ($_.Length / 1MB) })"
        if ($binariesFound) {
            Write-Host "`nBundled binaries included in executable:" @Green
            if (Test-Path "chdman.exe") { Write-Host "  - chdman.exe" @Green }
            if (Test-Path "maxcso.exe") { Write-Host "  - maxcso.exe" @Green }
            Write-Host "`nThe executable is fully self-contained!" @Green
        } else {
            Write-Host "`nNo binaries bundled - chdman will be downloaded on first run if needed" @Yellow
        }
    Write-Host "`nNext Steps:"
        Write-Host "  1. Run dist\ROM_Converter.exe directly"
        if (-not $binariesFound) {
            Write-Host "  2. (Optional) Place chdman.exe and maxcso.exe next to ROM_Converter.exe to avoid downloading"
        }
        Write-Host "`nThe executable is portable and can be moved anywhere!"
} else {
    Write-Host "Build verification failed - ROM_Converter.exe not found" @Red
    exit 1
}

# Optional: Open build folder
Write-Host "Opening dist folder..."
Invoke-Item "dist"