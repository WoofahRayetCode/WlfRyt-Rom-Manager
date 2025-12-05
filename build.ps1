# Build script for ROM Converter using PyInstaller
# Usage: .\build.ps1

param(
    [switch]$Clean,
    [switch]$IncludeBinaries
)

# Colors for output
$Green = @{ ForegroundColor = 'Green' }
$Yellow = @{ ForegroundColor = 'Yellow' }
$Red = @{ ForegroundColor = 'Red' }

Write-Host "🔨 ROM Converter Build Script" @Yellow
Write-Host "======================================`n"

# Check if we're in the right directory
if (-not (Test-Path "rom_converter.py")) {
    Write-Host "❌ Error: rom_converter.py not found in root folder" @Red
    Write-Host "Please run this script from the project root directory"
    exit 1
}

# Clean previous builds if requested
if ($Clean) {
    Write-Host "🧹 Cleaning previous builds..."
    Remove-Item -Path "dist" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "build" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "rom_converter.spec" -Force -ErrorAction SilentlyContinue
    Write-Host "✓ Clean complete`n" @Green
}

# Check if PyInstaller is installed
Write-Host "📦 Checking dependencies..."
try {
    $PyInstaller = python -m pip list | Select-String -Pattern "pyinstaller"
    if (-not $PyInstaller) {
        Write-Host "⬇️  Installing PyInstaller..."
        python -m pip install pyinstaller -q
        if ($LASTEXITCODE -ne 0) {
            Write-Host "❌ Failed to install PyInstaller" @Red
            exit 1
        }
    }
    Write-Host "✓ PyInstaller available`n" @Green
} catch {
    Write-Host "❌ Error checking dependencies: $_" @Red
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

# Add binaries if requested
if ($IncludeBinaries) {
    Write-Host "📝 Including binaries in executable..."
    if (Test-Path "chdman.exe") {
        Write-Host "  ✓ Found chdman.exe"
        $BuildArgs += "--add-data=chdman.exe:."
    } else {
        Write-Host "  ⚠️  chdman.exe not found (optional)"
    }
    
    if (Test-Path "maxcso.exe") {
        Write-Host "  ✓ Found maxcso.exe"
        $BuildArgs += "--add-data=maxcso.exe:."
    } else {
        Write-Host "  ⚠️  maxcso.exe not found (optional)`n"
    }
}

# Build the executable
Write-Host "🏗️  Building executable...`n"
try {
    python -m PyInstaller $BuildArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Host "❌ Build failed" @Red
        exit 1
    }
} catch {
    Write-Host "❌ Build error: $_" @Red
    exit 1
}

# Verify build
if (Test-Path "dist\ROM_Converter.exe") {
    Write-Host "`n✅ Build successful!" @Green
    Write-Host "======================================`n"
    Write-Host "📦 Output: dist\ROM_Converter.exe"
    Write-Host "📋 Size: $(Get-Item dist\ROM_Converter.exe | ForEach-Object { "{0:N2} MB" -f ($_.Length / 1MB) })"
    Write-Host "`n📌 Next Steps:"
    Write-Host "  1. Copy dist\ROM_Converter.exe to your desired location"
    Write-Host "  2. Place chdman.exe and maxcso.exe in the same folder"
    Write-Host "  3. Run ROM_Converter.exe`n"
    Write-Host "💡 Tip: Create a folder like 'ROM_Converter_v1.0' and put the .exe + binaries there"
} else {
    Write-Host "❌ Build verification failed - ROM_Converter.exe not found" @Red
    exit 1
}

# Optional: Open build folder
Write-Host "🔗 Opening dist folder..."
Invoke-Item "dist"
