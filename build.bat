@echo off
REM Build script for ROM Converter using PyInstaller
REM Run this to create ROM_Converter.exe

setlocal enabledelayedexpansion

echo.
echo ==================================================
echo  ROM Converter Build Script
echo ==================================================
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python not found. Please install Python first.
    pause
    exit /b 1
)

REM Check if rom_converter.py exists
if not exist "rom_converter.py" (
    echo Error: rom_converter.py not found in root folder
    pause
    exit /b 1
)

REM Capture build timestamp for embedding into the binary
for /f %%i in ('powershell -NoProfile -Command "Get-Date -UFormat %%s"') do set BUILD_TIMESTAMP=%%i
set RUNTIME_HOOK=build_timestamp_hook.py
echo import os>"%RUNTIME_HOOK%"
echo os.environ.setdefault("BUILD_TIMESTAMP","%BUILD_TIMESTAMP%")>>"%RUNTIME_HOOK%"

REM Check if tkinter is available
echo Checking for tkinter...
python -c "import tkinter" >nul 2>&1
if errorlevel 1 (
    echo tkinter not found - attempting to install...
    python -m pip install tk -q >nul 2>&1
    python -c "import tkinter" >nul 2>&1
    if errorlevel 1 (
        echo.
        echo Error: tkinter is not available.
        echo Please reinstall Python and ensure "tcl/tk and IDLE" is checked
        echo during installation, or repair your Python install via the installer.
        pause
        exit /b 1
    )
    echo tkinter installed successfully
) else (
    echo tkinter available
)

REM Install PyInstaller if not present
echo Checking for PyInstaller...
python -m pip list | findstr pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    python -m pip install pyinstaller -q
    if errorlevel 1 (
        echo Error: Failed to install PyInstaller
        pause
        exit /b 1
    )
)

REM Install psutil if not present
echo Checking for psutil...
python -m pip list | findstr psutil >nul 2>&1
if errorlevel 1 (
    echo Installing psutil...
    python -m pip install psutil -q
    if errorlevel 1 (
        echo Warning: Failed to install psutil - resource monitoring may be limited
    )
) else (
    echo psutil available
)

REM Install requests if not present
echo Checking for requests...
python -m pip list | findstr requests >nul 2>&1
if errorlevel 1 (
    echo Installing requests...
    python -m pip install requests -q
    if errorlevel 1 (
        echo Warning: Failed to install requests - Internet Archive login will be disabled
    )
) else (
    echo requests available
)

REM Clean previous builds
echo Cleaning previous builds...
if exist dist rmdir /s /q dist >nul 2>&1
if exist build rmdir /s /q build >nul 2>&1
if exist rom_converter.spec del rom_converter.spec >nul 2>&1

REM Build the executable
echo.
echo Building executable...
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name=ROM_Converter ^
    --distpath=dist ^
    --hidden-import=psutil ^
    --hidden-import=requests ^
    --runtime-hook=%RUNTIME_HOOK% ^
    rom_converter.py

if errorlevel 1 (
    echo Error: Build failed
    pause
    exit /b 1
)

REM Verify build
if exist dist\ROM_Converter.exe (
    echo.
    echo ==================================================
    echo  Build Successful!
    echo ==================================================
    echo.
    echo Copying build to OneDrive Desktop...
    set "DEST_DIR=%USERPROFILE%\OneDrive\Desktop\ROM Manager"
    if not exist "!DEST_DIR!" mkdir "!DEST_DIR!"
    copy /y "dist\ROM_Converter.exe" "!DEST_DIR!\" >nul
    
    echo Output: dist\ROM_Converter.exe
    echo Copied to: !DEST_DIR!
    echo.
    echo Next Steps:
    echo   1. The app is ready on your Desktop in the "ROM Manager" folder
    echo   2. Place chdman.exe and maxcso.exe in that same folder
    echo   3. Run ROM_Converter.exe
    echo.
    echo ==================================================
    echo.
    pause
) else (
    echo Error: Build verification failed
    pause
    exit /b 1
)

REM Clean up build-time hook
if exist "%RUNTIME_HOOK%" del /q "%RUNTIME_HOOK%" >nul 2>&1
