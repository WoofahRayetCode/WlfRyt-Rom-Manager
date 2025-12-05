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
    echo Output: dist\ROM_Converter.exe
    echo.
    echo Next Steps:
    echo   1. Copy dist\ROM_Converter.exe to your desired location
    echo   2. Place chdman.exe and maxcso.exe in the same folder
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
