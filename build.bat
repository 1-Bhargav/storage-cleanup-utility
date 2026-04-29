@echo off
REM ============================================================
REM  Storage Cleanup Utility - Local Build Script
REM ============================================================
REM  Run this on any 64-bit Windows 10/11 machine that has
REM  Python 3.11+ installed (https://www.python.org/downloads/).
REM
REM  Double-click this file, or run it from a Command Prompt.
REM  The finished .exe will appear in the "dist" folder.
REM ============================================================

echo.
echo Installing build dependencies...
python -m pip install --upgrade pip
python -m pip install pyinstaller==6.11.1 send2trash==1.8.3 pywin32==308
if errorlevel 1 (
    echo.
    echo ERROR: Could not install dependencies.
    echo Make sure Python 3.11+ is installed and on PATH.
    pause
    exit /b 1
)

echo.
echo Building Storage_Cleanup_Utility.exe ...
pyinstaller --onefile --windowed --noconfirm ^
    --name "Storage_Cleanup_Utility" ^
    --hidden-import "send2trash" ^
    --hidden-import "win32com" ^
    --hidden-import "win32com.client" ^
    --hidden-import "pywintypes" ^
    --collect-all "send2trash" ^
    storage_cleanup_utility.py

if errorlevel 1 (
    echo.
    echo ERROR: Build failed.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Build successful!
echo  Your executable is at:  dist\Storage_Cleanup_Utility.exe
echo ============================================================
echo.
pause
