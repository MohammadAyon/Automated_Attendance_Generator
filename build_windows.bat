@echo off
setlocal

echo ============================================
echo   Attendance Generator - Build Script
echo ============================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python was not found on this computer.
    echo Install it from https://www.python.org/downloads/
    echo IMPORTANT: during setup, check "Add python.exe to PATH".
    pause
    exit /b 1
)

echo Installing build tools ^(openpyxl, pyinstaller^)...
python -m pip install --upgrade pip >nul
python -m pip install openpyxl pyinstaller
if errorlevel 1 (
    echo Retrying with --break-system-packages ^(needed on some Python installs^)...
    python -m pip install --break-system-packages openpyxl pyinstaller
    if errorlevel 1 (
        echo ERROR: pip install failed. Check your internet connection and try again.
        pause
        exit /b 1
    )
)

echo.
echo Building AttendanceGenerator.exe ...
pyinstaller --onefile --windowed --name "AttendanceGenerator" attendance_gui.py

if errorlevel 1 (
    echo ERROR: Build failed. Scroll up for details.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Done! Your executable is at:
echo   dist\AttendanceGenerator.exe
echo.
echo   Give that single file to anyone else who
echo   needs it -- they do NOT need Python installed.
echo ============================================
pause
