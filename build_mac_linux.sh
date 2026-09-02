#!/usr/bin/env bash
set -e

echo "============================================"
echo "  Attendance Generator - Build Script (macOS/Linux)"
echo "============================================"
echo

if ! command -v python3 &> /dev/null; then
    echo "ERROR: python3 not found. Install it from https://www.python.org/downloads/"
    exit 1
fi

echo "Installing build tools (openpyxl, pyinstaller)..."
python3 -m pip install --upgrade pip 2>/dev/null || true
if ! python3 -m pip install openpyxl pyinstaller; then
    echo "Retrying with --break-system-packages (needed on some Homebrew/Linux Pythons)..."
    python3 -m pip install --break-system-packages openpyxl pyinstaller
fi

echo
echo "Building AttendanceGenerator..."
pyinstaller --onefile --windowed --name "AttendanceGenerator" attendance_gui.py

echo
echo "============================================"
echo "  Done! Your executable is at: dist/AttendanceGenerator"
echo
echo "  Give that single file to anyone else who"
echo "  needs it -- they do NOT need Python installed."
echo "  (macOS: they may need to right-click -> Open the"
echo "  first time, since it isn't code-signed.)"
echo "============================================"
