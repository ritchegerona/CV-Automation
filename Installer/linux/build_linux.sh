#!/bin/bash
# Shell script to build standalone Linux executable using PyInstaller
set -e

echo "=============================================="
echo "Building Linux Executable for CV Automation"
echo "=============================================="

# Resolve project root directory (two levels up from this script's directory)
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
PROJECT_ROOT="$( cd "$SCRIPT_DIR/../.." && pwd )"
cd "$PROJECT_ROOT"

# Check if .venv exists
if [ ! -d ".venv" ]; then
    echo "[ERROR] Python virtual environment (.venv) not found in project root."
    echo "Please create it first: python3 -m venv .venv"
    exit 1
fi

echo "Activating virtual environment..."
source .venv/bin/activate

echo "Checking if PyInstaller is installed..."
if ! pip show pyinstaller > /dev/null 2>&1; then
    echo "PyInstaller not found. Installing now..."
    pip install pyinstaller
fi

# Clean up previous builds
echo "Cleaning up previous builds..."
rm -rf Installer/linux/build Installer/linux/dist

# Run PyInstaller compile
echo "Running PyInstaller compilation..."
pyinstaller --workpath Installer/linux/build --distpath Installer/linux/dist --clean -y Installer/launcher.spec

echo "=============================================="
echo "Linux Executable Build Complete!"
echo "Standalone executable: Installer/linux/dist/CV_Automation"
echo "=============================================="
