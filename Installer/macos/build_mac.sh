#!/bin/bash
# Shell script to build standalone macOS executable and package into a .dmg installer
set -e

echo "=============================================="
echo "Building macOS DMG Installer for CV Automation"
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
rm -rf Installer/macos/build Installer/macos/dist Installer/macos/temp_dmg

# Run PyInstaller compile
echo "Running PyInstaller compilation..."
pyinstaller --workpath Installer/macos/build --distpath Installer/macos/dist --clean -y Installer/launcher.spec

# Package into DMG using manual fixed-size filesystem to avoid hdiutil space allocation bugs
echo "Preparing Disk Image (DMG)..."
DMG_PATH="Installer/macos/CV_Automation_Setup.dmg"
TEMP_DMG="Installer/macos/temp_uncompressed.dmg"
MNT_DIR="Installer/macos/mnt"

rm -f "$DMG_PATH" "$TEMP_DMG"
rm -rf "$MNT_DIR"
mkdir -p "$MNT_DIR"

# 1. Create a blank HFS+ formatted disk image of 200MB (large enough for binary + templates)
echo "Creating blank temporary disk image..."
hdiutil create -size 200m -fs HFS+ -volname "CV Automation Studio" -ov "$TEMP_DMG"

# 2. Attach the temporary disk image to a local mount point
echo "Mounting disk image..."
hdiutil attach "$TEMP_DMG" -mountpoint "$MNT_DIR"

# 3. Copy application binary and resources to the mounted drive
echo "Copying files to mounted disk image..."
cp "Installer/macos/dist/CV_Automation" "$MNT_DIR/"
cp "MSR_CV_Template.docx" "$MNT_DIR/"
cp "MSR_CV_Template.pdf" "$MNT_DIR/"
cp "Installer/macos/install_mac_deps.sh" "$MNT_DIR/"

# Ensure helper script is executable
chmod +x "$MNT_DIR/install_mac_deps.sh"

# Create symbolic link to /Applications for drag-and-drop
ln -s /Applications "$MNT_DIR/Applications"

# 4. Detach/unmount the disk image
echo "Unmounting disk image..."
hdiutil detach "$MNT_DIR"

# 5. Convert temporary DMG to compressed read-only DMG (UDZO format)
echo "Compressing and generating final DMG..."
hdiutil convert "$TEMP_DMG" -format UDZO -imagekey zlib-level=9 -o "$DMG_PATH"

# Clean up temporary artifacts
rm -f "$TEMP_DMG"
rm -rf "$MNT_DIR"

echo "=============================================="
echo "macOS Installer Build Complete!"
echo "File created: $DMG_PATH"
echo "=============================================="
