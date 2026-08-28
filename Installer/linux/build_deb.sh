#!/bin/bash
# Shell script to build a Debian package (.deb) for Debian/Ubuntu distributions
set -e

echo "=============================================="
echo "Building Debian (.deb) Package for Linux"
echo "=============================================="

# Resolve project root directory (two levels up from this script's directory)
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
PROJECT_ROOT="$( cd "$SCRIPT_DIR/../.." && pwd )"
cd "$PROJECT_ROOT"

# 1. Ensure Linux executable is built
if [ ! -f "Installer/linux/dist/CV_Automation" ]; then
    echo "Linux executable not found. Running build_linux.sh first..."
    bash Installer/linux/build_linux.sh
fi

# 2. Set up Debian packaging directory structure
PKG_DIR="Installer/linux/cv-automation_1.0.0_amd64"
rm -rf "$PKG_DIR"
mkdir -p "$PKG_DIR/DEBIAN"
mkdir -p "$PKG_DIR/usr/bin"
mkdir -p "$PKG_DIR/usr/share/applications"

# 3. Copy files to package paths
cp "Installer/linux/dist/CV_Automation" "$PKG_DIR/usr/bin/cv-automation"
chmod 755 "$PKG_DIR/usr/bin/cv-automation"

# 4. Create Desktop Entry file (Application Menu Shortcut)
cat << 'EOF' > "$PKG_DIR/usr/share/applications/cv-automation.desktop"
[Desktop Entry]
Version=1.0
Type=Application
Name=MSR CV Processing Studio
Comment=Process and summarize candidate CVs using Gemini API
Exec=cv-automation
Icon=document-send
Terminal=true
Categories=Office;Utility;
EOF
chmod 644 "$PKG_DIR/usr/share/applications/cv-automation.desktop"

# 5. Create Debian control file with declared APT dependencies
cat << 'EOF' > "$PKG_DIR/DEBIAN/control"
Package: cv-automation
Version: 1.0.0
Section: utils
Priority: optional
Architecture: amd64
Depends: libreoffice-core, libreoffice-writer, python3
Maintainer: Ritche Gerona <chie.msr@example.com>
Description: MSR CV Processing Studio
 A Streamlit-based web application to automate parsing, summarization, and
 report generation of CVs using the Gemini API. Automatically converts doc
 and pdf formats.
EOF

# 6. Build the debian package
echo "Running dpkg-deb to compile package..."
dpkg-deb --build "$PKG_DIR" "Installer/linux/cv-automation_1.0.0_amd64.deb"

# Clean up build directory
rm -rf "$PKG_DIR"

echo "=============================================="
echo "Debian Package Build Complete!"
echo "Package file: Installer/linux/cv-automation_1.0.0_amd64.deb"
echo "To install, run:"
echo "  sudo apt update && sudo apt install ./Installer/linux/cv-automation_1.0.0_amd64.deb"
echo "=============================================="
