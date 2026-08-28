#!/bin/bash
# Helper script to install system dependencies on macOS (Homebrew & LibreOffice)

echo "=============================================="
echo "Checking dependencies for CV Automation..."
echo "=============================================="

# Check if LibreOffice is already installed at standard locations
if [ -d "/Applications/LibreOffice.app" ] || [ -f "/opt/homebrew/bin/soffice" ] || [ -f "/usr/local/bin/soffice" ]; then
    echo "[INFO] LibreOffice is already installed on this machine."
    exit 0
fi

# Check if Homebrew is installed
if ! command -v brew &> /dev/null; then
    echo "[INFO] Homebrew not found. Installing Homebrew first..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    
    # Enable brew in current terminal session
    if [[ $(uname -m) == "arm64" ]]; then
        eval "$(/opt/homebrew/bin/brew shellenv)"
    else
        eval "$(/usr/local/bin/brew shellenv)"
    fi
else
    echo "[INFO] Homebrew is already installed."
fi

echo "[INFO] Installing LibreOffice (Community Edition) via Homebrew Cask..."
brew install --cask libreoffice

# Verify installation
if [ -d "/Applications/LibreOffice.app" ]; then
    echo "=============================================="
    echo "Success! LibreOffice was successfully installed."
    echo "=============================================="
else
    echo "=============================================="
    echo "[ERROR] Installation completed but LibreOffice.app was not found."
    echo "Please download and install it manually from https://www.libreoffice.org/"
    echo "=============================================="
fi
