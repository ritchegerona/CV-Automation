# CV Automation Installer Packaging Guides

This folder contains the build files, configurations, and scripts required to compile and bundle the **MSR CV Processing Studio** (Streamlit app) into native installer formats for Windows, macOS, and Linux.

---

## Folder Contents
- [launcher.py](file:///Users/ritchegerona/Documents/PROJECTS/CV%20Automation/Installer/launcher.py): Programmatic Python script that starts the Streamlit server from inside the PyInstaller bundle and copies template files to the current directory on launch.
- [launcher.spec](file:///Users/ritchegerona/Documents/PROJECTS/CV%20Automation/Installer/launcher.spec): PyInstaller specification configuration that collects and packages all Streamlit, docx, pypdf, and template assets.
- [windows/](file:///Users/ritchegerona/Documents/PROJECTS/CV%20Automation/Installer/windows/): 
  - `build_win.bat`: Builds the Windows `.exe` using PyInstaller.
  - `installer_setup.iss`: Inno Setup file to create the setup `.exe` installer (installs the app and offers silent download/install of LibreOffice if missing).
- [macos/](file:///Users/ritchegerona/Documents/PROJECTS/CV%20Automation/Installer/macos/):
  - `build_mac.sh`: Builds the macOS executable and wraps it inside a `.dmg` installer.
  - `install_mac_deps.sh`: Installs Homebrew and LibreOffice on client macOS systems.
- [linux/](file:///Users/ritchegerona/Documents/PROJECTS/CV%20Automation/Installer/linux/):
  - `build_linux.sh`: Builds the Linux executable using PyInstaller.
  - `build_deb.sh`: Packages the executable into a `.deb` file with declared dependencies (`libreoffice-core`, `libreoffice-writer`).

---

## 🪟 Windows Compilation & Installation

### How to Build (for Developers):
1. Install **Inno Setup** on your Windows build system (download from [jrsoftware.org](https://www.jrsoftware.org/isdl.php)).
2. Open PowerShell or Command Prompt in the project folder and run the build batch file:
   ```cmd
   .\Installer\windows\build_win.bat
   ```
   *This compiles the application and outputs `CV_Automation.exe` into `Installer/windows/dist/`.*
3. Open Inno Setup Compiler, load `Installer\windows\installer_setup.iss`, and click **Build -> Compile**.
4. The setup installer will be generated at `Installer\windows\dist\MSR_CV_Automation_Setup.exe`.

### How to Install (for Users):
1. Double-click `MSR_CV_Automation_Setup.exe`.
2. The installer will automatically scan the system for LibreOffice. 
3. If LibreOffice is **not** found, it prompts the user to download it:
   - Clicking **Yes** will trigger a silent download and background installation of LibreOffice.
4. The installer copies the CV Automation app and sets up Desktop/Start Menu shortcuts.

---

## 🍎 macOS Compilation & Installation

### How to Build (for Developers):
On your macOS machine, run the build shell script:
```bash
./Installer/macos/build_mac.sh
```
This will:
- Activate your Python virtual environment.
- Package the application using PyInstaller.
- Produce a drag-and-drop Disk Image installer at `Installer/macos/CV_Automation_Setup.dmg`.

### How to Install (for Users):
1. Double-click `CV_Automation_Setup.dmg`.
2. Drag the `CV_Automation` application into your `Applications` shortcut folder.
3. If you do not have LibreOffice installed, double-click the `install_mac_deps.sh` script included in the DMG to automatically install Homebrew and LibreOffice via terminal, or install it manually.

---

## 🐧 Linux Compilation & Installation

### How to Build (for Developers):
On your Linux system, run the Debian package builder script:
```bash
./Installer/linux/build_deb.sh
```
This will:
- Activate the virtual environment and build the Linux executable.
- Set up the Debian package directory structure.
- Declare APT package dependencies (`libreoffice-core`, `libreoffice-writer`).
- Compile the package into `Installer/linux/cv-automation_1.0.0_amd64.deb`.

### How to Install (for Users):
Run the following commands in the terminal to install the `.deb` package and let the package manager automatically fetch and install its dependencies:
```bash
sudo apt update
sudo apt install ./cv-automation_1.0.0_amd64.deb
```
*APT will automatically download and configure LibreOffice and Python if they are not already installed on the machine.*
*Once installed, you can launch the app from your application menu or run `cv-automation` from the command line.*

---

## Runtime Usage Notes
Upon running the installer-created executable on any operating system:
1. The app automatically creates a template `.env` file in the user's current directory (if not already present). The user should edit this file and add their `GEMINI_API_KEY`.
2. The app copies `MSR_CV_Template.docx` and `MSR_CV_Template.pdf` to the folder it runs from.
3. It sets up the `Candidate CV Summary/` directory automatically to store output summaries.
