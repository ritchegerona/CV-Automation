# -*- mode: python ; coding: utf-8 -*-
import sys
import os
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# SPECPATH is a global variable injected by PyInstaller containing the directory of this spec file.
# The project root is the parent directory of SPECPATH.
project_root = os.path.dirname(SPECPATH)

# Collect all resources from required packages
datas_st, binaries_st, hiddenimports_st = collect_all('streamlit')
datas_docx, binaries_docx, hiddenimports_docx = collect_all('docx')
datas_pypdf, binaries_pypdf, hiddenimports_pypdf = collect_all('pypdf')
datas_dotenv, binaries_dotenv, hiddenimports_dotenv = collect_all('dotenv')

datas = [
    (os.path.join(project_root, 'app.py'), '.'),
    (os.path.join(project_root, '.streamlit'), '.streamlit'),
    (os.path.join(project_root, 'MSR_CV_Template.docx'), '.'),
    (os.path.join(project_root, 'MSR_CV_Template.pdf'), '.'),
] + datas_st + datas_docx + datas_pypdf + datas_dotenv

binaries = binaries_st + binaries_docx + binaries_pypdf + binaries_dotenv

hiddenimports = [
    'streamlit.web.bootstrap',
] + hiddenimports_st + hiddenimports_docx + hiddenimports_pypdf + hiddenimports_dotenv

a = Analysis(
    [os.path.join(SPECPATH, 'launcher.py')],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CV_Automation',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Set to True so terminal output is visible for debugging, or False for windowed app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
