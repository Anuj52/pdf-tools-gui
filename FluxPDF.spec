# -*- mode: python ; coding: utf-8 -*-
import sys
import os
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs

# --- Handle tkinterdnd2 paths helper ---
def get_tkinterdnd2_datas():
    try:
        import tkinterdnd2
        tkdnd_path = os.path.dirname(tkinterdnd2.__file__)
        return [(tkdnd_path, 'tkinterdnd2')]
    except ImportError:
        return []

block_cipher = None

# --- Configuration ---
script_name = 'pdf_tools_tabbed_word_improved.py'  # Your main python file
app_name = 'FluxPDF'
icon_path = 'app.ico' # Ensure this file exists in repo root

# Combine datas
my_datas = []
if os.path.exists(icon_path):
    my_datas.append((icon_path, '.'))
my_datas += get_tkinterdnd2_datas()

a = Analysis(
    [script_name],
    pathex=[],
    binaries=[],
    datas=my_datas,
    hiddenimports=['docx2pdf', 'comtypes', 'win32com', 'tkinterdnd2', 'pandas'],
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
    [],
    exclude_binaries=True, # True for ONEDIR (folder) build
    name=app_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False, # Set to True if you want to see errors in a black box
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path if os.path.exists(icon_path) else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name=app_name,
)