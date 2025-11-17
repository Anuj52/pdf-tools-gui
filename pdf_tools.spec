# PyInstaller spec example for pdf_tools_tabbed_word_improved.py
# Save this text to pdf_tools.spec and run:
#   pyinstaller --onefile pdf_tools.spec

block_cipher = None

a = Analysis(
    ['pdf_tools_tabbed_word_improved.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['docx2pdf'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='pdf_tools',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='your_icon.ico'  # optional
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='pdf_tools_dist'
)
