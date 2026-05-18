# -*- mode: python ; coding: utf-8 -*-
# bir2307.spec — PyInstaller build spec for BIR Form 2307 Generator
#
# Build command:
#   pyinstaller bir2307.spec --clean
#
# Output: dist/BIR2307Generator/BIR2307Generator.exe

import sys
from pathlib import Path

block_cipher = None

# ── Data files bundled into the exe ──────────────────────────────────────────
added_files = [
    # (source_path, dest_folder_in_bundle)
    ("assets/BIR2307_template.xlsx", "assets"),
    ("config/app_config.json",       "config"),
    ("config/cell_mapping.json",     "config"),
]

# ── Hidden imports (modules discovered at runtime) ────────────────────────────
hidden = [
    "customtkinter",
    "openpyxl",
    "pandas",
    "pyodbc",
    "win32com",
    "win32com.client",
    "pywintypes",
    "tkinter",
    "tkinter.messagebox",
    "tkinter.filedialog",
]

a = Analysis(
    ["main.py"],
    pathex=["."],
    binaries=[],
    datas=added_files,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib", "scipy", "numpy.testing", "pytest"],
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
    exclude_binaries=True,
    name="BIR2307Generator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,         # No console window in production
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon="assets/icon.ico",   # Uncomment after adding icon.ico
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="BIR2307Generator",
)
