# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for SSA PDF Studio.

PyInstaller is not a cross-compiler, so this spec must be run once per target
OS: on Windows it produces a single-file dist/SSA PDF Studio.exe, and on macOS
it produces dist/SSA PDF Studio.app. The build.yml workflow runs both.
"""

import os
import sys

IS_MAC = sys.platform == "darwin"

APP_NAME = "SSA PDF Studio"
APP_VERSION = "1.0"

# Pdf.py pulls these in lazily inside functions, so PyInstaller's static
# analysis cannot always see them. PyMuPDF is importable under both names.
hidden_imports = [
    "pymupdf",
    "fitz",
    "pdfplumber",
    "openpyxl",
    "docx",
    "pandas",
    "reportlab.pdfgen.canvas",
    "PIL._tkinter_finder",
]

# Keeps the bundle from ballooning with scientific/GUI stacks pandas can drag in.
excluded_modules = [
    "matplotlib",
    "scipy",
    "PyQt5",
    "PyQt6",
    "PySide2",
    "PySide6",
    "IPython",
    "notebook",
    "pytest",
]

# Optional branding: used only when the file is actually present, so a missing
# icon can never fail the build.
_icon_candidate = os.path.join("assets", "icon.icns" if IS_MAC else "icon.ico")
icon_path = _icon_candidate if os.path.isfile(_icon_candidate) else None

a = Analysis(
    ["Pdf.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excluded_modules,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

if IS_MAC:
    # A .app must be onedir: a onefile binary re-extracts to a temp dir on every
    # launch, which is slow and trips Gatekeeper path checks.
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name=APP_NAME,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        # UPX is not reliably available on macOS and can corrupt signed dylibs.
        upx=False,
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=True,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=icon_path,
    )
    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=False,
        upx_exclude=[],
        name=APP_NAME,
    )
    app = BUNDLE(
        coll,
        name=f"{APP_NAME}.app",
        icon=icon_path,
        bundle_identifier="com.ssa.pdfstudio",
        version=APP_VERSION,
        info_plist={
            "CFBundleName": APP_NAME,
            "CFBundleDisplayName": APP_NAME,
            "CFBundleShortVersionString": APP_VERSION,
            "CFBundleVersion": APP_VERSION,
            "NSHighResolutionCapable": True,
            "LSMinimumSystemVersion": "11.0",
            # Without this the Tk file dialogs cannot reach ~/Documents etc.
            "NSDesktopFolderUsageDescription": "Open and save your PDF files.",
            "NSDocumentsFolderUsageDescription": "Open and save your PDF files.",
            "NSDownloadsFolderUsageDescription": "Open and save your PDF files.",
        },
    )
else:
    # Windows keeps the existing single-file .exe behaviour.
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        [],
        name=APP_NAME,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        upx_exclude=[],
        runtime_tmpdir=None,
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=icon_path,
    )
