# -*- mode: python ; coding: utf-8 -*-

import os

project_dir = os.path.abspath(os.path.dirname(SPEC))
tesseract_dir = os.path.join(project_dir, "Tesseract-OCR")

a = Analysis(
    ["main.py"],
    pathex=[project_dir],

    binaries=[],

    datas=[
        (tesseract_dir, "Tesseract-OCR"),
        ("CHANGELOG.txt", "."),
    ],

    hiddenimports=[
        "pdfplumber",
        "openpyxl",
        "PIL",
        "PIL.Image",
        "PIL.ImageOps",
        "pytesseract",
        "history_db",
        "template_config",
        "settings_manager",
        "settings_dialog",
        "update_manager",
        "version",
    ],

    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Allisons Data Extractor",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon="app_icon.ico",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="Allisons Data Extractor",
)