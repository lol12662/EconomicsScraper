# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for WSJ Treasury GUI.
Build with:   pyinstaller WSJ_Treasury_GUI.spec
"""

import sys
from pathlib import Path

block_cipher = None

a = Analysis(
    ['wsj_treasury_gui_fixed.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        # Bundle the scraper module alongside the GUI
        ('wsj_treasury_scraper_fixed.py', '.'),
    ],
    hiddenimports=[
        # Playwright
        'playwright',
        'playwright.sync_api',
        'playwright._impl._api_types',
        'playwright._impl._browser',
        'playwright._impl._browser_context',
        'playwright._impl._page',
        # pandas / openpyxl
        'pandas',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.writer.excel',
        # requests
        'requests',
        'urllib3',
        'certifi',
        'charset_normalizer',
        # stdlib bits that sometimes get missed
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'json',
        're',
        'threading',
        'traceback',
        'pathlib',
        'datetime',
        'calendar',
    ],
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
    exclude_binaries=True,   # onedir mode — much more reliable on Mac
    name='WSJ_Treasury_GUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,           # no terminal window
    disable_windowed_traceback=False,
    argv_emulation=True,     # needed on macOS for GUI apps
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='WSJ_Treasury_GUI',
)

# macOS: wrap in a proper .app bundle
if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='WSJ_Treasury_GUI.app',
        icon=None,
        bundle_identifier='com.treasury.scraper',
        info_plist={
            'NSHighResolutionCapable': True,
            'LSBackgroundOnly': False,
        },
    )
