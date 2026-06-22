# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for WSJ Treasury GUI.
Build:   pyinstaller WSJ_Treasury_GUI.spec
"""

import sys
from pathlib import Path

block_cipher = None

a = Analysis(
    ['wsj_treasury_gui_fixed.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('wsj_treasury_scraper_fixed.py', '.'),
    ],
    hiddenimports=[
        'playwright',
        'playwright.sync_api',
        'playwright._impl._api_types',
        'playwright._impl._browser',
        'playwright._impl._browser_context',
        'playwright._impl._page',
        'playwright._impl._playwright',
        'playwright._impl._connection',
        'playwright._impl._helper',
        'playwright._impl._network',
        'playwright._impl._element_handle',
        'playwright._impl._js_handle',
        'playwright._impl._frame',
        'playwright._impl._errors',
        'playwright._impl._input',
        'playwright._impl._transport',
        'pandas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.timezones',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.writer.excel',
        'openpyxl.reader.excel',
        'requests',
        'urllib3',
        'certifi',
        'charset_normalizer',
        'idna',
        'matplotlib',
        'matplotlib.backends.backend_tkagg',
        'matplotlib.figure',
        'matplotlib.dates',
        'matplotlib.ticker',
        'yfinance',
        'fredapi',
        'statsmodels',
        'statsmodels.api',
        'statsmodels.formula.api',
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
        'subprocess',
        'argparse',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['ms-playwright'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ── Windows / Linux: onedir bundle ────────────────────────────────────────────
if sys.platform != 'darwin':
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,      # onedir: libs go in the COLLECT folder
        name='WSJ_Treasury_GUI',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        console=False,              # no terminal window
        disable_windowed_traceback=False,
        argv_emulation=False,       # Windows/Linux: must be False
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
        upx=False,
        upx_exclude=[],
        name='WSJ_Treasury_GUI',    # → dist/WSJ_Treasury_GUI/WSJ_Treasury_GUI.exe
    )

# ── macOS: onedir → .app bundle ───────────────────────────────────────────────
else:
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name='WSJ_Treasury_GUI',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        console=False,
        disable_windowed_traceback=False,
        argv_emulation=True,        # macOS only: enables drag-open
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
        upx=False,
        upx_exclude=[],
        name='WSJ_Treasury_GUI',
    )
    app = BUNDLE(
        coll,
        name='WSJ_Treasury_GUI.app',
        icon=None,
        bundle_identifier='com.treasury.scraper',
        info_plist={
            'NSHighResolutionCapable': True,
            'LSBackgroundOnly': False,
            'NSRequiresAquaSystemAppearance': False,
        },
    )
