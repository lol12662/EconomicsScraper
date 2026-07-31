# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for WSJ Treasury GUI.
Build:   pyinstaller WSJ_Treasury_GUI.spec
"""

import sys
import os
from pathlib import Path

block_cipher = None

# ── Collect Tcl/Tk data files explicitly (fixes _tcl_data not found on Windows)
def get_tcl_tk_datas():
    """Return datas tuples for Tcl/Tk so Tkinter works in the frozen bundle."""
    import tkinter
    import _tkinter
    datas = []

    # Find the Tcl/Tk library root from _tkinter
    tcl_lib = getattr(_tkinter, 'TCL_LIBRARY', None)
    tk_lib  = getattr(_tkinter, 'TK_LIBRARY',  None)

    # Fallback: walk up from tkinter.__file__
    if not tcl_lib or not tk_lib:
        tk_dir = Path(tkinter.__file__).parent
        python_dir = tk_dir.parent
        for candidate in [python_dir, python_dir.parent]:
            for sub in candidate.iterdir() if candidate.exists() else []:
                name = sub.name.lower()
                if name.startswith('tcl') and sub.is_dir():
                    tcl_lib = str(sub)
                if name.startswith('tk') and sub.is_dir():
                    tk_lib = str(sub)

    if tcl_lib and Path(tcl_lib).exists():
        datas.append((str(tcl_lib), 'tcl'))
    if tk_lib and Path(tk_lib).exists():
        datas.append((str(tk_lib), 'tk'))

    # Also collect DLLs directory on Windows (tcl86t.dll, tk86t.dll etc.)
    if sys.platform == 'win32':
        python_root = Path(sys.executable).parent
        for dll_pattern in ['tcl*.dll', 'tk*.dll', '_tkinter*.pyd']:
            import glob
            for dll in glob.glob(str(python_root / dll_pattern)):
                datas.append((dll, '.'))
        # Also check DLLs subfolder
        dlls_dir = python_root / 'DLLs'
        if dlls_dir.exists():
            for dll_pattern in ['tcl*.dll', 'tk*.dll', '_tkinter*.pyd']:
                for dll in glob.glob(str(dlls_dir / dll_pattern)):
                    datas.append((dll, '.'))

    return datas


tcl_tk_datas = get_tcl_tk_datas()

a = Analysis(
    ['wsj_treasury_gui_fixed.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('wsj_treasury_scraper_fixed.py', '.'),
        *tcl_tk_datas,
    ],
    hiddenimports=[
        # Tkinter — explicit to ensure nothing is missed
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.font',
        '_tkinter',
        # Playwright
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
        # pandas / openpyxl
        'pandas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.timezones',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.writer.excel',
        'openpyxl.reader.excel',
        'openpyxl.drawing.image',
        # requests
        'requests',
        'urllib3',
        'certifi',
        'charset_normalizer',
        'idna',
        # matplotlib
        'matplotlib',
        'matplotlib.backends.backend_tkagg',
        'matplotlib.figure',
        'matplotlib.dates',
        'matplotlib.ticker',
        'matplotlib.gridspec',
        # data / stats
        'yfinance',
        'fredapi',
        'statsmodels',
        'statsmodels.api',
        'statsmodels.formula.api',
        # stdlib
        'json',
        're',
        'threading',
        'traceback',
        'pathlib',
        'datetime',
        'calendar',
        'subprocess',
        'argparse',
        'io',
        'glob',
    ],
    hookspath=[],
    hooksconfig={
        # Tell PyInstaller's matplotlib hook which backend we use
        'matplotlib': {'backends': ['TkAgg']},
    },
    runtime_hooks=[],
    excludes=['ms-playwright'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ── Windows / Linux ────────────────────────────────────────────────────────────
if sys.platform != 'darwin':
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
        argv_emulation=False,
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

# ── macOS ──────────────────────────────────────────────────────────────────────
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
        argv_emulation=True,
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
