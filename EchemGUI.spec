# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_submodules

# Collect everything from packages that embed data files
mpl_datas,  mpl_binaries,  mpl_hiddens  = collect_all('matplotlib')
opxl_datas, opxl_binaries, opxl_hiddens = collect_all('openpyxl')
pd_datas,   pd_binaries,   pd_hiddens   = collect_all('pandas')

# Anaconda keeps several runtime DLLs in Library\bin that PyInstaller's
# dependency scanner misses.  Add them explicitly so the frozen exe finds them.
ANACONDA_BIN = r'C:\Users\Mefford\anaconda3\Library\bin'
extra_binaries = [
    (fr'{ANACONDA_BIN}\ffi-8.dll',   '.'),
    (fr'{ANACONDA_BIN}\ffi-7.dll',   '.'),
    (fr'{ANACONDA_BIN}\ffi.dll',     '.'),
    (fr'{ANACONDA_BIN}\tcl86t.dll',  '.'),
    (fr'{ANACONDA_BIN}\tk86t.dll',   '.'),
    (fr'{ANACONDA_BIN}\sqlite3.dll', '.'),
]

a = Analysis(
    ['run_echem.py'],
    pathex=[],
    binaries=mpl_binaries + opxl_binaries + pd_binaries + extra_binaries,
    datas=mpl_datas + opxl_datas + pd_datas,
    hiddenimports=(
        mpl_hiddens + opxl_hiddens + pd_hiddens
        + [
            # matplotlib TkAgg backend
            'matplotlib.backends.backend_tkagg',
            'matplotlib.backends._backend_tk',
            'matplotlib.backends.backend_agg',
            # tkinter sub-modules used at runtime
            'tkinter',
            'tkinter.ttk',
            'tkinter.messagebox',
            'tkinter.filedialog',
            'tkinter.simpledialog',
            'tkinter.colorchooser',
            # stdlib used in plotting.py
            'colorsys',
        ]
    ),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'wx', 'IPython', 'jupyter'],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='EchemGUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,     # no console window — GUI app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='EchemGUI',
)
