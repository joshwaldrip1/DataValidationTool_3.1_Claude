# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

hiddenimports = ['win32com.client', 'pythoncom']
hiddenimports += collect_submodules('win32com')
hiddenimports += collect_submodules('tkinterdnd2')
hiddenimports += collect_submodules('pandas')
hiddenimports += collect_submodules('openpyxl')
hiddenimports += collect_submodules('ezdxf')
hiddenimports += collect_submodules('pyproj')

# pyproj needs its data files (proj.db, etc.) for CRS lookups
datas = [('config.json', '.')]
datas += collect_data_files('pyproj')
datas += collect_data_files('ezdxf')

# Try to include pyogrio if installed
try:
    hiddenimports += collect_submodules('pyogrio')
    datas += collect_data_files('pyogrio')
except Exception:
    pass

a = Analysis(
    ['DataValidationTool-v3.2.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='DataValidationTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
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
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DataValidationTool',
)
