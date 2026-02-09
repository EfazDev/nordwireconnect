# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
import os
try:
    from PyInstaller.building.api import *
    from PyInstaller.building.build_main import *
    from PyInstaller.building.osx import *
except:
    print("Disabled Visual Studio Code Mode")

block_cipher = None
a = Analysis(
    ['Main.py', 'Service.py'],
    pathex=[],
    binaries=[],
    datas=collect_data_files("NordWireConnect") + collect_data_files("NordWireConnectService"),
    hiddenimports=['PyKits', 'win32timezone', 'win32cred', 'pywintypes'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    [a.scripts[-2]],
    [],
    exclude_binaries=True,
    name='NordWireConnect',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=True,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['Resources\\app_icon.ico'],
)
exe2 = EXE(
    pyz,
    [a.scripts[-1]],
    [],
    exclude_binaries=True,
    name='NordWireConnectService',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['Resources\\app_icon.ico'],
)
coll = COLLECT(
    exe,
    exe2,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='NordWireConnect',
)
