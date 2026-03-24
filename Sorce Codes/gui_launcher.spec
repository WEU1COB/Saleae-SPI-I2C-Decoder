# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# Collect all submodules and data for saleae package
hidden_imports = collect_submodules('saleae') + collect_submodules('openpyxl')
data_files = collect_data_files('saleae') + collect_data_files('openpyxl')

a = Analysis(
    ['gui_launcher.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('chip_icon.ico', '.'),  # include your icon
        *data_files              # saleae + openpyxl data files
    ],
    hiddenimports=hidden_imports + [
        'decode_spi_core',
        'decode_i2c_core',
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
    name='SaleaeDecoderTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # False = no console window (GUI only)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='chip_icon.ico',  # use icon for exe
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SPI_I2C_Decoder',
)
