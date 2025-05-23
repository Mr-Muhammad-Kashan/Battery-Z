# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['Battery_Z.py'],
    pathex=['C:\\Users\\MKATW\\OneDrive\\Desktop\\Laptop-Apps\\Project'],
    binaries=[],
    datas=[
        ('logo.svg', '.'),
        ('logo.ico', '.')
    ],
    hiddenimports=[
        'psutil',
        'pywin32',
        'wmi',
        'pandas',
        'matplotlib',
        'beautifulsoup4',
        'PyQt5',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'win32api',
        'win32con'
    ],
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
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Battery-Z',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='logo.ico',
    uac_admin=True
)