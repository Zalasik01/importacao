# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
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
<<<<<<< HEAD
    a.binaries,
    a.datas,
    [],
=======
    [],
    exclude_binaries=True,
>>>>>>> 3a7936e93850be968fe2721dc6f65bcb8fbb4106
    name='app',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
<<<<<<< HEAD
    upx_exclude=[],
    runtime_tmpdir=None,
=======
>>>>>>> 3a7936e93850be968fe2721dc6f65bcb8fbb4106
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
<<<<<<< HEAD
=======
    icon=['C:\\Users\\Usuario\\Desktop\\importacao\\planilha.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='app',
>>>>>>> 3a7936e93850be968fe2721dc6f65bcb8fbb4106
)
