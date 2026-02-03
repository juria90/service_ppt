# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['service_ppt/__main__.py'],
    pathex=["."],
    binaries=[],
    datas=[("service_ppt/image24", "service_ppt/image24"), ("service_ppt/image32", "service_ppt/image32"), ("service_ppt/locale", "service_ppt/locale"), ("sample", "sample")],
    hiddenimports=[],
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
    exclude_binaries=True,
    name='service_ppt',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='service_ppt',
)
