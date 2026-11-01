# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['C:/Users/QQQQQ/.dsh/tools/wx-export/安装器.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('C:/Users/QQQQQ/.dsh/tools/wx-export/payload.zip', '.'),
        ('C:/Users/QQQQQ/.dsh/tools/wx-export/app.ico', '.'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['numpy', 'PIL'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='微信导出工具_安装包',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    uac_admin=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='C:/Users/QQQQQ/.dsh/tools/wx-export/app.ico',
)
