# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['sqlcipher3', 'docx', 'pinyin', 'Crypto']
hiddenimports += collect_submodules('Crypto')


a = Analysis(
    ['C:/Users/QQQQQ/.dsh/tools/wx-export/微信导出App.pyw'],
    pathex=[],
    binaries=[],
    datas=[('C:/Users/QQQQQ/.dsh/tools/wx-export/vendor', 'vendor'),
           ('C:/Users/QQQQQ/.dsh/tools/wx-export/app.ico', '.'),
           # 微信 4.x 的图片是 wxgf(HEVC) 封装，转成 JPG 必须靠 ffmpeg
           ('C:/Users/QQQQQ/.dsh/tools/wx-export/ffmpeg.exe', '.')],
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
    name='微信导出工具',
    debug=False,
    icon='C:/Users/QQQQQ/.dsh/tools/wx-export/app.ico',
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    uac_admin=True,
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
    name='微信导出工具',
)
