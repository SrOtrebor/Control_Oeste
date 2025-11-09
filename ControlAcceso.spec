# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


# ControlAcceso.spec

a = Analysis(['app.py'],
             pathex=[],
             binaries=[],
             # AQUÍ AGREGAMOS NUESTRAS CARPETAS Y ARCHIVOS
             datas=[
                 ('templates', 'templates'),
                 ('static', 'static')
             ],
             # AQUÍ AGREGAMOS LIBRERÍAS OCULTAS QUE PANDAS USA
             hiddenimports=['pandas._libs.tslibs.base', 'openpyxl'],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ControlAcceso',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
