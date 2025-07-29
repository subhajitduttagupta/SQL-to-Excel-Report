# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['SQL_to_REPORT.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('Report_Template.xlsx', '.'),
        ('logo.png', '.')
    ],
    hiddenimports=[
        'pyodbc',
        'sqlalchemy.dialects.mssql',
        'sqlalchemy.dialects.mssql.pyodbc'
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
    a.binaries,
    a.datas,
    [],
    name='SQL_to_Report',
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
