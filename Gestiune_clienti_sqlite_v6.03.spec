# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['gestiune_clienti_sqlite.py'],
    pathex=[],
    binaries=[],
    datas=[('baza_date.db', '.'), ('mapping_tehnician.xlsx', '.'), ('mapping_model_amef.xlsx', '.'), ('template', 'template'), ('semnaturi', 'semnaturi'), ('icons', 'icons')],
    hiddenimports=['tkcalendar', 'python_dotenv', 'dateutil', 'cryptography', 'reportlab', 'pypdf', 'pandas', 'docxtpl', 'openpyxl', 'jinja2', 'xlrd', 'pytz', 'babel', 'babel.numbers', 'babel.localedata', 'babel.dates', 'PIL', 'win32com', 'win32com.client', 'pywin32'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Gestiune_clienti_sqlite_v6.03',
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
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Gestiune_clienti_sqlite_v6.03',
)
