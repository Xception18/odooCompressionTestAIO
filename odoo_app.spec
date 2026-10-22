# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Collect data files
added_files = [
    ('database_schema.sql', '.'),
    ('.env.example', '.'),
    ('README.md', '.'),
]

# Include Excel templates
import os
if os.path.exists('Data Input Pengujian/data.xlsx'):
    added_files.append(('Data Input Pengujian/data.xlsx', 'Data Input Pengujian'))
if os.path.exists('Pengujian/PENGUJIAN.xlsx'):
    added_files.append(('Pengujian/PENGUJIAN.xlsx', 'Pengujian'))
if os.path.exists('Rencana/RENCANA.xlsx'):
    added_files.append(('Rencana/RENCANA.xlsx', 'Rencana'))

# Include chromedriver if exists
if os.path.exists('chromedriver.exe'):
    added_files.append(('chromedriver.exe', '.'))

# Hidden imports for dynamic imports
hidden_imports = [
    'customtkinter',
    'PIL._tkinter_finder',
    'playwright.sync_api',
    'selenium.webdriver.chrome.service',
    'psycopg2',
    'openpyxl',
    'pandas',
    'xlwings',
    'sqlalchemy.dialects.postgresql',
    'sqlalchemy.pool',
    'dotenv',
    'requests',
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pytest', 'pytest-cov', 'pytest-mock', 'pytest-asyncio', 'tkinter.test'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Include modules directory
a.datas += Tree('modules', prefix='modules', excludes=['__pycache__', '*.pyc'])

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='OdooCompressionTest',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Set to True for debugging
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='icon.ico',  # Uncomment if you have an icon
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='OdooCompressionTest',
)
