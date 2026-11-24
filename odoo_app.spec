# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Collect data files
added_files = [
    ('database_schema.sql', '.'),
    ('.env.example', '.'),
    ('README.md', '.'),
    ('INSTALL_PLAN.md', '.'),
]

# Include Excel templates
import os
for folder, filename in [
    ('Data Input Pengujian', 'data.xlsx'),
    ('Pengujian', 'PENGUJIAN.xlsx'),
    ('Rencana', 'RENCANA.xlsx'),
]:
    path = os.path.join(folder, filename)
    if os.path.exists(path):
        added_files.append((path, folder))

# Include chromedriver if exists (Selenium fallback)
if os.path.exists('chromedriver.exe'):
    added_files.append(('chromedriver.exe', '.'))

# Hidden imports for dynamic imports
hidden_imports = [
    'customtkinter',
    'PIL._tkinter_finder',
    'playwright.sync_api',
    'playwright._impl',
    'playwright._impl._driver',
    'selenium.webdriver.chrome.service',
    'psycopg2',
    'psycopg2._psycopg',
    'openpyxl',
    'pandas',
    'pandas._libs',
    'xlwings',
    'sqlalchemy.dialects.postgresql',
    'sqlalchemy.dialects.postgresql.psycopg2',
    'sqlalchemy.pool',
    'dotenv',
    'requests',
    'json',
    'modules.beban_config',
    'modules.excel_handler_config',
    'modules.first_run_setup',
    'modules.shutdown_manager',
    'modules.db_pool',
    'modules.daemon_sync',
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
    excludes=[
        'pytest', 'pytest-cov', 'pytest-mock', 'pytest-asyncio',
        'tkinter.test', 'unittest', 'test',
	# Exclude large unused ML/CV packages
        'torch', 'torchvision', 'torchaudio',
        'tensorflow', 'keras',
        'cv2', 'opencv',
        'scipy', 'scipy.stats', 'scipy.optimize',
        'sklearn', 'scikit-learn',
        'transformers', 'tokenizers',
        'onnx', 'onnxruntime',
        'numba', 'llvmlite',
        'av', 'ffmpeg',
        'pyarrow',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Include modules directory (catch-all)
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
    console=False,  # Hide console in production
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='icon.ico',
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
