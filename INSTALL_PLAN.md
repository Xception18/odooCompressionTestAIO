# Plan Pembuatan Installer .exe — Odoo Compression Test

## Overview

Membuat installer `.exe` all-in-one untuk Windows yang meng-bundle:
1. **Aplikasi utama** (PyInstaller → single folder executable)
2. **PostgreSQL** (embedded/portable atau auto-install)
3. **Python runtime** (embedded di PyInstaller bundle)
4. **Chromium browser** (untuk Playwright web automation)

Target output: `OdooCompressionTestSetup.exe` (~300-500MB) yang bisa di-install di PC Windows 10/11 tanpa perlu install Python, PostgreSQL, atau browser manual.

---

## Arsitektur Installer

```
OdooCompressionTestSetup.exe (Inno Setup)
├── OdooCompressionTest/          ← PyInstaller output (app + Python runtime)
│   ├── OdooCompressionTest.exe   ← Main executable
│   ├── _internal/                ← Python libs, DLLs
│   └── modules/                  ← App modules
├── pgsql/                        ← PostgreSQL portable (opsional)
│   ├── bin/
│   ├── lib/
│   └── data/
├── playwright-browsers/          ← Chromium binary
│   └── chromium-*/
├── .env.example                  ← Template konfigurasi
├── database_schema.sql           ← Schema DB
└── Excel templates/              ← Template Excel
```

---

## Tahap 1: Persiapan Build Environment

### 1.1 Prerequisites di Mesin Build

```
Tool                     Versi Minimum    Download
─────────────────────────────────────────────────────────────────
Python                   3.10+            python.org
PyInstaller              6.x              pip install pyinstaller
Inno Setup 6             6.2+             jrsoftware.org/isdl.php
Git                      any              git-scm.com
```

### 1.2 Setup Virtual Environment

```bash
cd odooCompressionTest
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pip install pyinstaller
```

### 1.3 Verifikasi Dependencies

```bash
# Test app berjalan normal
python main.py

# Pastikan playwright browser ter-install
playwright install chromium

# Catat lokasi chromium binary
python -c "from playwright._impl._driver import compute_driver_executable; print(compute_driver_executable())"
```

---

## Tahap 2: Download & Siapkan Bundled Components

### 2.1 PostgreSQL Portable

**Opsi A: PostgreSQL Embedded (Recommended)**

Download PostgreSQL binaries-only (tanpa installer) dari:
- URL: `https://get.enterprisedb.com/postgresql/postgresql-16.9-1-windows-x64-binaries.zip`
- Atau: https://www.postgresql.org/download/ → "Zip Archive" (Windows x86-64)

```bash
# Extract ke folder project
mkdir vendor\pgsql
# Extract isi zip ke vendor\pgsql\ (pastikan bin/, lib/, share/ ada)
```

Struktur yang dibutuhkan dari zip:
```
vendor/pgsql/
├── bin/
│   ├── postgres.exe
│   ├── pg_ctl.exe
│   ├── psql.exe
│   ├── initdb.exe
│   ├── pg_isready.exe
│   └── *.dll
├── lib/
│   └── *.dll
└── share/
    ├── postgresql.conf.sample
    └── pg_hba.conf.sample
```

**Opsi B: Auto-Download saat Install (Lighter installer)**

Installer hanya download PostgreSQL saat install jika belum ada. Lebih kecil tapi butuh internet.

### 2.2 Chromium untuk Playwright

**Opsi A: Bundle di installer (Recommended — offline install)**

```bash
# Lokasi default Playwright browsers di Windows:
# %LOCALAPPDATA%\ms-playwright\chromium-*\

# Copy folder chromium ke vendor
xcopy "%LOCALAPPDATA%\ms-playwright\chromium-*" vendor\playwright-browsers\ /E /I
```

Atau download langsung:
```bash
# Cek versi yang dibutuhkan playwright
python -c "from playwright._impl._driver import compute_driver_executable; import subprocess; subprocess.run([compute_driver_executable(), 'install', '--dry-run', 'chromium'])"
```

**Opsi B: First-run download (sudah ada di `first_run_setup.py`)**

Existing code sudah handle ini — `setup_playwright()` download Chromium ~170MB saat pertama kali app dijalankan. Cocok jika ingin installer lebih kecil.

### 2.3 Excel Templates

Pastikan file-file ini ada:
```
Data Input Pengujian/data.xlsx
Pengujian/PENGUJIAN.xlsx
Rencana/RENCANA.xlsx
```

---

## Tahap 3: Konfigurasi PyInstaller

### 3.1 Update `odoo_app.spec`

File `.spec` perlu di-update untuk include bundled components:

```python
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

added_files = [
    ('database_schema.sql', '.'),
    ('.env.example', '.'),
    ('README.md', '.'),
    ('INSTALL_PLAN.md', '.'),
    # beban config
    ('beban_config.json', '.'),
]

# Excel templates
import os
for folder, filename in [
    ('Data Input Pengujian', 'data.xlsx'),
    ('Pengujian', 'PENGUJIAN.xlsx'),
    ('Rencana', 'RENCANA.xlsx'),
]:
    path = os.path.join(folder, filename)
    if os.path.exists(path):
        added_files.append((path, folder))

# Chromedriver (Selenium fallback)
if os.path.exists('chromedriver.exe'):
    added_files.append(('chromedriver.exe', '.'))

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
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Include modules directory (catch-all untuk modules yang mungkin terlewat)
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
    console=False,      # GUI app, no console
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='icon.ico',  # Uncomment jika ada icon
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
```

### 3.2 Build Executable

```bash
# Clean build
rmdir /s /q build dist

# Build
pyinstaller odoo_app.spec

# Test executable
dist\OdooCompressionTest\OdooCompressionTest.exe
```

### 3.3 Troubleshooting Build

| Problem | Solusi |
|---------|--------|
| `ModuleNotFoundError: xlwings` | Tambah `'xlwings'` ke `hiddenimports` |
| `playwright._impl not found` | Tambah `'playwright._impl._driver'` ke `hiddenimports` |
| `psycopg2` DLL error | Pastikan pakai `psycopg2-binary`, bukan `psycopg2` |
| App crash tanpa error | Ubah `console=True` di `.spec` untuk lihat traceback |
| `.env` not found | `first_run_setup.py` handle ini otomatis |
| `customtkinter` theme missing | CTk biasanya bundle otomatis; jika tidak, tambah `Tree('customtkinter')` |

---

## Tahap 4: Konfigurasi Inno Setup Installer

### 4.1 Script `installer.iss`

```iss
; Inno Setup Script for Odoo Compression Test
; Full installer dengan PostgreSQL + Chromium

#define MyAppName "Odoo Compression Test"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Adhimix RMC"
#define MyAppExeName "OdooCompressionTest.exe"

[Setup]
AppId={{E8F9A2B3-4C5D-6E7F-8901-2A3B4C5D6E7F}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=output
OutputBaseFilename=OdooCompressionTestSetup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
UninstallDisplayIcon={app}\{#MyAppExeName}
SetupIconFile=icon.ico               ; Opsional: icon installer
; WizardImageFile=wizard.bmp         ; Opsional: gambar wizard
; WizardSmallImageFile=wizard_sm.bmp ; Opsional: gambar kecil wizard

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "indonesian"; MessagesFile: "compiler:Languages\Indonesian.isl"

[Types]
Name: "full"; Description: "Full Installation (Recommended)"
Name: "compact"; Description: "Compact (tanpa PostgreSQL)"
Name: "custom"; Description: "Custom Installation"; Flags: iscustom

[Components]
Name: "app"; Description: "Odoo Compression Test Application"; Types: full compact custom; Flags: fixed
Name: "pgsql"; Description: "PostgreSQL 16 Database Server (~180MB)"; Types: full
Name: "chromium"; Description: "Chromium Browser untuk Web Automation (~170MB)"; Types: full

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"
Name: "initdb"; Description: "Initialize database (create alatujidb)"; Components: pgsql
Name: "pgservice"; Description: "Register PostgreSQL as Windows Service"; Components: pgsql

[Files]
; === APLIKASI UTAMA ===
Source: "dist\OdooCompressionTest\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: app

; === POSTGRESQL PORTABLE ===
Source: "vendor\pgsql\*"; DestDir: "{app}\pgsql"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: pgsql

; === CHROMIUM BROWSER ===
Source: "vendor\playwright-browsers\*"; DestDir: "{localappdata}\ms-playwright"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: chromium

; === CONFIG & DOCS ===
Source: ".env.example"; DestDir: "{app}"; Flags: ignoreversion; Components: app
Source: "database_schema.sql"; DestDir: "{app}"; Flags: ignoreversion; Components: app
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion isreadme; Components: app
Source: "INSTALL_PLAN.md"; DestDir: "{app}"; Flags: ignoreversion; Components: app

; === EXCEL TEMPLATES ===
Source: "Data Input Pengujian\data.xlsx"; DestDir: "{userappdata}\{#MyAppName}\Data Input Pengujian"; Flags: ignoreversion onlyifdoesntexist; Components: app
Source: "Pengujian\PENGUJIAN.xlsx"; DestDir: "{userappdata}\{#MyAppName}\Pengujian"; Flags: ignoreversion onlyifdoesntexist; Components: app
Source: "Rencana\RENCANA.xlsx"; DestDir: "{userappdata}\{#MyAppName}\Rencana"; Flags: ignoreversion onlyifdoesntexist; Components: app

[Dirs]
Name: "{userappdata}\{#MyAppName}"; Permissions: users-modify
Name: "{userappdata}\{#MyAppName}\Output"; Permissions: users-modify
Name: "{userappdata}\{#MyAppName}\Rencana"; Permissions: users-modify
Name: "{userappdata}\{#MyAppName}\Pengujian"; Permissions: users-modify
Name: "{userappdata}\{#MyAppName}\Data Input Pengujian"; Permissions: users-modify
Name: "{app}\pgsql\data"; Permissions: users-modify; Components: pgsql

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Edit Configuration"; Filename: "notepad.exe"; Parameters: """{userappdata}\{#MyAppName}\.env"""
Name: "{group}\Open Data Folder"; Filename: "{userappdata}\{#MyAppName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Initialize PostgreSQL database cluster
Filename: "{app}\pgsql\bin\initdb.exe"; Parameters: "-D ""{app}\pgsql\data"" -U postgres -E UTF8 --locale=C"; StatusMsg: "Initializing PostgreSQL..."; Tasks: initdb; Flags: runhidden waituntilterminated
; Start PostgreSQL, create database, import schema
Filename: "{app}\pgsql\bin\pg_ctl.exe"; Parameters: "start -D ""{app}\pgsql\data"" -l ""{app}\pgsql\postgresql.log"" -w"; StatusMsg: "Starting PostgreSQL..."; Tasks: initdb; Flags: runhidden waituntilterminated
Filename: "{app}\pgsql\bin\psql.exe"; Parameters: "-U postgres -c ""CREATE DATABASE alatujidb;"""; StatusMsg: "Creating database..."; Tasks: initdb; Flags: runhidden waituntilterminated
Filename: "{app}\pgsql\bin\psql.exe"; Parameters: "-U postgres -d alatujidb -f ""{app}\database_schema.sql"""; StatusMsg: "Importing schema..."; Tasks: initdb; Flags: runhidden waituntilterminated
; Register as Windows Service (opsional)
Filename: "{app}\pgsql\bin\pg_ctl.exe"; Parameters: "register -N ""OdooCompressionTestPG"" -D ""{app}\pgsql\data"" -U ""NT AUTHORITY\NetworkService"""; StatusMsg: "Registering PostgreSQL service..."; Tasks: pgservice; Flags: runhidden waituntilterminated
; Launch app
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: postinstall nowait skipifsilent runascurrentuser

[UninstallRun]
; Stop dan unregister PostgreSQL service
Filename: "{app}\pgsql\bin\pg_ctl.exe"; Parameters: "stop -D ""{app}\pgsql\data"""; Flags: runhidden; RunOnceId: "StopPG"
Filename: "{app}\pgsql\bin\pg_ctl.exe"; Parameters: "unregister -N ""OdooCompressionTestPG"""; Flags: runhidden; RunOnceId: "UnregPG"

[Code]
var
  ConfigPage: TInputQueryWizardPage;
  OdooConfigPage: TInputQueryWizardPage;

procedure InitializeWizard;
begin
  // Halaman 1: Database Configuration
  ConfigPage := CreateInputQueryPage(wpSelectTasks,
    'Database Configuration', 'Konfigurasi koneksi PostgreSQL',
    'Masukkan detail koneksi database:');

  ConfigPage.Add('Database Host:', False);
  ConfigPage.Add('Database Port:', False);
  ConfigPage.Add('Database Name:', False);
  ConfigPage.Add('Database User:', False);
  ConfigPage.Add('Database Password:', True);

  ConfigPage.Values[0] := 'localhost';
  ConfigPage.Values[1] := '5432';
  ConfigPage.Values[2] := 'alatujidb';
  ConfigPage.Values[3] := 'postgres';
  ConfigPage.Values[4] := '';

  // Halaman 2: Odoo Configuration
  OdooConfigPage := CreateInputQueryPage(ConfigPage.ID,
    'Odoo Configuration', 'Konfigurasi koneksi ke Odoo RMC',
    'Masukkan detail akses Odoo (opsional — bisa diisi nanti via .env):');

  OdooConfigPage.Add('Session Token:', False);
  OdooConfigPage.Add('Odoo URL:', False);

  OdooConfigPage.Values[0] := '';
  OdooConfigPage.Values[1] := 'https://rmc.adhimix.web.id';
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  EnvFile: string;
  EnvLines: TArrayOfString;
  I: Integer;
begin
  if CurStep = ssPostInstall then
  begin
    EnvFile := ExpandConstant('{userappdata}\{#MyAppName}\.env');

    if LoadStringsFromFile(ExpandConstant('{app}\.env.example'), EnvLines) then
    begin
      for I := 0 to GetArrayLength(EnvLines) - 1 do
      begin
        // Database config
        if Pos('DB_HOST=', EnvLines[I]) = 1 then
          EnvLines[I] := 'DB_HOST=' + ConfigPage.Values[0];
        if Pos('DB_PORT=', EnvLines[I]) = 1 then
          EnvLines[I] := 'DB_PORT=' + ConfigPage.Values[1];
        if Pos('DB_NAME=', EnvLines[I]) = 1 then
          EnvLines[I] := 'DB_NAME=' + ConfigPage.Values[2];
        if Pos('DB_USER=', EnvLines[I]) = 1 then
          EnvLines[I] := 'DB_USER=' + ConfigPage.Values[3];
        if Pos('DB_PASSWORD=', EnvLines[I]) = 1 then
          EnvLines[I] := 'DB_PASSWORD=' + ConfigPage.Values[4];
        // Odoo config
        if Pos('SESSION_TOKEN=', EnvLines[I]) = 1 then
          EnvLines[I] := 'SESSION_TOKEN=' + OdooConfigPage.Values[0];
        if Pos('ENVIRONMENT=', EnvLines[I]) = 1 then
          EnvLines[I] := 'ENVIRONMENT=production';
      end;

      SaveStringsToFile(EnvFile, EnvLines, False);
    end;

    // Buat marker agar first_run_setup skip playwright download jika sudah bundle
    if IsComponentSelected('chromium') then
      SaveStringToFile(ExpandConstant('{userdocs}\.odoo_app_playwright_setup'), '', False);
  end;
end;

function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  Result := True;

  // Cek apakah PostgreSQL sudah ter-install (skip warning jika user pilih bundle pgsql)
  // Warning hanya muncul jika memilih compact install tanpa PostgreSQL
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  // Cek disk space: app ~150MB + pgsql ~180MB + chromium ~170MB = ~500MB
  // Inno Setup sudah handle disk space check otomatis
end;
```

### 4.2 Penjelasan Komponen Installer

| Komponen | Ukuran | Deskripsi |
|----------|--------|-----------|
| `app` (fixed) | ~150MB | Executable + Python runtime + semua pip packages |
| `pgsql` (opsional) | ~180MB | PostgreSQL 16 portable binaries |
| `chromium` (opsional) | ~170MB | Chromium browser untuk Playwright |
| **Total (full)** | **~500MB** | Semua komponen |
| **Total (compact)** | **~150MB** | Hanya app (user install PostgreSQL sendiri) |

---

## Tahap 5: Build Script Otomatis

### 5.1 `build_installer.bat` (Updated)

```bat
@echo off
setlocal enabledelayedexpansion

echo ========================================
echo Odoo Compression Test - Full Build
echo ========================================
echo.

REM ── 0. Checks ──────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found in PATH
    pause & exit /b 1
)

set INNO_PATH=C:\Program Files (x86)\Inno Setup 6\ISCC.exe
if not exist "%INNO_PATH%" (
    echo [ERROR] Inno Setup 6 not found
    echo Download: https://jrsoftware.org/isdl.php
    pause & exit /b 1
)

REM ── 1. Clean ────────────────────────────
echo [1/6] Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist output rmdir /s /q output
echo [OK] Cleaned
echo.

REM ── 2. Verify vendor files ─────────────
echo [2/6] Checking vendor components...

if exist vendor\pgsql\bin\postgres.exe (
    echo   [OK] PostgreSQL portable found
) else (
    echo   [WARN] PostgreSQL portable NOT found in vendor\pgsql\
    echo          Download dari: https://get.enterprisedb.com/postgresql/
    echo          Extract ke vendor\pgsql\
)

if exist vendor\playwright-browsers (
    echo   [OK] Playwright browsers found
) else (
    echo   [WARN] Playwright browsers NOT found in vendor\playwright-browsers\
    echo          Jalankan: xcopy "%LOCALAPPDATA%\ms-playwright\*" vendor\playwright-browsers\ /E /I
)
echo.

REM ── 3. Install PyInstaller ─────────────
echo [3/6] Ensuring PyInstaller...
pip install pyinstaller >nul 2>&1
echo [OK] PyInstaller ready
echo.

REM ── 4. Build executable ────────────────
echo [4/6] Building executable (3-5 minutes)...
pyinstaller odoo_app.spec
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed
    pause & exit /b 1
)
echo [OK] Built: dist\OdooCompressionTest\
echo.

REM ── 5. Quick smoke test ────────────────
echo [5/6] Smoke test...
if exist dist\OdooCompressionTest\OdooCompressionTest.exe (
    echo [OK] Executable exists
) else (
    echo [ERROR] Executable not found!
    pause & exit /b 1
)
echo.

REM ── 6. Build installer ────────────────
echo [6/6] Building installer...
"%INNO_PATH%" installer.iss
if errorlevel 1 (
    echo [ERROR] Inno Setup build failed
    pause & exit /b 1
)

REM ── Done ───────────────────────────────
echo.
echo ========================================
echo BUILD COMPLETE
echo ========================================
echo.
echo Executable: dist\OdooCompressionTest\OdooCompressionTest.exe
echo Installer:  output\OdooCompressionTestSetup.exe
echo.

for %%F in ("output\OdooCompressionTestSetup.exe") do (
    set size=%%~zF
    set /a size_mb=!size! / 1048576
    echo Installer size: !size_mb! MB
)

echo.
echo CHECKLIST SEBELUM DISTRIBUSI:
echo   [ ] Test installer di PC bersih (tanpa Python/PostgreSQL)
echo   [ ] Verifikasi app jalan tanpa error
echo   [ ] Verifikasi koneksi DB berhasil
echo   [ ] Verifikasi web automation (Playwright) berfungsi
echo   [ ] Verifikasi Excel processing berjalan
echo   [ ] Test uninstall bersih
echo.
pause
```

---

## Tahap 6: Post-Install Flow (Apa yang Terjadi di PC User)

### 6.1 Installer Wizard Flow

```
[1] Welcome Screen
      │
[2] License Agreement (opsional)
      │
[3] Select Destination → default: C:\Program Files\Odoo Compression Test\
      │
[4] Select Components
      │  ☑ Application (required)
      │  ☑ PostgreSQL 16 (recommended)
      │  ☑ Chromium Browser (recommended)
      │
[5] Select Tasks
      │  ☑ Create Desktop Shortcut
      │  ☑ Initialize Database
      │  ☑ Register PostgreSQL as Service
      │
[6] Database Configuration
      │  Host: localhost
      │  Port: 5432
      │  Database: alatujidb
      │  User: postgres
      │  Password: ________
      │
[7] Odoo Configuration
      │  Session Token: (opsional)
      │  Odoo URL: https://rmc.adhimix.web.id
      │
[8] Install Progress
      │  → Copy files
      │  → initdb (create PostgreSQL data cluster)
      │  → Start PostgreSQL
      │  → CREATE DATABASE alatujidb
      │  → Import database_schema.sql
      │  → Register pg service (opsional)
      │  → Generate .env di AppData
      │  → Create playwright marker (jika bundle chromium)
      │
[9] Finish → Launch Application
```

### 6.2 First Run (di `main.py`)

```
App Launch
  │
  ├─ sys.frozen == True → run_first_time_setup()
  │   ├─ ensure_env_file()     → .env di %APPDATA%\OdooCompressionTest\
  │   ├─ get_data_directories() → Output/, Rencana/, Pengujian/, Data Input Pengujian/
  │   └─ setup_playwright()    → SKIP jika marker exists (installer bundle chromium)
  │                            → DOWNLOAD jika marker absent (compact install)
  │
  ├─ initialize_db_pool()      → Connect ke PostgreSQL via .env config
  │
  └─ ExcelProcessorGUI(root)   → GUI ready
```

---

## Tahap 7: Testing Checklist

### 7.1 Test di Mesin Build

```
[ ] pyinstaller build sukses tanpa error
[ ] dist\OdooCompressionTest\OdooCompressionTest.exe jalan
[ ] App bisa connect ke database
[ ] Excel processing (copy data, generate CSV) berfungsi
[ ] Web automation (Playwright) berfungsi
[ ] Inno Setup compile sukses
[ ] Installer .exe tergenerate di output\
```

### 7.2 Test di PC Bersih (VM/mesin tanpa Python)

```
[ ] Installer jalan tanpa error
[ ] PostgreSQL ter-install dan service running (jika pilih full)
[ ] Database alatujidb ter-create dengan schema benar
[ ] .env tergenerate dengan config dari wizard
[ ] App launch dari desktop shortcut
[ ] App connect ke database
[ ] Copy Data dari Excel berfungsi
[ ] Generate CSV berfungsi (separator ; dan decimal ,)
[ ] Upload to Database berfungsi
[ ] Web Automation (Input Rencana Benda Uji) berfungsi
[ ] Daemon Sync ke Odoo berfungsi
[ ] Uninstall bersih (service dihapus, files dihapus)
```

### 7.3 Edge Cases

```
[ ] Install tanpa PostgreSQL component → app tetap jalan, error message jelas
[ ] Install tanpa Chromium component → first_run_setup download otomatis
[ ] Upgrade install (install ulang di atas versi lama) → data .env tidak tertimpa
[ ] Install di path dengan spasi (C:\Program Files\...)
[ ] Install di path dengan karakter Indonesia/Unicode
[ ] User cancel saat install → rollback bersih
[ ] PostgreSQL port 5432 sudah dipakai → error handling jelas
```

---

## Tahap 8: Distribusi

### 8.1 Output Files

```
output/
└── OdooCompressionTestSetup.exe   (~300-500MB, tergantung components)
```

### 8.2 Distribusi ke User

1. Copy `OdooCompressionTestSetup.exe` ke USB/shared drive
2. User jalankan sebagai Administrator
3. Ikuti wizard → selesai
4. App siap digunakan dari Desktop shortcut

### 8.3 Update Workflow

Untuk update versi baru:
1. Bump version di `installer.iss` (`#define MyAppVersion`)
2. Jalankan `build_installer.bat`
3. Distribusikan installer baru
4. User install ulang → file `.env` dan data Excel di AppData **tidak tertimpa** (`onlyifdoesntexist`)

---

## Referensi

| Resource | URL |
|----------|-----|
| PyInstaller docs | https://pyinstaller.org/en/stable/ |
| Inno Setup docs | https://jrsoftware.org/ishelp/ |
| PostgreSQL binaries | https://www.enterprisedb.com/download-postgresql-binaries |
| Playwright install | https://playwright.dev/python/docs/intro |
