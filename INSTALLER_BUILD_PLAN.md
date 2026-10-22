# Executable Installer Build Plan

## Goal
Package Python app into standalone Windows installer (.exe) that:
- Installs app without requiring Python
- Bundles all dependencies
- Runs setup wizard (DB config, .env setup)
- Creates desktop shortcut
- ~100-150MB final installer

## Architecture

```
Build Process:
1. PyInstaller → OdooCompressionTest.exe (Python app + libs)
2. Inno Setup → OdooCompressionTestSetup.exe (installer wizard)

Final Structure:
C:\Program Files\OdooCompressionTest\
├── OdooCompressionTest.exe    ← Main executable
├── _internal\                  ← PyInstaller dependencies
├── database_schema.sql
├── .env.example
├── chromedriver.exe
├── config\                     ← User config directory
└── unins000.exe               ← Uninstaller
```

## Phase 1: App Executable (PyInstaller)

### 1.1 Install PyInstaller
```bash
pip install pyinstaller
```

### 1.2 Create PyInstaller Spec File

**File: `odoo_app.spec`**
```python
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Collect data files
added_files = [
    ('database_schema.sql', '.'),
    ('.env.example', '.'),
    ('chromedriver.exe', '.'),
    ('modules', 'modules'),
    ('Data Input Pengujian/data.xlsx', 'Data Input Pengujian'),
    ('Pengujian/PENGUJIAN.xlsx', 'Pengujian'),
    ('Rencana/RENCANA.xlsx', 'Rencana'),
]

# Hidden imports for dynamic imports
hidden_imports = [
    'customtkinter',
    'PIL._tkinter_finder',
    'playwright.sync_api',
    'selenium.webdriver.chrome.service',
    'psycopg2',
    'openpyxl',
    'pandas',
    'sqlalchemy.dialects.postgresql',
    'dotenv',
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
    excludes=['pytest', 'pytest-cov', 'pytest-mock', 'pytest-asyncio'],
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
    name='OdooCompressionTest',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',  # Add if you have an icon
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

### 1.3 Build Executable
```bash
# Clean previous builds
rmdir /s /q build dist

# Build
pyinstaller odoo_app.spec

# Test executable
dist\OdooCompressionTest\OdooCompressionTest.exe
```

**Output:** `dist\OdooCompressionTest\` folder (~80-120MB)

### 1.4 Handle Playwright Browsers

**Problem:** Playwright browsers not bundled by PyInstaller (170MB Chromium)

**Solution 1: Bundle browser (increases size to ~300MB)**
```python
# In odoo_app.spec, add:
import os
import pathlib

playwright_browsers = str(pathlib.Path.home() / 'AppData/Local/ms-playwright/chromium-*')

added_files = [
    # ... existing files
    (playwright_browsers, 'ms-playwright/chromium'),
]
```

**Solution 2: Download on first run (recommended)**
Create `first_run_setup.py`:
```python
import subprocess
import sys
from pathlib import Path

def setup_playwright():
    """Download Playwright browser on first run"""
    marker = Path.home() / '.odoo_app_setup_done'
    if marker.exists():
        return
    
    print("First-time setup: Installing browser components...")
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"])
    marker.touch()
    print("Setup complete!")
```

Call from `main.py`:
```python
if __name__ == "__main__":
    from first_run_setup import setup_playwright
    setup_playwright()
    main()
```

## Phase 2: Installer Wizard (Inno Setup)

### 2.1 Install Inno Setup
Download: https://jrsoftware.org/isdl.php (free, 2MB)

### 2.2 Create Inno Setup Script

**File: `installer.iss`**
```ini
; Inno Setup Script for Odoo Compression Test

#define MyAppName "Odoo Compression Test"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Your Company"
#define MyAppExeName "OdooCompressionTest.exe"
#define MyAppURL "https://your-website.com"

[Setup]
AppId={{YOUR-GUID-HERE}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=output
OutputBaseFilename=OdooCompressionTestSetup
SetupIconFile=icon.ico
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "indonesian"; MessagesFile: "compiler:Languages\Indonesian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

[Files]
; Main application files
Source: "dist\OdooCompressionTest\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Config template
Source: ".env.example"; DestDir: "{app}"; Flags: ignoreversion
; Database schema
Source: "database_schema.sql"; DestDir: "{app}"; Flags: ignoreversion
; Excel templates
Source: "Data Input Pengujian\data.xlsx"; DestDir: "{userappdata}\{#MyAppName}\Data Input Pengujian"; Flags: ignoreversion onlyifdoesntexist
Source: "Pengujian\PENGUJIAN.xlsx"; DestDir: "{userappdata}\{#MyAppName}\Pengujian"; Flags: ignoreversion onlyifdoesntexist
Source: "Rencana\RENCANA.xlsx"; DestDir: "{userappdata}\{#MyAppName}\Rencana"; Flags: ignoreversion onlyifdoesntexist
; Documentation
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion isreadme
Source: "INSTALL_PLAN.md"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
; Create user data directories
Name: "{userappdata}\{#MyAppName}\Output"; Permissions: users-modify
Name: "{userappdata}\{#MyAppName}\Rencana"; Permissions: users-modify
Name: "{userappdata}\{#MyAppName}\Pengujian"; Permissions: users-modify
Name: "{userappdata}\{#MyAppName}\Data Input Pengujian"; Permissions: users-modify

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Configuration"; Filename: "notepad.exe"; Parameters: """{userappdata}\{#MyAppName}\.env"""
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Run first-time configuration wizard
Filename: "{app}\{#MyAppExeName}"; Description: "Run configuration wizard"; Flags: postinstall nowait skipifsilent runascurrentuser

[Code]
var
  ConfigPage: TInputQueryWizardPage;
  DBHostEdit, DBPortEdit, DBNameEdit, DBUserEdit, DBPasswordEdit: TEdit;
  SessionTokenEdit: TEdit;

procedure InitializeWizard;
begin
  { Configuration page }
  ConfigPage := CreateInputQueryPage(wpSelectTasks,
    'Database Configuration', 'Enter your PostgreSQL database details',
    'The application needs a PostgreSQL database. Enter your connection details below.');
  
  ConfigPage.Add('Database Host:', False);
  ConfigPage.Add('Database Port:', False);
  ConfigPage.Add('Database Name:', False);
  ConfigPage.Add('Database User:', False);
  ConfigPage.Add('Database Password:', True);
  ConfigPage.Add('Session Token (optional):', False);
  
  { Default values }
  ConfigPage.Values[0] := 'localhost';
  ConfigPage.Values[1] := '5432';
  ConfigPage.Values[2] := 'alatujidb';
  ConfigPage.Values[3] := 'postgres';
  ConfigPage.Values[4] := 'adhimix';
  ConfigPage.Values[5] := '';
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  EnvFile: string;
  EnvContent: TArrayOfString;
  I: Integer;
begin
  if CurStep = ssPostInstall then
  begin
    { Create .env file from user input }
    EnvFile := ExpandConstant('{userappdata}\{#MyAppName}\.env');
    
    if LoadStringsFromFile(ExpandConstant('{app}\.env.example'), EnvContent) then
    begin
      { Replace placeholders with user values }
      for I := 0 to GetArrayLength(EnvContent) - 1 do
      begin
        if Pos('DB_HOST=', EnvContent[I]) = 1 then
          EnvContent[I] := 'DB_HOST=' + ConfigPage.Values[0];
        if Pos('DB_PORT=', EnvContent[I]) = 1 then
          EnvContent[I] := 'DB_PORT=' + ConfigPage.Values[1];
        if Pos('DB_NAME=', EnvContent[I]) = 1 then
          EnvContent[I] := 'DB_NAME=' + ConfigPage.Values[2];
        if Pos('DB_USER=', EnvContent[I]) = 1 then
          EnvContent[I] := 'DB_USER=' + ConfigPage.Values[3];
        if Pos('DB_PASSWORD=', EnvContent[I]) = 1 then
          EnvContent[I] := 'DB_PASSWORD=' + ConfigPage.Values[4];
        if (Pos('SESSION_TOKEN=', EnvContent[I]) = 1) and (ConfigPage.Values[5] <> '') then
          EnvContent[I] := 'SESSION_TOKEN=' + ConfigPage.Values[5];
      end;
      
      SaveStringsToFile(EnvFile, EnvContent, False);
    end;
  end;
end;

function InitializeSetup(): Boolean;
begin
  { Check PostgreSQL installed }
  if not FileExists('C:\Program Files\PostgreSQL\bin\psql.exe') and
     not FileExists('C:\Program Files\PostgreSQL\14\bin\psql.exe') and
     not FileExists('C:\Program Files\PostgreSQL\15\bin\psql.exe') and
     not FileExists('C:\Program Files\PostgreSQL\16\bin\psql.exe') then
  begin
    if MsgBox('PostgreSQL not detected. Install PostgreSQL before continuing?' + #13#10 +
              'The application requires PostgreSQL to run.' + #13#10#13#10 +
              'Yes: Open PostgreSQL download page' + #13#10 +
              'No: Continue anyway (not recommended)', 
              mbConfirmation, MB_YESNO) = IDYES then
    begin
      ShellExec('open', 'https://www.postgresql.org/download/windows/', '', '', SW_SHOW, ewNoWait, Result);
      Result := False;
      Exit;
    end;
  end;
  
  Result := True;
end;
```

### 2.3 Build Installer
```bash
# Compile with Inno Setup
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss

# Output: output\OdooCompressionTestSetup.exe
```

## Phase 3: Configuration Handling

### 3.1 Modify app to use AppData config

**Update `modules/config.py`:**
```python
import os
from pathlib import Path

def get_config_dir():
    """Get user config directory"""
    if getattr(sys, 'frozen', False):
        # Running as executable
        appdata = Path(os.getenv('APPDATA'))
        config_dir = appdata / 'OdooCompressionTest'
        config_dir.mkdir(exist_ok=True)
        return config_dir
    else:
        # Running as script
        return Path(__file__).parent.parent

def get_env_path():
    """Get .env file path"""
    config_dir = get_config_dir()
    env_path = config_dir / '.env'
    
    # Copy .env.example on first run
    if not env_path.exists():
        example = Path(__file__).parent.parent / '.env.example'
        if example.exists():
            import shutil
            shutil.copy(example, env_path)
    
    return str(env_path)

# Use in load_dotenv:
from dotenv import load_dotenv
load_dotenv(get_env_path())
```

### 3.2 First-Run Setup Dialog

**Create `modules/first_run_dialog.py`:**
```python
import customtkinter as ctk
from pathlib import Path

class FirstRunDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("First-Time Setup")
        self.geometry("500x400")
        
        # Make modal
        self.transient(parent)
        self.grab_set()
        
        self.create_widgets()
        
    def create_widgets(self):
        label = ctk.CTkLabel(self, text="Welcome! Configure your database:")
        label.pack(pady=20)
        
        # DB fields
        self.db_host = self.add_field("DB Host:", "localhost")
        self.db_port = self.add_field("DB Port:", "5432")
        self.db_name = self.add_field("DB Name:", "alatujidb")
        self.db_user = self.add_field("DB User:", "postgres")
        self.db_pass = self.add_field("DB Password:", "", show="*")
        
        # Buttons
        btn = ctk.CTkButton(self, text="Save & Continue", command=self.save)
        btn.pack(pady=20)
        
    def add_field(self, label_text, default="", show=""):
        frame = ctk.CTkFrame(self)
        frame.pack(pady=5, padx=20, fill="x")
        
        label = ctk.CTkLabel(frame, text=label_text, width=120)
        label.pack(side="left")
        
        entry = ctk.CTkEntry(frame, show=show)
        entry.insert(0, default)
        entry.pack(side="left", fill="x", expand=True, padx=10)
        
        return entry
    
    def save(self):
        # Save to .env
        from modules.config import get_env_path
        env_path = Path(get_env_path())
        
        config_lines = [
            f"DB_HOST={self.db_host.get()}",
            f"DB_PORT={self.db_port.get()}",
            f"DB_NAME={self.db_name.get()}",
            f"DB_USER={self.db_user.get()}",
            f"DB_PASSWORD={self.db_pass.get()}",
            "LOG_LEVEL=INFO",
            "ENVIRONMENT=production",
        ]
        
        env_path.write_text('\n'.join(config_lines))
        self.destroy()
```

**Call from `main.py`:**
```python
def main():
    from modules.config import get_env_path
    env_path = Path(get_env_path())
    
    # Check if config exists
    if not env_path.exists() or env_path.stat().st_size < 50:
        root = tk.Tk()
        root.withdraw()  # Hide main window
        
        from modules.first_run_dialog import FirstRunDialog
        dialog = FirstRunDialog(root)
        root.wait_window(dialog)
    
    # Continue normal startup
    ...
```

## Phase 4: Build Automation

### 4.1 Build Script

**File: `build_installer.bat`**
```batch
@echo off
echo ========================================
echo Building Odoo Compression Test Installer
echo ========================================

REM Clean
echo [1/4] Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist output rmdir /s /q output

REM Install build tools
echo [2/4] Checking build tools...
pip install pyinstaller >nul 2>&1

REM Build executable
echo [3/4] Building executable...
pyinstaller odoo_app.spec
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed
    pause
    exit /b 1
)
echo [OK] Executable built: dist\OdooCompressionTest\

REM Build installer
echo [4/4] Building installer...
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
if errorlevel 1 (
    echo [ERROR] Inno Setup build failed
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build Complete!
echo ========================================
echo Installer: output\OdooCompressionTestSetup.exe
echo Size: ~100-150MB
echo.
pause
```

### 4.2 Test installer
```bash
# Build
build_installer.bat

# Install to test VM or clean machine
output\OdooCompressionTestSetup.exe
```

## Deployment Checklist

- [ ] Icon created: `icon.ico` (256x256)
- [ ] Version number set in `installer.iss`
- [ ] Code signing certificate (optional, $200/year)
- [ ] Test on clean Windows machine
- [ ] Test without Python installed
- [ ] Test without PostgreSQL (should warn)
- [ ] Verify .env creation in AppData
- [ ] Verify desktop shortcut works
- [ ] Verify uninstaller removes everything
- [ ] Create installer changelog
- [ ] Upload to distribution server

## Troubleshooting

### PyInstaller: Module not found
Add to `hiddenimports` in spec file

### Executable crashes on startup
Run with console: `console=True` in spec to see errors

### .env not found
Check `get_config_dir()` returns correct AppData path

### Database connection fails
Verify installer creates .env with correct permissions

### Large file size
- Exclude test dependencies
- Use UPX compression
- Don't bundle Playwright browser (download on first run)

## Future Enhancements

1. **Auto-update mechanism**: Check GitHub releases on startup
2. **Portable version**: No-install ZIP distribution
3. **Silent install**: `/SILENT` flag for enterprise deployment
4. **Custom themes**: Bundle theme files in installer
5. **Multi-language**: Indonesian + English installer
6. **Database auto-setup**: Run `database_schema.sql` during install
7. **Bundled PostgreSQL**: Include portable PostgreSQL (~50MB extra)

## File Size Breakdown

```
Component                Size
─────────────────────────────
Python runtime           15MB
Dependencies (pandas)    25MB
Tkinter/CustomTkinter    10MB
Playwright libs          8MB
Database drivers         5MB
App code                 2MB
Excel templates          0.5MB
─────────────────────────────
Executable total         ~66MB

Installer compression    ~36MB
Final installer size     ~36-50MB
```

Without Playwright browser bundle. With browser: ~200MB.

## Cost & Tools

All tools are free:
- PyInstaller: MIT License (free)
- Inno Setup: Free for commercial use
- Total cost: $0

Optional:
- Code signing cert: ~$200-400/year (DigiCert, Sectigo)
- Icon design: $0-50 (or use free tools)

## Build Time

- First build: 10-15 minutes
- Subsequent builds: 3-5 minutes
- Testing: 10 minutes per machine
