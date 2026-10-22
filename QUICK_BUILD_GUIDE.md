# Quick Start: Building Windows Installer

## Prerequisites
1. Python 3.8+ installed
2. Download Inno Setup: https://jrsoftware.org/isdl.php

## Build Steps

```bash
# 1. Run automated build
build_installer.bat

# Output: output\OdooCompressionTestSetup.exe (~35-50MB)
```

That's it! The script does:
- ✓ Installs PyInstaller
- ✓ Builds executable (3-5 min)
- ✓ Creates installer with Inno Setup
- ✓ Shows file size

## Test Installer

```bash
# Install on clean Windows machine (no Python required)
output\OdooCompressionTestSetup.exe

# Installer will:
# 1. Prompt for PostgreSQL credentials
# 2. Create .env in AppData\Roaming\OdooCompressionTest
# 3. Create desktop shortcut
# 4. Launch app after install
```

## Files Created

```
Project files (for building):
├── odoo_app.spec              ← PyInstaller config
├── installer.iss              ← Inno Setup script
├── build_installer.bat        ← Build automation
├── modules/first_run_setup.py ← First-run logic
└── Excel templates:
    ├── Data Input Pengujian/data.xlsx
    ├── Pengujian/PENGUJIAN.xlsx
    └── Rencana/RENCANA.xlsx

Build output:
├── dist\OdooCompressionTest\  ← Standalone app folder
└── output\OdooCompressionTestSetup.exe ← Final installer
```

## Manual Build (if batch fails)

```bash
# Step 1: Build executable
pip install pyinstaller
pyinstaller odoo_app.spec

# Step 2: Build installer
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
```

## Troubleshooting

**Error: Module not found during build**
→ Add to `hidden_imports` in `odoo_app.spec`

**Error: Inno Setup not found**
→ Install from https://jrsoftware.org/isdl.php

**Executable crashes on launch**
→ Change `console=False` to `console=True` in spec to see errors

**Large file size (>100MB)**
→ Normal. Without Playwright browser bundle: ~35-50MB
→ With browser bundled: ~200MB (not recommended)

## Distribution

**End users need:**
- Windows 10/11
- PostgreSQL installed
- ~50MB disk space

**End users DON'T need:**
- Python
- pip packages
- Git

## Post-Install User Experience

1. Run `OdooCompressionTestSetup.exe`
2. Installer asks for DB credentials
3. App installs to `C:\Program Files\OdooCompressionTest\`
4. Config saved to `%APPDATA%\OdooCompressionTest\.env`
5. Desktop shortcut created
6. Excel templates copied to user data folder
7. First launch downloads Playwright browser (~170MB, one-time)
8. App ready to use

## Build Checklist

- [x] `odoo_app.spec` created
- [x] `installer.iss` created
- [x] `build_installer.bat` created
- [x] `modules/first_run_setup.py` created
- [x] `main.py` updated for frozen exe
- [ ] Optional: Add `icon.ico` (256x256)
- [ ] Optional: Code signing certificate
- [ ] Test on clean machine

## Next Steps

See `INSTALLER_BUILD_PLAN.md` for:
- Advanced customization
- Multi-language support
- Auto-update mechanism
- Silent install for enterprise
- Custom themes
- Troubleshooting details
