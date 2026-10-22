@echo off
setlocal enabledelayedexpansion

echo ========================================
echo Odoo Compression Test - Installer Build
echo ========================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found
    pause
    exit /b 1
)

REM Clean previous builds
echo [1/5] Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist output rmdir /s /q output
echo [OK] Cleaned
echo.

REM Install PyInstaller
echo [2/5] Installing PyInstaller...
pip install pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Failed to install PyInstaller
    pause
    exit /b 1
)
echo [OK] PyInstaller ready
echo.

REM Build executable with PyInstaller
echo [3/5] Building executable (this may take 3-5 minutes)...
pyinstaller odoo_app.spec
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed
    echo Check errors above
    pause
    exit /b 1
)
echo [OK] Executable built: dist\OdooCompressionTest\OdooCompressionTest.exe
echo.

REM Check Inno Setup
echo [4/5] Checking Inno Setup...
set INNO_PATH=C:\Program Files (x86)\Inno Setup 6\ISCC.exe
if not exist "%INNO_PATH%" (
    echo [ERROR] Inno Setup not found at: %INNO_PATH%
    echo.
    echo Download from: https://jrsoftware.org/isdl.php
    echo Install, then re-run this script
    pause
    exit /b 1
)
echo [OK] Inno Setup found
echo.

REM Build installer with Inno Setup
echo [5/5] Building installer...
"%INNO_PATH%" installer.iss
if errorlevel 1 (
    echo [ERROR] Inno Setup build failed
    pause
    exit /b 1
)
echo [OK] Installer built
echo.

REM Show results
echo ========================================
echo Build Complete!
echo ========================================
echo.
echo Executable: dist\OdooCompressionTest\OdooCompressionTest.exe
echo Installer:  output\OdooCompressionTestSetup.exe
echo.

REM Get file size
for %%F in ("output\OdooCompressionTestSetup.exe") do (
    set size=%%~zF
    set /a size_mb=!size! / 1048576
    echo Installer size: !size_mb! MB
)

echo.
echo Next steps:
echo   1. Test installer on clean machine
echo   2. Verify app runs without Python installed
echo   3. Check .env creation in AppData
echo.
pause
