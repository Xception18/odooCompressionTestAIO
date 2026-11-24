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
OutputDir=installer_build\output
OutputBaseFilename=OdooCompressionTestSetup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Types]
Name: "full"; Description: "Full Installation (Recommended)"
Name: "compact"; Description: "Compact (tanpa PostgreSQL & Browser)"
Name: "custom"; Description: "Custom Installation"; Flags: iscustom

[Components]
Name: "app"; Description: "Odoo Compression Test Application"; Types: full compact custom; Flags: fixed
Name: "pgsql"; Description: "PostgreSQL 17 Database Server (~140MB)"; Types: full
Name: "chromium"; Description: "Chromium Browser untuk Web Automation (~690MB)"; Types: full

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"
Name: "initdb"; Description: "Initialize database (create alatujidb)"; Components: pgsql
Name: "pgservice"; Description: "Register PostgreSQL as Windows Service"; Components: pgsql

[Files]
Source: "dist\OdooCompressionTest\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: app
Source: "installer_build\vendor\pgsql\*"; DestDir: "{app}\pgsql"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: pgsql
Source: "installer_build\vendor\playwright-browsers\*"; DestDir: "{localappdata}\ms-playwright"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: chromium
Source: ".env.example"; DestDir: "{app}"; Flags: ignoreversion; Components: app
Source: "database_schema.sql"; DestDir: "{app}"; Flags: ignoreversion; Components: app
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion isreadme; Components: app
Source: "INSTALL_PLAN.md"; DestDir: "{app}"; Flags: ignoreversion; Components: app
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
Filename: "{app}\pgsql\bin\initdb.exe"; Parameters: "-D ""{app}\pgsql\data"" -U postgres -E UTF8 --locale=C"; StatusMsg: "Initializing PostgreSQL..."; Tasks: initdb; Flags: runhidden waituntilterminated
Filename: "{app}\pgsql\bin\pg_ctl.exe"; Parameters: "start -D ""{app}\pgsql\data"" -l ""{app}\pgsql\postgresql.log"" -w"; StatusMsg: "Starting PostgreSQL..."; Tasks: initdb; Flags: runhidden waituntilterminated
Filename: "{app}\pgsql\bin\psql.exe"; Parameters: "-U postgres -c ""CREATE DATABASE alatujidb;"""; StatusMsg: "Creating database..."; Tasks: initdb; Flags: runhidden waituntilterminated
Filename: "{app}\pgsql\bin\psql.exe"; Parameters: "-U postgres -d alatujidb -f ""{app}\database_schema.sql"""; StatusMsg: "Importing schema..."; Tasks: initdb; Flags: runhidden waituntilterminated
Filename: "{app}\pgsql\bin\pg_ctl.exe"; Parameters: "register -N ""OdooCompressionTestPG"" -D ""{app}\pgsql\data"""; StatusMsg: "Registering PostgreSQL service..."; Tasks: pgservice; Flags: runhidden waituntilterminated
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: postinstall nowait skipifsilent runascurrentuser

[UninstallRun]
Filename: "{app}\pgsql\bin\pg_ctl.exe"; Parameters: "stop -D ""{app}\pgsql\data"""; Flags: runhidden; RunOnceId: "StopPG"
Filename: "{app}\pgsql\bin\pg_ctl.exe"; Parameters: "unregister -N ""OdooCompressionTestPG"""; Flags: runhidden; RunOnceId: "UnregPG"

[Code]
var
  ConfigPage: TInputQueryWizardPage;
  OdooConfigPage: TInputQueryWizardPage;

procedure InitializeWizard;
begin
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
  ConfigPage.Values[4] := 'adhimix';

  OdooConfigPage := CreateInputQueryPage(ConfigPage.ID,
    'Odoo Configuration', 'Konfigurasi koneksi ke Odoo RMC',
    'Masukkan detail akses Odoo (opsional, bisa diisi nanti via .env):');
  OdooConfigPage.Add('Session Token:', False);
  OdooConfigPage.Add('Odoo URL:', False);
  OdooConfigPage.Values[0] := 'f3953bfa431242bc8c90f4b69917a7a1d3d24856';
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
        if Pos('SESSION_TOKEN=', EnvLines[I]) = 1 then
          EnvLines[I] := 'SESSION_TOKEN=' + OdooConfigPage.Values[0];
        if Pos('ENVIRONMENT=', EnvLines[I]) = 1 then
          EnvLines[I] := 'ENVIRONMENT=production';
      end;
      SaveStringsToFile(EnvFile, EnvLines, False);
    end;
    if IsComponentSelected('chromium') then
      SaveStringToFile(ExpandConstant('{userdocs}\.odoo_app_playwright_setup'), '', False);
  end;
end;

function InitializeSetup(): Boolean;
begin
  Result := True;
end;
