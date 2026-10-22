; Inno Setup Script for Odoo Compression Test

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

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "indonesian"; MessagesFile: "compiler:Languages\Indonesian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "dist\OdooCompressionTest\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: ".env.example"; DestDir: "{app}"; Flags: ignoreversion
Source: "database_schema.sql"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion isreadme
Source: "INSTALL_PLAN.md"; DestDir: "{app}"; Flags: ignoreversion
; Excel templates
Source: "Data Input Pengujian\data.xlsx"; DestDir: "{userappdata}\{#MyAppName}\Data Input Pengujian"; Flags: ignoreversion onlyifdoesntexist
Source: "Pengujian\PENGUJIAN.xlsx"; DestDir: "{userappdata}\{#MyAppName}\Pengujian"; Flags: ignoreversion onlyifdoesntexist
Source: "Rencana\RENCANA.xlsx"; DestDir: "{userappdata}\{#MyAppName}\Rencana"; Flags: ignoreversion onlyifdoesntexist

[Dirs]
Name: "{userappdata}\{#MyAppName}"; Permissions: users-modify
Name: "{userappdata}\{#MyAppName}\Output"; Permissions: users-modify
Name: "{userappdata}\{#MyAppName}\Rencana"; Permissions: users-modify
Name: "{userappdata}\{#MyAppName}\Pengujian"; Permissions: users-modify
Name: "{userappdata}\{#MyAppName}\Data Input Pengujian"; Permissions: users-modify

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Edit Configuration"; Filename: "notepad.exe"; Parameters: """{userappdata}\{#MyAppName}\.env"""
Name: "{group}\Open Data Folder"; Filename: "{userappdata}\{#MyAppName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: postinstall nowait skipifsilent runascurrentuser

[Code]
var
  ConfigPage: TInputQueryWizardPage;

procedure InitializeWizard;
begin
  ConfigPage := CreateInputQueryPage(wpSelectTasks,
    'Database Configuration', 'Enter your PostgreSQL database details',
    'The application requires PostgreSQL. Enter connection details:');
  
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
        if Pos('ENVIRONMENT=', EnvLines[I]) = 1 then
          EnvLines[I] := 'ENVIRONMENT=production';
      end;
      
      SaveStringsToFile(EnvFile, EnvLines, False);
    end;
  end;
end;

function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  Result := True;
  
  if not FileExists('C:\Program Files\PostgreSQL\bin\psql.exe') and
     not FileExists('C:\Program Files\PostgreSQL\14\bin\psql.exe') and
     not FileExists('C:\Program Files\PostgreSQL\15\bin\psql.exe') and
     not FileExists('C:\Program Files\PostgreSQL\16\bin\psql.exe') then
  begin
    if MsgBox('PostgreSQL is not detected on this system.' + #13#10 +
              'The application requires PostgreSQL to function.' + #13#10#13#10 +
              'Would you like to download PostgreSQL now?', 
              mbConfirmation, MB_YESNO) = IDYES then
    begin
      ShellExec('open', 'https://www.postgresql.org/download/windows/', '', '', SW_SHOW, ewNoWait, ResultCode);
      Result := False;
    end;
  end;
end;
