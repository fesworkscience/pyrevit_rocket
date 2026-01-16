#define MyAppName "CPSK Tools for Revit"
#define MyAppPublisher "CPSK"
#define MyAppURL "https://rocket-tools.ru"

[Setup]
AppId={{E8A5C7B2-4F3D-4E9A-B6C8-1A2B3C4D5E6F}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={localappdata}\CPSK
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename=CPSK_Tools_v{#MyAppVersion}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
DisableDirPage=yes
DisableReadyPage=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
FinishedLabel=Setup has finished installing [name] on your computer.%n%n%nIMPORTANT: Please restart Revit to activate CPSK Tools.%n%nIf Revit was running during installation, close and reopen it.

[Files]
Source: "..\build\CPSK.extension\*"; DestDir: "{app}\CPSK.extension"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\build\Install-PyRevit.ps1"; DestDir: "{app}\tools"; Flags: ignoreversion
Source: "..\build\Register-Extension.ps1"; DestDir: "{app}\tools"; Flags: ignoreversion
Source: "..\build\Uninstall-CPSK.ps1"; DestDir: "{app}\tools"; Flags: ignoreversion
Source: "..\build\requirements.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\build\version.yaml"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\build\config.py"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
Name: "{app}\logs"

[Registry]
Root: HKCU; Subkey: "Software\CPSK\Tools"; ValueType: string; ValueName: "Version"; ValueData: "{#MyAppVersion}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\CPSK\Tools"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"; Flags: uninsdeletekey

[Code]
var
  CheckPage: TWizardPage;
  ConfigPage: TWizardPage;
  PyRevitStatusLabel: TNewStaticText;
  PluginStatusLabel: TNewStaticText;
  RevitStatusLabel: TNewStaticText;
  ActionLabel: TNewStaticText;
  ConfigStatusLabel: TNewStaticText;
  ConfigProgressBar: TNewProgressBar;
  PyRevitInstalled: Boolean;
  PluginInstalled: Boolean;
  RevitRunning: Boolean;
  PluginVersion: String;
  LogFile: String;

procedure WriteLog(Msg: String);
var
  S: String;
begin
  S := GetDateTimeString('yyyy-mm-dd hh:nn:ss', '-', ':') + ' - ' + Msg;
  SaveStringToFile(LogFile, S + #13#10, True);
end;

function BoolToStr(B: Boolean): String;
begin
  if B then Result := 'True' else Result := 'False';
end;

function IsRevitRunning: Boolean;
var
  ResultCode: Integer;
begin
  Result := False;
  WriteLog('Checking if Revit is running...');
  if Exec('tasklist.exe', '/FI "IMAGENAME eq Revit.exe" /NH', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
  begin
    // Check via PowerShell for more reliable detection
    if Exec('powershell.exe', '-NoProfile -Command "if (Get-Process -Name Revit -ErrorAction SilentlyContinue) { exit 1 } else { exit 0 }"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    begin
      Result := (ResultCode = 1);
      WriteLog('Revit running check result: ' + BoolToStr(Result));
    end;
  end;
end;

function IsPyRevitInstalled: Boolean;
var
  Path: String;
begin
  Result := False;
  WriteLog('Checking pyRevit installation...');

  if RegQueryStringValue(HKEY_CURRENT_USER, 'Software\pyRevit', 'InstallPath', Path) then
  begin
    WriteLog('  Found in registry HKCU\Software\pyRevit: ' + Path);
    Result := True;
  end
  else
    WriteLog('  Not found in registry HKCU\Software\pyRevit');

  if DirExists(ExpandConstant('{userappdata}\pyRevit-Master')) then
  begin
    WriteLog('  Found directory: ' + ExpandConstant('{userappdata}\pyRevit-Master'));
    Result := True;
  end
  else
    WriteLog('  Directory not found: ' + ExpandConstant('{userappdata}\pyRevit-Master'));

  if DirExists(ExpandConstant('{commonappdata}\pyRevit')) then
  begin
    WriteLog('  Found directory: ' + ExpandConstant('{commonappdata}\pyRevit'));
    Result := True;
  end
  else
    WriteLog('  Directory not found: ' + ExpandConstant('{commonappdata}\pyRevit'));

  if FileExists(ExpandConstant('{pf}\pyRevit CLI\pyrevit.exe')) then
  begin
    WriteLog('  Found file: ' + ExpandConstant('{pf}\pyRevit CLI\pyrevit.exe'));
    Result := True;
  end
  else
    WriteLog('  File not found: ' + ExpandConstant('{pf}\pyRevit CLI\pyrevit.exe'));

  WriteLog('pyRevit installed: ' + BoolToStr(Result));
end;

function GetInstalledPluginVersion: String;
var
  Version: String;
begin
  Result := '';
  if RegQueryStringValue(HKEY_CURRENT_USER, 'Software\CPSK\Tools', 'Version', Version) then
  begin
    Result := Version;
    WriteLog('Found installed CPSK Tools version: ' + Version);
  end
  else
    WriteLog('CPSK Tools not found in registry');
end;

function IsPluginInstalled: Boolean;
begin
  Result := (GetInstalledPluginVersion <> '');
end;

procedure UpdateConfigStatus(Msg: String; Progress: Integer);
begin
  ConfigStatusLabel.Caption := Msg;
  ConfigProgressBar.Position := Progress;
  WizardForm.Refresh;
  WriteLog('Status: ' + Msg + ' (' + IntToStr(Progress) + '%)');
end;

procedure InitializeWizard;
var
  TitleLabel: TNewStaticText;
begin
  CheckPage := CreateCustomPage(wpWelcome, 'System Check', 'Checking your system before installation...');

  TitleLabel := TNewStaticText.Create(CheckPage);
  TitleLabel.Parent := CheckPage.Surface;
  TitleLabel.Left := 0;
  TitleLabel.Top := 10;
  TitleLabel.Caption := 'Current system status:';
  TitleLabel.Font.Style := [fsBold];

  PyRevitStatusLabel := TNewStaticText.Create(CheckPage);
  PyRevitStatusLabel.Parent := CheckPage.Surface;
  PyRevitStatusLabel.Left := 20;
  PyRevitStatusLabel.Top := 40;
  PyRevitStatusLabel.Width := CheckPage.SurfaceWidth - 40;
  PyRevitStatusLabel.Height := 40;
  PyRevitStatusLabel.AutoSize := False;
  PyRevitStatusLabel.WordWrap := True;

  PluginStatusLabel := TNewStaticText.Create(CheckPage);
  PluginStatusLabel.Parent := CheckPage.Surface;
  PluginStatusLabel.Left := 20;
  PluginStatusLabel.Top := 90;
  PluginStatusLabel.Width := CheckPage.SurfaceWidth - 40;
  PluginStatusLabel.Height := 40;
  PluginStatusLabel.AutoSize := False;
  PluginStatusLabel.WordWrap := True;

  RevitStatusLabel := TNewStaticText.Create(CheckPage);
  RevitStatusLabel.Parent := CheckPage.Surface;
  RevitStatusLabel.Left := 20;
  RevitStatusLabel.Top := 140;
  RevitStatusLabel.Width := CheckPage.SurfaceWidth - 40;
  RevitStatusLabel.Height := 40;
  RevitStatusLabel.AutoSize := False;
  RevitStatusLabel.WordWrap := True;

  ActionLabel := TNewStaticText.Create(CheckPage);
  ActionLabel.Parent := CheckPage.Surface;
  ActionLabel.Left := 0;
  ActionLabel.Top := 190;
  ActionLabel.Width := CheckPage.SurfaceWidth;
  ActionLabel.Height := 60;
  ActionLabel.AutoSize := False;
  ActionLabel.WordWrap := True;
  ActionLabel.Font.Style := [fsBold];

  ConfigPage := CreateCustomPage(wpInstalling, 'Configuring', 'Setting up CPSK Tools...');

  ConfigStatusLabel := TNewStaticText.Create(ConfigPage);
  ConfigStatusLabel.Parent := ConfigPage.Surface;
  ConfigStatusLabel.Left := 0;
  ConfigStatusLabel.Top := 20;
  ConfigStatusLabel.Width := ConfigPage.SurfaceWidth;
  ConfigStatusLabel.Height := 50;
  ConfigStatusLabel.AutoSize := False;
  ConfigStatusLabel.WordWrap := True;
  ConfigStatusLabel.Caption := 'Preparing...';

  ConfigProgressBar := TNewProgressBar.Create(ConfigPage);
  ConfigProgressBar.Parent := ConfigPage.Surface;
  ConfigProgressBar.Left := 0;
  ConfigProgressBar.Top := 80;
  ConfigProgressBar.Width := ConfigPage.SurfaceWidth;
  ConfigProgressBar.Height := 20;
  ConfigProgressBar.Min := 0;
  ConfigProgressBar.Max := 100;
  ConfigProgressBar.Position := 0;
end;

procedure CurPageChanged(CurPageID: Integer);
var
  PyRevitScript, RegisterScript, ExtensionPath, LogPath: String;
  ResultCode: Integer;
  CmdLine: String;
begin
  if CurPageID = CheckPage.ID then
  begin
    WriteLog('=== System Check Page ===');
    PyRevitInstalled := IsPyRevitInstalled;
    if PyRevitInstalled then
      PyRevitStatusLabel.Caption := '[OK] pyRevit is installed'
    else
      PyRevitStatusLabel.Caption := '[!] pyRevit is NOT installed - will be installed automatically';

    PluginVersion := GetInstalledPluginVersion;
    PluginInstalled := (PluginVersion <> '');
    if PluginInstalled then
      PluginStatusLabel.Caption := '[OK] CPSK Tools v' + PluginVersion + ' is installed - will be updated to v{#MyAppVersion}'
    else
      PluginStatusLabel.Caption := '[!] CPSK Tools is NOT installed - will be installed';

    RevitRunning := IsRevitRunning;
    if RevitRunning then
      RevitStatusLabel.Caption := '[!] Revit is RUNNING - please close Revit and restart it after installation'
    else
      RevitStatusLabel.Caption := '[OK] Revit is not running';

    if not PyRevitInstalled then
      ActionLabel.Caption := 'Click Next to install pyRevit and CPSK Tools. This may take several minutes.'
    else if PluginInstalled then
      ActionLabel.Caption := 'Click Next to update CPSK Tools from v' + PluginVersion + ' to v{#MyAppVersion}.'
    else
      ActionLabel.Caption := 'Click Next to install CPSK Tools.';
  end;

  if CurPageID = ConfigPage.ID then
  begin
    WriteLog('=== Configuration Page ===');
    WizardForm.NextButton.Enabled := False;
    WizardForm.BackButton.Enabled := False;

    PyRevitScript := ExpandConstant('{app}\tools\Install-PyRevit.ps1');
    RegisterScript := ExpandConstant('{app}\tools\Register-Extension.ps1');
    ExtensionPath := ExpandConstant('{app}\CPSK.extension');
    LogPath := ExpandConstant('{app}\logs');

    WriteLog('PyRevitScript: ' + PyRevitScript);
    WriteLog('RegisterScript: ' + RegisterScript);
    WriteLog('ExtensionPath: ' + ExtensionPath);
    WriteLog('LogPath: ' + LogPath);

    if not PyRevitInstalled then
    begin
      UpdateConfigStatus('Installing pyRevit... This may take several minutes.', 20);
      CmdLine := '-NoProfile -ExecutionPolicy Bypass -File "' + PyRevitScript + '" -RequiredVersion "{#PyRevitVersion}" -DownloadUrl "{#PyRevitUrl}"';
      WriteLog('Executing: powershell.exe ' + CmdLine);
      if Exec('powershell.exe', CmdLine, '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
        WriteLog('Install-PyRevit.ps1 completed with exit code: ' + IntToStr(ResultCode))
      else
        WriteLog('ERROR: Failed to execute Install-PyRevit.ps1');
      UpdateConfigStatus('pyRevit installation step complete.', 60);
      Sleep(500);
    end
    else
    begin
      UpdateConfigStatus('pyRevit is already installed. Skipping...', 50);
      Sleep(500);
    end;

    UpdateConfigStatus('Registering CPSK extension in pyRevit...', 80);
    CmdLine := Format('-NoProfile -ExecutionPolicy Bypass -File "%s" -ExtensionPath "%s" -LogPath "%s"', [RegisterScript, ExtensionPath, LogPath]);
    WriteLog('Executing: powershell.exe ' + CmdLine);
    if Exec('powershell.exe', CmdLine, '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
      WriteLog('Register-Extension.ps1 completed with exit code: ' + IntToStr(ResultCode))
    else
      WriteLog('ERROR: Failed to execute Register-Extension.ps1');

    UpdateConfigStatus('Configuration complete! Please restart Revit to use CPSK Tools.', 100);
    WriteLog('=== Configuration Complete ===');
    WriteLog('Log file saved to: ' + LogFile);
    Sleep(1500);

    WizardForm.NextButton.Enabled := True;
    WizardForm.NextButton.OnClick(WizardForm.NextButton);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  UninstallScript, ExtensionPath: String;
  ResultCode: Integer;
begin
  if CurUninstallStep = usUninstall then
  begin
    UninstallScript := ExpandConstant('{app}\tools\Uninstall-CPSK.ps1');
    ExtensionPath := ExpandConstant('{app}\CPSK.extension');
    Exec('powershell.exe', Format('-NoProfile -ExecutionPolicy Bypass -File "%s" -ExtensionPath "%s"', [UninstallScript, ExtensionPath]), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  end;
end;

function InitializeSetup: Boolean;
begin
  LogFile := ExpandConstant('{localappdata}\CPSK\logs\install_') + GetDateTimeString('yyyymmdd_hhnnss', '-', '-') + '.log';
  ForceDirectories(ExtractFilePath(LogFile));
  WriteLog('=== CPSK Tools Installation Started ===');
  WriteLog('Version: {#MyAppVersion}');
  WriteLog('Log file: ' + LogFile);
  WriteLog('Install path: ' + ExpandConstant('{localappdata}\CPSK'));
  Result := True;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  UninstallerPath: String;
  ResultCode: Integer;
  ExtensionPath: String;
begin
  Result := '';
  NeedsRestart := False;

  // Check for existing uninstaller
  UninstallerPath := ExpandConstant('{localappdata}\CPSK\unins000.exe');
  WriteLog('Checking for existing installation at: ' + UninstallerPath);

  if FileExists(UninstallerPath) then
  begin
    WriteLog('Found existing installation. Running uninstaller...');

    // Run uninstaller silently
    if Exec(UninstallerPath, '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    begin
      WriteLog('Uninstaller completed with exit code: ' + IntToStr(ResultCode));
      // Wait a moment for files to be released
      Sleep(1000);
    end
    else
    begin
      WriteLog('ERROR: Failed to run uninstaller');
    end;
  end
  else
  begin
    WriteLog('No existing uninstaller found');
  end;

  // Clean up old extension folder if it exists (in case uninstaller didn't remove it)
  // Note: cpsk_settings.yaml is in root folder and will NOT be deleted
  ExtensionPath := ExpandConstant('{localappdata}\CPSK\CPSK.extension');
  if DirExists(ExtensionPath) then
  begin
    WriteLog('Cleaning up old extension folder: ' + ExtensionPath);
    if not DelTree(ExtensionPath, True, True, True) then
    begin
      WriteLog('WARNING: Could not fully remove old extension folder');
    end
    else
    begin
      WriteLog('Old extension folder removed successfully');
    end;
  end;
end;

function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo, MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
begin
  Result := 'CPSK Tools v{#MyAppVersion} will be installed.' + NewLine + NewLine;
  if PyRevitInstalled then
    Result := Result + '- pyRevit: Already installed' + NewLine
  else
    Result := Result + '- pyRevit: Will be installed' + NewLine;
  if PluginInstalled then
    Result := Result + '- CPSK Tools: Update from v' + PluginVersion + NewLine
  else
    Result := Result + '- CPSK Tools: New installation' + NewLine;
  Result := Result + NewLine + 'Click Install to continue.';
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    WriteLog('=== Post-Install Step ===');
    // All user configs are stored in cpsk_settings.yaml in root folder
    // and are preserved during upgrade (not inside CPSK.extension)
    if RevitRunning then
      WriteLog('Revit was running during installation - user needs to restart Revit');
  end;
end;

function GetCustomSetupExitCode: Integer;
begin
  Result := 0;
end;
