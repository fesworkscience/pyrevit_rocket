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

[Files]
Source: "..\build\CPSK.extension\*"; DestDir: "{app}\CPSK.extension"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\build\Install-PyRevit.ps1"; DestDir: "{app}\tools"; Flags: ignoreversion
Source: "..\build\Register-Extension.ps1"; DestDir: "{app}\tools"; Flags: ignoreversion
Source: "..\build\Uninstall-CPSK.ps1"; DestDir: "{app}\tools"; Flags: ignoreversion
Source: "..\build\version.txt"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
Root: HKCU; Subkey: "Software\CPSK\Tools"; ValueType: string; ValueName: "Version"; ValueData: "{#MyAppVersion}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\CPSK\Tools"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"; Flags: uninsdeletekey

[Code]
var
  CheckPage: TWizardPage;
  ConfigPage: TWizardPage;
  PyRevitStatusLabel: TNewStaticText;
  PluginStatusLabel: TNewStaticText;
  ActionLabel: TNewStaticText;
  ConfigStatusLabel: TNewStaticText;
  ConfigProgressBar: TNewProgressBar;
  PyRevitInstalled: Boolean;
  PluginInstalled: Boolean;
  PluginVersion: String;

function IsPyRevitInstalled: Boolean;
var
  Path: String;
begin
  Result := False;
  if RegQueryStringValue(HKEY_CURRENT_USER, 'Software\pyRevit', 'InstallPath', Path) then
    Result := True
  else if DirExists(ExpandConstant('{userappdata}\pyRevit-Master')) then
    Result := True
  else if DirExists(ExpandConstant('{commonappdata}\pyRevit')) then
    Result := True
  else if FileExists(ExpandConstant('{pf}\pyRevit CLI\pyrevit.exe')) then
    Result := True;
end;

function GetInstalledPluginVersion: String;
var
  Version: String;
begin
  Result := '';
  if RegQueryStringValue(HKEY_CURRENT_USER, 'Software\CPSK\Tools', 'Version', Version) then
    Result := Version;
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

  ActionLabel := TNewStaticText.Create(CheckPage);
  ActionLabel.Parent := CheckPage.Surface;
  ActionLabel.Left := 0;
  ActionLabel.Top := 150;
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
  PyRevitScript, RegisterScript, ExtensionPath: String;
  ResultCode: Integer;
begin
  if CurPageID = CheckPage.ID then
  begin
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

    if not PyRevitInstalled then
      ActionLabel.Caption := 'Click Next to install pyRevit and CPSK Tools. This may take several minutes.'
    else if PluginInstalled then
      ActionLabel.Caption := 'Click Next to update CPSK Tools from v' + PluginVersion + ' to v{#MyAppVersion}.'
    else
      ActionLabel.Caption := 'Click Next to install CPSK Tools.';
  end;

  if CurPageID = ConfigPage.ID then
  begin
    WizardForm.NextButton.Enabled := False;
    WizardForm.BackButton.Enabled := False;

    PyRevitScript := ExpandConstant('{app}\tools\Install-PyRevit.ps1');
    RegisterScript := ExpandConstant('{app}\tools\Register-Extension.ps1');
    ExtensionPath := ExpandConstant('{app}\CPSK.extension');

    if not PyRevitInstalled then
    begin
      UpdateConfigStatus('Installing pyRevit... This may take several minutes.', 20);
      Exec('powershell.exe', Format('-NoProfile -ExecutionPolicy Bypass -File "%s" -RequiredVersion "{#PyRevitVersion}" -DownloadUrl "{#PyRevitUrl}"', [PyRevitScript]), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
      UpdateConfigStatus('pyRevit installed successfully!', 60);
      Sleep(500);
    end
    else
    begin
      UpdateConfigStatus('pyRevit is already installed. Skipping...', 50);
      Sleep(500);
    end;

    UpdateConfigStatus('Registering CPSK extension in pyRevit...', 80);
    Exec('powershell.exe', Format('-NoProfile -ExecutionPolicy Bypass -File "%s" -ExtensionPath "%s"', [RegisterScript, ExtensionPath]), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

    UpdateConfigStatus('Configuration complete! Please restart Revit to use CPSK Tools.', 100);
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
  Result := True;
end;
