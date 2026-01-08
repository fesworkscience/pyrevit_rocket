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
DisableDirPage=no
DisableReadyPage=no
InfoBeforeFile=info.txt

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
  StatusLabel: TNewStaticText;
  ProgressBar: TNewProgressBar;
  OutputPage: TWizardPage;

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
    Result := True;
end;

procedure UpdateStatus(Msg: String; Progress: Integer);
begin
  StatusLabel.Caption := Msg;
  ProgressBar.Position := Progress;
  WizardForm.Refresh;
end;

procedure InitializeWizard;
begin
  OutputPage := CreateCustomPage(wpInstalling, 'Configuring CPSK Tools', 'Please wait while setup configures the extension...');
  
  StatusLabel := TNewStaticText.Create(OutputPage);
  StatusLabel.Parent := OutputPage.Surface;
  StatusLabel.Left := 0;
  StatusLabel.Top := 20;
  StatusLabel.Width := OutputPage.SurfaceWidth;
  StatusLabel.Height := 40;
  StatusLabel.AutoSize := False;
  StatusLabel.WordWrap := True;
  StatusLabel.Caption := 'Preparing...';
  
  ProgressBar := TNewProgressBar.Create(OutputPage);
  ProgressBar.Parent := OutputPage.Surface;
  ProgressBar.Left := 0;
  ProgressBar.Top := 80;
  ProgressBar.Width := OutputPage.SurfaceWidth;
  ProgressBar.Height := 20;
  ProgressBar.Min := 0;
  ProgressBar.Max := 100;
  ProgressBar.Position := 0;
end;

procedure CurPageChanged(CurPageID: Integer);
var
  PyRevitScript, RegisterScript, ExtensionPath: String;
  ResultCode: Integer;
begin
  if CurPageID = OutputPage.ID then
  begin
    WizardForm.NextButton.Enabled := False;
    WizardForm.BackButton.Enabled := False;
    
    PyRevitScript := ExpandConstant('{app}\tools\Install-PyRevit.ps1');
    RegisterScript := ExpandConstant('{app}\tools\Register-Extension.ps1');
    ExtensionPath := ExpandConstant('{app}\CPSK.extension');
    
    UpdateStatus('Checking if pyRevit is installed...', 10);
    Sleep(500);
    
    if not IsPyRevitInstalled then
    begin
      UpdateStatus('pyRevit not found. Downloading and installing...' + #13#10 + 'This may take several minutes.', 20);
      Exec('powershell.exe', Format('-NoProfile -ExecutionPolicy Bypass -File "%s" -RequiredVersion "{#PyRevitVersion}" -DownloadUrl "{#PyRevitUrl}"', [PyRevitScript]), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
      UpdateStatus('pyRevit installation complete.', 60);
    end
    else
    begin
      UpdateStatus('pyRevit is already installed. Skipping...', 50);
    end;
    
    Sleep(500);
    UpdateStatus('Registering CPSK extension in pyRevit...', 70);
    Exec('powershell.exe', Format('-NoProfile -ExecutionPolicy Bypass -File "%s" -ExtensionPath "%s"', [RegisterScript, ExtensionPath]), '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    
    UpdateStatus('Configuration complete!', 100);
    Sleep(1000);
    
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
