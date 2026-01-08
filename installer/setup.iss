#define MyAppName "CPSK Tools for Revit"
#define MyAppPublisher "CPSK"
#define MyAppURL "https://rocket-tools.ru"
#define PyRevitURL "https://pyrevitlabs.io"

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
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
russian.WelcomeLabel2=��������� ���������� {#MyAppName} �� ��� ���������.%n%n��� ������ ��������� pyRevit - ���������� ��������� ��� ���������� Revit.%n%n���� pyRevit �� ����������, �� ����� ���������� ������������ ��.
english.WelcomeLabel2=This will install {#MyAppName} on your computer.%n%nThis extension requires pyRevit - a free platform for Revit extensions.%n%nIf pyRevit is not installed, it will be installed automatically.

[Files]
Source: "build\CPSK.extension\*"; DestDir: "{app}\CPSK.extension"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "build\Install-PyRevit.ps1"; DestDir: "{app}\tools"; Flags: ignoreversion
Source: "build\Register-Extension.ps1"; DestDir: "{app}\tools"; Flags: ignoreversion
Source: "build\Uninstall-CPSK.ps1"; DestDir: "{app}\tools"; Flags: ignoreversion
Source: "build\version.txt"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
Root: HKCU; Subkey: "Software\CPSK\Tools"; ValueType: string; ValueName: "Version"; ValueData: "{#MyAppVersion}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\CPSK\Tools"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"; Flags: uninsdeletekey

[Code]
var
  ProgressPage: TOutputProgressWizardPage;

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

function RunPowerShellScript(ScriptPath, Arguments, StatusMsg: String): Boolean;
var
  ResultCode: Integer;
  CmdLine: String;
begin
  ProgressPage.SetText(StatusMsg, '');
  ProgressPage.SetProgress(ProgressPage.ProgressBar.Position + 10, 100);

  CmdLine := Format('-NoProfile -ExecutionPolicy Bypass -File "%s" %s', [ScriptPath, Arguments]);
  Result := Exec('powershell.exe', CmdLine, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

procedure InitializeWizard;
begin
  ProgressPage := CreateOutputProgressPage('����������� ���������', '���������, ����������...');
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  PyRevitScript, RegisterScript: String;
  ExtensionPath: String;
begin
  if CurStep = ssPostInstall then
  begin
    ProgressPage.Show;
    ProgressPage.SetProgress(0, 100);

    PyRevitScript := ExpandConstant('{app}\tools\Install-PyRevit.ps1');
    RegisterScript := ExpandConstant('{app}\tools\Register-Extension.ps1');
    ExtensionPath := ExpandConstant('{app}\CPSK.extension');

    // Install pyRevit if needed
    ProgressPage.SetText('��������� ������� pyRevit...', '');
    ProgressPage.SetProgress(10, 100);

    if not IsPyRevitInstalled then
    begin
      ProgressPage.SetText('��������� pyRevit...', '��� ����� ������ ��������� �����');
      ProgressPage.SetProgress(20, 100);
      RunPowerShellScript(PyRevitScript,
        Format('-RequiredVersion "{#PyRevitVersion}" -DownloadUrl "{#PyRevitUrl}"', []),
        '��������� pyRevit...');
      ProgressPage.SetProgress(70, 100);
    end
    else
    begin
      ProgressPage.SetText('pyRevit ��� ����������', '');
      ProgressPage.SetProgress(50, 100);
    end;

    // Register extension
    ProgressPage.SetText('����������� ����������...', '');
    ProgressPage.SetProgress(80, 100);
    RunPowerShellScript(RegisterScript,
      Format('-ExtensionPath "%s"', [ExtensionPath]),
      '����������� ���������� � pyRevit...');

    ProgressPage.SetProgress(100, 100);
    ProgressPage.SetText('��������� ���������!', '');
    Sleep(500);
    ProgressPage.Hide;
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

    Exec('powershell.exe',
      Format('-NoProfile -ExecutionPolicy Bypass -File "%s" -ExtensionPath "%s"',
        [UninstallScript, ExtensionPath]),
      '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  end;
end;

function InitializeSetup: Boolean;
begin
  Result := True;
end;
