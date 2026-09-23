; Figured installer (Inno Setup 6). Built by build.ps1, which passes /DAppVersion.
#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif

[Setup]
AppId={{F5ACDB33-A493-4D01-85F1-28335DE2B456}
AppName=Figured
AppVersion={#AppVersion}
AppVerName=Figured {#AppVersion}
AppPublisher=Mohamed Alkhmis
AppPublisherURL=https://github.com/MAK340/figured
AppSupportURL=https://github.com/MAK340/figured/issues
AppUpdatesURL=https://github.com/MAK340/figured/releases
VersionInfoVersion={#AppVersion}
; per-user install: no admin prompt, lands in %LOCALAPPDATA%\Programs\Figured
PrivilegesRequired=lowest
DefaultDirName={autopf}\Figured
DisableProgramGroupPage=yes
DisableDirPage=auto
OutputDir=dist
OutputBaseFilename=Figured-{#AppVersion}-setup
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\Figured.exe
UninstallDisplayName=Figured
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=no

[Tasks]
Name: "autostart"; Description: "Start Figured when Windows starts"
Name: "desktopicon"; Description: "Create a desktop shortcut"; Flags: unchecked

[InstallDelete]
; drop libraries left by an older version before copying the new ones
Type: filesandordirs; Name: "{app}\_internal"
; the pre-installer copy from when the app was called SumSelect
Type: filesandordirs; Name: "{localappdata}\SumSelect"

[Files]
Source: "dist\Figured\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\Figured"; Filename: "{app}\Figured.exe"
Name: "{autodesktop}\Figured"; Filename: "{app}\Figured.exe"; Tasks: desktopicon

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "Figured"; ValueData: """{app}\Figured.exe"""; Tasks: autostart; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: none; ValueName: "SumSelect"; Flags: deletevalue

[Run]
Filename: "{app}\Figured.exe"; Description: "Launch Figured"; Flags: nowait postinstall

[UninstallRun]
Filename: "{sys}\taskkill.exe"; Parameters: "/F /IM Figured.exe"; Flags: runhidden; RunOnceId: "StopFigured"

[Code]
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  Rc: Integer;
begin
  { close a running copy (current or old name) so its files can be replaced }
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM Figured.exe', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM SumSelect.exe', '', SW_HIDE, ewWaitUntilTerminated, Rc);
  Sleep(700);
  Result := '';
end;