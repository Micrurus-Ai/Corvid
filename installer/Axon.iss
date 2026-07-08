; Axon intelligence — the floating dot, Windows installer (per-user, no admin).
; Installs ONLY the desktop dot (PyInstaller build) + the bundled open-computer-use. The Outlook
; add-in is a SEPARATE installer (AxonOutlook.iss / Axon-Outlook-Setup.exe).
;
; Build:  "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\Axon.iss
; Output: installer\Output\Axon-Dot-Setup.exe

#define AppName     "Axon intelligence"
#define AppVersion  "1.0.0"
#define AppPublisher "Axon Group"
#define AppExe      "AxonIntelligence.exe"
; Source folder (relative to this .iss in installer\)
#define DistDir   "..\assistant\dist\AxonIntelligence"

[Setup]
AppId={{A1C7F0E2-9D44-4B1A-8E55-7C2D9E3F6B41}}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={localappdata}\Axon
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=lowest
OutputDir=Output
OutputBaseFilename=Axon-Dot-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
; Axon branding: the setup .exe icon, and the Add/Remove Programs icon (the app exe, which carries
; the same icon via PyInstaller --icon).
SetupIconFile=axon.ico
UninstallDisplayIcon={app}\{#AppExe}

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"
Name: "startup";     Description: "Start Axon when I sign in to Windows"; GroupDescription: "Startup:"

[Files]
; The dot exe + its _internal deps + the bundled open-computer-use (ocu\) all live under DistDir.
Source: "{#DistDir}\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
; AppUserModelID must match APP_ID in axon/notify.py so Windows toasts show under "Axon intelligence".
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"; AppUserModelID: "AxonIntelligence.Dot"
Name: "{autodesktop}\{#AppName}";  Filename: "{app}\{#AppExe}"; Tasks: desktopicon; AppUserModelID: "AxonIntelligence.Dot"

[Registry]
; Start the dot at sign-in (optional task).
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "AxonIntelligence"; ValueData: """{app}\{#AppExe}"""; Flags: uninsdeletevalue; Tasks: startup

[Run]
Filename: "{app}\{#AppExe}"; Description: "Start Axon intelligence now"; Flags: nowait postinstall skipifsilent
; Silent self-update path: the dot runs the setup with /relaunch=1, so reopen the app afterwards.
Filename: "{app}\{#AppExe}"; Flags: nowait; Check: RelaunchRequested

[Code]
{ True when the dot triggered a silent self-update (Axon-Dot-Setup.exe /relaunch=1), so [Run] reopens
  the app once the in-place upgrade finishes. }
function RelaunchRequested(): Boolean;
begin
  Result := ExpandConstant('{param:relaunch|0}') = '1';
end;
