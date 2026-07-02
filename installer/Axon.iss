; Axon intelligence — Windows installer (per-user, no admin).
; Installs the dot (PyInstaller build), the bundled open-computer-use, and the Outlook add-in,
; and registers the add-in so the Move/Download buttons appear in Outlook automatically.
;
; Build:  "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\Axon.iss
; Output: installer\Output\Axon-Setup.exe

#define AppName     "Axon intelligence"
#define AppVersion  "1.0.0"
#define AppPublisher "Axon Group"
#define AppExe      "AxonIntelligence.exe"
#define AddinClsid  "{{7B2C9E14-6A3D-4F58-9C21-3E5A1B7D4F60}}"
#define AddinProgId "Axon.OutlookAddin"
; Source folders (relative to this .iss in installer\)
#define DistDir   "..\assistant\dist\AxonIntelligence"
#define AddinDir  "..\outlook_addin"

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
OutputBaseFilename=Axon-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"
Name: "startup";     Description: "Start Axon when I sign in to Windows"; GroupDescription: "Startup:"

[Files]
; The dot exe + its _internal deps + the bundled open-computer-use (ocu\) all live under DistDir.
Source: "{#DistDir}\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion
; The Outlook add-in DLL + icons.
Source: "{#AddinDir}\AxonAddin.dll";     DestDir: "{app}\addin"; Flags: ignoreversion
Source: "{#AddinDir}\axon-move.png";      DestDir: "{app}\addin"; Flags: ignoreversion
Source: "{#AddinDir}\axon-download.png";  DestDir: "{app}\addin"; Flags: ignoreversion
Source: "{#AddinDir}\axon-summarize.png"; DestDir: "{app}\addin"; Flags: ignoreversion
Source: "{#AddinDir}\axon-reply.png";     DestDir: "{app}\addin"; Flags: ignoreversion
Source: "{#AddinDir}\axon-schedule.png";  DestDir: "{app}\addin"; Flags: ignoreversion
Source: "{#AddinDir}\axon-followup.png";  DestDir: "{app}\addin"; Flags: ignoreversion
Source: "{#AddinDir}\axon-sendlater.png"; DestDir: "{app}\addin"; Flags: ignoreversion

[Icons]
; AppUserModelID must match APP_ID in axon/notify.py so Windows toasts show under "Axon intelligence".
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"; AppUserModelID: "AxonIntelligence.Dot"
Name: "{autodesktop}\{#AppName}";  Filename: "{app}\{#AppExe}"; Tasks: desktopicon; AppUserModelID: "AxonIntelligence.Dot"

[Registry]
; Start the dot at sign-in (optional task). The Outlook add-in COM registration is done in [Code]
; below (plain strings — Inno's [Registry] mishandles the GUID braces in CLSID subkeys).
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "AxonIntelligence"; ValueData: """{app}\{#AppExe}"""; Flags: uninsdeletevalue; Tasks: startup

[Run]
Filename: "{app}\{#AppExe}"; Description: "Start Axon intelligence now"; Flags: nowait postinstall skipifsilent

[Code]
const
  CLSID  = '{7B2C9E14-6A3D-4F58-9C21-3E5A1B7D4F60}';
  PROGID = 'Axon.OutlookAddin';
  ASM    = 'AxonAddin, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null';

{ Register the managed COM Outlook add-in (per-user). Plain strings here avoid the [Registry]
  GUID-brace parsing problem, and on a 64-bit OS this writes the 64-bit view Outlook reads. }
procedure RegisterAddin;
var
  code, clsKey, inproc, addins: String;
begin
  code := ExpandConstant('{app}\addin\AxonAddin.dll');
  StringChangeEx(code, '\', '/', True);
  code := 'file:///' + code;

  RegWriteStringValue(HKCU, 'Software\Classes\' + PROGID + '\CLSID', '', CLSID);

  clsKey := 'Software\Classes\CLSID\' + CLSID;
  inproc := clsKey + '\InprocServer32';
  RegWriteStringValue(HKCU, inproc, '', ExpandConstant('{sys}\mscoree.dll'));
  RegWriteStringValue(HKCU, inproc, 'ThreadingModel', 'Both');
  RegWriteStringValue(HKCU, inproc, 'Class', 'Axon.OutlookAddin.Connect');
  RegWriteStringValue(HKCU, inproc, 'Assembly', ASM);
  RegWriteStringValue(HKCU, inproc, 'RuntimeVersion', 'v4.0.30319');
  RegWriteStringValue(HKCU, inproc, 'CodeBase', code);
  RegWriteStringValue(HKCU, inproc + '\0.0.0.0', 'Class', 'Axon.OutlookAddin.Connect');
  RegWriteStringValue(HKCU, inproc + '\0.0.0.0', 'Assembly', ASM);
  RegWriteStringValue(HKCU, inproc + '\0.0.0.0', 'RuntimeVersion', 'v4.0.30319');
  RegWriteStringValue(HKCU, inproc + '\0.0.0.0', 'CodeBase', code);
  RegWriteStringValue(HKCU, clsKey + '\ProgId', '', PROGID);

  addins := 'Software\Microsoft\Office\Outlook\AddIns\' + PROGID;
  RegWriteStringValue(HKCU, addins, 'FriendlyName', 'Axon intelligence');
  RegWriteStringValue(HKCU, addins, 'Description', 'File and download emails with Axon intelligence');
  RegWriteDWordValue(HKCU, addins, 'LoadBehavior', 3);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    RegisterAddin;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\CLSID\' + CLSID);
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\' + PROGID);
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Microsoft\Office\Outlook\AddIns\' + PROGID);
  end;
end;
