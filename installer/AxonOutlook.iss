; Axon Outlook add-in — standalone installer (per-user, no admin).
; Installs ONLY the Outlook add-in (no floating dot). Registers the managed COM add-in and writes its
; config (Mistral primary, OpenAI backup). For colleagues who just want the email features.
;
; Build:  ISCC.exe /DMistralKey=... /DOpenAIKey=... installer\AxonOutlook.iss
; Output: installer\Output\Axon-Outlook-Setup.exe

#define AppName     "Axon Outlook add-in"
#define AppVersion  "1.0.0"
#define AppPublisher "Axon Group"
#define AddinProgId "Axon.OutlookAddin"
; Keys for the add-in config.json, passed at build time (empty by default so the committed script
; holds no secret). build.ps1 reads them from assistant\.env. Standalone add-in has no dot .env to
; read from, so BOTH keys are baked directly into config.json.
#ifndef MistralKey
  #define MistralKey ""
#endif
#ifndef OpenAIKey
  #define OpenAIKey ""
#endif
#define AddinDir  "..\outlook_addin"

[Setup]
AppId={{B2D8F1A3-4E77-4C9B-A1E6-8D3F2B9C7E52}}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
; Machine-wide: C:\Program Files\Axon\Disassist, registered under HKLM so it loads for EVERY user.
DefaultDirName={commonpf}\Axon\Disassist
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=admin
OutputDir=Output
OutputBaseFilename=Axon-Outlook-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
SetupIconFile=axon.ico

[Files]
; The add-in DLL + its ribbon icons must sit in the same folder (the add-in loads icons from its
; own DLL directory).
Source: "{#AddinDir}\AxonAddin.dll";      DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-move.png";      DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-download.png";  DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-summarize.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-reply.png";     DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-schedule.png";  DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-followup.png";  DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-sendlater.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-write.png";     DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-attach.png";    DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddinDir}\axon-settings.png";  DestDir: "{app}"; Flags: ignoreversion

[UninstallDelete]
; WriteAddinConfig creates config.json at run time, so Setup doesn't track it and uninstall would
; otherwise leave the API keys sitting in Program Files. Remove it, then the now-empty folders.
Type: files;     Name: "{app}\config.json"
Type: dirifempty; Name: "{app}"
Type: dirifempty; Name: "{commonpf}\Axon"

[Code]
const
  CLSID  = '{7B2C9E14-6A3D-4F58-9C21-3E5A1B7D4F60}';
  PROGID = 'Axon.OutlookAddin';
  ASM    = 'AxonAddin, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null';

{ Register the managed COM Outlook add-in MACHINE-WIDE (HKLM), so it loads for every user on the PC.
  Plain strings avoid the [Registry] GUID-brace parsing problem, and on a 64-bit OS this writes the
  64-bit view Outlook reads. }
procedure RegisterAddin;
var
  code, clsKey, inproc, addins: String;
begin
  code := ExpandConstant('{app}\AxonAddin.dll');
  StringChangeEx(code, '\', '/', True);
  code := 'file:///' + code;

  RegWriteStringValue(HKLM, 'Software\Classes\' + PROGID + '\CLSID', '', CLSID);

  clsKey := 'Software\Classes\CLSID\' + CLSID;
  inproc := clsKey + '\InprocServer32';
  RegWriteStringValue(HKLM, inproc, '', ExpandConstant('{sys}\mscoree.dll'));
  RegWriteStringValue(HKLM, inproc, 'ThreadingModel', 'Both');
  RegWriteStringValue(HKLM, inproc, 'Class', 'Axon.OutlookAddin.Connect');
  RegWriteStringValue(HKLM, inproc, 'Assembly', ASM);
  RegWriteStringValue(HKLM, inproc, 'RuntimeVersion', 'v4.0.30319');
  RegWriteStringValue(HKLM, inproc, 'CodeBase', code);
  RegWriteStringValue(HKLM, inproc + '\0.0.0.0', 'Class', 'Axon.OutlookAddin.Connect');
  RegWriteStringValue(HKLM, inproc + '\0.0.0.0', 'Assembly', ASM);
  RegWriteStringValue(HKLM, inproc + '\0.0.0.0', 'RuntimeVersion', 'v4.0.30319');
  RegWriteStringValue(HKLM, inproc + '\0.0.0.0', 'CodeBase', code);
  RegWriteStringValue(HKLM, clsKey + '\ProgId', '', PROGID);

  addins := 'Software\Microsoft\Office\Outlook\AddIns\' + PROGID;
  RegWriteStringValue(HKLM, addins, 'FriendlyName', 'Axon intelligence');
  RegWriteStringValue(HKLM, addins, 'Description', 'Summarize, reply, schedule, file and download emails with Axon');
  RegWriteDWordValue(HKLM, addins, 'LoadBehavior', 3);
end;

{ Write the add-in config with both keys baked, NEXT TO THE DLL in the install folder. A machine-wide
  install cannot write into every user's %APPDATA%, so the keys live beside the assembly and serve all
  users. A user who wants their own model server can still place a config.json in %APPDATA%\AxonOutlook,
  which wins. }
procedure WriteAddinConfig;
var
  path, json, key, okey: String;
begin
  key := '{#MistralKey}';
  okey := '{#OpenAIKey}';
  path := ExpandConstant('{app}\config.json');
  if key = '' then
    json := '{"api_base": "https://api.openai.com/v1", "api_key": "' + okey + '", "model": "gpt-4o"}'
  else
    json := '{"api_base": "https://api.mistral.ai/v1", "api_key": "' + key + '", "model": "mistral-medium-latest",' +
            ' "backup_api_base": "https://api.openai.com/v1", "backup_api_key": "' + okey + '", "backup_model": "gpt-4o"}';
  SaveStringToFile(path, json, False);
end;

{ Earlier builds installed PER-USER (HKCU + %LOCALAPPDATA%\AxonOutlook). HKCU COM registration takes
  precedence over HKLM, so a leftover per-user copy would shadow this machine-wide one and the user
  would keep running the old add-in. Scrub it. }
procedure RemoveOldPerUserInstall;
var
  old: String;
begin
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\CLSID\' + CLSID);
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\' + PROGID);
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Microsoft\Office\Outlook\AddIns\' + PROGID);
  old := ExpandConstant('{localappdata}\AxonOutlook');
  if DirExists(old) then DelTree(old, True, True, True);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    RemoveOldPerUserInstall;
    RegisterAddin;
    WriteAddinConfig;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    RegDeleteKeyIncludingSubkeys(HKLM, 'Software\Classes\CLSID\' + CLSID);
    RegDeleteKeyIncludingSubkeys(HKLM, 'Software\Classes\' + PROGID);
    RegDeleteKeyIncludingSubkeys(HKLM, 'Software\Microsoft\Office\Outlook\AddIns\' + PROGID);
  end;
end;
