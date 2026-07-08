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
DefaultDirName={localappdata}\AxonOutlook
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=lowest
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

[Code]
const
  CLSID  = '{7B2C9E14-6A3D-4F58-9C21-3E5A1B7D4F60}';
  PROGID = 'Axon.OutlookAddin';
  ASM    = 'AxonAddin, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null';

{ Register the managed COM Outlook add-in (per-user). Plain strings avoid the [Registry] GUID-brace
  parsing problem, and on a 64-bit OS this writes the 64-bit view Outlook reads. }
procedure RegisterAddin;
var
  code, clsKey, inproc, addins: String;
begin
  code := ExpandConstant('{app}\AxonAddin.dll');
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
  RegWriteStringValue(HKCU, addins, 'Description', 'Summarize, reply, schedule, file and download emails with Axon');
  RegWriteDWordValue(HKCU, addins, 'LoadBehavior', 3);
end;

{ Write the add-in config with both keys baked (no dot .env alongside a standalone add-in). Never
  clobbers an existing config. }
procedure WriteAddinConfig;
var
  dir, path, json, key, okey: String;
begin
  key := '{#MistralKey}';
  okey := '{#OpenAIKey}';
  dir := ExpandConstant('{userappdata}\AxonOutlook');
  ForceDirectories(dir);
  path := dir + '\config.json';
  if FileExists(path) then exit;
  if key = '' then
    json := '{"api_base": "https://api.openai.com/v1", "api_key": "' + okey + '", "model": "gpt-4o"}'
  else
    json := '{"api_base": "https://api.mistral.ai/v1", "api_key": "' + key + '", "model": "mistral-medium-latest",' +
            ' "backup_api_base": "https://api.openai.com/v1", "backup_api_key": "' + okey + '", "backup_model": "gpt-4o"}';
  SaveStringToFile(path, json, False);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    RegisterAddin;
    WriteAddinConfig;
  end;
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
