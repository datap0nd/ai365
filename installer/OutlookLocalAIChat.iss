#define AppName "Outlook Local AI Chat"
#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif
#define AppPublisher "Outlook Local AI Chat"
#define AppProgId "OutlookLocalAIChat.AddIn"
#define AppClsid "{{0D6E56F9-BE2D-4B94-B5E4-4C2DB0FD13E7}"
#define ManagedCategory "{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
#define AssemblyName "OutlookLocalAIChat, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"

[Setup]
AppId={{6BA7BCA9-F17E-4B50-8734-242063264160}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={localappdata}\Programs\Outlook Local AI Chat
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x86 x64compatible
OutputDir=..\artifacts
OutputBaseFilename=OutlookLocalAIChatSetup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
CloseApplicationsFilter=outlook.exe
RestartApplications=no
UninstallDisplayName={#AppName}
VersionInfoVersion={#AppVersion}
VersionInfoDescription=Local read-and-draft-only AI chat add-in for Outlook
VersionInfoProductName={#AppName}
VersionInfoCompany={#AppPublisher}

[Files]
Source: "..\src\OutlookLocalAIChat\bin\Release\OutlookLocalAIChat.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; 32-bit COM registration. Required for 32-bit Office, including on 64-bit Windows.
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey 32bit
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey 32bit
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Flags: 32bit
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.AddIn"; Flags: 32bit
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Flags: 32bit
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Flags: 32bit
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Flags: 32bit
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.AddIn"; Flags: 32bit
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Flags: 32bit
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Flags: 32bit
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Flags: 32bit
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#AppProgId}"; Flags: 32bit
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Flags: 32bit
Root: HKCU; Subkey: "Software\Classes\{#AppProgId}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey 32bit
Root: HKCU; Subkey: "Software\Classes\{#AppProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#AppClsid}"; Flags: 32bit
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey 32bit
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Local AI chat for the selected email with unsent draft creation."; Flags: 32bit
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: 32bit
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Flags: 32bit

; 64-bit COM registration. Written only on 64-bit Windows.
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Flags: 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.AddIn"; Flags: 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Flags: 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Flags: 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Flags: 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.AddIn"; Flags: 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Flags: 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Flags: 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.0.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Flags: 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#AppProgId}"; Flags: 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\CLSID\{#AppClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Flags: 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\{#AppProgId}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Classes\{#AppProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#AppClsid}"; Flags: 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Local AI chat for the selected email with unsent draft creation."; Flags: 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: 64bit; Check: IsWin64
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Flags: 64bit; Check: IsWin64

[Code]
function GetAssemblyCodeBase(Param: String): String;
var
  Path: String;
begin
  Path := ExpandConstant('{app}\OutlookLocalAIChat.dll');
  StringChangeEx(Path, '\', '/', True);
  Result := 'file:///' + Path;
end;
