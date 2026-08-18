#define AppName "MetoAI"
#ifndef AppVersion
  #define AppVersion "1.10.3"
#endif
#define AppPublisher "MetoAI contributors"
#define AppProgId "OutlookLocalAIChat.AddIn"
#define AppClsid "{{0D6E56F9-BE2D-4B94-B5E4-4C2DB0FD13E7}"
#define PaneProgId "OutlookLocalAIChat.ChatPane"
#define PaneClsid "{{14D24FA1-4342-442F-B68B-B68D7372794C}"
#define ExcelProgId "AI365.ExcelAddIn"
#define ExcelClsid "{{C0ABFA36-9854-434D-A542-DD834938737F}"
#define PptProgId "AI365.PowerPointAddIn"
#define PptClsid "{{69FAE812-274F-43F8-8F45-1B4EB22B5248}"
#define OfficePaneProgId "AI365.OfficePane"
#define OfficePaneClsid "{{BC9047E7-9AFE-4F75-BBBC-27241B1DE2FA}"
#define ManagedCategory "{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
#define ControlCategory "{{40FC6ED4-2438-11CF-A3DB-080036F12502}"
#define LockbackInterface "{{000C0601-0000-0000-C000-000000000046}"
#define AssemblyName "OutlookLocalAIChat, Version=1.1.0.0, Culture=neutral, PublicKeyToken=f51b005bfa6d7cc3"

[Setup]
AppId={{6BA7BCA9-F17E-4B50-8734-242063264160}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={localappdata}\Programs\OutlookLocalAIChat
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
CloseApplicationsFilter=outlook.exe,excel.exe,powerpnt.exe
RestartApplications=no
UninstallDisplayName={#AppName}
VersionInfoVersion={#AppVersion}
VersionInfoDescription=Local mailbox AI chat with one linked unsent Outlook draft
VersionInfoProductName={#AppName}
VersionInfoCompany={#AppPublisher}

[Files]
Source: "..\src\OutlookLocalAIChat\bin\Release\OutlookLocalAIChat.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\src\OutlookLocalAIChat\bin\Release\Microsoft.Web.WebView2.Core.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\src\OutlookLocalAIChat\bin\Release\Microsoft.Web.WebView2.WinForms.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\src\OutlookLocalAIChat\bin\Release\Microsoft.Web.WebView2.Wpf.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\src\OutlookLocalAIChat\bin\Release\runtimes\win-x86\native\WebView2Loader.dll"; DestDir: "{app}\runtimes\win-x86\native"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\src\OutlookLocalAIChat\bin\Release\runtimes\win-x64\native\WebView2Loader.dll"; DestDir: "{app}\runtimes\win-x64\native"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\src\OutlookLocalAIChat\bin\Release\runtimes\win-arm64\native\WebView2Loader.dll"; DestDir: "{app}\runtimes\win-arm64\native"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\src\OutlookLocalAIChat\bin\Release\WebView2Loader.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; 32-bit COM registration. Required for 32-bit Office, including on 64-bit Windows.
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.AddIn"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.AddIn"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#AppProgId}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""
Root: HKCU32; Subkey: "Software\Classes\{#AppProgId}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\{#AppProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#AppClsid}"
Root: HKCU32; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Local mailbox AI chat with one linked unsent draft."
Root: HKCU32; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"
Root: HKCU32; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"

; Managed ActiveX control hosted by Office as the native Outlook sidebar.
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}"; ValueType: string; ValueName: ""; ValueData: "{#AppName} Sidebar"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.UI.ChatPane"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.UI.ChatPane"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#PaneProgId}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Implemented Categories\{#ControlCategory}"; ValueType: string; ValueName: ""; ValueData: ""
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Programmable"; ValueType: string; ValueName: ""; ValueData: ""
Root: HKCU32; Subkey: "Software\Classes\{#PaneProgId}"; ValueType: string; ValueName: ""; ValueData: "{#AppName} Sidebar"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\{#PaneProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#PaneClsid}"
Root: HKCU32; Subkey: "Software\Classes\Interface\{#LockbackInterface}"; ValueType: string; ValueName: ""; ValueData: "Office .NET Framework Lockback Bypass Key"

; 64-bit COM registration. Written only on 64-bit Windows.
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.AddIn"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.AddIn"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#AppProgId}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\{#AppProgId}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\{#AppProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#AppClsid}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Local mailbox AI chat with one linked unsent draft."; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Check: IsWin64

Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}"; ValueType: string; ValueName: ""; ValueData: "{#AppName} Sidebar"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.UI.ChatPane"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.UI.ChatPane"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#PaneProgId}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Implemented Categories\{#ControlCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Programmable"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\{#PaneProgId}"; ValueType: string; ValueName: ""; ValueData: "{#AppName} Sidebar"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\{#PaneProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#PaneClsid}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\Interface\{#LockbackInterface}"; ValueType: string; ValueName: ""; ValueData: "Office .NET Framework Lockback Bypass Key"; Check: IsWin64


; AI365 for Excel add-in (32-bit).
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}"; ValueType: string; ValueName: ""; ValueData: "AI365 for Excel"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.ExcelAddIn"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.ExcelAddIn"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#ExcelProgId}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""
Root: HKCU32; Subkey: "Software\Classes\{#ExcelProgId}"; ValueType: string; ValueName: ""; ValueData: "AI365 for Excel"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\{#ExcelProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#ExcelClsid}"
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "AI365"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Chat with your workbook. AI365 never saves, deletes, or sends."
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"

; AI365 for PowerPoint add-in (32-bit).
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}"; ValueType: string; ValueName: ""; ValueData: "AI365 for PowerPoint"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.PowerPointAddIn"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.PowerPointAddIn"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#PptProgId}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""
Root: HKCU32; Subkey: "Software\Classes\{#PptProgId}"; ValueType: string; ValueName: ""; ValueData: "AI365 for PowerPoint"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\{#PptProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#PptClsid}"
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "AI365"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Chat with your presentation. AI365 never saves, deletes, or sends."
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"

; Managed ActiveX control hosted as the Excel/PowerPoint sidebar (32-bit).
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}"; ValueType: string; ValueName: ""; ValueData: "AI365 Sidebar"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.UI.OfficeChatPane"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.UI.OfficeChatPane"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#OfficePaneProgId}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Implemented Categories\{#ControlCategory}"; ValueType: string; ValueName: ""; ValueData: ""
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Programmable"; ValueType: string; ValueName: ""; ValueData: ""
Root: HKCU32; Subkey: "Software\Classes\{#OfficePaneProgId}"; ValueType: string; ValueName: ""; ValueData: "AI365 Sidebar"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\{#OfficePaneProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#OfficePaneClsid}"

; AI365 for Excel add-in (64-bit).
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}"; ValueType: string; ValueName: ""; ValueData: "AI365 for Excel"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.ExcelAddIn"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.ExcelAddIn"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#ExcelProgId}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\{#ExcelProgId}"; ValueType: string; ValueName: ""; ValueData: "AI365 for Excel"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\{#ExcelProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#ExcelClsid}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "AI365"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Chat with your workbook. AI365 never saves, deletes, or sends."; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Check: IsWin64

; AI365 for PowerPoint add-in (64-bit).
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}"; ValueType: string; ValueName: ""; ValueData: "AI365 for PowerPoint"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.PowerPointAddIn"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.PowerPointAddIn"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#PptProgId}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\{#PptProgId}"; ValueType: string; ValueName: ""; ValueData: "AI365 for PowerPoint"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\{#PptProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#PptClsid}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "AI365"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Chat with your presentation. AI365 never saves, deletes, or sends."; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Check: IsWin64

; Managed ActiveX control hosted as the Excel/PowerPoint sidebar (64-bit).
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}"; ValueType: string; ValueName: ""; ValueData: "AI365 Sidebar"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.UI.OfficeChatPane"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.UI.OfficeChatPane"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#OfficePaneProgId}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Implemented Categories\{#ControlCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Programmable"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\{#OfficePaneProgId}"; ValueType: string; ValueName: ""; ValueData: "AI365 Sidebar"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\{#OfficePaneProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#OfficePaneClsid}"; Check: IsWin64

[Code]
function GetAssemblyCodeBase(Param: String): String;
var
  Path: String;
begin
  Path := ExpandConstant('{app}\OutlookLocalAIChat.dll');
  StringChangeEx(Path, '\', '/', True);
  Result := 'file:///' + Path;
end;
