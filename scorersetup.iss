#define MyAppName "乞分君"
#define MyAppVersion "7.4.2"
#define MyAppPublisher "丐帮集团第一院·物理版象棋开发与研究院™"
#define MyAppURL "https://bggp.dpdns.org/1/scorer/"
#define MyAppExeName "scorerlauncher.exe"
#define SourcePath "D:\a"

[Setup]
AppId={{D971EE51-CCD8-4B4A-B7EC-23AE555B97BC}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName=D:\Program Files\bggp\1\scorer
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=C:\Users\鸿合HiteVision\OneDrive\桌面
OutputBaseFilename=Scorer_Installer_v7.4.2
SetupIconFile=C:\Users\鸿合HiteVision\OneDrive\文档\dec.ico
SolidCompression=yes
WizardStyle=modern dynamic windows11

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Default.isl"
Name: "chinesetraditional"; MessagesFile: "compiler:Languages\ChineseTraditional.isl"
Name: "english"; MessagesFile: "compiler:Languages\English.isl"
Name: "french"; MessagesFile: "compiler:Languages\French.isl"
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; 主程序和核心文件
Source: "{#SourcePath}\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\kei.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\notice.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}\ico.ico"; DestDir: "{app}"; Flags: ignoreversion
; 所有其余依赖（子文件夹、DLL 等）
Source: "{#SourcePath}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
; 密码管理快捷方式（带 -kc 参数）
Name: "{autoprograms}\{#MyAppName} 密码管理"; Filename: "{app}\{#MyAppExeName}"; Parameters: "-kc"

[Run]
; 1. 用记事本打开 kei.json 并等待用户编辑
Filename: "{win}\notepad.exe"; Parameters: "{app}\kei.json"; Flags: waituntilterminated; Description: "正在编辑配置文件······"
; 2. 编辑完成后自动启动主程序（不等待）
Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Flags: nowait; Description: "启动主程序"