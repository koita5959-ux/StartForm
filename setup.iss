; StartForm Inno Setup Script
; 便利アプリシリーズ 第二弾

#define MyAppName "StartForm"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "便利アプリシリーズ"
#define MyAppExeName "StartForm.exe"
#define MyAppDescription "自分が決めた作業環境を、いつでもワンアクションで整える"

[Setup]
AppId={{B8F3D4E2-A1C7-4F5E-9D6B-3E8A2C1F0D5E}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=installer_output
OutputBaseFilename=StartForm_Setup_1.0.0
SetupIconFile=I.ico
UninstallDisplayIcon={app}\StartForm.exe
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
DisableProgramGroupPage=yes

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "デスクトップにショートカットを作成"; GroupDescription: "追加タスク:"; Flags: unchecked

[Files]
Source: "publish\*"; Excludes: "*.pdb"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{group}\{#MyAppName} をアンインストール"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "StartFormを起動する"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
Type: filesandordirs; Name: "{userappdata}\StartForm"
