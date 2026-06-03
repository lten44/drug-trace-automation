; 药品批发企业追朔码自动处理软件 — Inno Setup 安装脚本
; 使用前请先运行 PyInstaller:
;   pyinstaller --onedir --name "药品批发企业追朔码自动处理软件" gui.py

#define MyAppName "药品批发企业追朔码自动处理软件"
#define MyAppVersion "v3.0"
#define MyAppPublisher "药品批发企业"
#define MyAppURL ""
#define MyAppExeName "药品批发企业追朔码自动处理软件.exe"

[Setup]
; 基本信息
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
; 唯一标识（用于检测已安装、自动升级/卸载）
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
; 升级时不重复询问安装目录和开始菜单
DisableDirPage=auto
DisableProgramGroupPage=auto

; 安装目录：Program Files
DefaultDirName={commonpf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes

; 输出安装包
OutputDir=.
OutputBaseFilename=Setup-{#MyAppName}-{#MyAppVersion}

; 安装包图标
SetupIconFile=icon.ico

; 压缩设置
Compression=lzma2/ultra64
SolidCompression=yes
LZMAUseSeparateProcess=yes

; 权限
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog

; 卸载相关
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName} {#MyAppVersion}

; 忽略 APPDATA 清理警告（安装后删除用户配置是预期行为）
UsedUserAreasWarning=no

; 语言
LanguageDetectionMethod=locale
ShowLanguageDialog=no

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "快捷方式："; Flags: checkedonce

[Files]
; 主程序文件（所有 PyInstaller 输出）
Source: "dist\{#MyAppName}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; 图标文件
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion

; 版本信息文件
Source: "version.json"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; 开始菜单
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\icon.ico"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"; WorkingDir: "{app}"

; 桌面快捷方式（可选）
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon; IconFilename: "{app}\icon.ico"

[Run]
; 安装完成后运行（可选）
Filename: "{app}\{#MyAppExeName}"; Description: "运行 {#MyAppName}"; Flags: postinstall nowait skipifsilent shellexec

[UninstallRun]
; 卸载时清理 %APPDATA% 中的用户配置
Filename: "{cmd}"; Parameters: "/c rmdir /s /q ""{userappdata}\{#MyAppName}"""; Flags: runhidden; RunOnceId: "CleanAppData"

[Code]
function InitializeSetup: Boolean;
begin
  Result := True;
end;
