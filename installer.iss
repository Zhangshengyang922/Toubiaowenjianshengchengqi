; ═══════════════════════════════════════════════════════
;  Inno Setup 安装脚本 —— 生成专业 Windows 安装包
;
;  使用方法:
;    1. 安装 Inno Setup: https://jrsoftware.org/isdl.php
;    2. 先用 build.py 打包出 .exe
;    3. 修改下方 #define 中的版本号和路径
;    4. 在 Inno Setup Compiler 中打开此文件，点击 Compile
;
;  注意: 将 icon.ico 放到项目根目录作为安装包图标
; ═══════════════════════════════════════════════════════

#define MyAppName "招投标文件自动生成系统"
#define MyAppNameEn "BiddingDocGen"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "ZhangShengYang"
#define MyAppURL "https://github.com/Zhangshengyang922/Toubiaowenjianshengchengqi"
#define MyAppExeName "BiddingDocGen_v1.0.0.exe"
; 以下路径请根据实际情况修改
#define MySourceDir "."

[Setup]
; 安装包基本信息
AppId={{B5D8E3F1-9A2C-4D6E-8F01-3C7A9B5D2E4F}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}

; 默认安装目录
DefaultDirName={autopf}\{#MyAppNameEn}

; 开始菜单文件夹
DefaultGroupName={#MyAppName}

; 是否允许用户修改安装目录
DisableDirPage=no

; 安装包输出
OutputDir={#MySourceDir}\dist
OutputBaseFilename={#MyAppNameEn}_Setup_v{#MyAppVersion}

; 压缩方式
Compression=lzma2/ultra64
SolidCompression=yes

; 安装包图标
SetupIconFile={#MySourceDir}\icon.ico

; 管理员权限（安装到 Program Files 需要）
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog

; 界面设置
WizardStyle=modern
WizardSizePercent=120
WindowVisible=yes
WindowShowCaption=yes

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "其他:"; Flags: unchecked

[Files]
; 主程序
Source: "{#MySourceDir}\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

; 默认数据目录（空模板）
Source: "{#MySourceDir}\data\*"; DestDir: "{app}\data"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#MySourceDir}\templates\*"; DestDir: "{app}\templates"; Flags: ignoreversion recursesubdirs createallsubdirs

; 创建默认 input/output 目录
[Dirs]
Name: "{app}\input"
Name: "{app}\output"

[Icons]
; 开始菜单快捷方式
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
; 桌面快捷方式
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon
; 卸载快捷方式
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"

[Run]
; 安装完成后运行程序
Filename: "{app}\{#MyAppExeName}"; Description: "启动 {#MyAppName}"; Flags: nowait postinstall skipifsilent; WorkingDir: "{app}"

[Code]
// 检查是否已安装旧版本
function InitializeSetup(): Boolean;
var
  OldVersion: string;
begin
  Result := True;
  if RegQueryStringValue(HKEY_LOCAL_MACHINE,
    'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#MyAppNameEn}_is1',
    'DisplayVersion', OldVersion) then
  begin
    if MsgBox('检测到已安装 v' + OldVersion + #13#10 +
              '是否覆盖安装为新版本 v{#MyAppVersion}？',
              mbConfirmation, MB_YESNO) = IDNO then
    begin
      Result := False;
    end;
  end;
end;
