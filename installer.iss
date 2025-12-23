; Inno Setup 脚本 - 送货单对账单生成工具
; 需要先用PyInstaller打包成EXE后再运行此脚本

#define MyAppName "送货单对账单工具"
#define MyAppVersion "1.0"
#define MyAppPublisher "百惠行"
#define MyAppExeName "送货单对账单工具.exe"

[Setup]
; 基本信息
AppId={{8F2B3C4D-5E6F-7A8B-9C0D-1E2F3A4B5C6D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=installer_output
OutputBaseFilename=送货单对账单工具_Setup_v{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern

; 权限设置
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

; 许可协议（可选）
;LicenseFile=LICENSE.txt

; 信息文件（可选）
;InfoBeforeFile=README.txt

; 安装界面设置
DisableProgramGroupPage=yes
DisableWelcomePage=no

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; 主程序
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; 说明文档
Source: "使用说明.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion
; 示例文件夹
Source: "raw-data\*"; DestDir: "{app}\raw-data"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; 开始菜单图标
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\使用说明"; Filename: "{app}\使用说明.txt"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"
; 桌面图标
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; 安装完成后运行
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 卸载时删除生成的文件
Type: filesandordirs; Name: "{app}\output"

[Code]
// 安装前检查
function InitializeSetup(): Boolean;
begin
  Result := True;
  if MsgBox('欢迎安装 送货单对账单工具！' + #13#10#13#10 +
            '此工具用于自动生成客户对账单。' + #13#10#13#10 +
            '点击"是"继续安装，点击"否"取消安装。',
            mbConfirmation, MB_YESNO) = IDNO then
  begin
    Result := False;
  end;
end;

// 安装完成后提示
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    // 创建output文件夹
    CreateDir(ExpandConstant('{app}\output'));
  end;
end;
