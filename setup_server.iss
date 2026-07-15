; Script cai dat cho MAY CHU (goi rieng, chay Windows Service - xem
; README.md muc "May chu chia se mang LAN"). TACH RIENG voi setup.iss
; (may tram) - 2 file cai dat doc lap, dong phien ban rieng.
; Bien MyAppVersion duoc truyen tu ngoai vao khi build:
;   ISCC.exe /DMyAppVersion=0.1.0 setup_server.iss
; Neu khong truyen, mac dinh lay "0.0.0-dev".
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0-dev"
#endif

#define MyAppName "May chu chia se du lieu - Quan ly benh nhan THA"
#define MyServiceExeName "QuanLyBenhNhanTHA-Service.exe"
#define MyTrayExeName "QuanLyBenhNhanTHA-Tray.exe"
#define MyAppPublisher "Monsterph6"

[Setup]
AppId={{B6C9B6B0-6E7E-4E9B-9B1B-QLBNTHASRV001}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
; Can quyen Administrator de cai/bat Windows Service - khac voi setup.iss
; (may tram) khong can quyen Admin.
DefaultDirName={autopf}\QuanLyBenhNhanTHA-Server
DefaultGroupName=Quan ly benh nhan THA - May chu
DisableProgramGroupPage=yes
PrivilegesRequired=admin
OutputDir=setup_output
OutputBaseFilename=QuanLyBenhNhanTHA-Server-Setup-{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyServiceExeName}
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; dist_server\ (tao boi build_server.bat) da gom san Service+Tray+
; install_server.bat/uninstall_server.bat/update_server.bat/
; update_server.ps1/VERSION_SERVER.txt/README.md - khong can liet ke rieng.
Source: "dist_server\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Dirs]
; Thu muc cai dat can duoc ghi duoc (luu benh_nhan.db, lan_config.json,
; backups\, update_token.txt, ...)
Name: "{app}"; Permissions: users-modify

[Icons]
Name: "{group}\May chu - Xem trạng thái (khay hệ thống)"; Filename: "{app}\{#MyTrayExeName}"
Name: "{group}\Kiểm tra cập nhật máy chủ"; Filename: "{app}\update_server.bat"
Name: "{group}\Gỡ cài đặt"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\{#MyTrayExeName}"; Description: "Mở bảng điều khiển (xem địa chỉ IP:cổng để cung cấp cho máy trạm)"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "{app}\{#MyServiceExeName}"; Parameters: "stop"; Flags: runhidden waituntilterminated; RunOnceId: "StopService"
Filename: "{app}\{#MyServiceExeName}"; Parameters: "remove"; Flags: runhidden waituntilterminated; RunOnceId: "RemoveService"

[Code]
var
  PortPage: TInputQueryWizardPage;

procedure InitializeWizard;
begin
  PortPage := CreateInputQueryPage(wpSelectDir,
    'Cổng chia sẻ qua mạng LAN',
    'Các máy trạm trong mạng sẽ kết nối tới máy này qua cổng dưới đây',
    'Mặc định là 8765 - chỉ cần đổi nếu cổng này đang được dùng cho việc khác ' +
    'trong mạng, hoặc quản trị mạng yêu cầu dùng cổng khác. Ghi lại địa chỉ IP ' +
    'của máy này (xem ở bước cuối, hoặc menu chuột phải icon khay hệ thống sau ' +
    'khi cài) để nhập vào các máy trạm.');
  PortPage.Add('Cổng (mặc định 8765):', False);
  PortPage.Values[0] := '8765';
end;

function GetLanConfigJson(): String;
var
  Port: String;
begin
  Port := Trim(PortPage.Values[0]);
  if Port = '' then Port := '8765';
  Result := '{"port": ' + Port + '}';
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  VersionFile: String;
  ConfigFile: String;
  ResultCode: Integer;
begin
  if CurStep = ssPostInstall then
  begin
    VersionFile := ExpandConstant('{app}\VERSION_SERVER.txt');
    SaveStringToFile(VersionFile, '{#MyAppVersion}', False);

    ConfigFile := ExpandConstant('{app}\lan_config.json');
    if not FileExists(ConfigFile) then
      SaveStringToFile(ConfigFile, GetLanConfigJson(), False);

    Exec(ExpandConstant('{app}\{#MyServiceExeName}'), '--startup auto install', '',
      SW_HIDE, ewWaitUntilTerminated, ResultCode);
    Exec(ExpandConstant('{app}\{#MyServiceExeName}'), 'start', '',
      SW_HIDE, ewWaitUntilTerminated, ResultCode);
  end;
end;
