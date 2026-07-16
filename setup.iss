; Script cai dat cho Quan ly & Loc trung danh sach benh nhan THA.
; Mot file cai dat duy nhat cho ca ung dung chinh lan thanh phan May chu
; (Windows Service chia se qua mang LAN) - luc cai dat se hoi chon vai tro:
; Mot may / May tram / May chu. Bien MyAppVersion duoc truyen tu ngoai vao
; khi build:
;   ISCC.exe /DMyAppVersion=1.0.0 setup.iss
; Neu khong truyen, mac dinh lay "0.0.0-dev".
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0-dev"
#endif

#define MyAppName "Quan ly & Loc trung danh sach benh nhan THA"
#define MyAppExeName "QuanLyBenhNhanTHA.exe"
#define MyServiceExeName "QuanLyBenhNhanTHA-Service.exe"
#define MyTrayExeName "QuanLyBenhNhanTHA-Tray.exe"
#define MyServiceName "QuanLyBenhNhanTHA_Server"
#define MyAppPublisher "Monsterph6"

[Setup]
AppId={{B6C9B6B0-6E7E-4E9B-9B1B-QLBNTHA00001}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
; Luon can quyen Administrator: neu nguoi dung chon vai tro "May chu" thi
; trinh cai dat phai cai/bat duoc Windows Service ngay trong buoc nay -
; ap dung chung cho ca 3 vai tro de chi co 1 file cai dat duy nhat (khong
; phai hoi quyen lai tu dau neu sau nay doi sang vai tro May chu).
DefaultDirName={autopf}\QuanLyBenhNhanTHA
DefaultGroupName=Quan ly benh nhan THA
DisableProgramGroupPage=yes
PrivilegesRequired=admin
OutputDir=setup_output
OutputBaseFilename=QuanLyBenhNhanTHA-Setup-{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Tao shortcut ngoai Desktop"; GroupDescription: "Shortcut bo sung:"; Flags: unchecked

[Files]
; dist\QuanLyBenhNhanTHA\ (tao boi build.bat) da gom san ung dung chinh +
; QuanLyBenhNhanTHA-Service.exe + QuanLyBenhNhanTHA-Tray.exe - khong can
; liet ke rieng, luon cai du ca 3 (chi kich hoat Windows Service neu vai
; tro duoc chon la "May chu").
Source: "dist\QuanLyBenhNhanTHA\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "update.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "update.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
; Thu muc cai dat can duoc ghi duoc (luu benh_nhan.db, lan_config.json,
; backups\, update_token.txt, ...) ke ca khi cai vao Program Files va
; ung dung sau nay chay boi 1 nguoi dung thuong (khong phai Admin).
Name: "{app}"; Permissions: users-modify

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\May chu - Xem trạng thái (khay hệ thống)"; Filename: "{app}\{#MyTrayExeName}"
Name: "{group}\Kiem tra cap nhat"; Filename: "{app}\update.bat"
Name: "{group}\Go cai dat"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Chay {#MyAppName} ngay"; Flags: nowait postinstall skipifsilent
Filename: "{app}\{#MyTrayExeName}"; Description: "Mở bảng điều khiển Máy chủ (xem địa chỉ IP:cổng để cung cấp cho máy trạm)"; Flags: nowait postinstall skipifsilent; Check: IsServerRole

[UninstallRun]
; Chi co tac dung neu Windows Service da tung duoc cai (vai tro May chu) -
; ServiceIsInstalled tu kiem tra qua "sc query", an toan goi ngay ca khi
; chua tung cai dich vu.
Filename: "{app}\{#MyServiceExeName}"; Parameters: "stop"; Flags: runhidden waituntilterminated; RunOnceId: "StopService"; Check: ServiceIsInstalled
Filename: "{app}\{#MyServiceExeName}"; Parameters: "remove"; Flags: runhidden waituntilterminated; RunOnceId: "RemoveService"; Check: ServiceIsInstalled

[Code]
const
  ROLE_SINGLE = 0;
  ROLE_CLIENT = 1;
  ROLE_SERVER = 2;

var
  RolePage: TInputOptionWizardPage;
  ClientPage: TInputQueryWizardPage;
  ServerPage: TInputQueryWizardPage;

procedure InitializeWizard;
begin
  RolePage := CreateInputOptionPage(wpSelectDir,
    'Vai trò của máy này',
    'Chọn cách máy này sẽ dùng dữ liệu bệnh nhân',
    'Có thể xem/đổi lại địa chỉ máy chủ (vai trò Máy trạm) sau trong tab ' +
    '"Mạng LAN" của ứng dụng. Riêng vai trò "Máy chủ" gắn với việc cài/gỡ ' +
    'Windows Service ngay trong bước cài đặt này - muốn đổi sang/khỏi vai ' +
    'trò Máy chủ sau này thì chạy lại trình cài đặt.',
    True, False);
  RolePage.Add('Một máy - dùng độc lập, không chia sẻ qua mạng (mặc định)');
  RolePage.Add('Máy trạm - kết nối tới 1 máy chủ khác đã có sẵn trong mạng LAN');
  RolePage.Add('Máy chủ - chia sẻ dữ liệu của máy này cho các máy khác trong mạng LAN');
  RolePage.SelectedValueIndex := ROLE_SINGLE;

  ClientPage := CreateInputQueryPage(RolePage.ID,
    'Kết nối máy chủ chia sẻ mạng LAN',
    'Nhập địa chỉ máy chủ đã có sẵn trong mạng nội bộ',
    'Ví dụ 192.168.1.10:8765 - nếu chưa biết chính xác ngay bây giờ, cứ để ' +
    'trống rồi bấm Next, sau đó nhập/sửa lại trong tab "Mạng LAN" của ứng dụng.');
  ClientPage.Add('Địa chỉ máy chủ:', False);

  ServerPage := CreateInputQueryPage(RolePage.ID,
    'Cổng chia sẻ qua mạng LAN',
    'Các máy trạm trong mạng sẽ kết nối tới máy này qua cổng dưới đây',
    'Mặc định là 8765 - chỉ cần đổi nếu cổng này đang được dùng cho việc khác ' +
    'trong mạng, hoặc quản trị mạng yêu cầu dùng cổng khác. Sau khi cài xong, ' +
    'xem địa chỉ IP của máy này qua menu chuột phải icon khay hệ thống để ' +
    'cung cấp cho các máy trạm.');
  ServerPage.Add('Cổng (mặc định 8765):', False);
  ServerPage.Values[0] := '8765';
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := False;
  if PageID = ClientPage.ID then
    Result := RolePage.SelectedValueIndex <> ROLE_CLIENT;
  if PageID = ServerPage.ID then
    Result := RolePage.SelectedValueIndex <> ROLE_SERVER;
end;

function IsServerRole(): Boolean;
begin
  Result := RolePage.SelectedValueIndex = ROLE_SERVER;
end;

function ServiceIsInstalled(): Boolean;
var
  ResultCode: Integer;
begin
  Result := Exec('sc.exe', 'query "{#MyServiceName}"', '', SW_HIDE,
    ewWaitUntilTerminated, ResultCode) and (ResultCode = 0);
end;

function BuildServerUrl(Addr: String): String;
begin
  if (Pos('http://', Addr) = 1) or (Pos('https://', Addr) = 1) then
    Result := Addr
  else
    Result := 'http://' + Addr;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  VersionFile: String;
  ConfigFile: String;
  ServerAddr: String;
  Port: String;
  ResultCode: Integer;
begin
  if CurStep = ssPostInstall then
  begin
    VersionFile := ExpandConstant('{app}\VERSION.txt');
    SaveStringToFile(VersionFile, '{#MyAppVersion}', False);

    ConfigFile := ExpandConstant('{app}\lan_config.json');
    case RolePage.SelectedValueIndex of
      ROLE_CLIENT:
        begin
          ServerAddr := Trim(ClientPage.Values[0]);
          if ServerAddr <> '' then
            SaveStringToFile(ConfigFile,
              '{"role": "client", "server_url": "' + BuildServerUrl(ServerAddr) + '"}', False);
        end;
      ROLE_SERVER:
        begin
          // Vai tro May chu: ung dung chinh tren may nay van doc/ghi truc
          // tiep benh_nhan.db cuc bo nhu 1 may don le ("role": "single") -
          // viec chia se qua mang la do Windows Service (doc lap) dam
          // nhiem, dua vao "port" ben duoi (xem service.py / core.py).
          Port := Trim(ServerPage.Values[0]);
          if Port = '' then Port := '8765';
          SaveStringToFile(ConfigFile, '{"role": "single", "port": ' + Port + '}', False);

          Exec(ExpandConstant('{app}\{#MyServiceExeName}'), '--startup auto install', '',
            SW_HIDE, ewWaitUntilTerminated, ResultCode);
          Exec(ExpandConstant('{app}\{#MyServiceExeName}'), 'start', '',
            SW_HIDE, ewWaitUntilTerminated, ResultCode);
        end;
    end;
  end;
end;
