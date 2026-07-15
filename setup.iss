; Script cai dat cho Quan ly & Loc trung danh sach benh nhan THA.
; Bien MyAppVersion duoc truyen tu ngoai vao khi build:
;   ISCC.exe /DMyAppVersion=1.0.0 setup.iss
; Neu khong truyen, mac dinh lay "0.0.0-dev".
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0-dev"
#endif

#define MyAppName "Quan ly & Loc trung danh sach benh nhan THA"
#define MyAppExeName "QuanLyBenhNhanTHA.exe"
#define MyAppPublisher "Monsterph6"

[Setup]
AppId={{B6C9B6B0-6E7E-4E9B-9B1B-QLBNTHA00001}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
; Cai vao thu muc rieng cua user (khong can quyen Admin, ghi du lieu benh_nhan.db
; ngay ben canh file .exe nen khong the cai vao Program Files thong thuong).
DefaultDirName={localappdata}\Programs\QuanLyBenhNhanTHA
DefaultGroupName=Quan ly benh nhan THA
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
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
Source: "dist\QuanLyBenhNhanTHA\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "update.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "update.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
; Thu muc cai dat can duoc ghi duoc (luu benh_nhan.db, update_token.txt, ...)
Name: "{app}"; Permissions: users-modify

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Kiem tra cap nhat"; Filename: "{app}\update.bat"
Name: "{group}\Go cai dat"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Chay {#MyAppName} ngay"; Flags: nowait postinstall skipifsilent

[Code]
var
  ServerPage: TInputQueryWizardPage;

procedure InitializeWizard;
begin
  ServerPage := CreateInputQueryPage(wpSelectDir,
    'Kết nối máy chủ chia sẻ mạng LAN (tuỳ chọn)',
    'Chỉ điền nếu máy này dùng chung dữ liệu với 1 máy chủ khác trong mạng nội bộ',
    'Nếu máy này chỉ dùng ĐỘC LẬP (không chia sẻ qua mạng), để trống ô bên dưới ' +
    'rồi bấm Next.'#13#10#13#10 +
    'Nếu đã có máy chủ chia sẻ (xem gói cài đặt "QuanLyBenhNhanTHA-Server-Setup") ' +
    'đang chạy trong cùng mạng LAN, nhập địa chỉ IP:cổng của máy chủ đó (ví dụ ' +
    '192.168.1.10:8765) - có thể xem/đổi lại sau trong tab "Mạng LAN" của ứng dụng.');
  ServerPage.Add('Địa chỉ máy chủ (để trống nếu dùng 1 máy):', False);
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
begin
  if CurStep = ssPostInstall then
  begin
    VersionFile := ExpandConstant('{app}\VERSION.txt');
    SaveStringToFile(VersionFile, '{#MyAppVersion}', False);

    ServerAddr := Trim(ServerPage.Values[0]);
    if ServerAddr <> '' then
    begin
      ConfigFile := ExpandConstant('{app}\lan_config.json');
      SaveStringToFile(ConfigFile,
        '{"role": "client", "server_url": "' + BuildServerUrl(ServerAddr) + '"}', False);
    end;
  end;
end;
