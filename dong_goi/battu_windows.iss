; Inno Setup – bộ cài Windows 10/11 (64-bit), có gỡ cài đặt. Dựng: iscc dong_goi\battu_windows.iss
#define TenPM "Bát Tự Thiệu Vĩ Hoa"
#define PhienBan "1.9.7"
[Setup]
AppId={{6C1E3A52-8B7D-4E0B-9C61-3B2F7A9D6E31}
AppName={#TenPM}
AppVersion={#PhienBan} (quy tắc 6.9)
AppVerName={#TenPM} {#PhienBan}
AppPublisher=Bát Tự Thiệu Vĩ Hoa
DefaultDirName={autopf}\BatTuThieuViHoa
DefaultGroupName={#TenPM}
DisableProgramGroupPage=yes
OutputDir=..\dist_cai
OutputBaseFilename=BatTuThieuViHoa_Setup_{#PhienBan}_Windows
SetupIconFile=bieu_tuong.ico
UninstallDisplayIcon={app}\BatTuThieuViHoa.exe
UninstallDisplayName={#TenPM} {#PhienBan}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequiredOverridesAllowed=dialog
MinVersion=10.0
[Languages]
Name: "vi"; MessagesFile: "Vietnamese.isl"
[Tasks]
Name: "desktopicon"; Description: "Tạo biểu tượng trên màn hình nền"; GroupDescription: "Biểu tượng:"
[Files]
Source: "..\dist\BatTuThieuViHoa\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
[Icons]
Name: "{group}\{#TenPM}"; Filename: "{app}\BatTuThieuViHoa.exe"
Name: "{group}\Gỡ cài đặt {#TenPM}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#TenPM}"; Filename: "{app}\BatTuThieuViHoa.exe"; Tasks: desktopicon
[Run]
Filename: "{app}\BatTuThieuViHoa.exe"; Description: "Mở {#TenPM}"; Flags: nowait postinstall skipifsilent
; Dữ liệu lá số người dùng (%APPDATA%\BatTuThieuViHoa\la_so.sqlite) được GIỮ LẠI khi gỡ để không mất lá đã lưu.
