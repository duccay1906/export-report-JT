[Setup]
; Thông tin phần mềm
AppName=JT Export Report
AppVersion=1.1
AppPublisher=Tuyet Trinh Xinh Dep
DefaultDirName={pf}\JT ExportReport
DefaultGroupName=JT Export Report
OutputBaseFilename=setup_v1.1
Compression=lzma
SolidCompression=yes

[Files]
; Copy tất cả file trong folder dist\export sang thư mục cài đặt
Source: "D:\JT_Report\dist\export\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

[Icons]
; Tạo shortcut trong Start Menu và ngoài Desktop
Name: "{group}\JT Export Report"; Filename: "{app}\export.exe"
Name: "{commondesktop}\JT Export Report"; Filename: "{app}\export.exe"

[Run]
; Chạy ứng dụng sau khi cài (tùy chọn)
Filename: "{app}\export.exe"; Description: "Run JT Export Report"; Flags: nowait postinstall skipifsilent
