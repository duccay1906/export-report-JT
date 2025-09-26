[Setup]
; Thông tin phần mềm
AppName=JT Export Report
AppVersion=1.2
AppPublisher=Tuyet Trinh Xinh Dep
DefaultDirName={pf}\JT ExportReport
DefaultGroupName=JT Export Report
OutputBaseFilename=setup_v1.2
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

[Code]
var
  UpdatePage: TOutputMsgWizardPage;

procedure InitializeWizard;
begin
  { Tạo trang mới ngay sau trang Welcome }
  UpdatePage := CreateOutputMsgPage(
    wpWelcome,
    'Thông tin cập nhật',
    'Những tính năng mới trong phiên bản 1.2:',
    '• Thêm xuất báo cáo tự động theo ngày' + #13#10 +
    '• Fix lỗi DIV/0 trong công thức Excel' + #13#10 +
    '• Tối ưu tốc độ xử lý dữ liệu lớn' + #13#10 +
    '• Giao diện bảng Summary rõ ràng, dễ đọc hơn'
  );
end;