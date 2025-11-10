; =============== Setup ===============
[Setup]
AppId={{3EAA1E9C-85E7-4E70-9E36-3A2E8E0D6E51}}
AppName=MinhTam
AppVersion=1.0.0
AppPublisher=MinhTam
DefaultDirName={pf}\MinhTam Tool
DefaultGroupName=MinhTam Tool
OutputDir=C:\Users\longn\OneDrive\Desktop\Tool_Check
OutputBaseFilename=MinhTamTool-Setup
Compression=lzma
SolidCompression=yes
; SetupIconFile=assets\app.ico
UninstallDisplayIcon={app}\FindJapanCharacters.exe
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
VersionInfoVersion=1.0.0

; =============== Files ===============
[Files]
Source: "assets\app.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\longn\source\repos\FindJapanCharacters\FindJapanCharacters\bin\Release\net9.0-windows\FindJapanCharacters.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\longn\source\repos\FindJapanCharacters\FindJapanCharacters\bin\Release\net9.0-windows\*"; \
    Excludes: "FindJapanCharacters.exe"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; =============== Shortcuts ===============
[Icons]
Name: "{commondesktop}\MinhTam Tool"; Filename: "{app}\FindJapanCharacters.exe"; IconFilename: "{app}\app.ico"

; =============== Post-install ===============
[Run]
; 1) Tải & cài .NET 9 Desktop Runtime nếu thiếu (silent)
Filename: "{tmp}\dotnet-runtime-installer.exe"; \
    Parameters: "/install /quiet /norestart"; \
    StatusMsg: "Đang tải và cài .NET Desktop Runtime 9..."; \
    Check: not IsDotNet9DesktopInstalled; \
    Flags: runhidden

; 2) Chạy ứng dụng
Filename: "{app}\FindJapanCharacters.exe"; Description: "Chạy MinhTam Tool"; Flags: nowait postinstall skipifsilent

; =============== Code ===============
[Code]
const
  DotNet9DesktopURL = 'https://builds.dotnet.microsoft.com/dotnet/WindowsDesktop/9.0.10/windowsdesktop-runtime-9.0.10-win-x64.exe';

function IsDotNet9DesktopInstalled: Boolean;
var
  FR: TFindRec;
  SharedDir: string;
begin
  Result := False;
  SharedDir := ExpandConstant('{pf}') + '\dotnet\shared\Microsoft.WindowsDesktop.App';
  if not DirExists(SharedDir) then
    exit;

  if FindFirst(SharedDir + '\*', FR) then
  begin
    try
      repeat
        { Chỉ lấy thư mục, tên bắt đầu bằng "9." }
        if ((FR.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0) and (Pos('9.', FR.Name) = 1) then
        begin
          Result := True;
          break;
        end;
      until not FindNext(FR);
    finally
      FindClose(FR);
    end;
  end;
end;

function DownloadTemporaryFile(const URL, DestPath: string): Boolean;
var
  WinHttp: Variant;
  Stream: Variant;
begin
  Result := False;
  try
    WinHttp := CreateOleObject('WinHttp.WinHttpRequest.5.1');
    WinHttp.Open('GET', URL, False);
    WinHttp.Send();
    if WinHttp.Status <> 200 then exit;

    Stream := CreateOleObject('ADODB.Stream');
    Stream.Type_ := 1;   { binary }
    Stream.Open();
    Stream.Write(WinHttp.ResponseBody);
    Stream.SaveToFile(DestPath, 2);  { adSaveCreateOverWrite }
    Stream.Close();

    Result := True;
  except
    Result := False;
  end;
end;

function InitializeSetup(): Boolean;
var
  NetRuntimePath: string;
begin
  Result := True;
  if not IsDotNet9DesktopInstalled then
  begin
    NetRuntimePath := ExpandConstant('{tmp}\dotnet-runtime-installer.exe');
    MsgBox('Ứng dụng cần .NET 9 Desktop Runtime. Trình cài sẽ tải tự động từ Microsoft.', mbInformation, MB_OK);
    if not DownloadTemporaryFile(DotNet9DesktopURL, NetRuntimePath) then
    begin
      MsgBox('Không thể tải .NET 9 Desktop Runtime. Kiểm tra kết nối mạng rồi chạy lại.', mbCriticalError, MB_OK);
      Result := False;  { hủy setup nếu không tải được }
    end;
  end;
end;
