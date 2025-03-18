; CyberTrackerInstaller.iss

[Setup]
AppName=CyberTracker
AppVersion=1.0
; 改成 C:\CyberTracker
DefaultDirName=C:\CyberTracker
DefaultGroupName=CyberTracker
OutputBaseFilename=CyberTrackerInstaller
Compression=lzma
SolidCompression=yes
; 設定為最低權限 (若程式不需要寫入系統敏感區域)
PrivilegesRequired=admin

[Files]
Source: "C:\Users\<inputYourUserName>\CyberTracker\CyberTracker.exe"; DestDir: "{app}"
Source: "C:\Users\<inputYourUserName>\CyberTracker\merge_csv.exe";   DestDir: "{app}"
Source: "C:\Users\<inputYourUserName>\CyberTracker\web_capture.exe";   DestDir: "{app}"
Source: "C:\Users\<inputYourUserName>\CyberTracker\csv_to_xlsx.exe";   DestDir: "{app}"

[Dirs]
Name: "{app}\all_csv"

[Icons]
Name: "{commondesktop}\CyberTracker"; Filename: "{app}\CyberTracker.exe"
Name: "{group}\CyberTracker"; Filename: "{app}\CyberTracker.exe"

[Run]
Filename: "{app}\CyberTracker.exe"; Description: "立即啟動 CyberTracker"; Flags: nowait postinstall skipifsilent
