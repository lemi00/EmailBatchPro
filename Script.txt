[Setup]
AppName=EmailBatchPro
AppVersion=1.0
DefaultDirName={autopf}\EmailBatchPro
DefaultGroupName=EmailBatchPro
OutputDir=Output
OutputBaseFilename=EmailBatchPro_Setup
Compression=lzma
SolidCompression=yes
SetupIconFile=EmailBatchPro.ico

[Files]
Source: "dist\EmailBatchPro.exe"; DestDir: "{app}"
Source: "EmailBatchPro.ico"; DestDir: "{app}"

[Dirs]
Name: "{app}\logs"
Name: "{app}\modelos"

[Icons]
Name: "{group}\EmailBatchPro"; Filename: "{app}\EmailBatchPro.exe"
Name: "{autodesktop}\EmailBatchPro"; Filename: "{app}\EmailBatchPro.exe"

[Run]
Filename: "{app}\EmailBatchPro.exe"; Description: "Iniciar EmailBatchPro"; Flags: nowait postinstall skipifsilent