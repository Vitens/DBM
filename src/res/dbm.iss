#define AppName "Dynamic Bandwidth Monitor"
#define AppVersion GetFileVersion('..\..\build\DBM.dll')
[Setup]
AppId={#AppName}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=J.H. Fitié, Vitens N.V.
SetupIconFile=..\..\src\res\dbm.ico
LicenseFile=..\..\LICENSE
DefaultDirName={pf}\{#AppName}
DisableDirPage=yes
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
DisableReadyPage=yes
[Files]
Source: "..\..\*"; Excludes: "\.git,\build"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs
[Icons]
Name: "{group}\About"; Filename: "{app}\build\DBMAbout.exe"
Name: "{group}\Build"; Filename: "{app}\build.bat"
Name: "{group}\Files"; Filename: "{app}\"
[Run]
Filename: "{app}\build.bat"
[UninstallDelete]
Type: filesandordirs; Name: "{app}\build"
