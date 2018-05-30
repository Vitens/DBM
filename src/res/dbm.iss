; iscc /O"%TEMP%" /Fdbmsetup /Qp -
; "%ProgramFiles%\Inno Setup 5\ISCC.exe" /O"%TEMP%" /Fdbmsetup src\res\dbm.iss && %TEMP%\dbmsetup.exe
#define MyAppName "Dynamic Bandwidth Monitor"
#define MyAppVersion GetFileVersion('..\..\build\DBM.dll')
[Setup]
AppId={#MyAppName}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=J.H. Fitié, Vitens N.V.
SetupIconFile=..\..\src\res\dbm.ico
LicenseFile=..\..\LICENSE
DefaultDirName={pf}\{#MyAppName}
DisableDirPage=yes
DefaultGroupName={#MyAppName}
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
