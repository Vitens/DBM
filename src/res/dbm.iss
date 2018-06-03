#define Application "Dynamic Bandwidth Monitor"
#define Version GetFileVersion("..\..\build\DBM.dll")
#define Company GetFileCompany("..\..\build\DBM.dll")
#define Copyright GetFileCopyright("..\..\build\DBM.dll")

[Setup]
AppId={#Application}
AppName={#Application}
AppVersion={#Version}
AppPublisher={#Company}
AppCopyright={#Copyright}
VersionInfoVersion={#Version}
SetupIconFile=..\..\src\res\dbm.ico
WizardImageFile=WizModernImage.bmp
WizardSmallImageFile=WizModernSmallImage.bmp
LicenseFile=..\..\LICENSE
DefaultDirName={pf}\{#Application}
DisableDirPage=yes
DefaultGroupName={#Application}
DisableProgramGroupPage=yes
DisableReadyPage=yes

[Files]
Source: "..\..\*"; Excludes: "\.git,\build"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\Files"; Filename: "{app}\"; IconFilename: "{app}\src\res\dbm.ico"
Name: "{group}\Build"; Filename: "{app}\build.bat"; IconFilename: "{app}\src\res\dbm.ico"
Name: "{group}\About"; Filename: "{app}\build\DBMAbout.exe"
Name: "{group}\Uninstall"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\build.bat"; StatusMsg: "Building {#Application} v{#Version}…"; Flags: runhidden

[UninstallDelete]
Type: filesandordirs; Name: "{app}\build"
