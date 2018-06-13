#define Product GetEnv("product")
#define Version GetEnv("version")
#define Commit GetEnv("commit")
#define Company GetFileCompany("..\..\build\DBM.dll")
#define Copyright GetFileCopyright("..\..\build\DBM.dll")

[Setup]
AppId={#Product}
AppName={#Product}
AppVersion={#Version}+{#Commit}
AppVerName={#Product} v{#SetupSetting("AppVersion")}
AppPublisher={#Company}
AppCopyright={#Copyright}
VersionInfoVersion={#Version}
SetupIconFile=dbm.ico
WizardImageFile=WizModernImage.bmp
WizardSmallImageFile=WizModernSmallImage.bmp
LicenseFile=..\..\LICENSE
DefaultDirName={pf}\{#Company}\{#Product}
DisableDirPage=yes
DefaultGroupName={#Company}\{#Product}
DisableProgramGroupPage=yes
DisableReadyPage=yes
UninstallDisplayIcon={app}\src\res\dbm.ico

[Files]
Source: "..\..\*"; Excludes: "\.git,\build"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\Files"; Filename: "{app}"; IconFilename: "{app}\src\res\dbm.ico"
Name: "{group}\Build"; Filename: "{cmd}"; Parameters: "/k build.bat"; WorkingDir: "{app}"; IconFilename: "{app}\src\res\dbm.ico"
Name: "{group}\About"; Filename: "{cmd}"; Parameters: "/k DBMAbout.exe"; WorkingDir: "{app}\build"; IconFilename: "{app}\src\res\dbm.ico"
Name: "{group}\Uninstall"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\build.bat"; StatusMsg: "Building {#Product}..."; Flags: runhidden

[UninstallDelete]
Type: filesandordirs; Name: "{app}\build"
Type: filesandordirs; Name: "{app}\samples\sample?.csv"
