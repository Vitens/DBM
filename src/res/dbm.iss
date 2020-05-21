#define Product GetEnv("product")
#define Version GetEnv("version")
#define Commit GetEnv("commit")
#define Company GetEnv("company")
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
Source: "..\..\*"; Excludes: "\.git"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\Files"; Filename: "{app}"; IconFilename: "{app}\src\res\dbm.ico"
Name: "{group}\About"; Filename: "{cmd}"; Parameters: "/k DBMAbout.exe"; WorkingDir: "{app}\build"; IconFilename: "{app}\src\res\dbm.ico"
Name: "{group}\Uninstall"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\registerpiafdr.bat"; StatusMsg: "Registering PI AF Data Reference on AF server ..."; Flags: runhidden

[UninstallDelete]
Type: files; Name: "{app}\samples\sample?.csv"
