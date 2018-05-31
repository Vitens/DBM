#define Application "Dynamic Bandwidth Monitor"
#define Version GetFileVersion('..\..\build\DBM.dll')
#define Publisher "J.H. Fitié, Vitens N.V."

[Setup]
AppId={#Application}
AppName={#Application}
AppVersion={#Version}
AppPublisher={#Publisher}
AppCopyright=Copyright (C) 2014-2018  {#Publisher}
VersionInfoVersion={#Version}
SetupIconFile=..\..\src\res\dbm.ico
LicenseFile=..\..\LICENSE
WizardImageFile=WizModernImage.bmp
WizardSmallImageFile=WizModernSmallImage.bmp
DefaultDirName={pf}\{#Application}
DisableDirPage=yes
DefaultGroupName={#Application}
DisableProgramGroupPage=yes
DisableReadyPage=yes

[Files]
Source: "..\..\*"; Excludes: "\.git,\build"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\Files"; Filename: "{app}\"
Name: "{group}\Build"; Filename: "{app}\build.bat"
Name: "{group}\About"; Filename: "{app}\build\DBMAbout.exe"

[Run]
Filename: "{app}\build.bat"

[UninstallDelete]
Type: filesandordirs; Name: "{app}\build"
