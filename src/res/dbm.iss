; Dynamic Bandwidth Monitor
; Leak detection method implemented in a real-time data historian
; Copyright (C) 2014-2021  J.H. Fitié, Vitens N.V.
;
; This file is part of DBM.
;
; DBM is free software: you can redistribute it and/or modify
; it under the terms of the GNU General Public License as published by
; the Free Software Foundation, either version 3 of the License, or
; (at your option) any later version.
;
; DBM is distributed in the hope that it will be useful,
; but WITHOUT ANY WARRANTY; without even the implied warranty of
; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
; GNU General Public License for more details.
;
; You should have received a copy of the GNU General Public License
; along with DBM.  If not, see <http://www.gnu.org/licenses/>.

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
Source: "..\..\*"; Excludes: "\.git,\enc,\appveyor-tools"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\Files"; Filename: "{app}"; IconFilename: "{app}\src\res\dbm.ico"
Name: "{group}\About"; Filename: "{cmd}"; Parameters: "/k DBMAbout.exe"; WorkingDir: "{app}\build"; IconFilename: "{app}\src\res\dbm.ico"
Name: "{group}\Uninstall"; Filename: "{uninstallexe}"
