@echo off

rem DBM
rem Dynamic Bandwidth Monitor
rem Leak detection method implemented in a real-time data historian
rem
rem Copyright (C) 2014, 2015, 2016, 2017  J.H. Fitié, Vitens N.V.
rem
rem This file is part of DBM.
rem
rem DBM is free software: you can redistribute it and/or modify
rem it under the terms of the GNU General Public License as published by
rem the Free Software Foundation, either version 3 of the License, or
rem (at your option) any later version.
rem
rem DBM is distributed in the hope that it will be useful,
rem but WITHOUT ANY WARRANTY; without even the implied warranty of
rem MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
rem GNU General Public License for more details.
rem
rem You should have received a copy of the GNU General Public License
rem along with DBM.  If not, see <http://www.gnu.org/licenses/>.

%~d0
cd %~dp0

if not exist build mkdir build

del /Q build\*

set vbc="%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\Vbc.exe" /win32icon:res\dbm.ico /optimize+ /nologo /novbruntimeref
set PICheck="%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll"
set PIRefs=%PICheck%,"%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDKCommon.dll"
set PIACECheck="%PIHOME%\ACE\OSISoft.PIACENet.dll"
set PIACERefs=%PIACECheck%,"%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PITimeServer.dll"

%vbc% /target:library /out:build\DBMDriverNull.dll src\shared\*.vb src\dbm\*.vb src\dbm\driver\DBMDriverNull.vb
if not exist build\DBMDriverNull.dll goto ExitBuild

%vbc% /target:library /out:build\DBMDriverCSV.dll src\shared\*.vb src\dbm\*.vb src\dbm\driver\DBMDriverCSV.vb
if not exist build\DBMDriverCSV.dll goto ExitBuild

%vbc% /reference:build\DBMDriverCSV.dll /out:build\DBMTester.exe src\shared\*.vb src\dbmtester\*.vb
if not exist build\DBMTester.exe goto ExitBuild

if exist %PICheck% (

    %vbc% /reference:%PIRefs% /target:library /out:build\DBMDriverOSIsoftPI.dll src\shared\*.vb src\dbm\*.vb src\dbm\driver\DBMDriverOSIsoftPI.vb
    if not exist build\DBMDriverOSIsoftPI.dll goto ExitBuild

    if exist %PIACECheck% (

        %vbc% /reference:%PIRefs%,%PIACERefs%,build\DBMDriverOSIsoftPI.dll /rootnamespace:PIACE.DBMRt /target:library /out:build\DBMRt.dll src\shared\*.vb src\PIACENet\*.vb
        if not exist build\DBMRt.dll goto ExitBuild

    )

)

build\DBMTester.exe

call docs\calculatesamples.bat

:ExitBuild
pause
