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

rem Variables
set vbc="%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\Vbc.exe" /win32icon:res\dbm.ico /optimize+ /nologo /novbruntimeref
set PIHOME=%CD%\3rdParty\PILibraries\pidll
set PICheck="%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll"
set PIRefs=%PICheck%,"%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDKCommon.dll"
set PIACECheck="%PIHOME%\ACE\OSISoft.PIACENet.dll"
set PIACERefs=%PIACECheck%,"%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PITimeServer.dll"

rem Set up build directory
if not exist build mkdir build
del /Q build\*
copy LICENSE build > NUL

rem Apply patches
for /f "delims=" %%i in ('git rev-parse --short HEAD') do set commit=%%i
powershell -Command "(Get-Content src\dbm\DBM.vb) -replace 'Const GITHASH As String = \".*?\"', 'Const GITHASH As String = \"%commit%\"' | Set-Content src\dbm\DBM.vb"

Rem Build
%vbc% /target:library /out:build\DBM.dll src\shared\*.vb src\dbm\*.vb
%vbc% /reference:build\DBM.dll /target:library /out:build\DBMDriverCSV.dll src\shared\*.vb src\dbm\driver\DBMDriverCSV.vb
%vbc% /reference:build\DBM.dll,build\DBMDriverCSV.dll /out:build\DBMTester.exe src\shared\*.vb src\dbmtester\*.vb
if exist %PICheck% (
  %vbc% /reference:%PIRefs%,build\DBM.dll /target:library /out:build\DBMDriverOSIsoftPI.dll src\shared\*.vb src\dbm\driver\DBMDriverOSIsoftPI.vb
  if exist %PIACECheck% (
    %vbc% /reference:%PIRefs%,%PIACERefs%,build\DBM.dll,build\DBMDriverOSIsoftPI.dll /target:library /out:build\DBMRt.dll src\shared\*.vb src\PIACENet\*.vb
  )
)

rem Output version, copyright and license information and unit and integration test results
build\DBMTester.exe

:ExitBuild
