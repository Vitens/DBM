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

set vbc="%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\Vbc.exe" /win32icon:res\dbm.ico /optimize+ /nologo /novbruntimeref
if not defined PIHOME set PIHOME=%CD%\3rdParty\PILibraries
set PIRefs="%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll","%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDKCommon.dll"
set PIAFDir=%PIHOME%\AF
set PIAFRefs="%PIAFDir%\PublicAssemblies\4.0\OSIsoft.AFSDK.dll"
set PIACERefs="%PIHOME%\ACE\OSISoft.PIACENet.dll","%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PITimeServer.dll"

if not exist build mkdir build
del /Q build\*
copy LICENSE build > NUL
copy gpl-v3-nl-101.pdf build > NUL

if "%CI%" == "True" for /f "delims=" %%i in ('git rev-parse --short HEAD') do set commit=%%i
if "%CI%" == "True" powershell -Command "(Get-Content src\dbm\DBM.vb) -replace 'Const GITHASH As String = \".*?\"', 'Const GITHASH As String = \"%commit%\"' | Set-Content src\dbm\DBM.vb"

%vbc% /target:library /out:build\DBM.dll src\shared\*.vb src\dbm\*.vb

%vbc% /reference:build\DBM.dll /target:library /out:build\DBMPointDriverCSV.dll src\shared\*.vb src\dbm\driver\DBMPointDriverCSV.vb
%vbc% /reference:%PIRefs%,build\DBM.dll /target:library /out:build\DBMPointDriverOSIsoftPI.dll src\shared\*.vb src\dbm\driver\DBMPointDriverOSIsoftPI.vb
%vbc% /reference:%PIAFRefs%,build\DBM.dll /target:library /out:build\DBMPointDriverOSIsoftPIAF.dll src\shared\*.vb src\dbm\driver\DBMPointDriverOSIsoftPIAF.vb

%vbc% /reference:build\DBM.dll,build\DBMPointDriverCSV.dll /out:build\DBMTester.exe src\shared\*.vb src\dbmtester\*.vb

%vbc% /reference:%PIRefs%,%PIACERefs%,build\DBM.dll,build\DBMPointDriverOSIsoftPI.dll /target:library /out:build\DBMRt.dll src\shared\*.vb src\PIACENet\*.vb
%vbc% /reference:%PIAFRefs%,build\DBM.dll,build\DBMPointDriverOSIsoftPIAF.dll /target:library /out:build\DBMDataRef.dll src\shared\*.vb src\PIAFDataRef\*.vb

if exist "%PIAFDir%\regplugin.exe" (
 if exist "%PIAFDir%\DBMDataRef.dll" "%PIAFDir%\regplugin.exe" /Unregister "%PIAFDir%\DBMDataRef.dll"
 copy build\DBMDataRef.dll "%PIAFDir%"
 copy build\DBM.dll "%PIAFDir%"
 copy build\DBMPointDriverOSIsoftPIAF.dll "%PIAFDir%"
 cd "%PIAFDir%\"
 "%PIAFDir%\regplugin.exe" DBMDataRef.dll
 "%PIAFDir%\regplugin.exe" /Owner:DBMDataRef.dll DBM.dll
 "%PIAFDir%\regplugin.exe" /Owner:DBMDataRef.dll DBMPointDriverOSIsoftPIAF.dll
 cd %~dp0
)

build\DBMTester.exe

:ExitBuild
