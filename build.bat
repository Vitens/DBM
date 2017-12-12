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

cd /d %~dp0

rem Variables
set vbc="%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\Vbc.exe" /optimize+ /nologo /novbruntimeref
set PIAFDir=%PIHOME%\AF
set PIAFRef=%PIAFDir%\PublicAssemblies\4.0\OSIsoft.AFSDK.dll

rem Set up build directory
if not exist build mkdir build
del /Q build\*
copy LICENSE build > NUL

rem Apply patches
if "%CI%" == "True" for /f "delims=" %%i in ('git rev-parse --short HEAD') do set commit=%%i
if "%CI%" == "True" powershell -Command "(Get-Content src\dbm\DBM.vb) -replace 'Const GITHASH As String = \".*?\"', 'Const GITHASH As String = \"%commit%\"' | Set-Content src\dbm\DBM.vb"

rem Build
%vbc% /target:library /out:build\DBM.dll src\shared\*.vb src\dbm\*.vb
%vbc% /reference:build\DBM.dll /out:build\DBMAbout.exe src\shared\*.vb src\dbmabout\*.vb
%vbc% /reference:build\DBM.dll /target:library /out:build\DBMPointDriverCSV.dll src\shared\*.vb src\dbm\driver\DBMPointDriverCSV.vb
%vbc% /reference:build\DBM.dll,build\DBMPointDriverCSV.dll /out:build\DBMTester.exe src\shared\*.vb src\dbmtester\*.vb
if exist "%PIAFRef%" (
 %vbc% /reference:"%PIAFRef%",build\DBM.dll /target:library /out:build\DBMPointDriverOSIsoftPIAF.dll src\shared\*.vb src\dbm\driver\DBMPointDriverOSIsoftPIAF.vb
 %vbc% /reference:"%PIAFRef%",build\DBM.dll,build\DBMPointDriverOSIsoftPIAF.dll /target:library /out:build\DBMDataRef.dll src\shared\*.vb src\PIAFDataRef\*.vb
 rem Register PI AF Data Reference on AF server
 tasklist | find "AFService.exe" && (
  if exist "%PIAFDir%\DBMDataRef.dll" "%PIAFDir%\regplugin.exe" /Unregister "%PIAFDir%\DBMDataRef.dll"
  copy build\DBMDataRef.dll "%PIAFDir%"
  copy build\DBM.dll "%PIAFDir%"
  copy build\DBMPointDriverOSIsoftPIAF.dll "%PIAFDir%"
  cd /d "%PIAFDir%"
  regplugin.exe DBMDataRef.dll
  regplugin.exe /Owner:DBMDataRef.dll DBM.dll
  regplugin.exe /Owner:DBMDataRef.dll DBMPointDriverOSIsoftPIAF.dll
  cd /d %~dp0
 )
)

rem Output version, copyright and license information and unit and integration test results
build\DBMAbout.exe

:ExitBuild
