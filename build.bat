@echo off

rem Dynamic Bandwidth Monitor
rem Leak detection method implemented in a real-time data historian
rem Copyright (C) 2014-2020  J.H. Fitié, Vitens N.V.
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
set vbc="%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\Vbc.exe" /win32icon:src\res\dbm.ico /optimize+ /nologo /novbruntimeref
set PIAFRef=%PIHOME%\AF\PublicAssemblies\4.0\OSIsoft.AFSDK.dll

rem Set up build directory
if not exist build mkdir build
del /Q build\*
copy LICENSE build > NUL

rem Build
%vbc% /target:library /out:build\DBM.dll src\dbm\*.vb
%vbc% /reference:build\DBM.dll /out:build\DBMAbout.exe src\dbm\DBMManifest.vb src\dbmabout\*.vb
%vbc% /reference:build\DBM.dll /target:library /out:build\DBMPointDriverCSV.dll src\dbm\DBMManifest.vb src\dbm\driver\DBMPointDriverCSV.vb
%vbc% /reference:build\DBM.dll,build\DBMPointDriverCSV.dll /out:build\DBMTester.exe src\dbm\DBMManifest.vb src\dbmtester\*.vb
if exist "%PIAFRef%" (
 %vbc% /reference:"%PIAFRef%",build\DBM.dll /target:library /out:build\DBMPointDriverOSIsoftPIAF.dll src\dbm\DBMManifest.vb src\dbm\driver\DBMPointDriverOSIsoftPIAF.vb
 %vbc% /reference:"%PIAFRef%",build\DBM.dll,build\DBMPointDriverOSIsoftPIAF.dll /target:library /out:build\DBMDataRef.dll src\dbm\DBMManifest.vb src\PIAFDataRef\*.vb
)
