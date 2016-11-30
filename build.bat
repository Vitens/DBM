@echo off

rem DBM
rem Dynamic Bandwidth Monitor
rem Leak detection method implemented in a real-time data historian
rem
rem Copyright (C) 2014, 2015, 2016 J.H. Fitié, Vitens N.V.
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

set vbc="%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\Vbc.exe"
set IncludeFiles=src\dbm\DBM.vb src\dbm\DBMCachedValue.vb src\dbm\DBMConstants.vb src\dbm\DBMDataManager.vb src\dbm\DBMFunctions.vb src\dbm\DBMMath.vb src\dbm\DBMPoint.vb src\dbm\DBMResult.vb src\dbm\DBMStatistics.vb
set PIRefs="%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll","%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDKCommon.dll"
set PIACERefs="%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PITimeServer.dll","%PIHOME%\ACE\OSISoft.PIACENet.dll"

%vbc% /target:library /out:build\DBMDriverArray.dll src\dbm\driver\DBMDriverArray.vb %IncludeFiles%
if not exist build\DBMDriverArray.dll goto ExitBuild

%vbc% /reference:build\DBMDriverArray.dll /out:build\DBMUnitTestsArray.exe /define:OfflineUnitTests=True test\unit\DBMUnitTests.vb
if not exist build\DBMUnitTestsArray.exe goto ExitBuild

if exist "%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll" (

    %vbc% /reference:%PIRefs% /target:library /out:build\DBMDriverOSIsoftPI.dll src\dbm\driver\DBMDriverOSIsoftPI.vb %IncludeFiles%
    if not exist build\DBMDriverOSIsoftPI.dll goto ExitBuild

    %vbc% /reference:%PIRefs%,build\DBMDriverOSIsoftPI.dll /out:build\DBMUnitTestsOSIsoftPI.exe test\unit\DBMUnitTests.vb
    if not exist build\DBMUnitTestsOSIsoftPI.exe goto ExitBuild

    if exist "%PIHOME%\ACE\OSISoft.PIACENet.dll" (

        %vbc% /reference:%PIRefs%,%PIACERefs%,build\DBMDriverOSIsoftPI.dll /rootnamespace:PIACE.DBMRt /target:library /out:build\DBMRt.dll src\PIACENet\DBMRt.vb src\PIACENet\DBMRtCalculator.vb src\PIACENet\DBMRtCorrelationPIPoint.vb src\PIACENet\DBMRtPIPoint.vb src\PIACENet\DBMRtPIServer.vb
        if not exist build\DBMRt.dll goto ExitBuild

    )

)

build\DBMUnitTestsArray.exe

:ExitBuild
pause
