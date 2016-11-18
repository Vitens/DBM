@echo off
%~d0
cd %~dp0

if not exist build mkdir build

call clean.bat

set IncludeFiles=src\dbm\dbm.vb src\dbm\DBMCachedValue.vb src\dbm\DBMConstants.vb src\dbm\DBMFunctions.vb src\dbm\DBMMath.vb src\dbm\DBMPoint.vb src\dbm\DBMResult.vb src\dbm\DBMStatistics.vb

"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\Vbc.exe" /out:build\DBMUnitTests-offline.exe /define:OfflineUnitTests=True %IncludeFiles% test\unit\DBMUnitTests.vb
if not %errorlevel% equ 0 pause
if exist "%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll" "%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\vbc.exe" /reference:"%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll","%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDKCommon.dll" /out:build\DBMUnitTests.exe %IncludeFiles% test\unit\DBMUnitTests.vb
if not %errorlevel% equ 0 pause
if exist "%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll" "%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\vbc.exe" /reference:"%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll","%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDKCommon.dll","%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PITimeServer.dll","%PIHOME%\ACE\OSISoft.PIACENet.dll" /rootnamespace:PIACE.DBMRt /target:library /out:build\DBMRt.dll %IncludeFiles% src\PIACENet\DBMRtConstants.vb src\PIACENet\DBMRt.vb
if not %errorlevel% equ 0 pause

if exist build\DBMUnitTests-offline.exe build\DBMUnitTests-offline.exe

pause
