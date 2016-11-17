@echo off
%~d0
cd %~dp0

if not exist build mkdir build

call clean.bat

set IncludeFiles=src\dbm\DBMConstants.vb src\dbm\DBMStatistics.vb src\dbm\DBMMath.vb src\dbm\DBMCachedValue.vb src\dbm\DBMPoint.vb src\dbm\dbm.vb

"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\Vbc.exe" /out:build\DBMUnitTests-offline.exe /define:OfflineUnitTests=True %IncludeFiles% test\unit\DBMUnitTests.vb
if exist "%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll" "%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\vbc.exe" /reference:"%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll","%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDKCommon.dll" /out:build\DBMUnitTests.exe %IncludeFiles% test\unit\DBMUnitTests.vb
if exist "%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll" "%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\vbc.exe" /reference:"%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll","%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDKCommon.dll","%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PITimeServer.dll","%PIHOME%\ACE\OSISoft.PIACENet.dll" /rootnamespace:PIACE.DBMRt /target:library /out:build\DBMRt.dll %IncludeFiles% src\PIACENet\DBMRtConstants.vb src\PIACENet\DBMRt.vb

build\DBMUnitTests-offline.exe

pause
