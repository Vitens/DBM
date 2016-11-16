@echo off
"%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\Vbc.exe" /out:build\DBMUnitTests-offline.exe /define:OfflineUnitTests=True src\dbm\Statistics.vb src\dbm\dbm.vb test\unit\DBMUnitTests.vb
if exist "%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll" "%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\vbc.exe" /reference:"%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll","%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDKCommon.dll" /out:build\DBMUnitTests.exe src\dbm\Statistics.vb src\dbm\dbm.vb test\unit\DBMUnitTests.vb
if exist "%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll" "%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\vbc.exe" /reference:"%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll","%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDKCommon.dll","%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PITimeServer.dll","%PIHOME%\ACE\OSISoft.PIACENet.dll" /rootnamespace:PIACE.DBMRt /target:library /out:build\DBMRt.dll src\dbm\Statistics.vb src\dbm\dbm.vb src\PIACENet\DBMRt.vb
build\DBMUnitTests-offline.exe
pause
