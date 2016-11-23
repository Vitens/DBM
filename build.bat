@echo off
%~d0
cd %~dp0

if not exist build mkdir build

call clean.bat

set vbc="%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\Vbc.exe"
set IncludeFiles=src\dbm\dbm.vb src\dbm\DBMCachedValue.vb src\dbm\DBMConstants.vb src\dbm\DBMDataManager.vb src\dbm\DBMFunctions.vb src\dbm\DBMMath.vb src\dbm\DBMPoint.vb src\dbm\DBMResult.vb src\dbm\DBMStatistics.vb
set PIRefs="%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll","%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDKCommon.dll"
set PIACERefs="%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PITimeServer.dll","%PIHOME%\ACE\OSISoft.PIACENet.dll"

%vbc% /target:library /out:build\DBMDriverArray.dll src\dbm\driver\DBMDriverArray.vb %IncludeFiles%
%vbc% /reference:build\DBMDriverArray.dll /out:build\DBMUnitTestsArray.exe /define:OfflineUnitTests=True test\unit\DBMUnitTests.vb

if exist "%PIHOME%\pisdk\PublicAssemblies\OSIsoft.PISDK.dll" (

    %vbc% /reference:%PIRefs% /target:library /out:build\DBMDriverOSIsoftPI.dll src\dbm\driver\DBMDriverOSIsoftPI.vb %IncludeFiles%
    %vbc% /reference:%PIRefs%,build\DBMDriverOSIsoftPI.dll /out:build\DBMUnitTestsOSIsoftPI.exe test\unit\DBMUnitTests.vb

    if exist "%PIHOME%\ACE\OSISoft.PIACENet.dll" (

        %vbc% /reference:%PIRefs%,%PIACERefs%,build\DBMDriverOSIsoftPI.dll /rootnamespace:PIACE.DBMRt /target:library /out:build\DBMRt.dll src\PIACENet\DBMRt.vb src\PIACENet\DBMRtConstants.vb

    )

)

if exist build\DBMUnitTestsArray.exe build\DBMUnitTestsArray.exe

pause
