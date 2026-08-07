@echo off

echo. >NUL 2>&1
echo Starting Microsoft Defender for Endpoint onboarding process... >NUL 2>&1
echo. >NUL 2>&1

set errorCode=0
set lastError=0
set "troubleshootInfo=For more information, visit: https://go.microsoft.com/fwlink/p/?linkid=822807"
set "errorDescription="

echo Testing administrator privileges >NUL 2>&1

%windir%\System32\net.exe session >NUL 2>&1
if %ERRORLEVEL% NEQ 0 (
	@echo Script is running with insufficient privileges. Please run with administrator privileges> %WINDIR%\temp\senseTmp.txt
	set errorCode=65
 set lastError=%ERRORLEVEL%
	GOTO ERROR
)

echo Script is running with sufficient privileges >NUL 2>&1
echo. >NUL 2>&1
echo Performing onboarding operations >NUL 2>&1
echo. >NUL 2>&1

IF [%PROCESSOR_ARCHITEW6432%] EQU [] (
  set powershellPath=%windir%\System32\WindowsPowerShell\v1.0\powershell.exe
) ELSE (
  set powershellPath=%windir%\SysNative\WindowsPowerShell\v1.0\powershell.exe
)

set sdbin=0100048044000000540000000000000014000000020030000200000000001400FF0F120001010000000000051200000000001400E104120001010000000000050B0000000102000000000005200000002002000001020000000000052000000020020000 >NUL 2>&1
%windir%\System32\reg.exe add HKLM\SYSTEM\CurrentControlSet\Control\WMI\Security /v 14f8138e-3b61-580b-544b-2609378ae460 /t REG_BINARY /d %sdbin% /f >NUL 2>&1
%windir%\System32\reg.exe add HKLM\SYSTEM\CurrentControlSet\Control\WMI\Security /v cb2ff72d-d4e4-585d-33f9-f3a395c40be7 /t REG_BINARY /d %sdbin% /f >NUL 2>&1

%windir%\System32\reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection" /v DisableEnterpriseAuthProxy /t REG_DWORD /f /d 1 >NUL 2>&1

%powershellPath% -ExecutionPolicy Bypass -NoProfile -Command "Add-Type ' using System; using System.IO; using System.Runtime.InteropServices; using Microsoft.Win32.SafeHandles; using System.ComponentModel; public static class Elam{ [DllImport(\"Kernel32\", CharSet=CharSet.Auto, SetLastError=true)] public static extern bool InstallELAMCertificateInfo(SafeFileHandle handle); public static void InstallWdBoot(string path) { Console.Out.WriteLine(\"About to call create file on {0}\", path); var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read); var handle = stream.SafeFileHandle; Console.Out.WriteLine(\"About to call InstallELAMCertificateInfo on handle {0}\", handle.DangerousGetHandle()); if (!InstallELAMCertificateInfo(handle)) { Console.Out.WriteLine(\"Call failed.\"); throw new Win32Exception(Marshal.GetLastWin32Error()); } Console.Out.WriteLine(\"Call successful.\"); } } '; $driverPath = $env:SystemRoot + '\System32\Drivers\WdBoot.sys'; [Elam]::InstallWdBoot($driverPath) " >NUL 2>&1

%windir%\System32\reg.exe query "HKLM\SOFTWARE\Policies\Microsoft\Windows Advanced Threat Protection" /v 696C1FA1-4030-4FA4-8713-FAF9B2EA7C0A /reg:64 > %WINDIR%\temp\senseTmp.txt 2>&1
if %ERRORLEVEL% EQU 0 (  
    %windir%\System32\reg.exe delete "HKLM\SOFTWARE\Policies\Microsoft\Windows Advanced Threat Protection" /v 696C1FA1-4030-4FA4-8713-FAF9B2EA7C0A /f > %WINDIR%\temp\senseTmp.txt 2>&1
    if %ERRORLEVEL% NEQ 0 (
        set "errorDescription=Unable to delete previous offboarding information from registry."
        set errorCode=5
        set lastError=%ERRORLEVEL%
        GOTO ERROR
    )
)

%windir%\System32\reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows Advanced Threat Protection" /v OnboardingInfo /t REG_SZ /f /d "{\"body\":\"{\\\"previousOrgIds\\\":[],\\\"orgId\\\":\\\"be506892-7989-4df4-840a-bbbf3cdaef71\\\",\\\"geoLocationUrl\\\":\\\"https://edr-cus3.us.endpoint.security.microsoft.com/edr/\\\",\\\"datacenter\\\":\\\"CentralUs3\\\",\\\"vortexGeoLocation\\\":\\\"US\\\",\\\"vortexServerUrl\\\":\\\"https://us-v20.events.endpoint.security.microsoft.com/OneCollector/1.0\\\",\\\"vortexTicketUrl\\\":\\\"https://events.data.microsoft.com\\\",\\\"partnerGeoLocation\\\":\\\"GW_US\\\",\\\"version\\\":\\\"2.12\\\",\\\"packageGuid\\\":\\\"c7171b51-596f-4fd3-a8d0-f8b2d2f6934a\\\"}\",\"sig\":\"r5f92LSz+6t29Xl49ciHNjAI32LpjEoC5oxI3W49Alse/Uxanz/f/EHqhUC4ObSNpi8RwECsryHIzG/hA/fSXDiGonILn+Fs03+f6k2tPDrUM67yYhTqJxnPYA/m+2VZe2iBT1ia0m2PNFJGG05WsPXBRPae9M4hRXuWElR0Q16ZsSNEDTdZcKd8GPARIGjQ0lljARoftAb0/rcqRafyxP3NkJ0y3OB9G1SEZVM5kIQIBAxaydTj4cYMMJfakMCKkwbsAvaaozxt2nTvy/nKAiYx29cWHN66X6IDbWhYA/jCSXq3s6pkw0F6Z05Pg4kOHQy0cLaIAOatZwDLKOwNjA==\",\"sha256sig\":\"r5f92LSz+6t29Xl49ciHNjAI32LpjEoC5oxI3W49Alse/Uxanz/f/EHqhUC4ObSNpi8RwECsryHIzG/hA/fSXDiGonILn+Fs03+f6k2tPDrUM67yYhTqJxnPYA/m+2VZe2iBT1ia0m2PNFJGG05WsPXBRPae9M4hRXuWElR0Q16ZsSNEDTdZcKd8GPARIGjQ0lljARoftAb0/rcqRafyxP3NkJ0y3OB9G1SEZVM5kIQIBAxaydTj4cYMMJfakMCKkwbsAvaaozxt2nTvy/nKAiYx29cWHN66X6IDbWhYA/jCSXq3s6pkw0F6Z05Pg4kOHQy0cLaIAOatZwDLKOwNjA==\",\"cert\":\"MIIFjDCCA3SgAwIBAgITMwAABHxkRBN+f6t9xQAAAAAEfDANBgkqhkiG9w0BAQsFADB+MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgU2VjdXJlIFNlcnZlciBDQSAyMDExMB4XDTI2MDQxNjE5MjcxMloXDTI2MTAxNzE5MjcxMlowJzElMCMGA1UEAxMcbTM2NWRvb3AtU2V2aWxsZS5XaW5kb3dzLmNvbTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAO1NhDnVTVFaunPNq5QM7hs0NAp7gANPycyhGapj9UPFgrwmUgj8kWYcqzUOTAffbii1gK2rgwTh6ZA7u/ZCF+CXaHThIx1V6v+ziNApPjaLUe2nPNeRvdFVtAijL0C3MEAYXuugaB1eGlUgjpHkWydIrrNswEWEgN3gUogn1yD2OnG0Nfrl6+syyPMyS3oxT3FpQO5Y1ZeO/9tcqa6vNm8YaZhrD30Whg9TOfrMqttTuUvbOS/qOyL/kOHrpQBqfqcKxBqvqlG7M0P/I8uYO8lFT5gxkkA48BfW+w4QsN36u6a3wdvFSl3s4g5UvCyChTezTqFKoPd/Ed46zKurhjkCAwEAAaOCAVgwggFUMA4GA1UdDwEB/wQEAwIFIDAdBgNVHSUEFjAUBggrBgEFBQcDAQYIKwYBBQUHAwIwDAYDVR0TAQH/BAIwADAeBgNVHREEFzAVghNTZXZpbGxlLldpbmRvd3MuY29tMB0GA1UdDgQWBBRVZW0B1gDN+Uw/l3NT9cnRiMVZlzAfBgNVHSMEGDAWgBQ2VollSctbmy88rEIWUE2RuTPXkTBTBgNVHR8ETDBKMEigRqBEhkJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNTZWNTZXJDQTIwMTFfMjAxMS0xMC0xOC5jcmwwYAYIKwYBBQUHAQEEVDBSMFAGCCsGAQUFBzAChkRodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY1NlY1NlckNBMjAxMV8yMDExLTEwLTE4LmNydDANBgkqhkiG9w0BAQsFAAOCAgEAzutnSOqro/vJycNAWSRMKUCan82NGlTG2UBVK+KQkggWq/Sf73iLhUFF+GaHyqzTmrXZT0/B68LRgl1pMtqNsMS+RvrcLMnPP8Uu88Z66nYn2qbkLy9cgeGvb05iPdH2E2RGuOh0mVdfBWRBqqnzL6s8pNM5vk5rRfOfQqnce49FmN29bHRbsAnjP3HJOaSu8JN9rtPGoihiXFTuCPj+4cVlPJVW/SImgl3VVET1+QVqRlQpPbQaQch7nKkFE8o3QHcJA/YE5pnU/xeVqUy39QMvDETn1CREG+3WU8AP8mLXWHvkPon6bOz+5Mf3fh+5bgjvQ5JKMhcpOBrYHV+duyA6kb1GjvEW5DSGO1E1gbEA1Y8dhRc0reCZcz1mCrGm8QDQ3f1eQ6j5NYN8fxDIhunnL8fBZwREFBO92Zb9Y7kueBUWUIptC51XeL40OgWJY/b7qyZiuwzWkBkAfDc5Iy7qSo1BtYC46Uio53auRj5VTq7UTCr7+gZYc81OaV8+u3M4x9pOXGBQW2dKfJMNz204nYVhgkLytd9l3PFziSuULdiRzLyHIQQjm8CVUbtCHj8Bex3pXy8A6VckrqDrjgAvkjChmOoab6DHm8bdiOTH3x+cciT5Xtk2cXVyRUBdIM+d7pTt+FOJ2GCjYgxnNeE82cJO87tRzzygHVvOoBo=\",\"chain\":[\"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\"]}" > %WINDIR%\temp\senseTmp.txt 2>&1
if %ERRORLEVEL% NEQ 0 (
   set "errorDescription=Unable to write onboarding information to registry."
   set errorCode=10
   set lastError=%ERRORLEVEL%
   GOTO ERROR
)

echo Starting the service, if not already running >NUL 2>&1
echo. >NUL 2>&1
%windir%\System32\sc.exe query "SENSE" | %windir%\System32\find.exe /i "RUNNING" >NUL 2>&1
if %ERRORLEVEL% EQU 0 GOTO RUNNING

%windir%\System32\net.exe start sense > %WINDIR%\temp\senseTmp.txt 2>&1
if %ERRORLEVEL% NEQ 0 (
   echo Microsoft Defender for Endpoint Service has not started yet >NUL 2>&1
   GOTO WAIT_FOR_THE_SERVICE_TO_START
)
goto SUCCEEDED

:RUNNING
set "runningOutput=The Microsoft Defender for Endpoint Service is already running!"
echo %runningOutput% >NUL 2>&1
echo. >NUL 2>&1
%windir%\System32\eventcreate.exe /l Application /so WDATPOnboarding /t Information /id 10 /d "%runningOutput%" >NUL 2>&1
GOTO WAIT_FOR_THE_SERVICE_TO_START

:ERROR
Set /P errorMsg=<%WINDIR%\temp\senseTmp.txt
set "errorOutput=[Error Id: %errorCode%, Error Level: %lastError%] %errorDescription% Error message: %errorMsg%"
%powershellPath% -ExecutionPolicy Bypass -NoProfile -Command "Add-Type 'using System; using System.Diagnostics; using System.Diagnostics.Tracing; namespace Sense { [EventData(Name = \"Onboarding\")]public struct Onboarding{public string Message { get; set; }} public class Trace {public static EventSourceOptions TelemetryCriticalOption = new EventSourceOptions(){Level = EventLevel.Error, Keywords = (EventKeywords)0x0000200000000000, Tags = (EventTags)0x0200000}; public void WriteOnboardingMessage(string message){es.Write(\"OnboardingScript\", TelemetryCriticalOption, new Onboarding {Message = message});} private static readonly string[] telemetryTraits = { \"ETW_GROUP\", \"{5ECB0BAC-B930-47F5-A8A4-E8253529EDB7}\" }; private EventSource es = new EventSource(\"Microsoft.Windows.Sense.Client.Management\",EventSourceSettings.EtwSelfDescribingEventFormat,telemetryTraits);}}'; $logger = New-Object -TypeName Sense.Trace; $logger.WriteOnboardingMessage('%errorOutput%')" >NUL 2>&1
echo %errorOutput% >NUL 2>&1
echo. >NUL 2>&1
%windir%\System32\eventcreate.exe /l Application /so WDATPOnboarding /t Error /id %errorCode% /d "%errorOutput%" >NUL 2>&1
GOTO CLEANUP

:SUCCEEDED
echo Finished performing onboarding operations >NUL 2>&1
echo. >NUL 2>&1
GOTO WAIT_FOR_THE_SERVICE_TO_START

:WAIT_FOR_THE_SERVICE_TO_START
echo Waiting for the service to start >NUL 2>&1
echo. >NUL 2>&1

set /a counter=0

:SENSE_RUNNING_WAIT
%windir%\System32\sc.exe query "SENSE" | %windir%\System32\find.exe /i "RUNNING" >NUL 2>&1
if %ERRORLEVEL% NEQ 0 (
	IF %counter% EQU 4 (
		set "errorDescription=Unable to start Microsoft Defender for Endpoint Service."
		set errorCode=15
		set lastError=%ERRORLEVEL%
		GOTO ERROR
	)

	set /a counter=%counter%+1

	%windir%\System32\timeout.exe 5 >NUL 2>&1
	GOTO :SENSE_RUNNING_WAIT
)

set /a counter=0

:SENSE_ONBOARDED_STATUS_WAIT
%windir%\System32\reg.exe query "HKLM\SOFTWARE\Microsoft\Windows Advanced Threat Protection\Status" /v OnboardingState /reg:64 >NUL 2>&1
if %ERRORLEVEL% NEQ 0 (
	IF %counter% EQU 4 (
		@echo Microsoft Defender for Endpoint Service is not running as expected> %WINDIR%\temp\senseTmp.txt
		set errorCode=35
     set lastError=%ERRORLEVEL%
		GOTO ERROR
	)

	set /a counter=%counter%+1

	%windir%\System32\timeout.exe 5 >NUL 2>&1
	GOTO :SENSE_ONBOARDED_STATUS_WAIT
)

set /a counter=0

:SENSE_ONBOARDED_WAIT
%windir%\System32\reg.exe query "HKLM\SOFTWARE\Microsoft\Windows Advanced Threat Protection\Status" /v OnboardingState /reg:64 | %windir%\System32\find.exe /i "0x1" >NUL 2>&1
if %ERRORLEVEL% NEQ 0 (
	IF %counter% EQU 4 (
		@echo Microsoft Defender for Endpoint Service is not running as expected> %WINDIR%\temp\senseTmp.txt
		set errorCode=40
     set lastError=%ERRORLEVEL%
		GOTO ERROR
	)

	set /a counter=%counter%+1

	%windir%\System32\timeout.exe 5 >NUL 2>&1
	GOTO :SENSE_ONBOARDED_WAIT
)

set "successOutput=Successfully onboarded machine to Microsoft Defender for Endpoint"
echo %successOutput% >NUL 2>&1
echo.
%windir%\System32\eventcreate.exe /l Application /so WDATPOnboarding /t Information /id 20 /d "%successOutput%" >NUL 2>&1
%powershellPath% -ExecutionPolicy Bypass -NoProfile -Command "Add-Type 'using System; using System.Diagnostics; using System.Diagnostics.Tracing; namespace Sense { [EventData(Name = \"Onboarding\")]public struct Onboarding{public string Message { get; set; }} public class Trace {public static EventSourceOptions TelemetryCriticalOption = new EventSourceOptions(){Level = EventLevel.Informational, Keywords = (EventKeywords)0x0000200000000000, Tags = (EventTags)0x0200000}; public void WriteOnboardingMessage(string message){es.Write(\"OnboardingScript\", TelemetryCriticalOption, new Onboarding {Message = message});} private static readonly string[] telemetryTraits = { \"ETW_GROUP\", \"{5ECB0BAC-B930-47F5-A8A4-E8253529EDB7}\" }; private EventSource es = new EventSource(\"Microsoft.Windows.Sense.Client.Management\",EventSourceSettings.EtwSelfDescribingEventFormat,telemetryTraits);}}'; $logger = New-Object -TypeName Sense.Trace; $logger.WriteOnboardingMessage('%successOutput%')" >NUL 2>&1
"%PROGRAMFILES%\Windows Defender\MpCmdRun.exe" -ReloadEngine >NUL 2>&1

GOTO CLEANUP

:CLEANUP
if exist %WINDIR%\temp\senseTmp.txt del %WINDIR%\temp\senseTmp.txt
EXIT /B %errorCode%

