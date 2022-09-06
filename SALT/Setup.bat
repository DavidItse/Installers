@echo off
Title Download Setup file from github And run it.


if /i "%answer:~,1%" EQU "Y" goto InstallIt
echo This will install the SALT WorkPaper 
set /p answer=Do you want to continue (Y/N)? 
if /i "%answer:~,1%" EQU "Y"GoTo InstallIt
if /i "%answer:~,1%" EQU "N"exit /b
echo Please type Y for Yes Or N for No
:InstallIt
If Not "%minimized%"=="" GoTo :minimized
Set minimized=True
start /min cmd /C"%~dpnx0"
GoTo :EOF
:minimized
If Not exist "%LOCALAPPDATA%\HubSync Workpapers\SALT\" mkdir "%LOCALAPPDATA%\HubSync Workpapers\SALT\"
If Not exist "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\HubSync Workpapers\" mkdir "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\HubSync Workpapers"

Call :CreateConverter
:CheckForFile
IF EXIST "%LOCALAPPDATA%\HubSync Workpapers\SALT\converttoAnsiPC.VBS" GOTO FoundIt
GOTO CheckForFile
:FoundIt
If Not exist "%LOCALAPPDATA%\HubSync Workpapers\SALT\converttoAnsiPC.VBS" (
Pause
)

Set "url=https://raw.githubusercontent.com/DavidItse/Installers/main/SALT/SendEmailToAdmin.vbs"
For %%# in (%url%) do ( set "File=%tmp%\%%~n#.vbs")

Call :Download "%url%" "%File%"
If exist "%File%" (
   ( Type "%File%")>con
REM to save the contents in new text file
( Type "%File%" > "%LOCALAPPDATA%\HubSync Workpapers\SALT\SendEmailToAdmin.vbs")
del /q %File%
)

:CheckForSendEmailToAdmin.vbs
IF EXIST "%LOCALAPPDATA%\HubSync Workpapers\SALT\SendEmailToAdmin.vbs" GOTO FoundSendEmailToAdmin.vbs
GOTO CheckForSendEmailToAdmin.vbs
:FoundSendEmailToAdmin.vbs
If Not exist "%LOCALAPPDATA%\HubSync Workpapers\SALT\SendEmailToAdmin.vbs" (
Pause
)

wscript //NoLogo "%LOCALAPPDATA%\HubSync Workpapers\SALT\converttoAnsiPC.vbs" <"%LOCALAPPDATA%\HubSync Workpapers\SALT\SendEmailToAdmin.vbs" >"%LOCALAPPDATA%\HubSync Workpapers\SALT\SendEmailToAdmin.vbs.txt"
del /q "%LOCALAPPDATA%\HubSync Workpapers\SALT\SendEmailToAdmin.vbs"
ren "%LOCALAPPDATA%\HubSync Workpapers\SALT\SendEmailToAdmin.vbs.txt" "SendEmailToAdmin.vbs"

Set "url=https://raw.githubusercontent.com/DavidItse/Installers/main/SALT/SALT.xlsm"
For %%# in (%url%) do ( set "File=%tmp%\%%~n#.xlsm")

Call :Download "%url%" "%File%"
If exist "%File%" (
   ( Type "%File%")>con
REM to save the contents in new text file
( Type "%File%" > "%LOCALAPPDATA%\HubSync Workpapers\SALT\SALT.xlsm")
del /q %File%
)

:CheckForSALT.xlsm
IF EXIST "%LOCALAPPDATA%\HubSync Workpapers\SALT\SALT.xlsm" GOTO FoundSALT.xlsm
GOTO CheckForSALT.xlsm
:FoundSALT.xlsm
If Not exist "%LOCALAPPDATA%\HubSync Workpapers\SALT\SALT.xlsm" (
Pause
)

Set "url=https://raw.githubusercontent.com/DavidItse/Installers/main/SALT/Version.txt"
For %%# in (%url%) do ( set "File=%tmp%\%%~n#.txt")

Call :Download "%url%" "%File%"
If exist "%File%" (
   ( Type "%File%")>con
REM to save the contents in new text file
( Type "%File%" > "%LOCALAPPDATA%\HubSync Workpapers\SALT\Version.txt")
del /q %File%
)

:CheckForVersion.txt
IF EXIST "%LOCALAPPDATA%\HubSync Workpapers\SALT\Version.txt" GOTO FoundVersion.txt
GOTO CheckForVersion.txt
:FoundVersion.txt
If Not exist "%LOCALAPPDATA%\HubSync Workpapers\SALT\Version.txt" (
Pause
)

Set "url=https://raw.githubusercontent.com/DavidItse/Installers/main/SALT/Uninstall SALT.bat"
For %%# in (%url%) do ( set "File=%tmp%\%%~n#.bat")

Call :Download "%url%" "%File%"
If exist "%File%" (
   ( Type "%File%")>con
REM to save the contents in new text file
( Type "%File%" > "%LOCALAPPDATA%\HubSync Workpapers\SALT\Uninstall SALT.bat")
del /q %File%
)

:CheckForUninstall SALT.bat
IF EXIST "%LOCALAPPDATA%\HubSync Workpapers\SALT\Uninstall SALT.bat" GOTO FoundUninstall SALT.bat
GOTO CheckForUninstall SALT.bat
:FoundUninstall SALT.bat
If Not exist "%LOCALAPPDATA%\HubSync Workpapers\SALT\Uninstall SALT.bat" (
Pause
)

wscript //NoLogo "%LOCALAPPDATA%\HubSync Workpapers\SALT\converttoAnsiPC.vbs" <"%LOCALAPPDATA%\HubSync Workpapers\SALT\Uninstall SALT.bat" >"%LOCALAPPDATA%\HubSync Workpapers\SALT\Uninstall SALT.bat.txt"
del /q "%LOCALAPPDATA%\HubSync Workpapers\SALT\Uninstall SALT.bat"
ren "%LOCALAPPDATA%\HubSync Workpapers\SALT\Uninstall SALT.bat.txt" "Uninstall SALT.bat"

set SCRIPT="%TEMP%\%RANDOM%-%RANDOM%-%RANDOM%-%RANDOM%.vbs"

echo Set oWS = WScript.CreateObject("WScript.Shell") >> %SCRIPT%
echo sLinkFile = "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\HubSync Workpapers\SALT.lnk" >> %SCRIPT% 
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> %SCRIPT%
echo oLink.TargetPath = "%LOCALAPPDATA%\HubSync Workpapers\SALT\SALT.xlsm" >> %SCRIPT%
echo oLink.Save >> %SCRIPT%

wscript /nologo %SCRIPT%
del %SCRIPT%

set SCRIPT="%TEMP%\%RANDOM%-%RANDOM%-%RANDOM%-%RANDOM%.vbs"

echo Set oWS = WScript.CreateObject("WScript.Shell") >> %SCRIPT%
echo sLinkFile = "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\HubSync Workpapers\Uninstall SALT.lnk" >> %SCRIPT%
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> %SCRIPT%
echo oLink.TargetPath = "%LOCALAPPDATA%\HubSync Workpapers\SALT\Uninstall SALT.bat" >> %SCRIPT%
echo oLink.Save >> %SCRIPT%

wscript /nologo %SCRIPT%
del %SCRIPT%

Call :OpenSetup
timeout 2 > nul
@Echo off

Exit /b
ECHO
::*********************************************************************************
:CreateConverter
@echo off
DEL "%LOCALAPPDATA%\HubSync Workpapers\SALT\converttoAnsiPC.VBS"
echo Do Until WScript.StdIn.AtEndOfStream>> "%LOCALAPPDATA%\HubSync Workpapers\SALT\converttoAnsiPC.VBS"
echo WScript.StdOut.WriteLine WScript.StdIn.ReadLine>> "%LOCALAPPDATA%\HubSync Workpapers\SALT\converttoAnsiPC.VBS"
echo Loop>> "%LOCALAPPDATA%\HubSync Workpapers\SALT\converttoAnsiPC.VBS"

Exit /b
:Download <url> <File>
@ECHO off
Powershell.exe -command "(New-Object System.Net.WebClient).DownloadFile('%1','%2')"
exit /b
:DeleteSetup
@ECHO off
del "%LOCALAPPDATA%\HubSync Workpapers\SALT\SALT.xlsm" /s /f /q
exit /b
:OpenSetup
@ECHO off
Start "" "%LOCALAPPDATA%\HubSync Workpapers\SALT\SALT.xlsm"
exit /b
