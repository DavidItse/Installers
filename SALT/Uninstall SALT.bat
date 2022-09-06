Title Delete installed files. 

@echo off
setlocal
PROMPT
SET /P AREYOUSURE=Are you sure you want to uninstall SALT (Y/[N])?
If /I "%AREYOUSURE%" NEQ "Y" GOTO END

echo

If EXIST "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\HubSync Workpapers\" del /q "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\HubSync Workpapers\"

If EXIST "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\HubSync Workpapers\" rmdir /q "%USERPROFILE%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\HubSync Workpapers\"

If EXIST "%LOCALAPPDATA%\HubSync Workpapers\SALT\" del /q "%LOCALAPPDATA%\HubSync Workpapers\SALT\"

If EXIST "%LOCALAPPDATA%\HubSync Workpapers\SALT\" rmdir "%LOCALAPPDATA%\HubSync Workpapers\SALT\"

ECHO Uninstall Complete. 
Pause
RmDir /q "%LOCALAPPDATA%\HubSync Workpapers\SALT\" 

:END