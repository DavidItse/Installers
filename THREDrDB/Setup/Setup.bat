@echo off
Title THREDrDB Update

cls
echo Updating THREDrDB to the latest version...
echo.
set /p answer=Do you want to continue (Y/N)? 
if /i "%answer:~,1%" EQU "Y" goto InstallIt
if /i "%answer:~,1%" EQU "N" exit /b
echo Please type Y for Yes or N for No
goto :eof
:InstallIt
if not exist "C:\Users\david\AppData\Roaming\ThredrDB" mkdir "C:\Users\david\AppData\Roaming\ThredrDB"
echo Downloading update...
powershell.exe -Command "(New-Object System.Net.WebClient).DownloadFile('C:\Users\david\source\repos\Installers\THREDrDB\THREDrDB-v1.0.3.zip', '%TEMP%\THREDrDB-v1.0.3.zip')" || goto :Error
echo Extracting update...
powershell.exe -Command "Expand-Archive -Path '%TEMP%\THREDrDB-v1.0.3.zip' -DestinationPath 'C:\Users\david\AppData\Roaming\ThredrDB' -Force" || goto :Error
echo Cleaning up...
del "%TEMP%\THREDrDB-v1.0.3.zip" > nul 2>&1
echo Registering THREDrDB add-in with Excel...
echo Set objExcel = CreateObject("Excel.Application") > "%TEMP%\RegisterXLL.vbs"
echo objExcel.Visible = False >> "%TEMP%\RegisterXLL.vbs"
echo objExcel.Workbooks.Add >> "%TEMP%\RegisterXLL.vbs"
echo objExcel.AddIns.Add("C:\Users\david\AppData\Roaming\ThredrDB\ThredrDB_add-in-AddIn64-packed.xll").Installed = True >> "%TEMP%\RegisterXLL.vbs"
echo objExcel.Quit >> "%TEMP%\RegisterXLL.vbs"
cscript "%TEMP%\RegisterXLL.vbs" > nul 2>&1
del "%TEMP%\RegisterXLL.vbs" > nul 2>&1
echo Registration complete!
echo Adding entry to Add/Remove Programs...
reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\THREDrDB /v DisplayName /t REG_SZ /d "THREDrDB" /f > nul
reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\THREDrDB /v UninstallString /t REG_SZ /d "%APPDATA%\ThredrDB\UninstallTHREDr.exe" /f > nul
reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\THREDrDB /v Publisher /t REG_SZ /d "Thredr LLC" /f > nul
reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\THREDrDB /v DisplayVersion /t REG_SZ /d "1.0.3" /f > nul
reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\THREDrDB /v InstallLocation /t REG_SZ /d "%APPDATA%\ThredrDB" /f > nul
echo Update complete! Please restart Excel to load the new version.
echo This window will close in 5 seconds...
timeout 5 > nul
exit /b
:Error
echo An error occurred during the update. Please check your internet connection and try again.
echo If the problem persists, contact support@thredrdb.com.
echo This window will close in 10 seconds...
timeout 10 > nul
exit /b 1