[Setup]
AppName=THREDrDB
AppVersion=1.0.3
DefaultDirName={userappdata}\ThredrDB
DefaultGroupName=THREDrDB
OutputDir=C:\Users\david\source\repos\Installers\THREDrDB
OutputBaseFilename=SetupTHREDr
UninstallDisplayName=THREDrDB
UninstallDisplayIcon={app}\ThredrDB_add-in-AddIn64-packed.xll
PrivilegesRequired=lowest

[Files]
Source: "C:\Users\david\source\repos\Installers\THREDrDB\Setup\Setup.bat"; DestDir: "{app}"

[Run]
Filename: "cscript"; Parameters: """{app}\RegisterXLL.vbs"""; Flags: runhidden; Description: "Register THREDrDB XLL"

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  VBSFile: String;
  ResultCode: Integer;
  ZipPath: String;
  XLLPath: String;
  TempDir: String;
  InstallerPath: String;
  DestPath: String;
begin
  if CurStep = ssInstall then
  begin
    TempDir := ExpandConstant('{userappdata}');
    if TempDir = '' then
    begin
      MsgBox('Failed to resolve the temporary directory ({userappdata}).', mbError, MB_OK);
      Exit;
    end;
   
    ZipPath := TempDir + '\THREDrDB-v1.0.3.zip';

    if not Exec('powershell.exe', 
      '-WindowStyle Hidden -Command "(New-Object System.Net.WebClient).DownloadFile(''https://github.com/DavidItse/Installers/raw/refs/heads/main/THREDrDB/THREDrDB-v1.0.3.zip'', ''' + ZipPath + ''')"', 
      '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    begin
      MsgBox('Failed to download the ZIP file from GitHub. Error code: ' + IntToStr(ResultCode), mbError, MB_OK);
      Exit;
    end;

    if not FileExists(ZipPath) then
    begin
      MsgBox('The ZIP file could not be downloaded. Please check your internet connection and try again.', mbError, MB_OK);
      Exit;
    end;
	MsgBox('Excel will be closed to complete the installation.', mbInformation, MB_OK);
	if Exec('powershell.exe', 
      '-WindowStyle Hidden -Command "Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue"', 
      '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    begin
      
      
    end
    else
    begin
      MsgBox('Warning: Could not terminate Excel processes (Error code: ' + IntToStr(ResultCode) + '). Installation will proceed, but there might be file conflicts if Excel is running.', mbInformation, MB_OK);
    end;
	
    if not Exec('powershell.exe', 
      '-WindowStyle Hidden -Command "Expand-Archive -Path ''' + ZipPath + ''' -DestinationPath ''' + ExpandConstant('{app}') + ''' -Force"', 
      '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    begin
      MsgBox('Failed to extract the ZIP file. Error code: ' + IntToStr(ResultCode), mbError, MB_OK);
      Exit;
    end;

    DeleteFile(ZipPath);

    XLLPath := ExpandConstant('{app}\ThredrDB_add-in-AddIn64-packed.xll');
    if not FileExists(XLLPath) then
    begin
      MsgBox('The XLL file was not found in the ZIP archive. Please ensure the ZIP file contains ThredrDB_add-in-AddIn64-packed.xll.', mbError, MB_OK);
      Exit;
    end;

    if not FileExists(ExpandConstant('{app}\Version.txt')) then
    begin
      MsgBox('Version.txt was not found in the ZIP archive. Please ensure the ZIP file contains Version.txt.', mbError, MB_OK);
      Exit;
    end;

    VBSFile := ExpandConstant('{app}\RegisterXLL.vbs');
    SaveStringToFile(VBSFile, 
      'On Error Resume Next' + #13#10 +
      'Set objExcel = CreateObject("Excel.Application")' + #13#10 +
      'If Err.Number <> 0 Then' + #13#10 +
      '    WScript.Echo "Error: Could not create Excel instance. Please ensure Excel is installed."' + #13#10 +
      '    WScript.Quit 1' + #13#10 +
      'End If' + #13#10 +
      'objExcel.Visible = True' + #13#10 +
      'objExcel.Workbooks.Add' + #13#10 +
      'Set addIn = objExcel.AddIns.Add("' + ExpandConstant('{app}\ThredrDB_add-in-AddIn64-packed.xll') + '")' + #13#10 +
      'If Err.Number <> 0 Then' + #13#10 +
      '    WScript.Echo "Error: Failed to register the add-in. Ensure the XLL file exists and you have permissions."' + #13#10 +
      'End If' + #13#10 +
      'addIn.Installed = True' + #13#10 +
      'If Err.Number <> 0 Then' + #13#10 +
      '    WScript.Echo "Error: Failed to enable the add-in. Ensure you have permissions to modify Excel settings."' + #13#10 +
      'End If' + #13#10 +
      'For Each addIn In objExcel.AddIns' + #13#10 +
      '    If InStr(LCase(addIn.FullName), LCase("THREDrDB-packed.xll")) > 0 Then' + #13#10 +
      '        If Not addIn.Installed Then' + #13#10 +
      '            WScript.Echo "Warning: Add-in registered but not enabled. Please enable it manually in Excel."' + #13#10 +
      '        End If' + #13#10 +
      '        Exit For' + #13#10 +
      '    End If' + #13#10 +
      'Next', False);
  end;
  if CurStep = ssPostInstall then
  begin
    InstallerPath := ExpandConstant('{srcexe}'); 
    DestPath := ExpandConstant('{app}\SetupTHREDr.exe'); 
    if not FileCopy(InstallerPath, DestPath, False) and not FileExists(DestPath) then
    begin
      MsgBox('Failed to copy the installer to the installation directory.', mbError, MB_OK);
    end;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    DeleteFile(ExpandConstant('{app}\RegisterXLL.vbs'));
    DeleteFile(ExpandConstant('{app}\Version.txt'));
    DeleteFile(ExpandConstant('{app}\ThredrDB_add-in-AddIn64-packed.xll'));
	DeleteFile(ExpandConstant('{app}\SetupTHREDr.exe'));
	DeleteFile(ExpandConstant('{app}\SetupTHREDr.zip'));
  end;
end;