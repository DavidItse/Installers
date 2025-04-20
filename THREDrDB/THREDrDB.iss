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
Source: "C:\Users\david\source\repos\THREDrDBForExcel\THREDrDB_add-in\bin\Release\ThredrDB_add-in-AddIn64-packed.xll"; DestDir: "{app}"
Source: "C:\Users\david\source\repos\Installers\THREDrDB\Setup\Setup.bat"; DestDir: "{app}"
Source: "C:\Users\david\source\repos\Installers\THREDrDB\Version.txt"; DestDir: "{app}"
[Run]
Filename: "cscript"; Parameters: """{app}\RegisterXLL.vbs"""; Flags: runhidden; Description: "Register THREDrDB XLL"

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  VBSFile: String;
  InstallerPath: String;
  DestPath: String;
begin
  if CurStep = ssInstall then
  begin
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
    if not FileCopy(InstallerPath, DestPath, False) then
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
	DeleteFile(ExpandConstant('{app}\SetupTHREDr.exe')); 
  end;
end;