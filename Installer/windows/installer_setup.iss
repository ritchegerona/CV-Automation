; Inno Setup script for CV Automation App
#define MyAppName "MSR CV Processing Studio"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "MSR"
#define MyAppExeName "CV_Automation.exe"

[Setup]
AppId={{D37E60FE-A01F-4395-97F7-BE68953112E5}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DisableDirPage=no
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename=MSR_CV_Automation_Setup
OutputDir=..\dist
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\MSR_CV_Template.docx"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\MSR_CV_Template.pdf"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
function CheckLibreOfficeInstalled(): Boolean;
var
  RegistryPath: String;
begin
  // Fallback 1: Check typical default file paths
  if FileExists('C:\Program Files\LibreOffice\program\soffice.exe') or
     FileExists('C:\Program Files (x86)\LibreOffice\program\soffice.exe') then
  begin
    Result := True;
    Exit;
  end;

  // Fallback 2: Check standard App Paths Registry keys (both 64-bit and 32-bit registry)
  if RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\soffice.exe', '', RegistryPath) or
     RegQueryStringValue(HKCU64, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\soffice.exe', '', RegistryPath) or
     RegQueryStringValue(HKLM32, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\soffice.exe', '', RegistryPath) or
     RegQueryStringValue(HKCU32, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\soffice.exe', '', RegistryPath) then
  begin
    Result := True;
    Exit;
  end;

  Result := False;
end;

function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  Result := True;

  if not CheckLibreOfficeInstalled() then
  begin
    if MsgBox('LibreOffice (soffice) was not found on your system. It is required to convert Word (.doc) and PDF files.' + #13#10#13#10 +
              'Would you like the installer to download and install LibreOffice (Community Edition) silently now?' + #13#10 +
              '(This will run in the background and may take a few minutes depending on your internet connection.)', 
              mbConfirmation, MB_YESNO) = IDYES then
    begin
      // Run PowerShell to download and run LibreOffice MSI installer silently
      ExtractTemporaryFile('{#MyAppExeName}'); // dummy statement to ensure installer UI stays alive
      
      WizardForm.StatusLabel.Caption := 'Downloading and installing LibreOffice dependency... Please wait.';
      
      // We run PowerShell bypass to download LibreOffice 24.2.4 x64 MSI and run it silently
      if Exec('powershell.exe', 
         '-ExecutionPolicy Bypass -Command "& { ' +
         'Write-Host ""Downloading LibreOffice MSI...""; ' +
         '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; ' +
         '$msiPath = Join-Path $env:TEMP ""libreoffice_install.msi""; ' +
         'Invoke-WebRequest -Uri ""https://download.documentfoundation.org/libreoffice/stable/24.2.4/win/x86_64/LibreOffice_24.2.4_Win_x86-64.msi"" -OutFile $msiPath; ' +
         'Write-Host ""Running silent installer...""; ' +
         'Start-Process msiexec.exe -ArgumentList ""/i `""$msiPath`"" /qn /norestart"" -Wait; ' +
         'Remove-Item $msiPath -Force -ErrorAction SilentlyContinue; ' +
         '}"', 
         '', SW_SHOW, ewWaitUntilTerminated, ResultCode) then
      begin
        if CheckLibreOfficeInstalled() then
        begin
          MsgBox('LibreOffice was successfully installed!', mbInformation, MB_OK);
        end
        else
        begin
          MsgBox('Silent installation completed, but LibreOffice was still not detected.' + #13#10 +
                 'You may need to install LibreOffice manually from https://www.libreoffice.org/ later.', mbWarning, MB_OK);
        end;
      end
      else
      begin
        MsgBox('Failed to download LibreOffice. Please install it manually from https://www.libreoffice.org/.', mbError, MB_OK);
      end;
    end;
  end;
end;
