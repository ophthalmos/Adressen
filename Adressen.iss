#define MyAppName "Adressen"
#define MyAppVersion "1.0.0.1"

[Setup]
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
VersionInfoVersion={#MyAppVersion}
ArchitecturesAllowed=x64os
ArchitecturesInstallIn64BitMode=x64os
PrivilegesRequired=admin
AppPublisher=Wilhelm Happe
VersionInfoCopyright=(C) 2025, W. Happe
AppPublisherURL=https://www.netradio.info/
AppSupportURL=https://www.netradio.info/
AppUpdatesURL=https://www.netradio.info/
DefaultDirName={autopf}\{#MyAppName}
DisableWelcomePage=yes
DisableDirPage=no
DisableReadyPage=yes
CloseApplications=yes
WizardStyle=modern
WizardSizePercent=100
SetupIconFile=img\Journal.ico
UninstallDisplayIcon={app}\Adressen.exe
DefaultGroupName=Adressen
AppId=Adressen
TimeStampsInUTC=yes
OutputDir=.
OutputBaseFilename={#MyAppName}Setup
Compression=lzma2/max
SolidCompression=yes
DirExistsWarning=no
MinVersion=0,10.0
ChangesAssociations=yes

[Languages]
Name: "German"; MessagesFile: "compiler:Languages\German.isl"

[Files]
Source: "bin\x64\Release\net8.0-windows7.0\Adressen.exe"; DestDir: "{app}"; Permissions: users-modify; Flags: ignoreversion
Source: "bin\x64\Release\net8.0-windows7.0\{#MyAppName}.dll"; DestDir: "{app}"; Permissions: users-modify; Flags: ignoreversion
Source: "bin\x64\Release\net8.0-windows7.0\{#MyAppName}.runtimeconfig.json"; DestDir: "{app}"; Permissions: users-modify; Flags: ignoreversion
Source: "Adressen.pdf"; DestDir: "{app}"; Permissions: users-modify;
Source: "Lizenzvereinbarung.txt"; DestDir: "{app}"; Permissions: users-modify;
Source: "bin\x64\Release\net8.0-windows7.0\System.Data.SQLite.dll"; DestDir: "{app}"; Permissions: users-modify; Flags: ignoreversion
Source: "bin\x64\Release\net8.0-windows7.0\SQLite.Interop.dll"; DestDir: "{app}"; Permissions: users-modify; Flags: ignoreversion
Source: "bin\x64\Release\net8.0-windows7.0\Google.Apis.dll"; DestDir: "{app}"; Permissions: users-modify; Flags: ignoreversion
Source: "bin\x64\Release\net8.0-windows7.0\Google.Apis.Auth.dll"; DestDir: "{app}"; Permissions: users-modify; Flags: ignoreversion
Source: "bin\x64\Release\net8.0-windows7.0\Google.Apis.Core.dll"; DestDir: "{app}"; Permissions: users-modify; Flags: ignoreversion
Source: "bin\x64\Release\net8.0-windows7.0\Google.Apis.PeopleService.v1.dll"; DestDir: "{app}"; Permissions: users-modify; Flags: ignoreversion
Source: "bin\x64\Release\net8.0-windows7.0\Microsoft.Office.Interop.Word.dll"; DestDir: "{app}"; Permissions: users-modify; Flags: ignoreversion
Source: "bin\x64\Release\net8.0-windows7.0\Newtonsoft.Json.dll"; DestDir: "{app}"; Permissions: users-modify; Flags: ignoreversion
Source: "bin\x64\Release\net8.0-windows7.0\System.Management.dll"; DestDir: "{app}"; Permissions: users-modify; Flags: ignoreversion
Source: "bin\x64\Release\net8.0-windows7.0\client_secret.json"; DestDir: "{app}"; Permissions: users-modify; Flags: ignoreversion  
Source: "bin\x64\Release\net8.0-windows7.0\adb_file.ico"; DestDir: "{app}"; Permissions: users-modify; Flags: ignoreversion
Source: "bin\x64\Release\net8.0-windows7.0\LibreHelper\*.*"; DestDir: "{app}\LibreHelper"; Permissions: users-modify; Flags: ignoreversion  
Source: "MännlicheVornamen.txt"; DestDir: "{userappdata}\{#MyAppName}"; Flags: onlyifdoesntexist
Source: "WeiblicheVornamen.txt"; DestDir: "{userappdata}\{#MyAppName}"; Flags: onlyifdoesntexist

[Icons]
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppName}.exe"; Tasks: desktopicon
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppName}.exe"

[Registry]
Root: HKA; Subkey: "Software\Classes\.adb\OpenWithProgids"; ValueType: string; ValueName: "{#MyAppName}.adb"; ValueData: ""; Flags: uninsdeletevalue; Tasks: fileassoc
Root: HKA; Subkey: "Software\Classes\{#MyAppName}.adb"; ValueType: string; ValueName: ""; ValueData: "Adressen-Datenbank"; Flags: uninsdeletekey; Tasks: fileassoc
Root: HKA; Subkey: "Software\Classes\{#MyAppName}.adb\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\adb_file.ico,0"; Tasks: fileassoc
Root: HKA; Subkey: "Software\Classes\{#MyAppName}.adb\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppName}.exe"" ""%1"""; Tasks: fileassoc
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppName}.exe\SupportedTypes"; ValueType: string; ValueName: ".adb"; ValueData: ""; Tasks: fileassoc

[Tasks]
Name: fileassoc; Description: {cm:AssocFileExtension,{#MyAppName},.adb}
Name: desktopicon; Description: {cm:CreateDesktopIcon}


[Run]
Filename: "{app}\{#MyAppName}.exe"; Description: "Starte {#MyAppName}"; Flags: postinstall nowait skipifsilent shellexec

[Messages]
BeveledLabel=
WinVersionTooLowError=Das Programm erfordert eine höhere Windowsversion.
ConfirmUninstall=Möchten Sie '%1' von Ihrem PC entfernen? Eine Deinstallation ist vor einem Update nicht erforderlich.

[CustomMessages]
RemoveSettings=Möchten Sie die Einstellungsdateien ebenfalls entfernen?
Description=Adressen-Datenbank

[Code]
const
  SetupMutexName = 'AdressenSetupMutex';
  
function InitializeSetup(): Boolean; // only one instance of Inno Setup without prompting
begin
  Result := True;
  if CheckForMutexes(SetupMutexName) then
  begin
    Result := False; // Mutex exists, setup is running already, silently aborting
  end
    else
  begin
    CreateMutex(SetupMutexName); 
  end;
end;

procedure CurUninstallStepChanged (CurUninstallStep: TUninstallStep);
var
  mres : integer;
begin
  case CurUninstallStep of                   
    usPostUninstall:
      begin
        mres := MsgBox(CustomMessage('RemoveSettings'), mbConfirmation, MB_YESNO or MB_DEFBUTTON2)
        if mres = IDYES then
          begin
          DelTree(ExpandConstant('{userappdata}\{#MyAppName}'), True, True, True);
          RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Adressen');
          end;
      end;
  end;
end; 

procedure DeinitializeSetup();
var
  FilePath: string;
  BatchPath: string;
  S: TArrayOfString;
  ResultCode: Integer;
begin
  if ExpandConstant('{param:deleteSetup|false}') = 'true' then
  begin
    FilePath := ExpandConstant('{srcexe}');
    begin
      BatchPath := ExpandConstant('{%TEMP}\') + 'delete_' + ExtractFileName(ExpandConstant('{tmp}')) + '.bat';
      SetArrayLength(S, 7);
      S[0] := ':loop';
      S[1] := 'del "' + FilePath + '"';
      S[2] := 'if not exist "' + FilePath + '" goto end';
      S[3] := 'goto loop';
      S[4] := ':end';
      S[5] := 'rd "' + ExpandConstant('{tmp}') + '"';
      S[6] := 'del "' + BatchPath + '"';
      if SaveStringsToFile(BatchPath, S, True) then
      begin
        Exec(BatchPath, '', '', SW_HIDE, ewNoWait, ResultCode)
      end;
    end;
  end;
end;

procedure InitializeWizard;
var
  StaticText: TNewStaticText;
begin
  StaticText := TNewStaticText.Create(WizardForm);
  StaticText.Parent := WizardForm.FinishedPage;
  StaticText.Left := WizardForm.FinishedLabel.Left;
  StaticText.Top := WizardForm.FinishedLabel.Top + 120;
  StaticText.Font.Style := [fsBold];
  StaticText.Caption := 'Ein Zugang zu Google-Kontakten ist derzeit nur'#13'mit eigenen OAuth-Developer-Key möglich.'#13#13 + 
'Speichern Sie die Datei mit folgendem Pfadnamen:'#13'''…\AppData\Roaming\Adressen\client_secret.json''';
end;
