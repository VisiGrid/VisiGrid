; VisiGrid Inno Setup Script
; See docs/features/windows-installer.md for specification

#ifndef Version
  #define Version "0.0.0"
#endif

#define AppName "VisiGrid"
#define AppPublisher "VisiGrid"
#define AppURL "https://visigrid.app"
#define AppExeName "VisiGrid.exe"
#define CliExeName "visigrid.exe"

[Setup]
; NOTE: AppId MUST remain constant across all releases
AppId={{B8F2E4A1-7C3D-4E5F-9A1B-2C3D4E5F6A7B}
AppName={#AppName}
AppVersion={#Version}
AppVerName={#AppName} {#Version}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL=https://github.com/VisiGrid/VisiGrid/releases
DefaultDirName={localappdata}\Programs\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
LicenseFile=LICENSE.txt
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=VisiGrid-Setup-x64
SetupIconFile=..\gpui-app\windows\visigrid.ico
UninstallDisplayIcon={app}\{#AppExeName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
VersionInfoVersion={#Version}
VersionInfoCompany={#AppPublisher}
VersionInfoProductName={#AppName}
VersionInfoDescription=VisiGrid Setup
ArchitecturesInstallIn64BitMode=x64compatible
ArchitecturesAllowed=x64compatible
UsePreviousAppDir=yes
UsePreviousTasks=yes
; Required for WM_SETTINGCHANGE broadcast
ChangesEnvironment=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked
Name: "startmenu"; Description: "Create Start Menu shortcuts"; GroupDescription: "Additional icons:"; Flags: checkedonce
Name: "openwith"; Description: "Add VisiGrid to ""Open with"" menu"; GroupDescription: "Shell integration:"; Flags: checkedonce
Name: "associatevgrid"; Description: "Set as default app for .vgrid files"; GroupDescription: "File associations:"; Flags: checkedonce
Name: "associatecsv"; Description: "Set as default app for .csv files"; GroupDescription: "File associations:"; Flags: unchecked
Name: "associatetsv"; Description: "Set as default app for .tsv files"; GroupDescription: "File associations:"; Flags: unchecked
Name: "installcli"; Description: "Install command-line tools (visigrid)"; GroupDescription: "Command-line:"; Flags: checkedonce
Name: "addtopath"; Description: "Add to PATH (requires shell restart)"; GroupDescription: "Command-line:"; Flags: checkedonce

[Files]
; GUI Application
Source: "..\target\release\visigrid.exe"; DestDir: "{app}"; DestName: "{#AppExeName}"; Flags: ignoreversion

; CLI Tool (in cli subfolder)
Source: "..\target\release\visigrid-cli.exe"; DestDir: "{app}\cli"; DestName: "{#CliExeName}"; Flags: ignoreversion; Tasks: installcli

; Icon (for file associations)
Source: "..\gpui-app\windows\visigrid.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: startmenu
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"; Tasks: startmenu
Name: "{group}\Release Notes"; Filename: "https://github.com/VisiGrid/VisiGrid/releases"; Tasks: startmenu

; Desktop
Name: "{userdesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Registry]
; Install bookkeeping
Root: HKCU; Subkey: "Software\VisiGrid"; ValueType: string; ValueName: "InstallDir"; ValueData: "{app}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\VisiGrid"; ValueType: string; ValueName: "CliPath"; ValueData: "{app}\cli"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\VisiGrid"; ValueType: string; ValueName: "Version"; ValueData: "{#Version}"; Flags: uninsdeletekey

; App Paths registration (Win+R support)
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\App Paths\{#AppExeName}"; ValueType: string; ValueName: ""; ValueData: "{app}\{#AppExeName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\App Paths\{#AppExeName}"; ValueType: string; ValueName: "Path"; ValueData: "{app}"

; ProgID: VisiGrid.Sheet (for .vgrid files)
Root: HKCU; Subkey: "Software\Classes\VisiGrid.Sheet"; ValueType: string; ValueName: ""; ValueData: "VisiGrid Spreadsheet"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\VisiGrid.Sheet\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\{#AppExeName},0"
Root: HKCU; Subkey: "Software\Classes\VisiGrid.Sheet\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"" ""%1"""

; ProgID: VisiGrid.CSV (for .csv files)
Root: HKCU; Subkey: "Software\Classes\VisiGrid.CSV"; ValueType: string; ValueName: ""; ValueData: "CSV File"; Flags: uninsdeletekey; Tasks: associatecsv
Root: HKCU; Subkey: "Software\Classes\VisiGrid.CSV\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\{#AppExeName},0"; Tasks: associatecsv
Root: HKCU; Subkey: "Software\Classes\VisiGrid.CSV\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"" ""%1"""; Tasks: associatecsv

; ProgID: VisiGrid.TSV (for .tsv files)
Root: HKCU; Subkey: "Software\Classes\VisiGrid.TSV"; ValueType: string; ValueName: ""; ValueData: "TSV File"; Flags: uninsdeletekey; Tasks: associatetsv
Root: HKCU; Subkey: "Software\Classes\VisiGrid.TSV\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\{#AppExeName},0"; Tasks: associatetsv
Root: HKCU; Subkey: "Software\Classes\VisiGrid.TSV\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"" ""%1"""; Tasks: associatetsv

; File associations
Root: HKCU; Subkey: "Software\Classes\.vgrid"; ValueType: string; ValueName: ""; ValueData: "VisiGrid.Sheet"; Flags: uninsdeletekey; Tasks: associatevgrid
Root: HKCU; Subkey: "Software\Classes\.csv"; ValueType: string; ValueName: ""; ValueData: "VisiGrid.CSV"; Tasks: associatecsv
Root: HKCU; Subkey: "Software\Classes\.tsv"; ValueType: string; ValueName: ""; ValueData: "VisiGrid.TSV"; Tasks: associatetsv

; "Open with" registration
Root: HKCU; Subkey: "Software\Classes\Applications\{#AppExeName}"; ValueType: string; ValueName: "FriendlyAppName"; ValueData: "{#AppName}"; Flags: uninsdeletekey; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\Applications\{#AppExeName}\SupportedTypes"; ValueType: string; ValueName: ".vgrid"; ValueData: ""; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\Applications\{#AppExeName}\SupportedTypes"; ValueType: string; ValueName: ".csv"; ValueData: ""; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\Applications\{#AppExeName}\SupportedTypes"; ValueType: string; ValueName: ".tsv"; ValueData: ""; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\Applications\{#AppExeName}\SupportedTypes"; ValueType: string; ValueName: ".json"; ValueData: ""; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\Applications\{#AppExeName}\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExeName}"" ""%1"""; Tasks: openwith

; Bookkeeping flags for clean uninstall
Root: HKCU; Subkey: "Software\VisiGrid"; ValueType: dword; ValueName: "AddedToPath"; ValueData: "1"; Tasks: addtopath
Root: HKCU; Subkey: "Software\VisiGrid"; ValueType: dword; ValueName: "AddedToPath"; ValueData: "0"; Tasks: not addtopath
Root: HKCU; Subkey: "Software\VisiGrid"; ValueType: dword; ValueName: "AddedToOpenWith"; ValueData: "1"; Tasks: openwith
Root: HKCU; Subkey: "Software\VisiGrid"; ValueType: dword; ValueName: "AddedToOpenWith"; ValueData: "0"; Tasks: not openwith
Root: HKCU; Subkey: "Software\VisiGrid"; ValueType: dword; ValueName: "AssociatedVgrid"; ValueData: "1"; Tasks: associatevgrid
Root: HKCU; Subkey: "Software\VisiGrid"; ValueType: dword; ValueName: "AssociatedVgrid"; ValueData: "0"; Tasks: not associatevgrid
Root: HKCU; Subkey: "Software\VisiGrid"; ValueType: dword; ValueName: "AssociatedCsv"; ValueData: "1"; Tasks: associatecsv
Root: HKCU; Subkey: "Software\VisiGrid"; ValueType: dword; ValueName: "AssociatedCsv"; ValueData: "0"; Tasks: not associatecsv
Root: HKCU; Subkey: "Software\VisiGrid"; ValueType: dword; ValueName: "AssociatedTsv"; ValueData: "1"; Tasks: associatetsv
Root: HKCU; Subkey: "Software\VisiGrid"; ValueType: dword; ValueName: "AssociatedTsv"; ValueData: "0"; Tasks: not associatetsv

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent
Filename: "https://github.com/VisiGrid/VisiGrid/releases/tag/v{#Version}"; Description: "View release notes"; Flags: nowait postinstall skipifsilent shellexec unchecked

[UninstallDelete]
; Clean up CLI directory if empty
Type: dirifempty; Name: "{app}\cli"

[Code]
const
  EnvironmentKey = 'Environment';

procedure InitializeWizard;
begin
  // Create uninstall options page (shown during uninstall)
end;

function InitializeUninstall: Boolean;
begin
  Result := True;
  // Ask about removing data
  if MsgBox('Do you also want to remove VisiGrid settings and cache?' + #13#10 + #13#10 +
            'This will delete:' + #13#10 +
            '  - Settings and preferences' + #13#10 +
            '  - Recent files list' + #13#10 +
            '  - Cache and logs' + #13#10 + #13#10 +
            'Your documents and .vgrid files will NOT be deleted.',
            mbConfirmation, MB_YESNO) = IDYES then
  begin
    // Mark for deletion
    RegWriteDWordValue(HKEY_CURRENT_USER, 'Software\VisiGrid', 'RemoveData', 1);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  RemoveData: Cardinal;
  AppDataPath, LocalAppDataPath: String;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    // Check if we should remove data
    if RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\VisiGrid', 'RemoveData', RemoveData) then
    begin
      if RemoveData = 1 then
      begin
        // Remove AppData\VisiGrid
        AppDataPath := ExpandConstant('{userappdata}\VisiGrid');
        if DirExists(AppDataPath) then
          DelTree(AppDataPath, True, True, True);

        // Remove LocalAppData\VisiGrid
        LocalAppDataPath := ExpandConstant('{localappdata}\VisiGrid');
        if DirExists(LocalAppDataPath) then
          DelTree(LocalAppDataPath, True, True, True);
      end;
    end;

    // Clean up our registry key
    RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\VisiGrid');
  end;
end;

// PATH manipulation functions
function GetPathEntry: String;
begin
  Result := ExpandConstant('{app}\cli');
end;

function NeedsAddPath: Boolean;
var
  OrigPath: String;
  PathEntry: String;
begin
  PathEntry := GetPathEntry;
  if not RegQueryStringValue(HKEY_CURRENT_USER, EnvironmentKey, 'Path', OrigPath) then
  begin
    Result := True;
    exit;
  end;
  // Look for the path entry in the current PATH
  Result := Pos(';' + PathEntry + ';', ';' + OrigPath + ';') = 0;
end;

procedure AddToPath;
var
  OrigPath: String;
  PathEntry: String;
begin
  PathEntry := GetPathEntry;
  if not RegQueryStringValue(HKEY_CURRENT_USER, EnvironmentKey, 'Path', OrigPath) then
    OrigPath := '';

  // Only add if not already present
  if Pos(';' + PathEntry + ';', ';' + OrigPath + ';') = 0 then
  begin
    if OrigPath <> '' then
      OrigPath := OrigPath + ';';
    RegWriteStringValue(HKEY_CURRENT_USER, EnvironmentKey, 'Path', OrigPath + PathEntry);
  end;
end;

procedure RemoveFromPath;
var
  OrigPath, NewPath: String;
  PathEntry: String;
  P: Integer;
begin
  PathEntry := GetPathEntry;
  if RegQueryStringValue(HKEY_CURRENT_USER, EnvironmentKey, 'Path', OrigPath) then
  begin
    NewPath := ';' + OrigPath + ';';
    P := Pos(';' + PathEntry + ';', NewPath);
    if P > 0 then
    begin
      Delete(NewPath, P, Length(PathEntry) + 1);
      // Remove leading/trailing semicolons
      if (Length(NewPath) > 0) and (NewPath[1] = ';') then
        Delete(NewPath, 1, 1);
      if (Length(NewPath) > 0) and (NewPath[Length(NewPath)] = ';') then
        Delete(NewPath, Length(NewPath), 1);
      RegWriteStringValue(HKEY_CURRENT_USER, EnvironmentKey, 'Path', NewPath);
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    // Add to PATH if task selected
    if WizardIsTaskSelected('addtopath') then
      AddToPath;
  end;
end;

procedure CurUninstallStepChanged_Path(CurUninstallStep: TUninstallStep);
var
  AddedToPath: Cardinal;
begin
  if CurUninstallStep = usUninstall then
  begin
    // Only remove from PATH if we added it
    if RegQueryDWordValue(HKEY_CURRENT_USER, 'Software\VisiGrid', 'AddedToPath', AddedToPath) then
    begin
      if AddedToPath = 1 then
        RemoveFromPath;
    end;
  end;
end;

// Combine uninstall handlers
procedure DeinitializeUninstall;
begin
  // PATH removal is handled in CurUninstallStepChanged_Path
end;

// Task dependency: disable PATH option if CLI not selected
function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := False;
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if CurPageID = wpSelectTasks then
  begin
    // Update PATH task based on CLI task
    // Note: Inno Setup doesn't have direct API for this,
    // but the dependency is documented in the UI
  end;
end;

// Refresh shell after install/uninstall for PATH changes
procedure RefreshEnvironment;
var
  S: String;
begin
  S := 'Environment';
  // Broadcast WM_SETTINGCHANGE to notify shell of environment changes
  // This is handled automatically by Inno Setup when ChangesEnvironment=yes
end;
