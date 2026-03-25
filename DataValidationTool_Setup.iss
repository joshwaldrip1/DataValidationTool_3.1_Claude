; Inno Setup script for Data Validation Tool v3.1
; Default install dir: Program Files (admin) or %LOCALAPPDATA%\Programs (per-user)

#define AppName "Data Validation Tool"
#define AppVersion "3.1"
#define AppPublisher "Topographic Land Surveyors"
#define AppExeName "DataValidationTool.exe"
#define SourceDir "dist\DataValidationTool"

[Setup]
AppId={{A2F1C3D4-5E6F-7A8B-9C0D-1E2F3A4B5C6D}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL=
DefaultDirName={autodesktop}\DataValidationTool
DefaultGroupName={#AppName}
DisableDirPage=no
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename=DataValidationTool_v3.1_Setup
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
; Per-user install — no UAC prompt needed
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=
; Show a license page if you have one (optional — remove if not needed)
; LicenseFile=LICENSE.txt
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName} {#AppVersion}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"

[Files]
; Copy everything from the PyInstaller onedir output
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Start Menu shortcut
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"
; Desktop shortcut (optional, unchecked by default)
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
; Offer to launch the app after install
Filename: "{app}\{#AppExeName}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Remove config.json and any other files the app writes at runtime
Type: files; Name: "{app}\config.json"
Type: files; Name: "{app}\validation_log.json"
Type: files; Name: "{app}\crdb_watchlist.json"
