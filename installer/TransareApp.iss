#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif

#define MyAppName "TransareApp"
#define MyAppPublisher "TransareApp"
#define MyAppExeName "TransareApp.exe"
#define MyAppAssocName MyAppName + " Installer"

[Setup]
AppId={{B7F6D980-8B5A-4E4E-B84A-92A5A6F9AAB6}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\dist\installer
OutputBaseFilename=TransareApp-Setup
SetupIconFile=..\build\transareapp.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Creează o iconiță pe desktop"; GroupDescription: "Opțiuni suplimentare:"

[Files]
Source: "..\dist\TransareApp\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Lansează {#MyAppName}"; Flags: nowait postinstall skipifsilent
