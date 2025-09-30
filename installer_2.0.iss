; ============================
; Death Star 2.0 (World Honey Market)
; ============================

#define MyAppName "The Death Star"
#define MyAppVersion "2.0.0"
#define MyAppPublisher "World Honey Market"
#define MyAppExeName "Death Star.exe"

[Setup]
; Program installs per-user, no admin needed
AppId={{C9A2E9D8-7B1E-4B8E-9CBA-DS-STAR-0002}}   ; keep this stable for WHM upgrades
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\DeathStar
DisableDirPage=yes
DefaultGroupName=Death Star
DisableProgramGroupPage=yes
OutputDir=.\output
OutputBaseFilename=DeathStar2.0-Setup
SetupIconFile=Payload\Seed\deathstar.ico
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=lowest
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
; --- Program (the app EXE you copied into Payload) ---
Source: "Payload\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

; --- Seed user data (first run only; preserved on upgrades) ---
; Goes to: C:\Users\<User>\AppData\Roaming\DeathStarApp
Source: "Payload\Seed\templates.ini";     DestDir: "{userappdata}\DeathStarApp"; Flags: onlyifdoesntexist
Source: "Payload\Seed\kybercrystals.csv"; DestDir: "{userappdata}\DeathStarApp"; Flags: onlyifdoesntexist
Source: "Payload\Seed\README.txt";        DestDir: "{userappdata}\DeathStarApp"; Flags: onlyifdoesntexist

; Icon next to the program for shortcuts
Source: "Payload\Seed\deathstar.ico";     DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu entry
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\deathstar.ico"
; Optional Desktop shortcut (user checks the box during install)
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; IconFilename: "{app}\deathstar.ico"

[Run]
; Show Quick Start after install (optional; remove this line to skip)
Filename: "{userappdata}\DeathStarApp\README.txt"; Description: "View Quick Start Guide"; Flags: postinstall shellexec skipifsilent
; Offer to launch the app
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Keep user data by default
; Type: filesandordirs; Name: "{userappdata}\DeathStarApp"
