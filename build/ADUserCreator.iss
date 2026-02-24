#define AppName "ADUserCreator"
#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif

#ifndef SourceDist
  #define SourceDist "..\\dist"
#endif

#ifndef OutputDir
  #define OutputDir "..\\release"
#endif

[Setup]
AppId={{D7F277C5-4A52-4A9E-A82E-51D3CB9F3321}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Local
DefaultDirName=C:\Users\ADUserCreator
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir={#OutputDir}
OutputBaseFilename=ADUserCreator-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin
UninstallDisplayIcon={app}\ADUserCreator.exe

[Languages]
Name: "ukrainian"; MessagesFile: "compiler:Languages\Ukrainian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"; Flags: checkedonce

[Files]
Source: "{#SourceDist}\ADUserCreator.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDist}\main.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDist}\ad\*"; DestDir: "{app}\ad"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceDist}\common\*"; DestDir: "{app}\common"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceDist}\excel\*"; DestDir: "{app}\excel"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#SourceDist}\ui\*"; DestDir: "{app}\ui"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "install-prereqs.ps1"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Icons]
Name: "{autodesktop}\ADUserCreator"; Filename: "{app}\ADUserCreator.exe"; Tasks: desktopicon
Name: "{group}\ADUserCreator"; Filename: "{app}\ADUserCreator.exe"
Name: "{group}\Uninstall ADUserCreator"; Filename: "{uninstallexe}"

[Run]
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{tmp}\install-prereqs.ps1"""; StatusMsg: "Installing PowerShell prerequisites (RSAT/ImportExcel)..."; Flags: waituntilterminated runhidden
Filename: "{app}\ADUserCreator.exe"; Description: "Launch ADUserCreator"; Flags: nowait postinstall skipifsilent shellexec; Verb: runas
