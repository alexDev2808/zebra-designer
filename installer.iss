; Instalador de Etiquetas Zebra (Inno Setup 6.3 o superior).
; Requiere haber generado antes dist\EtiquetasZebra con build.bat.

#define AppName "Etiquetas Zebra"
#define AppVersion "1.0.0"
#define AppExe "EtiquetasZebra.exe"

[Setup]
AppId={{8F3C2A51-6B4E-4C2D-9A71-3E5D2B7C9F10}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher=J. Alexis
DefaultDirName={autopf}\Etiquetas Zebra
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=dist
OutputBaseFilename=EtiquetasZebra-Setup-{#AppVersion}
SetupIconFile=assets\icono.ico
UninstallDisplayIcon={app}\{#AppExe}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "dist\EtiquetasZebra\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "ejemplo.xlsx"; DestDir: "{app}\Ejemplos"; Flags: ignoreversion
Source: "plantilla_ejemplo.json"; DestDir: "{app}\Ejemplos"; Flags: ignoreversion
Source: "docs\MANUAL_DE_USO.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{group}\Desinstalar {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent
