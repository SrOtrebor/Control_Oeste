; Inno Setup Script para Control de Acceso - Al Oeste Shopping
#define MyAppName "Control de Acceso - Al Oeste Shopping"
#define MyAppVersion "2026.1"
#define MyAppPublisher "Roberto Laforcada"
#define MyAppExeName "ControlAcceso.exe"

[Setup]
AppId={{D37F2C0B-9B84-4B81-9F32-8415F49B90D8}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName=C:\Control de Acceso Oeste
DisableDirPage=no
DefaultGroupName=Control de Acceso Oeste
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename=Instalador_ControlAcceso_v2026
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Dirs]
Name: "{app}"; Permissions: users-full
Name: "{app}\registros_diarios"; Permissions: users-full
Name: "{app}\registros_fichajes"; Permissions: users-full
Name: "{app}\logs"; Permissions: users-full
Name: "{app}\backups"; Permissions: users-full

[Files]
; Ejecutable principal compilado
Source: "dist\ControlAcceso.exe"; DestDir: "{app}"; Flags: ignoreversion

; Archivos Excel de datos (no sobreescribir si ya existen para no borrar datos de guardias)
Source: "ListadoFAPs.xlsx"; DestDir: "{app}"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "ListadoFAOs.xlsx"; DestDir: "{app}"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "excepciones.xlsx"; DestDir: "{app}"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "nominas_persistentes.xlsx"; DestDir: "{app}"; Flags: onlyifdoesntexist uninsneveruninstall
Source: ".env.example"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
