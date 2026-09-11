#define MyAppName "PhotoArchiveCatalog"
#define MyAppVersion "6.0.0"
#define MyAppPublisher "Archive of the President of the Kyrgyz Republic"
#define MyAppExeName "PhotoArchiveCatalog.exe"

[Setup]
AppId={{B2C3B782-C47E-4D12-A742-8CFF4F6B2149}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\PhotoArchiveCatalog
DefaultGroupName=PhotoArchiveCatalog
DisableProgramGroupPage=yes
PrivilegesRequired=admin
OutputDir=installer_output
OutputBaseFilename=PhotoArchiveCatalog_Setup_v{#MyAppVersion}
SetupIconFile=assets\program_icon.ico
UninstallDisplayIcon={app}\PhotoArchiveCatalog.exe
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; GroupDescription: "Ярлыки:"; Flags: unchecked

[Files]
Source: "dist\PhotoArchiveCatalog.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\PhotoArchiveCatalogUpdater.exe"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
Name: "{userdocs}\PhotoArchiveCatalog"
Name: "{userdocs}\PhotoArchiveCatalog\Database"
Name: "{userdocs}\PhotoArchiveCatalog\Photos"
Name: "{userdocs}\PhotoArchiveCatalog\Backups"
Name: "{userdocs}\PhotoArchiveCatalog\Settings"
Name: "{userdocs}\PhotoArchiveCatalog\Updates"
Name: "{userdocs}\PhotoArchiveCatalog\Logs"
Name: "{userdocs}\PhotoArchiveCatalog\Cache"
Name: "{userdocs}\PhotoArchiveCatalog\Cache\Thumbs"

[Icons]
Name: "{group}\PhotoArchiveCatalog"; Filename: "{app}\PhotoArchiveCatalog.exe"; WorkingDir: "{app}"; IconFilename: "{app}\PhotoArchiveCatalog.exe"
Name: "{autodesktop}\PhotoArchiveCatalog"; Filename: "{app}\PhotoArchiveCatalog.exe"; WorkingDir: "{app}"; IconFilename: "{app}\PhotoArchiveCatalog.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\PhotoArchiveCatalog.exe"; Description: "Запустить PhotoArchiveCatalog"; Flags: nowait postinstall skipifsilent
