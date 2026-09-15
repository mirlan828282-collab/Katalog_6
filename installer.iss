#define MyAppName "ФОТОКАТАЛОГ"
#define MyAppVersion "6.0.8"
#define MyAppPublisher "Archive of the President of the Kyrgyz Republic"
#define MyAppExeName "PhotoArchiveCatalog.exe"

[Setup]
AppId={{B2C3B782-C47E-4D12-A742-8CFF4F6B2149}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\PhotoArchiveCatalog
DefaultGroupName=ФОТОКАТАЛОГ
DisableProgramGroupPage=yes
PrivilegesRequired=admin
OutputDir=installer_output
OutputBaseFilename=PhotoArchiveCatalog_Setup_v{#MyAppVersion}
SetupIconFile=assets\program_icon.ico
UninstallDisplayIcon={app}\program_icon.ico
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
Source: "assets\program_icon.ico"; DestDir: "{app}"; DestName: "program_icon.ico"; Flags: ignoreversion

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

[InstallDelete]
Type: files; Name: "{group}\PhotoArchiveCatalog.lnk"
Type: files; Name: "{group}\ФОТОКАТАЛОГ.lnk"
Type: files; Name: "{autodesktop}\PhotoArchiveCatalog.lnk"
Type: files; Name: "{autodesktop}\ФОТОКАТАЛОГ.lnk"

[Icons]
Name: "{group}\ФОТОКАТАЛОГ"; Filename: "{app}\PhotoArchiveCatalog.exe"; WorkingDir: "{app}"; IconFilename: "{app}\program_icon.ico"; IconIndex: 0
Name: "{autodesktop}\ФОТОКАТАЛОГ"; Filename: "{app}\PhotoArchiveCatalog.exe"; WorkingDir: "{app}"; IconFilename: "{app}\program_icon.ico"; IconIndex: 0; Tasks: desktopicon

[Run]
Filename: "{app}\PhotoArchiveCatalog.exe"; Description: "Запустить ФОТОКАТАЛОГ"; Flags: nowait postinstall skipifsilent
