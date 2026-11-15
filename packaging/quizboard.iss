; Inno Setup script: the Windows installer for Quizboard.
; Built by .github/workflows/release.yml from PyInstaller's dist\Quizboard folder:
;   iscc /DAppVersion=1.0.0 packaging\quizboard.iss
; Installs for the current user only (no administrator rights), adds a Start-menu
; entry and an uninstaller (Settings > Apps).

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif

[Setup]
AppId={{2267C286-FA80-5DF9-B156-DC76E0B95350}
AppName=Quizboard
AppVersion={#AppVersion}
AppPublisher=Zachary Klaas
AppPublisherURL=https://github.com/zklaas/quizboard
DefaultDirName={localappdata}\Programs\Quizboard
DefaultGroupName=Quizboard
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\release
OutputBaseFilename=Quizboard-Setup-Windows
SetupIconFile=..\icons\quizboard.ico
UninstallDisplayIcon={app}\Quizboard.exe
LicenseFile=..\LICENSE
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible

[Files]
Source: "..\dist\Quizboard\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\Quizboard"; Filename: "{app}\Quizboard.exe"
Name: "{userdesktop}\Quizboard"; Filename: "{app}\Quizboard.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; Flags: unchecked

[Run]
Filename: "{app}\Quizboard.exe"; Description: "Start Quizboard"; Flags: nowait postinstall skipifsilent
