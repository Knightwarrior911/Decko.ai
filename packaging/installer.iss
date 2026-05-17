[Setup]
AppName=Decko
AppVersion=1.0.0
DefaultDirName={autopf}\Decko
DefaultGroupName=Decko
OutputBaseFilename=Decko-Setup
Compression=lzma2
SolidCompression=yes

[Files]
Source: "..\dist\Decko.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Decko"; Filename: "{app}\Decko.exe"
Name: "{commondesktop}\Decko"; Filename: "{app}\Decko.exe"
