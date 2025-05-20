[Setup]
AppName=Conversor de PDF para Excel
AppVersion=1.0
DefaultDirName={pf}\MobiConversor
DefaultGroupName=Mobi
OutputBaseFilename=Instalador_Mobi
Compression=lzma
SolidCompression=yes
SetupIconFile=icon.ico

[Files]
Source: "PDF_To_EX.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Conversor de PDF para Excel"; Filename: "{app}\PDF_To_EX.exe"
Name: "{group}\Desinstalar Conversor"; Filename: "{uninstallexe}"
