; ----------------------------------------------
; FluxPDF Windows Installer
; ----------------------------------------------

[Setup]
AppName=FluxPDF
AppVersion=1.0.1
DefaultDirName={pf}\FluxPDF
DefaultGroupName=FluxPDF
OutputDir=installer
OutputBaseFilename=FluxPDF_Installer
SetupIconFile=app.ico
Compression=lzma
SolidCompression=yes

[Files]
Source: "dist\FluxPDF\FluxPDF.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{commondesktop}\FluxPDF"; Filename: "{app}\FluxPDF.exe"
Name: "{group}\FluxPDF"; Filename: "{app}\FluxPDF.exe"
