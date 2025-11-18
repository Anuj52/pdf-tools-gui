; ----------------------------------------------
; FluxPDF Windows Installer - FIXED
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
; Copy EVERYTHING inside dist\FluxPDF\
Source: "dist\FluxPDF\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{commondesktop}\FluxPDF"; Filename: "{app}\FluxPDF.exe"
Name: "{group}\FluxPDF"; Filename: "{app}\FluxPDF.exe"
