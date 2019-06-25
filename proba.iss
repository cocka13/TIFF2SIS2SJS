; SETUP za TIFF2SIS2SJS verzija 1.3.0
; 11.11.2005. Autor: Mladen Kolarek


[Setup]
AppName=TIFF2SIS2SJS
AppVerName=TIFF2SIS2SJS ver.1.3.0
DefaultGroupName=ZZF_Software\TIFF2SIS2SJS
AppPublisher=ZZF
AppPublisherURL=http://www.zzf.hr/
AppVersion=1.3.0
AllowNoIcons=yes
;InfoBeforeFile=Setup.txt
;InfoAfterFile=ReadMe.txt
WizardImageFile=ZZF_Setup.bmp
AppCopyright=c2005 Sva prava pridrzana / All rights reserved
PrivilegesRequired=none
OutputBaseFilename=Tiff2Sis2Sjs130
DefaultDirName={pf}\ZZF_Software\Tiff2Sis2Sjs
LicenseFile=license.txt

[Tasks]
Name: desktopicon; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"
Name: quicklaunchicon; Description: "Create a &Quick Launch icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "c:\program files\randem systems\innoscript\innoscript 2.2\vb6 runtime\olepro32.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver restartreplace sharedfile
Source: "c:\program files\randem systems\innoscript\innoscript 2.2\vb6 runtime\comcat.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver restartreplace sharedfile
Source: "c:\program files\randem systems\innoscript\innoscript 2.2\vb6 runtime\stdole2.tlb"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  uninsneveruninstall restartreplace sharedfile regtypelib
Source: "c:\program files\randem systems\innoscript\innoscript 2.2\vb6 runtime\asycfilt.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  restartreplace sharedfile
Source: "c:\program files\randem systems\innoscript\innoscript 2.2\vb6 runtime\oleaut32.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver restartreplace sharedfile
Source: "c:\program files\randem systems\innoscript\innoscript 2.2\vb6 runtime\msvbvm60.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver restartreplace sharedfile
Source: "c:\program files\randem systems\innoscript\innoscript 2.2\vb6 runtime\vb6stkit.dll"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  restartreplace sharedfile
Source: "c:\WINDOWS\system32\comdlg32.ocx"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "c:\WINDOWS\system32\tabctl32.ocx"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "c:\WINDOWS\system32\mscomctl.ocx"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "c:\WINDOWS\system32\richtx32.ocx"; DestDir: "{sys}"; MinVersion: 4.0,4.0; Flags:  regserver sharedfile
Source: "D:\Mladen_PRIVATE\VB_Project\Tiff2Sis2Sjs\tiff2sis2sjs.exe"; DestDir: "{app}"; MinVersion: 4.0,4.0; Flags:  ignoreversion
Source: "D:\Mladen_PRIVATE\VB_Project\Tiff2Sis2Sjs\s2j.bat"; DestDir: "{app}"; MinVersion: 4.0,4.0; Flags:  ignoreversion
Source: "D:\Mladen_PRIVATE\VB_Project\Tiff2Sis2Sjs\T2S.bat"; DestDir: "{app}"; MinVersion: 4.0,4.0; Flags:  ignoreversion
Source: "D:\Mladen_PRIVATE\VB_Project\Tiff2Sis2Sjs\SDI2JPEG.exe"; DestDir: "{app}"; MinVersion: 4.0,4.0; Flags:  ignoreversion
Source: "D:\Mladen_PRIVATE\VB_Project\Tiff2Sis2Sjs\TIFF2SDI.exe"; DestDir: "{app}"; MinVersion: 4.0,4.0; Flags:  ignoreversion
Source: "D:\Mladen_PRIVATE\VB_Project\Tiff2Sis2Sjs\translators.pdf"; DestDir: "{app}"; MinVersion: 4.0,4.0; Flags:  ignoreversion
Source: "D:\Mladen_PRIVATE\VB_Project\Tiff2Sis2Sjs\license.txt"; DestDir: "{app}"; MinVersion: 4.0,4.0; Flags:  ignoreversion

[INI]
Filename: "{app}\ZZF Web page.url"; Section: "InternetShortcut"; Key: "URL"; String: "http://www.zzf.hr"

[Icons]
Name: "{group}\License"; Filename: "{app}\license.txt"; WorkingDir: "{app}"
Name: "{group}\TIFF2SIS2SJS"; Filename: "{app}\Tiff2Sis2Sjs.exe"; WorkingDir: "{app}"
Name: "{group}\About ISM Translators"; Filename: "{app}\Translators.pdf"; WorkingDir: "{app}"
Name: "{group}\Zavod za fotogrametriju on the Web"; Filename: "{app}\ZZF Web page.url"
Name: "{group}\Uninstall TIFF2SIS2SJS"; Filename: "{uninstallexe}"
Name: "{userdesktop}\Tiff2Sis2Sjs"; Filename: "{app}\Tiff2Sis2Sjs.exe"; Tasks: desktopicon; WorkingDir: "{app}"
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\Tiff2Sis2Sjs"; Filename: "{app}\Tiff2Sis2Sjs.exe"; Tasks: quicklaunchicon; WorkingDir: "{app}"

[Run]
Filename: "{app}\Tiff2Sis2Sjs.exe"; Description: "Launch TIFF2SIS2SJS"; Flags: nowait postinstall skipifsilent; WorkingDir: "{app}"

[UninstallDelete]
Type: files; Name: "{app}\ZZF Web page.url"
Type: dirifempty; Name: "{app}"
