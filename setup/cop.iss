; ������ ������������ Inno Setup 5.4.2
; ��������� "������������ ��"
; �����: �������� �������� ����������, ����� �������
; E-mail: unihomelab@ya.ru
; Skype: viacheslavmezentsev
; www: http://vkontatke.ru/viacheslavmezentsev

#define Version GetFileVersion("..\cop.exe")
#define WizardImage "WizardImage.bmp"
#define ButtonImage "ButtonImage.bmp"

#include "Skin.iss"

[Setup]
AppName=������������ ��
AppVerName=������������ �� ������ {#Version}
DefaultDirName={pf}\������������ ��
DefaultGroupName=������������ ��
AppendDefaultDirName=true
DirExistsWarning=no
Compression=lzma/Max
SolidCompression=true
OutputDir=Output
OutputBaseFilename=Configurator-{#Version}-win32-setup
Uninstallable=true
UninstallDisplayIcon={app}\cop.exe
UninstallFilesDir={app}\Uninstall
RestartIfNeededByRun=false
CreateUninstallRegKey=false
ShowLanguageDialog=no
LanguageDetectionMethod=none
WizardImageFile={#WizardImage}
;InfoBeforeFile=MyInfoBefore.txt
VersionInfoVersion={#Version}
VersionInfoDescription=��������� "������������ ��"
VersionInfoProductName=������������ ��
VersionInfoProductVersion={#Version}

[Languages]
Name: ru; MessagesFile: compiler:Languages\Russian.isl

[Messages]
ru.BeveledLabel=Russian
BeveledLabel=Copyright 2012
LicenseLabel3=

[Files]
; begin VB system files
Source: "..\..\..\vbrun60sp5\stdole2.tlb";  DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: "..\..\..\vbrun60sp5\msvbvm60.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "..\..\..\vbrun60sp5\oleaut32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "..\..\..\vbrun60sp5\olepro32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "..\..\..\vbrun60sp5\asycfilt.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile
Source: "..\..\..\vbrun60sp5\comcat.dll";   DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
; End VB system files

; Specific control you included In your project(s)
; always use the following parameters For an OCX:
;  DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Windows\system32\comct332.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver; 
Source: "C:\Windows\system32\mscomctl.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver; 
Source: "C:\Windows\system32\mscomct2.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver; 
Source: "C:\Windows\system32\comdlg32.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver; 
Source: "C:\Windows\system32\msflxgrd.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver; 
Source: "C:\Windows\system32\tabctl32.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver; 

Source: "C:\Windows\system32\scrrun.dll";   DestDir: "{app}"; Flags: ignoreversion

; Some external controls I use all of the time
; always use the following parameters For an OCX:
;  DestDir: "{sys}"; Flags: restartreplace sharedfile regserver

; ����������� ����
Source: "..\cop.exe";       DestDir: "{app}"; Flags: ignoreversion
Source: "..\help\cop.chm";  DestDir: "{app}"; Flags: ignoreversion
Source: "..\limits.ini";    DestDir: "{app}"; Flags: ignoreversion

Source: {#ButtonImage}; DestDir: "{tmp}"; Flags: dontcopy

[Icons]
Name: "{group}\������������ ��"; Filename: "{app}\cop.exe"; WorkingDir: "{app}"
Name: "{group}\����������� ������������"; Filename: "{app}\cop.chm"; WorkingDir: "{app}"
Name: "{group}\�������"; Filename: "{app}\Uninstall\unins000"; WorkingDir: "{app}"

[Code]
procedure InitializeWizard;
begin

	InitializeSkin

end;

procedure CurPageChanged(CurPageID: Integer);
begin

	UpdateButton

	if CurPageID=wpLicense then begin

		If WizardForm.FindComponent('NextButton') is TButton then
			TButton(WizardForm.FindComponent('NextButton')).Caption:='����� >';

		If WizardForm.FindComponent('CancelButton') is TButton then
			TButton(WizardForm.FindComponent('CancelButton')).Caption:='������';

		If WizardForm.FindComponent('OuterNotebook') is TNewNotebook then
			TNewNotebook(WizardForm.FindComponent('OuterNotebook')).Height:=ScaleY(313);

	end

end;

[_ISToolPreCompile]
Name: D:\Projects\vbasic\Configurator\pack.bat; Parameters: ; Flags: runminimized

[InnoIDE_PreCompile]
Name: D:\Projects\vbasic\Projects\Configurator\pack.bat; Flags: RunMinimized AbortOnError; 
Name: D:\Projects\vbasic\Projects\Configurator\makehelp.bat; Flags: RunMinimized AbortOnError;
