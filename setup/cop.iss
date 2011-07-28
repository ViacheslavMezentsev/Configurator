; ������ ������������ Inno Setup 5.4.2
; ��������� "������������ ��"
; �����: �������� �������� ����������, ����� �������
; E-mail: unihomelab@ya.ru
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
Compression=lzma
SolidCompression=true
OutputDir=Output
OutputBaseFilename=Configurator_{#Version}_setup
Uninstallable=true
UninstallDisplayIcon={app}\cop.exe
UninstallFilesDir={app}\Uninstall
RestartIfNeededByRun=false
CreateUninstallRegKey=false
ShowLanguageDialog=no
LanguageDetectionMethod=none
WizardImageFile={#WizardImage}

[Languages]
Name: ru; MessagesFile: compiler:Languages\Russian.isl

[Messages]
ru.BeveledLabel=Russian
BeveledLabel =Copyright 2011
LicenseLabel3=

[Files]
Source: ..\cop.exe; DestDir: {app}; Flags: ignoreversion
Source: ..\help\cop.chm; DestDir: {app}; Flags: ignoreversion
Source: ..\limits.ini; DestDir: {app}; Flags: ignoreversion
Source: ..\msvbvm60.dll; DestDir: {app}; Flags: ignoreversion
Source: ..\scrrun.dll; DestDir: {app}; Flags: ignoreversion
Source: ..\comct332.ocx; DestDir: {app}; Flags: ignoreversion
Source: ..\mscomctl.ocx; DestDir: {app}; Flags: ignoreversion
Source: ..\mscomct2.ocx; DestDir: {app}; Flags: ignoreversion
Source: ..\comdlg32.ocx; DestDir: {app}; Flags: ignoreversion
Source: ..\msflxgrd.ocx; DestDir: {app}; Flags: ignoreversion
Source: ..\tabctl32.ocx; DestDir: {app}; Flags: ignoreversion
Source: {#ButtonImage}; DestDir: {tmp}; Flags: dontcopy

[Icons]
Name: {group}\������������ ��; Filename: {app}\cop.exe; WorkingDir: {app}
Name: {group}\����������� ������������; Filename: {app}\cop.chm; WorkingDir: {app}
Name: {group}\�������; Filename: {app}\Uninstall\unins000; WorkingDir: {app}

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
