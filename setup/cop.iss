; Скрипт инсталлятора Inno Setup 5.1.14
; Программа "Конфигуратор УП"
; Автор: Мезенцев Вячеслав Николаевич, Зыков Василий
; E-mail: unihomelab@ya.ru
; www: http://vkontatke.ru/viacheslavmezentsev

[Setup]
AppName=Конфигуратор УП
AppVerName=Конфигуратор УП версия 1.3.0
DefaultDirName={pf}\Конфигуратор УП
DefaultGroupName=Конфигуратор УП
UninstallDisplayIcon={app}\cop.exe
Compression=lzma
SolidCompression=true
OutputDir=Output
;LicenseFile=License.rtf
OutputBaseFilename=Configurator_1.3.0_setup

[Languages]
;Name: en; MessagesFile: compiler:Default.isl
Name: ru; MessagesFile: compiler:Languages\Russian.isl

[Messages]
;en.BeveledLabel=English
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
Source: ..\Comdlg32.ocx; DestDir: {app}; Flags: ignoreversion
Source: ..\Mscomctl.ocx; DestDir: {app}; Flags: ignoreversion
Source: ..\Msflxgrd.ocx; DestDir: {app}; Flags: ignoreversion
;Source: License.rtf; DestDir: {tmp}; Flags: dontcopy noencryption

[Icons]
Name: {group}\Конфигуратор УП; Filename: {app}\cop.exe

[Code]
var
  License: String;

procedure BevelLabelClick(Sender: TObject);
var ErrorCode: Integer;
begin
//  ShellExec('open', 'http://www.rospipe.ru', '', '', SW_SHOW, ewNoWait, ErrorCode );
end;

procedure InitializeWizard;
var
 ii: integer;
begin
  { Create the pages }

//  ExtractTemporaryFile( 'License.rtf' );
//  LoadStringFromFile(ExpandConstant( '{tmp}\License.rtf'), License );
//  with WizardForm.LicenseMemo do begin
//    Lines.Clear;
//    Left := 0; Top := WizardForm.LicenseLabel1.Top + WizardForm.LicenseLabel1.Height;
//    Width := WizardForm.LicensePage.Width;
//    Height := WizardForm.LicenseAcceptedRadio.Top - Top - 5;
//    RTFText := License;
//  end;

  with WizardForm.BeveledLabel do begin
    OnClick := @BevelLabelClick;
    Font.Color := clBlue;
    Enabled := True;
    Cursor := crHand;
//    Hint := 'Конфигуратор УП';
//    ShowHint := True;
  end;
//  WizardForm.LicenseAcceptedRadio.Visible := True;
//  WizardForm.LicenseNotAcceptedRadio.Visible := True;
//  WizardForm.LicenseAcceptedRadio.Checked := False;
//  WizardForm.LicenseNotAcceptedRadio.Checked := True;
end;

procedure CurPageChanged(CurPageID: Integer);
begin
if CurPageID=wpLicense then
 begin
  If WizardForm.FindComponent('NextButton') is TButton then
  TButton(WizardForm.FindComponent('NextButton')).Caption:='Далее >';
  If WizardForm.FindComponent('CancelButton') is TButton then
  TButton(WizardForm.FindComponent('CancelButton')).Caption:='Отмена';
  If WizardForm.FindComponent('OuterNotebook') is TNewNotebook then
  TNewNotebook(WizardForm.FindComponent('OuterNotebook')).Height:=ScaleY(313);
 end
end;
