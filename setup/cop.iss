; Скрипт инсталлятора Inno Setup 5.2.3
; Программа "Конфигуратор УП"
; Автор: Мезенцев Вячеслав Николаевич, Зыков Василий
; E-mail: unihomelab@ya.ru
; www: http://vkontatke.ru/viacheslavmezentsev

[Setup]
AppName=Конфигуратор УП
AppVerName=Конфигуратор УП версия 1.3.141
DefaultDirName={pf}\Конфигуратор УП
DefaultGroupName=Конфигуратор УП
Compression=lzma
SolidCompression=true
OutputDir=Output
OutputBaseFilename=Configurator_1.3.141_setup
Uninstallable=true
UninstallDisplayIcon={app}\cop.exe
UninstallFilesDir={app}\Uninstall
RestartIfNeededByRun=false
CreateUninstallRegKey=false
ShowLanguageDialog=no
LanguageDetectionMethod=none
WizardImageFile=WizardImage.bmp

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
Source: ButtonImage.bmp; DestDir: {tmp}; Flags: dontcopy

[Icons]
Name: {group}\Конфигуратор УП; Filename: {app}\cop.exe; WorkingDir: {app}
Name: {group}\Руководство пользователя; Filename: {app}\cop.chm; WorkingDir: {app}
Name: {group}\Удалить; Filename: {app}\Uninstall\unins000; WorkingDir: {app}

[Code]
Const
ButtonWidth = 77;       //размер кнопок
ButtonHeight = 25;
ButtonFontColor = $000000;   //цвет шрифта кнопок
PageColor = $ab663d;         //цвет страниц
FontColor = $ffffff;         //цвет шрифта
MainTextBackColor = $663300;  //цвет заднего фона текста сверху
BeveledLabelFontColor = clBlue;  //цвет текста в нижнем левом углу
bidBack = 0;
bidNext = 1;
bidCancel = 2;
bidDirBrowse = 3;
bidGroupBrowse = 4;

Var
ButtonPanel: array [0..4] of TPanel;
ButtonImage: array [0..4] of TBitmapImage;
ButtonLabel: array [0..4] of TLabel;
BeveledLabel: TLabel;
WizardButtonPanel,BrowseButtonPanel: TPanel;
WizardButtonImage,BrowseButtonImage: TBitmapImage;
WizardButtonLabel,WizardButtonLabel2,BrowseButtonLabel,BrowseButtonLabel2: TLabel;
LicenseAcceptedText,LicenseNotAcceptedText,NoIconsText,YesRadioText,NoRadioText: TNewStatictext;

Procedure LicenseAcceptedOnClick (Sender: TObject);
begin
WizardForm.LicenseAcceptedRadio.Checked:=True
ButtonLabel[bidNext].Enabled:=True
end;

Procedure LicenseNotAcceptedOnClick (Sender: TObject);
begin
WizardForm.LicenseNotAcceptedRadio.Checked:=True
ButtonLabel[bidNext].Enabled:=False
end;

Procedure NoIconsLabelOnClick (Sender: TObject);
begin
WizardForm.NoIconsCheck.Checked:=Not(WizardForm.NoIconsCheck.Checked)
end;

Procedure YesRadioOnClick (Sender: TObject);
begin
WizardForm.YesRadio.Checked:=True
end;

Procedure NoRadioOnClick (Sender: TObject);
begin
WizardForm.NoRadio.Checked:=True
end;

procedure WizardButtonOnClick(Sender: TObject);
var
Button: TButton;
begin
ButtonImage[TLabel(Sender).Tag].Left:=0
case TLabel(Sender).Tag of
bidBack: Button:=WizardForm.BackButton
bidNext: Button:=WizardForm.NextButton
bidCancel: Button:=WizardForm.CancelButton
end
Button.OnClick(Button)
end;

procedure BrowseButtonOnClick(Sender: TObject);
var
Button: TButton;
begin
ButtonImage[TLabel(Sender).Tag].Left:=0
case TLabel(Sender).Tag of
bidDirBrowse: Button:=WizardForm.DirBrowseButton
bidGroupBrowse: Button:=WizardForm.GroupBrowseButton
end
Button.OnClick(Button)
end;

procedure WizardButtonMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
ButtonImage[TLabel(Sender).Tag].Left:=-ButtonWidth
end;

procedure WizardButtonMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
ButtonImage[TLabel(Sender).Tag].Left:=0
end;

procedure BrowseButtonMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
ButtonImage[TLabel(Sender).Tag].Left:=-ButtonWidth
end;

procedure BrowseButtonMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
ButtonImage[TLabel(Sender).Tag].Left:=0
end;

procedure LoadWizardButtonImage(AButton: TButton; AButtonIndex: integer);
begin
WizardButtonPanel:=TPanel.Create(WizardForm)
WizardButtonPanel.Left:=AButton.Left
WizardButtonPanel.Top:=AButton.Top
WizardButtonPanel.Width:=AButton.Width
WizardButtonPanel.Height:=AButton.Height
WizardButtonPanel.Tag:=AButtonIndex
WizardButtonPanel.Parent:=AButton.Parent
ButtonPanel[AButtonIndex]:=WizardButtonPanel

WizardButtonImage:=TBitmapImage.Create(WizardForm)
WizardButtonImage.Autosize:=True
WizardButtonImage.Enabled:=False
WizardButtonImage.Bitmap.LoadFromFile(ExpandConstant('{tmp}\ButtonImage.bmp'))
WizardButtonImage.Parent:=WizardButtonPanel
ButtonImage[AButtonIndex]:=WizardButtonImage

WizardButtonLabel:=TLabel.Create(WizardForm)
WizardButtonLabel.Tag:=AButtonIndex
WizardButtonLabel.Width:=WizardButtonPanel.Width
WizardButtonLabel.Height:=WizardButtonPanel.Height
WizardButtonLabel.Autosize:=False
WizardButtonLabel.Transparent:=True
WizardButtonLabel.OnClick:=@WizardButtonOnClick
WizardButtonLabel.OnDblClick:=@WizardButtonOnClick
WizardButtonLabel.OnMouseDown:=@WizardButtonMouseDown
WizardButtonLabel.OnMouseUp:=@WizardButtonMouseUp
WizardButtonLabel.Parent:=WizardButtonPanel

WizardButtonLabel2:=TLabel.Create(WizardForm)
WizardButtonLabel2.Top:=6
WizardButtonLabel2.Width:=WizardButtonPanel.Width
WizardButtonLabel2.Height:=WizardButtonPanel.Height
WizardButtonLabel2.Autosize:=False
WizardButtonLabel2.Alignment:=taCenter
WizardButtonLabel2.Tag:=AButtonIndex
WizardButtonLabel2.Transparent:=True
WizardButtonLabel2.Font.Color:=ButtonFontColor
WizardButtonLabel2.Caption:=AButton.Caption
WizardButtonLabel2.OnClick:=@WizardButtonOnClick
WizardButtonLabel2.OnDblClick:=@WizardButtonOnClick
WizardButtonLabel2.OnMouseDown:=@WizardButtonMouseDown
WizardButtonLabel2.OnMouseUp:=@WizardButtonMouseUp
WizardButtonLabel2.Parent:=WizardButtonPanel
ButtonLabel[AButtonIndex]:=WizardButtonLabel2
end;

procedure LoadBrowseButtonImage(AButton: TButton; AButtonIndex: integer);
begin
BrowseButtonPanel:=TPanel.Create(WizardForm)
BrowseButtonPanel.Left:=AButton.Left
BrowseButtonPanel.Top:=AButton.Top
BrowseButtonPanel.Width:=AButton.Width
BrowseButtonPanel.Height:=AButton.Height
BrowseButtonPanel.Tag:=AButtonIndex
BrowseButtonPanel.Parent:=AButton.Parent
ButtonPanel[AButtonIndex]:=BrowseButtonPanel

BrowseButtonImage:=TBitmapImage.Create(WizardForm)
BrowseButtonImage.Top:=-ButtonHeight
BrowseButtonImage.Autosize:=True
BrowseButtonImage.Enabled:=False
BrowseButtonImage.Bitmap.LoadFromFile(ExpandConstant('{tmp}\ButtonImage.bmp'))
BrowseButtonImage.Parent:=BrowseButtonPanel
ButtonImage[AButtonIndex]:=BrowseButtonImage

BrowseButtonLabel:=TLabel.Create(WizardForm)
BrowseButtonLabel.Tag:=AButtonIndex
BrowseButtonLabel.Width:=BrowseButtonPanel.Width
BrowseButtonLabel.Height:=BrowseButtonPanel.Height
BrowseButtonLabel.Autosize:=False
BrowseButtonLabel.Transparent:=True
BrowseButtonLabel.OnClick:=@BrowseButtonOnClick
BrowseButtonLabel.OnDblClick:=@BrowseButtonOnClick
BrowseButtonLabel.OnMouseDown:=@BrowseButtonMouseDown
BrowseButtonLabel.OnMouseUp:=@BrowseButtonMouseUp
BrowseButtonLabel.Parent:=BrowseButtonPanel

BrowseButtonLabel2:=TLabel.Create(WizardForm)
BrowseButtonLabel2.Top:=6
BrowseButtonLabel2.Width:=BrowseButtonPanel.Width
BrowseButtonLabel2.Height:=BrowseButtonPanel.Height
BrowseButtonLabel2.Autosize:=False
BrowseButtonLabel2.Alignment:=taCenter
BrowseButtonLabel2.Tag:=AButtonIndex
BrowseButtonLabel2.Transparent:=True
BrowseButtonLabel2.Font.Color:=ButtonFontColor
BrowseButtonLabel2.Caption:=AButton.Caption
BrowseButtonLabel2.OnClick:=@BrowseButtonOnClick
BrowseButtonLabel2.OnDblClick:=@BrowseButtonOnClick
BrowseButtonLabel2.OnMouseDown:=@BrowseButtonMouseDown
BrowseButtonLabel2.OnMouseUp:=@BrowseButtonMouseUp
BrowseButtonLabel2.Parent:=BrowseButtonPanel
ButtonLabel[AButtonIndex]:=BrowseButtonLabel2
end;

procedure UpdateWizardButton(AButton: TButton;AButtonIndex: integer);
begin
ButtonLabel[AButtonIndex].Caption:=AButton.Caption
ButtonPanel[AButtonIndex].Visible:=AButton.Visible
ButtonLabel[AButtonIndex].Enabled:=Abutton.Enabled
end;

procedure UpdateButton();
begin
UpdateWizardButton(WizardForm.BackButton,bidBack)
UpdateWizardButton(WizardForm.NextButton,bidNext)
UpdateWizardButton(WizardForm.CancelButton,bidCancel)
end;

Procedure InitializeSkin();
begin
with WizardForm do
  with OuterNotebook do
    with InnerPage do
      with InnerNotebook do
        Color:=PageColor

ExtractTemporaryFile('ButtonImage.bmp')

//WizardForm
WizardForm.Bevel.Hide
WizardForm.Bevel1.Hide
WizardForm.MainPanel.Hide
WizardForm.WizardSmallBitmapImage.Hide
WizardForm.WizardBitmapImage2.Hide
WizardForm.BeveledLabel.Hide
WizardForm.BeveledLabel.Left:=700

WizardForm.ClientWidth:=690
WizardForm.ClientHeight:=496
WizardForm.Center

WizardForm.WizardBitmapImage.Left:=0
WizardForm.WizardBitmapImage.Top:=0
WizardForm.WizardBitmapImage.Width:=690
WizardForm.WizardBitmapImage.Height:=496
WizardForm.WizardBitmapImage.Parent:=WizardForm

WizardForm.BackButton.Left:=293
WizardForm.BackButton.Top:=462
WizardForm.BackButton.Width:=ButtonWidth
WizardForm.BackButton.Height:=ButtonHeight

WizardForm.NextButton.Left:=375
WizardForm.NextButton.Top:=462
WizardForm.NextButton.Width:=ButtonWidth
WizardForm.NextButton.Height:=ButtonHeight

WizardForm.CancelButton.Left:=600
WizardForm.CancelButton.Top:=462
WizardForm.CancelButton.Width:=ButtonWidth
WizardForm.CancelButton.Height:=ButtonHeight

WizardForm.OuterNotebook.Left:=200
WizardForm.OuterNotebook.Top:=80
WizardForm.OuterNotebook.Width:=460
WizardForm.OuterNotebook.Height:=354

WizardForm.InnerNotebook.Left:=0
WizardForm.InnerNotebook.Top:=0
WizardForm.InnerNotebook.Width:=460
WizardForm.InnerNotebook.Height:=354

WizardForm.PageNameLabel.Left:=15
WizardForm.PageNameLabel.Top:=7
WizardForm.PageNameLabel.Autosize:=True
WizardForm.PageNameLabel.WordWrap:=False
WizardForm.PageNameLabel.Color:=MainTextBackColor
WizardForm.PageNameLabel.Font.Color:=FontColor
WizardForm.PageNameLabel.Parent:=WizardForm

WizardForm.PageDescriptionLabel.Left:=25
WizardForm.PageDescriptionLabel.Top:=25
WizardForm.PageDescriptionLabel.Autosize:=True
WizardForm.PageDescriptionLabel.WordWrap:=False
WizardForm.PageDescriptionLabel.Color:=MainTextBackColor
WizardForm.PageDescriptionLabel.Font.Color:=FontColor
WizardForm.PageDescriptionLabel.Parent:=WizardForm

BeveledLabel:=TLabel.Create(WizardForm)
BeveledLabel.Left:=10
BeveledLabel.Top:=468
BeveledLabel.Transparent:=True
BeveledLabel.Font.Color:=BeveledLabelFontColor
BeveledLabel.Caption:=WizardForm.BeveledLabel.Caption
BeveledLabel.Parent:=WizardForm

//wpWelcome
WizardForm.WelcomePage.Color:=PageColor

WizardForm.WelcomeLabel1.Left:=0
WizardForm.WelcomeLabel1.Top:=110
WizardForm.WelcomeLabel1.Width:=465
WizardForm.WelcomeLabel1.Height:=28
WizardForm.WelcomeLabel1.Font.Size:=8
WizardForm.WelcomeLabel1.Font.Color:=FontColor

WizardForm.WelcomeLabel2.Left:=0
WizardForm.WelcomeLabel2.Top:=150
WizardForm.WelcomeLabel2.Width:=465
WizardForm.WelcomeLabel2.Height:=200
WizardForm.WelcomeLabel2.Font.Color:=FontColor

//wpLicense
WizardForm.LicenseLabel1.Left:=0
WizardForm.LicenseLabel1.Top:=0
WizardForm.LicenseLabel1.Width:=460
WizardForm.LicenseLabel1.Height:=28
WizardForm.LicenseLabel1.Font.Color:=FontColor

WizardForm.LicenseMemo.Left:=0
WizardForm.LicenseMemo.Top:=38
WizardForm.LicenseMemo.Width:=460
WizardForm.LicenseMemo.Height:=266

WizardForm.LicenseAcceptedRadio.Left:=0
WizardForm.LicenseAcceptedRadio.Top:=318
WizardForm.LicenseAcceptedRadio.Width:=17
WizardForm.LicenseAcceptedRadio.Height:=17

LicenseAcceptedText:=TNewStatictext.Create(WizardForm)
LicenseAcceptedText.Left:=17
LicenseAcceptedText.Top:=321
LicenseAcceptedText.Font.Color:=FontColor
LicenseAcceptedText.Caption:=WizardForm.LicenseAcceptedRadio.Caption
LicenseAcceptedText.OnClick:=@LicenseAcceptedOnClick
LicenseAcceptedText.Parent:=WizardForm.LicensePage

WizardForm.LicenseNotAcceptedRadio.Left:=0
WizardForm.LicenseNotAcceptedRadio.Top:=338
WizardForm.LicenseNotAcceptedRadio.Width:=17
WizardForm.LicenseNotAcceptedRadio.Height:=17

LicenseNotAcceptedText:=TNewStatictext.Create(WizardForm)
LicenseNotAcceptedText.Left:=17
LicenseNotAcceptedText.Top:=341
LicenseNotAcceptedText.Font.Color:=FontColor
LicenseNotAcceptedText.Caption:=WizardForm.LicenseNotAcceptedRadio.Caption
LicenseNotAcceptedText.OnClick:=@LicenseNotAcceptedOnClick
LicenseNotAcceptedText.Parent:=WizardForm.LicensePage

//wpPassword
WizardForm.PasswordLabel.Left:=0
WizardForm.PasswordLabel.Top:=0
WizardForm.PasswordLabel.Width:=460
WizardForm.PasswordLabel.Height:=28
WizardForm.PasswordLabel.Font.Color:=FontColor

WizardForm.PasswordEditLabel.Left:=0
WizardForm.PasswordEditLabel.Top:=34
WizardForm.PasswordEditLabel.Width:=460
WizardForm.PasswordEditLabel.Height:=14
WizardForm.PasswordEditLabel.Font.Color:=FontColor

WizardForm.PasswordEdit.Left:=0
WizardForm.PasswordEdit.Top:=50
WizardForm.PasswordEdit.Width:=460
WizardForm.PasswordEdit.Height:=21
WizardForm.PasswordEdit.Color:=$ffffff
WizardForm.PasswordEdit.Font.Color:=$000000

//wpInfoBefore
WizardForm.InfoBeforeClickLabel.Left:=0
WizardForm.InfoBeforeClickLabel.Top:=0
WizardForm.InfoBeforeClickLabel.Width:=460
WizardForm.InfoBeforeClickLabel.Height:=14
WizardForm.InfoBeforeClickLabel.Font.Color:=FontColor

WizardForm.InfoBeforeMemo.Left:=0
WizardForm.InfoBeforeMemo.Top:=24
WizardForm.InfoBeforeMemo.Width:=460
WizardForm.InfoBeforeMemo.Height:=327

//wpUserInfo
WizardForm.UserInfoNameLabel.Left:=0
WizardForm.UserInfoNameLabel.Top:=0
WizardForm.UserInfoNameLabel.Width:=460
WizardForm.UserInfoNameLabel.Height:=14
WizardForm.UserInfoNameLabel.Font.Color:=FontColor

WizardForm.UserInfoNameEdit.Left:=0
WizardForm.UserInfoNameEdit.Top:=16
WizardForm.UserInfoNameEdit.Width:=460
WizardForm.UserInfoNameEdit.Height:=21
WizardForm.UserInfoNameEdit.Color:=$ffffff
WizardForm.UserInfoNameEdit.Font.Color:=$000000

WizardForm.UserInfoOrgLabel.Left:=0
WizardForm.UserInfoOrgLabel.Top:=52
WizardForm.UserInfoOrgLabel.Width:=460
WizardForm.UserInfoOrgLabel.Height:=14
WizardForm.UserInfoOrgLabel.Font.Color:=FontColor

WizardForm.UserInfoOrgEdit.Left:=0
WizardForm.UserInfoOrgEdit.Top:=68
WizardForm.UserInfoOrgEdit.Width:=460
WizardForm.UserInfoOrgEdit.Height:=21
WizardForm.UserInfoOrgEdit.Color:=$ffffff
WizardForm.UserInfoOrgEdit.Font.Color:=$000000

WizardForm.UserInfoSerialLabel.Left:=0
WizardForm.UserInfoSerialLabel.Top:=104
WizardForm.UserInfoSerialLabel.Width:=460
WizardForm.UserInfoSerialLabel.Height:=14
WizardForm.UserInfoSerialLabel.Font.Color:=FontColor

WizardForm.UserInfoSerialEdit.Left:=0
WizardForm.UserInfoSerialEdit.Top:=120
WizardForm.UserInfoSerialEdit.Width:=460
WizardForm.UserInfoSerialEdit.Height:=21
WizardForm.UserInfoSerialEdit.Color:=$ffffff
WizardForm.UserInfoSerialEdit.Font.Color:=$000000

//wpSelectDir
WizardForm.SelectDirBitmapImage.Hide

WizardForm.SelectDirLabel.Left:=0
WizardForm.SelectDirLabel.Top:=0
WizardForm.SelectDirLabel.Width:=460
WizardForm.SelectDirLabel.Height:=14
WizardForm.SelectDirLabel.Font.Color:=FontColor

WizardForm.SelectDirBrowseLabel.Left:=0
WizardForm.SelectDirBrowseLabel.Top:=24
WizardForm.SelectDirBrowseLabel.Width:=460
WizardForm.SelectDirBrowseLabel.Height:=28
WizardForm.SelectDirBrowseLabel.Font.Color:=FontColor

WizardForm.DirEdit.Left:=0
WizardForm.DirEdit.Top:=290
WizardForm.DirEdit.Width:=370
WizardForm.DirEdit.Height:=21
WizardForm.DirEdit.Color:=$ffffff
WizardForm.DirEdit.Font.Color:=$000000

WizardForm.DirBrowseButton.Left:=383
WizardForm.DirBrowseButton.Top:=289
WizardForm.DirBrowseButton.Width:=ButtonWidth
WizardForm.DirBrowseButton.Height:=ButtonHeight

WizardForm.DiskSpaceLabel.Left:=0
WizardForm.DiskSpaceLabel.Top:=340
WizardForm.DiskSpaceLabel.Width:=460
WizardForm.DiskSpaceLabel.Height:=14
WizardForm.DiskSpaceLabel.Font.Color:=FontColor

//wpSelectComponents
WizardForm.SelectComponentsLabel.Left:=0
WizardForm.SelectComponentsLabel.Top:=0
WizardForm.SelectComponentsLabel.Width:=460
WizardForm.SelectComponentsLabel.Height:=14
WizardForm.SelectComponentsLabel.Font.Color:=FontColor

WizardForm.TypesCombo.Left:=0
WizardForm.TypesCombo.Top:=24
WizardForm.TypesCombo.Width:=460
WizardForm.TypesCombo.Height:=21
WizardForm.TypesCombo.Color:=$ffffff
WizardForm.TypesCombo.Font.Color:=$000000

WizardForm.ComponentsList.Left:=0
WizardForm.ComponentsList.Top:=48
WizardForm.ComponentsList.Width:=460
WizardForm.ComponentsList.Height:=275
WizardForm.ComponentsList.Color:=$ffffff
WizardForm.ComponentsList.Font.Color:=$000000

WizardForm.ComponentsDiskSpaceLabel.Left:=0
WizardForm.ComponentsDiskSpaceLabel.Top:=340
WizardForm.ComponentsDiskSpaceLabel.Width:=417
WizardForm.ComponentsDiskSpaceLabel.Height:=14
WizardForm.ComponentsDiskSpaceLabel.Font.Color:=FontColor

//wpSelectProgramGroup
WizardForm.SelectGroupBitmapImage.Hide

WizardForm.SelectStartMenuFolderLabel.Left:=0
WizardForm.SelectStartMenuFolderLabel.Top:=0
WizardForm.SelectStartMenuFolderLabel.Width:=460
WizardForm.SelectStartMenuFolderLabel.Height:=14
WizardForm.SelectStartMenuFolderLabel.Font.Color:=FontColor

WizardForm.SelectStartMenuFolderBrowseLabel.Left:=0
WizardForm.SelectStartMenuFolderBrowseLabel.Top:=24
WizardForm.SelectStartMenuFolderBrowseLabel.Width:=460
WizardForm.SelectStartMenuFolderBrowseLabel.Height:=28
WizardForm.SelectStartMenuFolderBrowseLabel.Font.Color:=FontColor

WizardForm.GroupEdit.Left:=0
WizardForm.GroupEdit.Top:=290
WizardForm.GroupEdit.Width:=370
WizardForm.GroupEdit.Height:=21
WizardForm.GroupEdit.Color:=$ffffff
WizardForm.GroupEdit.Font.Color:=$000000

WizardForm.GroupBrowseButton.Left:=383
WizardForm.GroupBrowseButton.Top:=289
WizardForm.GroupBrowseButton.Width:=ButtonWidth
WizardForm.GroupBrowseButton.Height:=ButtonHeight

WizardForm.NoIconsCheck.Left:=0
WizardForm.NoIconsCheck.Top:=337
WizardForm.NoIconsCheck.Width:=17
WizardForm.NoIconsCheck.Height:=17
WizardForm.NoIconsCheck.Visible:=True

NoIconsText:=TNewStatictext.Create(WizardForm)
NoIconsText.Left:=17
NoIconsText.Top:=340
NoIconsText.Font.Color:=FontColor
NoIconsText.OnClick:=@NoIconsLabelOnClick
NoIconsText.Caption:=WizardForm.NoIconsCheck.Caption
NoIconsText.Parent:=WizardForm.SelectProgramGroupPage

//wpSelectTasks
WizardForm.SelectTasksLabel.Left:=0
WizardForm.SelectTasksLabel.Top:=0
WizardForm.SelectTasksLabel.Width:=460
WizardForm.SelectTasksLabel.Height:=28
WizardForm.SelectTasksLabel.Font.Color:=FontColor

WizardForm.TasksList.Left:=0
WizardForm.TasksList.Top:=34
WizardForm.TasksList.Width:=460
WizardForm.TasksList.Height:=317
WizardForm.TasksList.Color:=PageColor
WizardForm.TasksList.Font.Color:=FontColor

//wpReady
WizardForm.ReadyLabel.Left:=0
WizardForm.ReadyLabel.Top:=0
WizardForm.ReadyLabel.Width:=460
WizardForm.ReadyLabel.Height:=28
WizardForm.ReadyLabel.Font.Color:=FontColor

WizardForm.ReadyMemo.Left:=0
WizardForm.ReadyMemo.Top:=34
WizardForm.ReadyMemo.Width:=460
WizardForm.ReadyMemo.Height:=317
WizardForm.ReadyMemo.Color:=PageColor
WizardForm.ReadyMemo.Font.Color:=FontColor

//wpInstalling
WizardForm.StatusLabel.Left:=0
WizardForm.StatusLabel.Top:=0
WizardForm.StatusLabel.Width:=460
WizardForm.StatusLabel.Height:=16
WizardForm.StatusLabel.Font.Color:=FontColor

WizardForm.FilenameLabel.Left:=0
WizardForm.FilenameLabel.Top:=16
WizardForm.FilenameLabel.Width:=460
WizardForm.FilenameLabel.Height:=16
WizardForm.FilenameLabel.Font.Color:=FontColor

WizardForm.ProgressGauge.Left:=0
WizardForm.ProgressGauge.Top:=42
WizardForm.ProgressGauge.Width:=460
WizardForm.ProgressGauge.Height:=21

//wpInfoAfter
WizardForm.InfoAfterClickLabel.Left:=0
WizardForm.InfoAfterClickLabel.Top:=0
WizardForm.InfoAfterClickLabel.Width:=460
WizardForm.InfoAfterClickLabel.Height:=14
WizardForm.InfoAfterClickLabel.Font.Color:=FontColor

WizardForm.InfoAfterMemo.Left:=0
WizardForm.InfoAfterMemo.Top:=24
WizardForm.InfoAfterMemo.Width:=460
WizardForm.InfoAfterMemo.Height:=327

//wpFinished
WizardForm.FinishedPage.Color:=PageColor

WizardForm.FinishedHeadingLabel.Left:=0
WizardForm.FinishedHeadingLabel.Top:=79
WizardForm.FinishedHeadingLabel.Width:=460
WizardForm.FinishedHeadingLabel.Height:=24
WizardForm.FinishedHeadingLabel.Font.Size:=8
WizardForm.FinishedHeadingLabel.Font.Color:=FontColor

WizardForm.FinishedLabel.Left:=0
WizardForm.FinishedLabel.Top:=119
WizardForm.FinishedLabel.Width:=460
WizardForm.FinishedLabel.Height:=53
WizardForm.FinishedLabel.Font.Color:=FontColor

WizardForm.RunList.Left:=0
WizardForm.RunList.Top:=199
WizardForm.RunList.Width:=460
WizardForm.RunList.Height:=149
WizardForm.RunList.Font.Color:=FontColor

WizardForm.YesRadio.Left:=0
WizardForm.YesRadio.Top:=199
WizardForm.YesRadio.Width:=460
WizardForm.YesRadio.Height:=17
WizardForm.YesRadio.OnClick:=@YesRadioOnClick

YesRadioText:=TNewStatictext.Create(WizardForm)
YesRadioText.Left:=16
YesRadioText.Top:=2
YesRadioText.Width:=445
YesRadioText.Height:=15
YesRadioText.Font.Color:=FontColor
YesRadioText.Caption:=WizardForm.YesRadio.Caption
YesRadioText.OnClick:=@YesRadioOnClick
YesRadioText.Parent:=WizardForm.YesRadio

WizardForm.NoRadio.Left:=0
WizardForm.NoRadio.Top:=227
WizardForm.NoRadio.Width:=460
WizardForm.NoRadio.Height:=17
WizardForm.NoRadio.OnClick:=@NoRadioOnClick

NoRadioText:=TNewStatictext.Create(WizardForm)
NoRadioText.Left:=16
NoRadioText.Top:=2
NoRadioText.Width:=445
NoRadioText.Height:=15
NoRadioText.Font.Color:=FontColor
NoRadioText.Caption:=WizardForm.NoRadio.Caption
NoRadioText.OnClick:=@NoRadioOnClick
NoRadioText.Parent:=WizardForm.NoRadio

LoadWizardButtonImage(WizardForm.BackButton,bidBack)
LoadWizardButtonImage(WizardForm.NextButton,bidNext)
LoadWizardButtonImage(WizardForm.CancelButton,bidCancel)
LoadBrowseButtonImage(WizardForm.DirBrowseButton,bidDirBrowse)
LoadBrowseButtonImage(WizardForm.GroupBrowseButton,bidGroupBrowse)
end;

procedure InitializeWizard;
begin

	InitializeSkin

end;

procedure CurPageChanged(CurPageID: Integer);
begin

	UpdateButton

	if CurPageID=wpLicense then begin

		If WizardForm.FindComponent('NextButton') is TButton then
			TButton(WizardForm.FindComponent('NextButton')).Caption:='Далее >';

		If WizardForm.FindComponent('CancelButton') is TButton then
			TButton(WizardForm.FindComponent('CancelButton')).Caption:='Отмена';

		If WizardForm.FindComponent('OuterNotebook') is TNewNotebook then
			TNewNotebook(WizardForm.FindComponent('OuterNotebook')).Height:=ScaleY(313);

	end

end;

[_ISToolPreCompile]
Name: D:\Projects\vbasic\Configurator\pack.bat; Parameters: ; Flags: runminimized
