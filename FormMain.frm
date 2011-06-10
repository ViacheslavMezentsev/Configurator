VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FormMain 
   Caption         =   "Шаблон проекта"
   ClientHeight    =   6420
   ClientLeft      =   2532
   ClientTop       =   1944
   ClientWidth     =   8976
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   8976
   Begin VB.Timer Timer1 
      Left            =   2040
      Top             =   5520
   End
   Begin VB.Frame SplitterLeft 
      BorderStyle     =   0  'None
      Height          =   5052
      Left            =   2400
      MousePointer    =   9  'Size W E
      TabIndex        =   5
      Top             =   360
      Width           =   60
   End
   Begin MSComDlg.CommonDialog SaveFileDialog 
      Left            =   1560
      Top             =   5520
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog OpenFileDialog 
      Left            =   1200
      Top             =   5520
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      DefaultExt      =   "*.bin"
      DialogTitle     =   "Открыть"
      Filter          =   "Файлы проекта|*.bin"
      FilterIndex     =   1
   End
   Begin VB.Frame SplitterRight 
      BorderStyle     =   0  'None
      Height          =   5052
      Left            =   6516
      MousePointer    =   9  'Size W E
      TabIndex        =   4
      Top             =   360
      Width           =   60
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   312
      Left            =   0
      TabIndex        =   3
      Top             =   6108
      Width           =   8976
      _ExtentX        =   15833
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10915
            MinWidth        =   1834
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1842
            MinWidth        =   1834
            Text            =   "Изменён"
            TextSave        =   "Изменён"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameMain 
      Caption         =   "Шаги"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5052
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   3972
      Begin VB.Frame FrameCodeView 
         BorderStyle     =   0  'None
         Height          =   2172
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Visible         =   0   'False
         Width           =   3732
         Begin VB.TextBox TextByte 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   120
            MaxLength       =   3
            TabIndex        =   13
            Top             =   1800
            Visible         =   0   'False
            Width           =   612
         End
         Begin MSFlexGridLib.MSFlexGrid CodeView 
            Height          =   1572
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   2772
            _ExtentX        =   4890
            _ExtentY        =   2773
            _Version        =   393216
            Cols            =   17
            GridLines       =   0
            BorderStyle     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame FrameGridView 
         BorderStyle     =   0  'None
         Height          =   1812
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3732
         Begin MSFlexGridLib.MSFlexGrid StepsView 
            Height          =   1452
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   2292
            _ExtentX        =   4043
            _ExtentY        =   2561
            _Version        =   393216
            Rows            =   16
            Cols            =   81
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            BorderStyle     =   0
            Appearance      =   0
         End
      End
   End
   Begin VB.Frame FrameLeft 
      Caption         =   "Программы"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5052
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2172
      Begin VB.TextBox TextName 
         BorderStyle     =   0  'None
         Height          =   288
         Left            =   120
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   4680
         Visible         =   0   'False
         Width           =   732
      End
      Begin MSFlexGridLib.MSFlexGrid ListPrograms 
         Height          =   4332
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1932
         _ExtentX        =   3408
         _ExtentY        =   7641
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollBars      =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         BorderStyle     =   0
      End
   End
   Begin VB.Frame FrameRight 
      Caption         =   "Свойства"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5052
      Left            =   6600
      TabIndex        =   0
      Top             =   360
      Width           =   2292
      Begin VB.ComboBox ComboCell 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   4560
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.TextBox TextCell 
         BorderStyle     =   0  'None
         Height          =   288
         Left            =   120
         TabIndex        =   11
         Top             =   4200
         Visible         =   0   'False
         Width           =   732
      End
      Begin MSFlexGridLib.MSFlexGrid PropertyTable 
         Height          =   3852
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2052
         _ExtentX        =   3620
         _ExtentY        =   6795
         _Version        =   393216
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         BorderStyle     =   0
      End
      Begin VB.Label LabelDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   1200
         TabIndex        =   16
         Top             =   4320
         Visible         =   0   'False
         Width           =   972
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   1320
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageListMainToolbar 
      Left            =   0
      Top             =   5520
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":6432
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":6786
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":6ADA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListSquares 
      Left            =   480
      Top             =   5520
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":6E2E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   288
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8976
      _ExtentX        =   15833
      _ExtentY        =   508
      ButtonWidth     =   487
      ButtonHeight    =   466
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageListMainToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Новый"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Открыть"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Сохранить"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Image ImageGrayed 
      Height          =   192
      Left            =   3120
      Picture         =   "FormMain.frx":7182
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image ImageChecked 
      Height          =   192
      Left            =   2880
      Picture         =   "FormMain.frx":74E2
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image ImageUnchecked 
      Appearance      =   0  'Flat
      Height          =   192
      Left            =   2640
      Picture         =   "FormMain.frx":7854
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Menu FileMainMenuItem 
      Caption         =   "&Файл"
      Begin VB.Menu NewMainMenuItem 
         Caption         =   "&Новый"
         Shortcut        =   ^N
      End
      Begin VB.Menu OpenMainMenuItem 
         Caption         =   "&Открыть..."
         Shortcut        =   ^O
      End
      Begin VB.Menu SaveMainMenuItem 
         Caption         =   "&Сохранить"
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveAsMainMenuItem 
         Caption         =   "Сохранить &как..."
      End
      Begin VB.Menu CloseMainMenuItem 
         Caption         =   "&Закрыть"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu Separator 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMainMenuItem 
         Caption         =   "&Выход"
      End
   End
   Begin VB.Menu PopupMenuPrograms 
      Caption         =   "П&рограмма"
      Begin VB.Menu PopupMenuListClear 
         Caption         =   "&Очистить"
      End
      Begin VB.Menu CopyMainMenuItem 
         Caption         =   "&Копировать..."
      End
      Begin VB.Menu Separator2 
         Caption         =   "-"
      End
      Begin VB.Menu PopupMenuListClearAll 
         Caption         =   "Очистить &все"
      End
   End
   Begin VB.Menu StepMainMenuItem 
      Caption         =   "&Шаг"
      Begin VB.Menu InsertStepMenuItem 
         Caption         =   "&Вставить"
      End
      Begin VB.Menu DeleteStepMenuItem 
         Caption         =   "&Удалить"
      End
   End
   Begin VB.Menu CodeMainMenuItem 
      Caption         =   "&Код"
      Visible         =   0   'False
   End
   Begin VB.Menu HelpMainMenuItem 
      Caption         =   "&Помощь"
      Begin VB.Menu HelpMainMenuSubItem 
         Caption         =   "&Справка"
         Shortcut        =   {F1}
      End
      Begin VB.Menu AboutMainMenuItem 
         Caption         =   "&О программе"
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TITLE_SECTION_NAME = "Title"

' Режим отображения средней панели
Private ViewMode As Integer
' Режимы отображения таблицы шагов
Private StepsViewMode As Integer

Private CurrentDir As String
Private FileName As String

Public IniFile As TIniFiles
Public Manager As TProgramManager

Public ModuleIdle As TModuleIdle
Public ModuleFill As TModuleFill
Public ModuleDTRG As TModuleDTRG
Public ModuleHeat As TModuleHeat
Public ModuleWashOrRinsOrJolt As TModuleWashOrRinsOrJolt
'Public ModulePause As TModulePause
Public ModuleDrain As TModuleDrain
Public ModuleSpin As TModuleSpin
Public ModuleCool As TModuleCool
Public ModuleTrin As TModuleTrin

' Настройки по умолчанию
' Дискретные типы
Private Const ENDSOUND_DEFAULT = 1
Private Const DOORUNLOCK_DEFAULT = 1

Public LimitsLoaded As Boolean

Dim EndSound As TYPE_BOOL_DESCRIPTION
Dim DoorUnlock As TYPE_BOOL_DESCRIPTION

Dim SplitterRightMoving As Boolean
Dim SplitterLeftMoving As Boolean
Dim BegX As Integer, BegY As Integer

Private Sub SavePlacement()
    ' Размеры формы
    IniFile.WriteInteger "Placement", "Left", Left
    IniFile.WriteInteger "Placement", "Top", Top
    IniFile.WriteInteger "Placement", "Width", Width
    IniFile.WriteInteger "Placement", "Height", Height
    
    ' Размеры и положение компонентов
    IniFile.WriteInteger "Placement", "FrameLeft.Left", FrameLeft.Left
    IniFile.WriteInteger "Placement", "FrameLeft.Top", FrameLeft.Top
    IniFile.WriteInteger "Placement", "FrameLeft.Width", FrameLeft.Width
    IniFile.WriteInteger "Placement", "FrameLeft.Height", FrameLeft.Height
    
    IniFile.WriteInteger "Placement", "SplitterLeft.Left", SplitterLeft.Left
    IniFile.WriteInteger "Placement", "SplitterLeft.Height", SplitterLeft.Height
    
    IniFile.WriteInteger "Placement", "FrameMain.Left", FrameMain.Left
    IniFile.WriteInteger "Placement", "FrameMain.Top", FrameMain.Top
    IniFile.WriteInteger "Placement", "FrameMain.Width", FrameMain.Width
    IniFile.WriteInteger "Placement", "FrameMain.Height", FrameMain.Height
    
    IniFile.WriteInteger "Placement", "SplitterRight.Left", SplitterRight.Left
    IniFile.WriteInteger "Placement", "SplitterRight.Height", SplitterRight.Height
    
    IniFile.WriteInteger "Placement", "FrameRight.Left", FrameRight.Left
    IniFile.WriteInteger "Placement", "FrameRight.Top", FrameRight.Top
    IniFile.WriteInteger "Placement", "FrameRight.Width", FrameRight.Width
    IniFile.WriteInteger "Placement", "FrameRight.Height", FrameRight.Height
    
    ' Прочие настройки
    IniFile.WriteString "Settings", "CurrentDir", CurrentDir
End Sub

Private Sub LoadPlacement()
    ' Размеры формы
    Left = IniFile.ReadInteger("Placement", "Left", 324)
    Top = IniFile.ReadInteger("Placement", "Top", 324)
    Width = IniFile.ReadInteger("Placement", "Width", 9072)
    Height = IniFile.ReadInteger("Placement", "Height", 7092)
    
    ' Размеры и положение компонентов
    FrameLeft.Left = IniFile.ReadInteger("Placement", "FrameLeft.Left", 0)
    FrameLeft.Top = IniFile.ReadInteger("Placement", "FrameLeft.Top", 360)
    FrameLeft.Width = IniFile.ReadInteger("Placement", "FrameLeft.Width", 2172)
    If FrameLeft.Width < 500 Then FrameLeft.Width = 500
    FrameLeft.Height = IniFile.ReadInteger("Placement", "FrameLeft.Height", 5052)
    
    SplitterLeft.Left = IniFile.ReadInteger("Placement", "SplitterLeft.Left", 2400)
    SplitterLeft.Height = IniFile.ReadInteger("Placement", "SplitterLeft.Height", 5052)
    
    FrameMain.Left = IniFile.ReadInteger("Placement", "FrameMain.Left", 2640)
    FrameMain.Top = IniFile.ReadInteger("Placement", "FrameMain.Top", 360)
    FrameMain.Width = IniFile.ReadInteger("Placement", "FrameMain.Width", 3612)
    FrameMain.Height = IniFile.ReadInteger("Placement", "FrameMain.Height", 5052)
    
    SplitterRight.Left = IniFile.ReadInteger("Placement", "SplitterRight.Left", 6396)
    SplitterRight.Height = IniFile.ReadInteger("Placement", "SplitterRight.Height", 5052)
    
    FrameRight.Left = IniFile.ReadInteger("Placement", "FrameRight.Left", 6600)
    FrameRight.Top = IniFile.ReadInteger("Placement", "FrameRight.Top", 360)
    FrameRight.Width = IniFile.ReadInteger("Placement", "FrameRight.Width", 2292)
    If FrameRight.Width < 500 Then FrameRight.Width = 500
    FrameRight.Height = IniFile.ReadInteger("Placement", "FrameRight.Height", 5052)
    
    ' Прочие настройки
    Dim Path As String
    Dim Result%
    
    Path = String$(255, 0)
    Result = GetModuleFileName(0, Path, 254)
    Path = MiscExtractPathName(Path, True)
    
    CurrentDir = IniFile.ReadString("Settings", "CurrentDir", Path)
End Sub

Private Sub RefreshComponents(ByVal FramesOnly As Boolean)
    If Me.WindowState = vbMinimized Then Exit Sub
    
    ' Обновление данных в компонентах
    If Not FramesOnly Then RefreshDataComponents
    
    FrameLeft.Top = Me.ScaleTop + Toolbar1.Top + Toolbar1.Height
    FrameLeft.Height = Me.ScaleHeight - (StatusBar.Height + Toolbar1.Top + Toolbar1.Height)
        
    SplitterLeft.Left = FrameLeft.Left + FrameLeft.Width
    SplitterLeft.Top = FrameLeft.Top + 100
    SplitterLeft.Height = FrameLeft.Height - 100
    
    FrameMain.Left = SplitterLeft.Left + SplitterLeft.Width
    FrameMain.Top = FrameLeft.Top
    FrameMain.Height = FrameLeft.Height
    FrameMain.Width = Me.ScaleWidth - FrameMain.Left - FrameRight.Width - SplitterRight.Width
    
    SplitterRight.Left = FrameMain.Left + FrameMain.Width
    SplitterRight.Top = FrameLeft.Top + 100
    SplitterRight.Height = FrameLeft.Height - 100
    
    FrameRight.Left = Me.ScaleWidth - FrameRight.Width
    FrameRight.Top = FrameLeft.Top
    FrameRight.Height = FrameLeft.Height
    
    ' Обновление вида внутренних компонентов
    RefreshForm
    RefreshMainMenu
    RefreshFrameLeft
    RefreshFrameRight
    RefreshFrameMain
    RefreshStatusBar
End Sub

Private Sub RefreshFrameLeft()
    ListPrograms.Left = 120
    ListPrograms.Width = FrameLeft.Width - ListPrograms.Left - 120
    ListPrograms.Height = FrameLeft.Height - ListPrograms.Top - 120
    
    ListPrograms.ColWidth(0) = ListPrograms.Width
End Sub

Private Sub RefreshFrameMain()
    Select Case ViewMode
        Case STEPS_VIEW
            FrameCodeView.Visible = False
            FrameGridView.Left = 120
            FrameGridView.Top = 240
            FrameGridView.Width = FrameMain.Width - FrameGridView.Left - 120
            FrameGridView.Height = FrameMain.Height - FrameGridView.Top - 120
    
            StepsView.Left = 0
            StepsView.Top = 0
            StepsView.Width = FrameGridView.Width
            StepsView.Height = FrameGridView.Height

            If Manager.FileLoaded Then
                FrameMain.Caption = "Шаги - [" & ListPrograms.Text & _
                                ".Шаг" & Manager.StepIndex + 1 & "]"
            Else
                FrameMain.Caption = "Шаги"
            End If

            FrameGridView.Visible = True

        Case CODE_VIEW
            FrameGridView.Visible = False
            FrameCodeView.Left = 120
            FrameCodeView.Top = 240
            FrameCodeView.Width = FrameMain.Width - FrameCodeView.Left - 120
            FrameCodeView.Height = FrameMain.Height - FrameCodeView.Top - 120
    
            CodeView.Left = 0
            CodeView.Top = 0
            CodeView.Width = FrameCodeView.Width
            CodeView.Height = FrameCodeView.Height
            
            If Manager.FileLoaded Then
                FrameMain.Caption = "Код - [" & ListPrograms.Text & "]"
            Else
                FrameMain.Caption = "Код"
            End If
            
            ' Отображение данных в CodeView зависит от видимости строк
            ' Поэтому нужно делать обновление после измненения размеров
            RefreshCodeView
            
            FrameCodeView.Visible = True
    End Select
    
    FrameMain.Enabled = Manager.FileLoaded
End Sub

Private Sub RefreshFrameRight()
    PropertyTable.Left = 120
    PropertyTable.Top = 240
    PropertyTable.Width = FrameRight.Width - PropertyTable.Left - 120
    PropertyTable.Height = FrameRight.Height - PropertyTable.Top - 120
    
    If PropertyTable.Width > PropertyTable.ColWidth(0) Then
        PropertyTable.ColWidth(1) = PropertyTable.Width - PropertyTable.ColWidth(0)
    End If
End Sub

Private Sub RefreshForm()
    SetCaption Manager.FileName
End Sub

Private Sub RefreshMainMenu()
    PopupMenuPrograms.Visible = Manager.FileLoaded
    StepMainMenuItem.Visible = Manager.FileLoaded And (ViewMode = STEPS_VIEW)
End Sub

Private Sub RefreshStatusBar()
    If Modified Then
        StatusBar.Panels(2).Text = "Изменён"
    Else
        StatusBar.Panels(2).Text = ""
    End If
End Sub

Private Sub AboutMainMenuItem_Click()
    FormAbout.Show (vbModal)
End Sub


Private Sub CloseMainMenuItem_Click()
    If Modified = True Then
        Dim vbRes%
        vbRes = MsgBox("Сохранить изменения в файле:" & _
           Chr(13) & Chr(13) & """" & _
           Manager.FileName & """?", vbYesNoCancel + vbExclamation, APP_NAME)
        
        Select Case vbRes
        Case vbYes
            SaveMainMenuItem_Click
            Manager.CloseFile
            SetModified False
            RefreshComponents (False)
        Case vbNo
            Manager.CloseFile
            SetModified False
            RefreshComponents (False)
        Case vbCancel
        End Select
    Else
        Manager.CloseFile
        RefreshComponents (False)
    End If
End Sub

Private Sub CodeView_Click()
    Dim X%, Y%
    Dim col%, row%
    
    CodeView.Visible = False
    
    X% = CodeView.col
    Y% = CodeView.row

    For col% = 1 To CodeView.Cols - 2
        CodeView.col = col%
        CodeView.row = 0
        CodeView.CellFontBold = False
    Next col%
    
    row% = CodeView.TopRow
    
    Do While CodeView.RowIsVisible(row%)
        CodeView.col = 0
        CodeView.row = row%
    
        CodeView.CellFontBold = False
        row% = row% + 1
        If row% > CodeView.Rows - 1 Then Exit Do
    Loop
    
    CodeView.row = 0
    CodeView.col = X%
    CodeView.CellFontBold = True
    
    CodeView.row = Y%
    CodeView.col = 0
    CodeView.CellFontBold = True
    
    CodeView.col = X%
    CodeView.row = Y%
    
    CodeView.Visible = True
    CodeView.SetFocus
End Sub

Private Sub CodeView_DblClick()
    CodeView_KeyDown VBRUN.KeyCodeConstants.vbKeyReturn, 0
End Sub

Private Sub CodeView_KeyDown(keyCode As Integer, Shift As Integer)
    ' При нажатии Enter в ячейке даём возможность редактировать _
    её содержимое
    'If Not (KeyCode = VBRUN.KeyCodeConstants.vbKeyReturn) Then Exit Sub
    ' Фильтруем не нужные клавиши
    Select Case keyCode
        Case Asc("a"), Asc("b"), Asc("c"), Asc("d"), Asc("e"), Asc("f"):
        Case Asc("A"), Asc("B"), Asc("C"), Asc("D"), Asc("E"), Asc("F"):
        Case Asc("0"), Asc("1"), Asc("2"), Asc("3"), Asc("4"), _
            Asc("5"), Asc("6"), Asc("7"), Asc("8"), Asc("9"):

        Case VBRUN.KeyCodeConstants.vbKeyReturn, _
            VBRUN.KeyCodeConstants.vbKeyDelete, _
            VBRUN.KeyCodeConstants.vbKeyBack, _
            VBRUN.KeyCodeConstants.vbKeySpace, _
            VBRUN.KeyCodeConstants.vbKeyNumpad0, _
            VBRUN.KeyCodeConstants.vbKeyNumpad1, _
            VBRUN.KeyCodeConstants.vbKeyNumpad2, _
            VBRUN.KeyCodeConstants.vbKeyNumpad3, _
            VBRUN.KeyCodeConstants.vbKeyNumpad4, _
            VBRUN.KeyCodeConstants.vbKeyNumpad5, _
            VBRUN.KeyCodeConstants.vbKeyNumpad6, _
            VBRUN.KeyCodeConstants.vbKeyNumpad7, _
            VBRUN.KeyCodeConstants.vbKeyNumpad8, _
            VBRUN.KeyCodeConstants.vbKeyNumpad9:
        
        Case Else
            Exit Sub
    End Select
    
    Dim col, row, FuncN As Integer
    
    col = CodeView.col
    row = CodeView.row
    
    ' На всякий случай пропускаем фиксированные ячейки
    If col = 0 Or row = 0 Then Exit Sub
    
    TextByte.Font = CodeView.Font
    TextByte.Left = CodeView.Left + CodeView.CellLeft
    TextByte.Top = CodeView.Top + CodeView.CellTop
    TextByte.Width = CodeView.CellWidth
    TextByte.Height = CodeView.CellHeight
    TextByte.Text = CodeView.Text
    TextByte.SelStart = 0
    TextByte.SelLength = Len(TextByte.Text)
    TextByte.Visible = True
    TextByte.SetFocus
End Sub

Private Sub CodeView_Scroll()
    RefreshCodeView
End Sub

Private Sub ComboCell_KeyDown(keyCode As Integer, Shift As Integer)
    If keyCode = VBRUN.KeyCodeConstants.vbKeyEscape Then
        ComboCell.Visible = False
        LabelDescription.Visible = False
        RefreshFrameRight
        PropertyTable.SetFocus
    End If
    
    If keyCode = VBRUN.KeyCodeConstants.vbKeyReturn Then
        Dim FuncN As Integer
        
        FuncN = Manager.GetFunctionType(Manager.ProgramIndex + 1, Manager.StepIndex + 1)
        
        ' Сохраняем изменённое значение
        If FuncN < 12 Then
            Select Case FuncN
                Case WPC_OPERATION_IDLE ' пропуск
                    ModuleIdle.SetComboPropertyForIdle Me
            
                Case WPC_OPERATION_FILL ' Налив
                    ModuleFill.SetComboPropertyForFill Me
                
                Case WPC_OPERATION_DTRG ' моющие
                    ModuleDTRG.SetComboPropertyForDTRG Me
                
                Case WPC_OPERATION_HEAT ' нагрев
                    ModuleHeat.SetComboPropertyForHeat Me
                    
                ' стирка, полоскание, расстряска
                Case WPC_OPERATION_WASH, WPC_OPERATION_RINS, WPC_OPERATION_JOLT, WPC_OPERATION_PAUS
                    ModuleWashOrRinsOrJolt.SetComboPropertyForWashOrRinsOrJolt Me
                    
'                Case WPC_OPERATION_PAUS ' пауза
'                    ModulePause.SetComboPropertyForPause Me
    
                Case WPC_OPERATION_DRAIN ' слив
                    ModuleDrain.SetComboPropertyForDrain Me
                    
                Case WPC_OPERATION_SPIN ' отжим
                    ModuleSpin.SetComboPropertyForSpin Me
                
                Case WPC_OPERATION_COOL ' охлаждение
                    ModuleCool.SetComboPropertyForCool Me
                    
                Case WPC_OPERATION_TRIN ' тех.полоскание
                    ModuleTrin.SetComboPropertyForTrin Me
            
                Case Else
    
            End Select
            
            ' Пересчитываем CRC поле записи программы
            Dim b As Byte
            b = Manager.CalculateCRC8(Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + 1, PROGRAM_SIZE_IN_BYTES - 1)
            Manager.SetByte Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES, b

            ComboCell.Visible = False
            LabelDescription.Visible = False
            Dim row As Integer
            row = PropertyTable.row
            RefreshComponents (False)
            If row < PropertyTable.Rows - 1 Then PropertyTable.row = row
            PropertyTable.SetFocus
        End If
    End If
End Sub

Private Sub ComboCell_LostFocus()
    ComboCell.Visible = False
End Sub


Private Sub CopyMainMenuItem_Click()
    Dim I As Integer
    
    FormCopy.List1.Clear
    FormCopy.List2.Clear
    
    For I = 1 To Manager.ProgramsCount
        FormCopy.List1.AddItem ListPrograms.TextArray(GetCellIndex(ListPrograms, I, 0))
        FormCopy.List2.AddItem ListPrograms.TextArray(GetCellIndex(ListPrograms, I, 0))
    Next I
    
    FormCopy.List1.ListIndex = 0
    FormCopy.List2.ListIndex = 0
    
    FormCopy.Show (vbModal)
    
    RefreshComponents (False)
End Sub



Private Sub DeleteStepMenuItem_Click()
    Manager.DeleteStep
    
    ' Пересчитываем CRC поле записи программы
    Dim b As Byte
    b = Manager.CalculateCRC8(Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + 1, PROGRAM_SIZE_IN_BYTES - 1)
    Manager.SetByte Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES, b
    
    SetModified True
    RefreshDataComponents
    
    StepsView.SetFocus
End Sub

Private Sub ExitMainMenuItem_Click()
    If Modified = True Then
        Dim vbRes%
        vbRes = MsgBox("Сохранить изменения в файле?", vbYesNoCancel + vbExclamation, APP_NAME)
        
        Select Case vbRes
        Case vbYes
            SaveMainMenuItem_Click
            UnHookKeyboard
            
        Case vbNo
            UnHookKeyboard
            
        Case vbCancel
            Exit Sub
            
        End Select
    End If
    
    ' Сохраняем настройки интерфейса
    SavePlacement
    
    End
End Sub

Private Sub FileMainMenuItem_Click()
    SaveMainMenuItem.Enabled = Modified
    SaveAsMainMenuItem.Enabled = Manager.FileLoaded
    CloseMainMenuItem.Enabled = Manager.FileLoaded
End Sub

Private Sub Form_KeyDown(keyCode As Integer, Shift As Integer)
    Dim col As Integer, row As Integer
    
    If keyCode = VBRUN.KeyCodeConstants.vbKeyF3 And Shift = 0 Then
        If Not Manager.FileLoaded Then Exit Sub
        
        Select Case ViewMode
            Case STEPS_VIEW
                CodeView.TopRow = (PROGRAM_SIZE_IN_BYTES * Manager.ProgramIndex + _
                    HEADER_SIZE_IN_BYTES + STEP_SIZE_IN_BYTES * Manager.StepIndex) / 16 + 1
                
                ViewMode = CODE_VIEW
                RefreshCodeView
            
            Case CODE_VIEW
                ViewMode = STEPS_VIEW
                RefreshStepsView
                
        End Select
        
        RefreshFrameMain
        RefreshMainMenu
        Exit Sub
    End If
    
    If keyCode = VBRUN.KeyCodeConstants.vbKeyF4 And Shift = 0 Then
        If ViewMode = CODE_VIEW Then Exit Sub
        
        Select Case StepsViewMode
            Case TEXT_VIEW: StepsViewMode = CHECKS_VIEW
            Case CHECKS_VIEW: StepsViewMode = TEXT_VIEW
        End Select
            
        row = StepsView.row
        col = StepsView.col
        
        RefreshStepsView
        RefreshMainMenu
        
        StepsView.row = row
        StepsView.col = col
        
        StepsView.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim itm As ListItem
    Dim sitm As ListSubItem
    Dim IniFilePath As String
    Dim Result%

    KeyPreview = True
    
    ' Среда разработки часто "вылетает" из-за кода внутри
    ' Поэтому его тестирование нужно проводить только на
    ' откомпилированном приложении
    Dim WE_ARE_IN_IDE As Boolean
    
    Debug.Assert MakeTrue(WE_ARE_IN_IDE)
    
    If WE_ARE_IN_IDE Then
        ' Код, выполняемый в runtime среды разработки
        DesignMode = True
    Else
        ' Код, который будет в скомпилированном файле
        DesignMode = False
        
        Timer1.Enabled = True
        Timer1.Interval = 0
    
        HookKeyboard Timer1
    End If
   
    ' Режим отображения средней панели
    ViewMode = STEPS_VIEW
    
    ' Режимы отображения таблицы шагов
    StepsViewMode = TEXT_VIEW
    
    SplitterRightMoving = False
    SplitterLeftMoving = False
    
    ' Формируем путь к файлу настроек
    IniFilePath = String$(255, 0)
    Result = GetModuleFileName(0, IniFilePath, 254)
    CurrentDir = MiscExtractPathName(IniFilePath, True)
    IniFilePath = StrConv(IniFilePath, vbLowerCase)
    IniFilePath = Replace(IniFilePath, ".exe", ".ini")

    ' Создаём экземпляр объекта
    Set IniFile = New TIniFiles
    IniFile.Create (IniFilePath)
    
    ' Создаём экземпляр объекта
    Set Manager = New TProgramManager
    Manager.Create
    
    App.HelpFile = CurrentDir & "\cop.chm"
    
    ' Начальные пути для диалоговых окон
    OpenFileDialog.InitDir = CurrentDir
    SaveFileDialog.InitDir = CurrentDir
    
    Set ModuleIdle = New TModuleIdle
    Set ModuleFill = New TModuleFill
    Set ModuleDTRG = New TModuleDTRG
    Set ModuleHeat = New TModuleHeat
    Set ModuleWashOrRinsOrJolt = New TModuleWashOrRinsOrJolt
    'Set ModulePause = New TModulePause
    Set ModuleDrain = New TModuleDrain
    Set ModuleSpin = New TModuleSpin
    Set ModuleCool = New TModuleCool
    Set ModuleTrin = New TModuleTrin
        
    IniFilePath = CurrentDir & "\limits.ini"
    
    LoadLimits IniFilePath
    ModuleIdle.LoadLimits IniFilePath
    ModuleFill.LoadLimits IniFilePath
    ModuleDTRG.LoadLimits IniFilePath
    ModuleHeat.LoadLimits IniFilePath
    ModuleWashOrRinsOrJolt.LoadLimits IniFilePath
'    ModulePause.LoadLimits IniFilePath
    ModuleDrain.LoadLimits IniFilePath
    ModuleSpin.LoadLimits IniFilePath
    ModuleCool.LoadLimits IniFilePath
    ModuleTrin.LoadLimits IniFilePath
    
    SetModified False
    
    ' Восстанавливаем положение формы и компонентов
    LoadPlacement
        
    Dim s As String
    Dim col%, row%
    
    StepsView.Visible = False
    
    StepsView.Cols = MAX_NUMBER_OF_STEPS + 1
    
    s = "<   |"
    For col% = 1 To StepsView.Cols - 1
        StepsView.ColWidth(col%) = 250
        If col% < StepsView.Cols - 1 Then
            s = s & col% & "|"
        Else
            s = s & col%
        End If
        StepsView.col = col%
        StepsView.row = 0
        StepsView.CellAlignment = flexAlignCenterCenter
    Next col%
    
    StepsView.FormatString = s
       
    s = ";|" _
        & "Клапан горячей воды" & "|" _
        & "Клапан холодной воды 1" & "|" _
        & "Клапан холодной воды 2" & "|" _
        & "Клапан МС 1" & "|" _
        & "Клапан МС 2" & "|" _
        & "Клапан МС 3" & "|" _
        & "Клапан МС 4" & "|" _
        & "Клапан МС 5" & "|" _
        & "Клапан МС 6" & "|" _
        & "Клапан МС 7" & "|" _
        & "Клапан МС 8" & "|" _
        & "Клапан МС 9" & "|" _
        & "Замок люка 1" & "|" _
        & "Замок люка 2" & "|" _
        & "Слив 1" & "|" _
        & "Слив 2" & "|" _
        & "Нагрев" & "|" _
        & "Мотор"
    
    StepsView.FormatString = s
    
    ' "Тушим" все ячейки таблицы
    For row% = 1 To StepsView.Rows - 1
        For col% = 1 To MAX_NUMBER_OF_STEPS
            StepsView.col = col%
            StepsView.row = row%
            StepsView.CellBackColor = &H8000000F
        Next col%
    Next row%
    
    StepsView.col = 1
    StepsView.row = 1
    
    StepsView.Visible = True
    
    ' Инициализируем окно кода
    s = "<   |"
    For col% = 1 To CodeView.Cols - 1
        CodeView.ColWidth(col%) = 250
        If col% < CodeView.Cols - 1 Then
            If col% < 11 Then
                s = s & "0" & col% - 1 & "|"
            Else
                s = s & "0" & Chr$(col% - 11 + 65) & "|"
            End If
        Else
            If col% < 11 Then
                s = s & "0" & col% - 1 & "|"
            Else
                s = s & "0" & Chr$(col% - 11 + 65) & "|"
            End If
        End If
        CodeView.col = col%
        CodeView.row = 0
        CodeView.CellAlignment = flexAlignCenterCenter
    Next col%
    
    CodeView.FormatString = s
    
    FunctionsStrings(0) = "Пропуск"
    FunctionsStrings(1) = "Налив"
    FunctionsStrings(2) = "Моющие"
    FunctionsStrings(3) = "Нагрев"
    FunctionsStrings(4) = "Стирка"
    FunctionsStrings(5) = "Полоскание"
    FunctionsStrings(6) = "Расстряска"
    FunctionsStrings(7) = "Пауза"
    FunctionsStrings(8) = "Слив"
    FunctionsStrings(9) = "Отжим"
    FunctionsStrings(10) = "Охлаждение"

    ' Обновляем вид
    RefreshComponents (False)
    
    ' Симулируем изменение размером формы для вызова Resize()
    Move Left, Top, Width, Height
End Sub

Private Sub Form_Resize()
    RefreshComponents (True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Modified = True Then
        Dim vbRes%
        vbRes = MsgBox("Сохранить изменения в файле?", vbYesNoCancel + vbExclamation, APP_NAME)
        
        Select Case vbRes
        Case vbYes
            SaveMainMenuItem_Click
            Unload Me
            Set FormMain = Nothing
            UnHookKeyboard
            
        Case vbNo
            Unload Me
            Set FormMain = Nothing
            UnHookKeyboard
            
        Case vbCancel
            Cancel = 1
        End Select
    End If
    
    ' Сохраняем настройки интерфейса
    SavePlacement
End Sub

Private Sub HelpMainMenuSubItem_Click()
    If DoesFileExist(App.HelpFile) Then
        Shell ("hh " & App.HelpFile), vbNormalFocus
    End If
End Sub

Private Sub InsertStepMenuItem_Click()
    Manager.InsertStep
            
    ' Пересчитываем CRC поле записи программы
    Dim b As Byte
    b = Manager.CalculateCRC8(Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + 1, PROGRAM_SIZE_IN_BYTES - 1)
    Manager.SetByte Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES, b
            
    SetModified True
    RefreshDataComponents

    StepsView.SetFocus
End Sub

Private Sub ListPrograms_Click()
    Manager.ProgramIndex = ListPrograms.row - 1
    
    Select Case ViewMode
        Case STEPS_VIEW
            RefreshFrameMain
            RefreshStepsView

        Case CODE_VIEW
            CodeView.TopRow = (PROGRAM_SIZE_IN_BYTES * Manager.ProgramIndex) / 16 + 1
            RefreshFrameMain
    End Select

    RefreshFrameLeft
    RefreshProperties
    RefreshFrameRight
End Sub

Private Sub ListPrograms_DblClick()
    ListPrograms_KeyDown VBRUN.KeyCodeConstants.vbKeyReturn, 0
End Sub

Private Sub ListPrograms_KeyDown(keyCode As Integer, Shift As Integer)
    If keyCode = VBRUN.KeyCodeConstants.vbKeyUp Or _
        keyCode = VBRUN.KeyCodeConstants.vbKeyDown Then
        
        ListPrograms_Click
    End If
    
    If keyCode = VBRUN.KeyCodeConstants.vbKeyReturn Then
        TextName.Left = ListPrograms.Left + ListPrograms.CellLeft
        TextName.Top = ListPrograms.Top + ListPrograms.CellTop
        TextName.Width = ListPrograms.CellWidth
        TextName.Height = ListPrograms.CellHeight
        
        TextName.Text = ListPrograms.Text
        TextName.Visible = True
        TextName.SetFocus
    End If
End Sub

Private Sub ListPrograms_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'проверка, нажата ли правая клавиша мыши
    If Button And vbRightButton Then PopupMenu PopupMenuPrograms
End Sub

Private Sub NewMainMenuItem_Click()
    Manager.CreateNewFile (DEFAULT_FILE_NAME)
    
    ' Очистить все программы из образа
    PopupMenuListClearAll_Click
End Sub
Public Sub SetModified(Value As Boolean)
    Modified = Value
End Sub

Private Sub PopupMenuListClear_Click()
    Dim StepPointer As Long
    
    ' Очищаем текущую программу
    Manager.ClearProgramN (ListPrograms.row)
    
    ' Устанавливаем заголовок по умолчанию
    If LimitsLoaded Then SetDefaultProgramTitle Manager.ProgramIndex + 1
    
    ' Пересчитываем CRC поле записи программы
    Dim b As Byte
    b = Manager.CalculateCRC8((ListPrograms.row - 1) * PROGRAM_SIZE_IN_BYTES + 1, PROGRAM_SIZE_IN_BYTES - 1)
    Manager.SetByte (ListPrograms.row - 1) * PROGRAM_SIZE_IN_BYTES, b
            
    SetModified True
    
    RefreshDataComponents
End Sub

Private Sub PopupMenuListClearAll_Click()
    Manager.ClearAll
    
    If LimitsLoaded Then
        Dim I As Integer
        
        For I = 1 To Manager.ProgramsCount
            SetDefaultProgramTitle I
        
            ' Пересчитываем CRC поле записи программы
            Dim b As Byte
            b = Manager.CalculateCRC8((I - 1) * PROGRAM_SIZE_IN_BYTES + 1, PROGRAM_SIZE_IN_BYTES - 1)
            Manager.SetByte (I - 1) * PROGRAM_SIZE_IN_BYTES, b
        Next I
    End If
    
    SetModified True
    
    RefreshComponents (False)
End Sub

Private Sub PropertyTable_DblClick()
    PropertyTable_KeyDown VBRUN.KeyCodeConstants.vbKeyReturn, 0
End Sub

Private Sub PropertyTable_KeyDown(keyCode As Integer, Shift As Integer)
    ' При нажатии Enter в ячейке даём возможность редактировать _
    её содержимое
    If Not (keyCode = VBRUN.KeyCodeConstants.vbKeyReturn) Then Exit Sub
    
    Dim col, row, FuncN As Integer
    
    col = PropertyTable.col
    row = PropertyTable.row
    
    ' На всякий случай пропускаем фиксированные ячейки
    If col = 0 Or row = 0 Then Exit Sub
       
    ' Действуем в зависимости от типа функции шага
    FuncN = Manager.GetFunctionType(Manager.ProgramIndex + 1, Manager.StepIndex + 1)
    
    If FuncN < 12 Then
        Select Case FuncN
            Case WPC_OPERATION_IDLE ' пропуск
                ModuleIdle.EditPropertyForIdle Me
        
            Case WPC_OPERATION_FILL ' Налив
                ModuleFill.EditPropertyForFill Me
            
            Case WPC_OPERATION_DTRG ' моющие
                ModuleDTRG.EditPropertyForDTRG Me
            
            Case WPC_OPERATION_HEAT ' нагрев
                ModuleHeat.EditPropertyForHeat Me
                
            ' стирка, полоскание, расстряска
            Case WPC_OPERATION_WASH, WPC_OPERATION_RINS, WPC_OPERATION_JOLT, WPC_OPERATION_PAUS
                ModuleWashOrRinsOrJolt.EditPropertyForWashOrRinsOrJolt Me
                
'            Case WPC_OPERATION_PAUS ' пауза
'                ModulePause.EditPropertyForPause Me

            Case WPC_OPERATION_DRAIN ' слив
                ModuleDrain.EditPropertyForDrain Me
                
            Case WPC_OPERATION_SPIN ' отжим
                ModuleSpin.EditPropertyForSpin Me
            
            Case WPC_OPERATION_COOL ' охлаждение
                ModuleCool.EditPropertyForCool Me
                
            Case WPC_OPERATION_TRIN ' тех.полоскание
                ModuleTrin.EditPropertyForTrin Me
        
            Case Else

        End Select
    End If
End Sub

Private Sub StepMainMenuItem_Click()
    InsertStepMenuItem = ActiveControl Is StepsView
    DeleteStepMenuItem.Enabled = ActiveControl Is StepsView
End Sub

Private Sub StepsView_Click()
    Dim X%, Y%
    Dim col%, row%
    
    StepsView.Visible = False
    
    X% = StepsView.col
    Y% = StepsView.row
       
    For col% = 1 To StepsView.Cols - 2
        StepsView.col = col%
        StepsView.row = 0
        StepsView.CellFontBold = False
    Next col%
    
    row% = StepsView.TopRow
    
    Do While StepsView.RowIsVisible(row%)
        StepsView.col = 0
        StepsView.row = row%
    
        StepsView.CellFontBold = False
        row% = row% + 1
        If row% > StepsView.Rows - 1 Then Exit Do
    Loop
       
    StepsView.row = 0
    StepsView.col = X%
    StepsView.CellFontBold = True
    
    StepsView.row = Y%
    StepsView.col = 0
    StepsView.CellFontBold = True
    
    StepsView.col = X%
    StepsView.row = Y%
    
    Manager.StepIndex = X% - 1
    
    CodeView.TopRow = (PROGRAM_SIZE_IN_BYTES * Manager.ProgramIndex + _
        HEADER_SIZE_IN_BYTES + STEP_SIZE_IN_BYTES * Manager.StepIndex) / 16 + 1
    
    StepsView.Visible = True
    StepsView.SetFocus
    
    ' Обновляем зависимые компоненты
    RefreshProperties
    RefreshFrameMain
    RefreshFrameRight
    RefreshCodeView
End Sub

Private Sub OpenMainMenuItem_Click()
    On Local Error GoTo errhandler
    Dim FileName As String

    OpenFileDialog.DialogTitle = "Открыть файл..."
    OpenFileDialog.DefaultExt = ".bin"
    OpenFileDialog.Filter = "Файлы проекта (*.bin)|*.bin"
    OpenFileDialog.FilterIndex = 1
    OpenFileDialog.MaxFileSize = 32767
    OpenFileDialog.InitDir = CurrentDir
    OpenFileDialog.CancelError = True
    OpenFileDialog.ShowOpen

    FileName = OpenFileDialog.FileName
        
    If FileName <> "" Then
        If Manager.FileLoaded Then
            CloseMainMenuItem_Click
        End If
        
        FileName = MiscExtractPathName(OpenFileDialog.FileName, False)
        Manager.LoadFromFile (OpenFileDialog.FileName)
        SetCaption (Manager.FileName)
        
        ViewMode = STEPS_VIEW
        RefreshComponents (False)
        'RefreshFrameRight
        
        SetModified False
        CurrentDir = MiscExtractPathName(OpenFileDialog.FileName, True)
    End If

    Exit Sub
    
errhandler:
End Sub

Private Sub SetCaption(FileName As String)
    If Manager.FileLoaded Then
        If DesignMode Then
            Caption = APP_NAME & " [DESIGN] - [" & FileName & "]"
        Else
            Caption = APP_NAME & " - [" & FileName & "]"
        End If
    Else
        If DesignMode Then
            Caption = APP_NAME & " [DESIGN]"
        Else
            Caption = APP_NAME & ""
        End If
    End If
End Sub

Private Sub SaveAsMainMenuItem_Click()
    On Local Error GoTo errhandler
    Dim FileName As String

    SaveFileDialog.FileName = Manager.FileName
    SaveFileDialog.DialogTitle = "Сохранить файл..."
    SaveFileDialog.DefaultExt = ".bin"
    SaveFileDialog.Filter = "Файлы проекта (*.bin)|*.bin"
    SaveFileDialog.FilterIndex = 1
    SaveFileDialog.MaxFileSize = 32767
    SaveFileDialog.InitDir = CurrentDir
    SaveFileDialog.CancelError = True
    SaveFileDialog.ShowSave

    FileName = SaveFileDialog.FileName
    
    If FileName <> "" Then
        FileName = MiscExtractPathName(SaveFileDialog.FileName, False)
        SetCaption (FileName)
        Manager.SaveToFile (SaveFileDialog.FileName)
        SetModified False
        CurrentDir = MiscExtractPathName(SaveFileDialog.FileName, True)
        RefreshDataComponents
    End If
    Exit Sub
    
errhandler:
End Sub

Private Sub SaveMainMenuItem_Click()
    If Modified Then
        If DoesFileExist(Manager.FileName) Then
            Manager.SaveToFile (Manager.FileName)
            SetModified False
            RefreshDataComponents
        Else
            SaveAsMainMenuItem_Click
        End If
    End If
End Sub

Private Sub SplitterLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Показываем разделительную линию
    SplitterLeft.BackColor = &H80000010
    
    BegX = X
    BegY = Y
    
    SplitterLeftMoving = True
End Sub

Private Sub SplitterLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SplitterLeftMoving = True Then
        SplitterLeft.Left = SplitterLeft.Left + X - BegX
        FrameLeft.Width = SplitterLeft.Left
        
        FrameMain.Left = SplitterLeft.Left + SplitterLeft.Width
        FrameMain.Width = SplitterRight.Left - FrameMain.Left
        
        RefreshFrameLeft
        RefreshFrameMain
        Refresh
    End If
End Sub

Private Sub SplitterLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SplitterLeft.BackColor = &H8000000F
    SplitterLeftMoving = False
End Sub

Private Sub SplitterRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Показываем разделительную линию
    SplitterRight.BackColor = &H80000010
    
    BegX = X
    BegY = Y
    
    SplitterRightMoving = True
End Sub

Private Sub SplitterRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SplitterRightMoving = True Then
        SplitterRight.Left = SplitterRight.Left + X - BegX
        
        FrameRight.Left = SplitterRight.Left + SplitterRight.Width
        FrameRight.Width = Me.ScaleWidth - FrameRight.Left
        
        FrameMain.Width = SplitterRight.Left - FrameMain.Left
        
        RefreshFrameRight
        RefreshFrameMain
        Refresh
    End If
End Sub

Private Sub SplitterRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SplitterRight.BackColor = &H8000000F
    SplitterRightMoving = False
End Sub

Private Sub StepsView_KeyDown(keyCode As Integer, Shift As Integer)
    If keyCode = VBRUN.KeyCodeConstants.vbKeyInsert Then
        InsertStepMenuItem_Click
        Exit Sub
    End If
        
    If keyCode = VBRUN.KeyCodeConstants.vbKeyDelete Then
        DeleteStepMenuItem_Click
        Exit Sub
    End If
        
    If keyCode = VBRUN.KeyCodeConstants.vbKeyLeft Or _
        keyCode = VBRUN.KeyCodeConstants.vbKeyRight Then
        
        StepsView_Click
    End If
End Sub

Private Sub StepsView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'проверка, нажата ли правая клавиша мыши
    If Button And vbRightButton Then PopupMenu StepMainMenuItem
End Sub

Private Sub TextByte_Change()
    TextByte.Text = Mid(TextByte.Text, 1, 2)
End Sub

Private Sub TextByte_KeyDown(keyCode As Integer, Shift As Integer)
    If keyCode = VBRUN.KeyCodeConstants.vbKeyEscape Then
        TextByte.Visible = False
        CodeView.SetFocus
    End If
    
    If keyCode = VBRUN.KeyCodeConstants.vbKeyReturn Then
        Dim OldByte, NewByte As Byte
        Dim ProgNum As Integer
        Dim Offset, TopRow As Long
        
        Offset = (CodeView.row - 1) * 16 + CodeView.col - 1
        
        ' Номер программы
        ProgNum = Offset / PROGRAM_SIZE_IN_BYTES
        ListPrograms.row = ProgNum + 1
        Manager.ProgramIndex = ProgNum
        
        OldByte = Manager.GetByte(Offset)
        NewByte = Val("&H" & TextByte.Text)
        
        If NewByte = OldByte Then
            TextByte.Visible = False
            CodeView.SetFocus
            Exit Sub
        End If
        
        Dim row, col As Long
        
        row = CodeView.row
        col = CodeView.col
        
        Manager.SetByte Offset, NewByte
        
        ' Пересчитываем CRC поле записи программы
        Dim b As Byte
        b = Manager.CalculateCRC8(Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + 1, PROGRAM_SIZE_IN_BYTES - 1)
        Manager.SetByte Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES, b

        SetModified True
        
        TextByte.Visible = False
        TopRow = CodeView.TopRow
        RefreshComponents (False)
        CodeView.TopRow = TopRow
        
        CodeView.row = row
        CodeView.col = col
        CodeView.SetFocus
    End If
End Sub

Private Sub TextByte_KeyPress(KeyAscii As Integer)
    ' Фильтруем не нужные клавиши
    Select Case KeyAscii
        Case Asc("a"), Asc("b"), Asc("c"), Asc("d"), Asc("e"), Asc("f"):
        Case Asc("A"), Asc("B"), Asc("C"), Asc("D"), Asc("E"), Asc("F"):
        Case Asc("0"), Asc("1"), Asc("2"), Asc("3"), Asc("4"), _
            Asc("5"), Asc("6"), Asc("7"), Asc("8"), Asc("9"):
        ' Enter и Del
        Case 13, 8:
            
        Case Else
            KeyAscii = 0
    End Select
    
    If KeyAscii = VBRUN.KeyCodeConstants.vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub TextByte_LostFocus()
    TextByte.Visible = False
End Sub

Private Sub TextCell_KeyDown(keyCode As Integer, Shift As Integer)
    If keyCode = VBRUN.KeyCodeConstants.vbKeyEscape Then
        TextCell.Visible = False
        LabelDescription.Visible = False
        RefreshFrameRight
        PropertyTable.SetFocus
    End If
    
    If keyCode = VBRUN.KeyCodeConstants.vbKeyReturn Then
        Dim FuncN As Integer
        
        FuncN = Manager.GetFunctionType(Manager.ProgramIndex + 1, Manager.StepIndex + 1)
        
        ' Сохраняем изменённое значение
        If FuncN < 12 Then
            Select Case FuncN
                Case WPC_OPERATION_IDLE ' пропуск
                    ModuleIdle.SetComboPropertyForIdle Me
            
                Case WPC_OPERATION_FILL ' Налив
                    ModuleFill.SetComboPropertyForFill Me
                
                Case WPC_OPERATION_DTRG ' моющие
                    ModuleDTRG.SetComboPropertyForDTRG Me
                
                Case WPC_OPERATION_HEAT ' нагрев
                    ModuleHeat.SetComboPropertyForHeat Me
                    
                ' стирка, полоскание, расстряска
                Case WPC_OPERATION_WASH, WPC_OPERATION_RINS, WPC_OPERATION_JOLT, WPC_OPERATION_PAUS
                    ModuleWashOrRinsOrJolt.SetComboPropertyForWashOrRinsOrJolt Me
                    
'                Case WPC_OPERATION_PAUS ' пауза
'                    ModulePause.SetComboPropertyForPause Me
    
                Case WPC_OPERATION_DRAIN ' слив
                    ModuleDrain.SetComboPropertyForDrain Me
                    
                Case WPC_OPERATION_SPIN ' отжим
                    ModuleSpin.SetComboPropertyForSpin Me
                
                Case WPC_OPERATION_COOL ' охлаждение
                    ModuleCool.SetComboPropertyForCool Me
                    
                Case WPC_OPERATION_TRIN ' тех.полоскание
                    ModuleTrin.SetComboPropertyForTrin Me
            
                Case Else
    
            End Select
            
            ' Пересчитываем CRC поле записи программы
            Dim b As Byte
            b = Manager.CalculateCRC8(Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + 1, PROGRAM_SIZE_IN_BYTES - 1)
            Manager.SetByte Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES, b
            
            TextCell.Visible = False
            LabelDescription.Visible = False
            Dim row As Integer
            row = PropertyTable.row
            RefreshComponents (False)
            If row < PropertyTable.Rows - 1 Then PropertyTable.row = row
            PropertyTable.SetFocus
        End If
    End If
End Sub

Private Sub TextCell_KeyPress(KeyAscii As Integer)
    If KeyAscii = VBRUN.KeyCodeConstants.vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub TextCell_LostFocus()
    TextCell.Visible = False
    RefreshFrameRight
End Sub

Private Sub TextName_KeyDown(keyCode As Integer, Shift As Integer)
    Dim I As Integer
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    
    If keyCode = VBRUN.KeyCodeConstants.vbKeyEscape Then
        TextName.Visible = False
        ListPrograms.SetFocus
    End If
    
    If keyCode = VBRUN.KeyCodeConstants.vbKeyReturn Then
        StepPointer = Manager.DataPointer + Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        
        For I = 1 To PROG_NAME_LENGTH - 1
            If I <= Len(TextName.Text) Then
                RecordTitle.ProgName(I) = Asc(Mid(TextName.Text, I, 1))
            Else
                RecordTitle.ProgName(I) = 0
            End If
        Next I
        RecordTitle.ProgName(PROG_NAME_LENGTH) = 0
        ' Сохраняем изменения
        CopyMemory ByVal StepPointer, RecordTitle, HEADER_SIZE_IN_BYTES
        SetModified True
        
        ' Пересчитываем CRC поле записи программы
        Dim b As Byte
        b = Manager.CalculateCRC8(Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + 1, PROGRAM_SIZE_IN_BYTES - 1)
        Manager.SetByte Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES, b
        
        TextName.Visible = False
        Dim row As Integer
        row = ListPrograms.row
        RefreshComponents (False)
        
        If row < ListPrograms.Rows - 1 Then ListPrograms.row = row
        ListPrograms.SetFocus
    End If
End Sub

Private Sub TextName_KeyPress(KeyAscii As Integer)
    If KeyAscii = VBRUN.KeyCodeConstants.vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub TextName_LostFocus()
    TextName.Visible = False
End Sub

Private Sub Timer1_Timer()
    ' Особый случай: мы ловим клавиши при помощи хуков
    ' Когда файлы не загружены нет активного элемента управления
    ' поэтому обращение к ActiveControl вызывает сбой программы
    If Me.ActiveControl Is Nothing Then Exit Sub

    If TypeOf Me.ActiveControl Is MSFlexGrid Then
        If Me.ActiveControl.Name = "StepsView" Then
            StepsView_KeyDown CInt(Timer1.Tag), 0
        End If
        
        If Me.ActiveControl.Name = "ListPrograms" Then
            ListPrograms_KeyDown CInt(Timer1.Tag), 0
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If (Button.Index = 1) Then NewMainMenuItem_Click
    If (Button.Index = 2) Then OpenMainMenuItem_Click
    If (Button.Index = 3) Then SaveMainMenuItem_Click
End Sub

Private Sub RefreshDataComponents()
    RefreshForm
    RefreshMainMenu
    RefreshList
    RefreshCodeView
    RefreshStepsView
    RefreshProperties
    RefreshFrameRight
    RefreshStatusBar
End Sub

Private Sub RefreshList()
    If Not Manager.FileLoaded Then
        FrameLeft.Caption = "Программы"
        ListPrograms.Clear
        ListPrograms.FormatString = "<Список"
        ListPrograms.Rows = 1
        FrameLeft.Enabled = False
        Exit Sub
    End If
    
    FrameLeft.Enabled = True
    FrameLeft.Caption = "Программы [" & Manager.ProgramsCount & "]"
    
    ListPrograms.Visible = False
    ListPrograms.Clear
    ListPrograms.Rows = 1
    ListPrograms.FormatString = "<Список"
    
    If Manager.ProgramsCount > 0 Then
        Dim b As Byte, N As Byte
        Dim Cnt As Integer
        Dim StepPointer As Long, Value As Long
        Dim s As String
        Dim RecordTitle As TYPE_WPC_TITLE
        
        For Cnt = 1 To Manager.ProgramsCount
            StepPointer = Manager.DataPointer + (Cnt - 1) * PROGRAM_SIZE_IN_BYTES
            CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
            
            Value = 0
            For N = 1 To PROG_NAME_LENGTH - 1
                Value = Value + CLng(RecordTitle.ProgName(N))
            Next N
            
            If Value = 0 Then
                ListPrograms.AddItem "Программа" & Cnt
            Else
                s = ""
                For N = 1 To PROG_NAME_LENGTH - 1
                    s = s & Chr(RecordTitle.ProgName(N))
                Next N
                ListPrograms.AddItem s
            End If
        Next Cnt
    
        ' Проверяем CRC для каждой из управляющих программ
        For Cnt = 0 To Manager.ProgramsCount - 1
            ListPrograms.row = Cnt + 1
            
            b = Manager.CalculateCRC8(Cnt% * PROGRAM_SIZE_IN_BYTES + 1, _
                PROGRAM_SIZE_IN_BYTES - 1)

            If Not b = Manager.GetByte(Cnt * PROGRAM_SIZE_IN_BYTES) Then
                ListPrograms.CellBackColor = &H8080FF
            End If
        Next Cnt
        
        ListPrograms.row = Manager.ProgramIndex + 1
    End If
    
    ListPrograms.ColWidth(0) = ListPrograms.Width
    ListPrograms.Visible = True
End Sub

Private Sub RefreshCodeView()
    Dim col, row As Integer
    Dim s As String
    
    ' Если файл не загружен, то выводить нечего,
    ' поэтому отображаем вид без данных
    If (Not Manager.FileLoaded) Then
        FrameMain.Caption = "Код"
        CodeView.Visible = False
        CodeView.Clear
        
        CodeView.Rows = 2
        ' Инициализируем окно кода
        s = "<   |"
        For col = 1 To CodeView.Cols - 2
            CodeView.ColWidth(col) = 250
            If col < CodeView.Cols - 1 Then
                If col < 11 Then
                    s = s & "0" & col - 1 & "|"
                Else
                    s = s & "0" & Chr$(col - 11 + 65) & "|"
                End If
            Else
                If col < 11 Then
                    s = s & "0" & col - 1 & "|"
                Else
                    s = s & "0" & Chr$(col - 11 + 65) & "|"
                End If
            End If
            CodeView.col = col
            CodeView.row = 0
            CodeView.CellAlignment = flexAlignCenterCenter
        Next col
        
        CodeView.FormatString = s
        CodeView.Visible = True
        
        FrameMain.Enabled = False
        Exit Sub
    End If
       
    CodeView.Visible = False
    CodeView.Clear
   
    ' Формируем заголовки столбцов
    s = "<   |"
    For col = 1 To CodeView.Cols - 2
        CodeView.ColWidth(col) = 250
        If col < CodeView.Cols - 1 Then
            If col < 11 Then
                s = s & "0" & col - 1 & "|"
            Else
                s = s & "0" & Chr$(col - 11 + 65) & "|"
            End If
        Else
            If col < 11 Then
                s = s & "0" & col - 1 & "|"
            Else
                s = s & "0" & Chr$(col - 11 + 65) & "|"
            End If
        End If
        CodeView.col = col
        CodeView.row = 0
        CodeView.CellAlignment = flexAlignCenterCenter
    Next col
    
    CodeView.FormatString = s
    
    ' Формируем заголовки строк
    Dim HexValue As Long
   
    CodeView.ColWidth(0) = 600
    CodeView.Rows = Manager.ImageSize / 16
    
    For row = 1 To CodeView.Rows - 1
        CodeView.col = 0
        CodeView.row = row
        
        HexValue = (row - 1) * 16
        
        If HexValue < &H10 Then
            CodeView.Text = "0000"
        Else
            If HexValue < &H100 Then
                CodeView.Text = "00" & Hex(HexValue)
            Else
                If HexValue < &H1000 Then
                    CodeView.Text = "0" & Hex(HexValue)
                Else
                    If HexValue < &H10000 Then
                        CodeView.Text = "" & Hex(HexValue)
                    End If
                End If
            End If
        End If
        
        CodeView.CellAlignment = flexAlignRightCenter
    Next row
    
    ' Выводим данные
    row = CodeView.TopRow
    
    Do While CodeView.RowIsVisible(row)
        For col = 1 To CodeView.Cols - 2
            CodeView.col = col
            CodeView.row = row
            CodeView.CellAlignment = flexAlignCenterCenter
            
            If ((16 * (row - 1) + col - 1) < Manager.ImageSize) Then
                HexValue = Manager.GetByte(16 * (row - 1) + col - 1)
            
                If (HexValue < &H10) Then
                    CodeView.Text = "0" & Hex(HexValue)
                Else
                    CodeView.Text = "" & Hex(HexValue)
                End If
            Else
                CodeView.Text = ""
            End If
        Next col
        
        row = row + 1
        If (row > CodeView.Rows - 1) Then Exit Do
    Loop
    
    ' После сделанных изменений можно визуализировать компонент
    CodeView.Visible = True
End Sub

Private Function ValveEnabled(col As Integer, row As Integer) As Boolean
    Dim FuncN As Integer
    
    FuncN = Manager.GetFunctionType(Manager.ProgramIndex + 1, col)
    
    ' Сохраняем изменённое значение
    If FuncN < 12 Then
        Select Case FuncN
            Case WPC_OPERATION_IDLE ' пропуск
                ValveEnabled = ModuleIdle.ValveEnabled(Me, col - 1, row)
                Exit Function
        
            Case WPC_OPERATION_FILL ' Налив
                ValveEnabled = ModuleFill.ValveEnabled(Me, col - 1, row)
                Exit Function
            
            Case WPC_OPERATION_DTRG ' моющие
                ValveEnabled = ModuleDTRG.ValveEnabled(Me, col - 1, row)
                Exit Function
            
            Case WPC_OPERATION_HEAT ' нагрев
                ValveEnabled = ModuleHeat.ValveEnabled(Me, col - 1, row)
                Exit Function
                
            ' стирка, полоскание, расстряска
            Case WPC_OPERATION_WASH, WPC_OPERATION_RINS, WPC_OPERATION_JOLT, WPC_OPERATION_PAUS
                ValveEnabled = ModuleWashOrRinsOrJolt.ValveEnabled(Me, col - 1, row)
                Exit Function
                
'                Case WPC_OPERATION_PAUS ' пауза
'                    ValveEnabled = ModulePause.SetComboPropertyForPause(Me, Col - 1, Row)
'                   Exit Function

            Case WPC_OPERATION_DRAIN ' слив
                ValveEnabled = ModuleDrain.ValveEnabled(Me, col - 1, row)
                Exit Function
                
            Case WPC_OPERATION_SPIN ' отжим
                ValveEnabled = ModuleSpin.ValveEnabled(Me, col - 1, row)
                Exit Function
            
            Case WPC_OPERATION_COOL ' охлаждение
                ValveEnabled = ModuleCool.ValveEnabled(Me, col - 1, row)
                Exit Function
                
            Case WPC_OPERATION_TRIN ' тех.полоскание
                ValveEnabled = ModuleTrin.ValveEnabled(Me, col - 1, row)
                Exit Function
        
            Case Else

        End Select
    End If
    
    ValveEnabled = False
End Function

Private Sub RefreshStepsView()
    Dim col%, row%, X%, Y%, FuncN%
    Dim s As String
    
    ' Выходим из процедуры, если программы не загружены или отсутствуют
    If Not Manager.FileLoaded Then
        FrameMain.Caption = "Шаги"
        StepsView.Visible = False
        StepsView.Clear
        
        StepsView.Cols = MAX_NUMBER_OF_STEPS + 1
        
        s = "<   |"
        For col% = 1 To StepsView.Cols - 1
            StepsView.ColWidth(col%) = 250
            If col% < StepsView.Cols - 1 Then
                If col% < 10 Then
                    s = s & "0" & col% & "|"
                Else
                    s = s & col% & "|"
                End If
            Else
                 s = s & col%
            End If
            StepsView.col = col%
            StepsView.row = 0
            StepsView.CellAlignment = flexAlignCenterCenter
        Next col%
        
        StepsView.FormatString = s
           
        s = ";|" _
            & "Клапан горячей воды" & "|" _
            & "Клапан холодной воды 1" & "|" _
            & "Клапан холодной воды 2" & "|" _
            & "Клапан МС 1" & "|" _
            & "Клапан МС 2" & "|" _
            & "Клапан МС 3" & "|" _
            & "Клапан МС 4" & "|" _
            & "Клапан МС 5" & "|" _
            & "Клапан МС 6" & "|" _
            & "Клапан МС 7" & "|" _
            & "Клапан МС 8" & "|" _
            & "Клапан МС 9" & "|" _
            & "Замок люка 1" & "|" _
            & "Замок люка 2" & "|" _
            & "Слив 1" & "|" _
            & "Слив 2" & "|" _
            & "Нагрев" & "|" _
            & "Мотор"
        
        StepsView.FormatString = s
        
        ' "Тушим" все ячейки таблицы
        For row% = 1 To StepsView.Rows - 1
            For col% = 1 To StepsView.Cols - 1
                StepsView.col = col%
                StepsView.row = row%
                StepsView.CellBackColor = &H8000000F
            Next col%
        Next row%
        
        StepsView.col = 1
        StepsView.row = 1

        StepsView.Visible = True
        FrameMain.Enabled = False
        Exit Sub
    End If
        
    StepsView.Visible = False
    
    X% = StepsView.col
    Y% = StepsView.row
    
    StepsView.Clear
    
    StepsView.Cols = MAX_NUMBER_OF_STEPS + 1
    
    s = "<   |"
    For col% = 1 To StepsView.Cols - 1
        StepsView.ColWidth(col%) = 250
        If col% < StepsView.Cols - 1 Then
            If col% < 10 Then
                s = s & "0" & col% & "|"
            Else
                s = s & col% & "|"
            End If
        Else
            s = s & col%
        End If
        StepsView.col = col%
        StepsView.row = 0
        StepsView.CellAlignment = flexAlignCenterCenter
    Next col%
    
    StepsView.FormatString = s
           
    s = ";|" _
        & "Клапан горячей воды" & "|" _
        & "Клапан холодной воды 1" & "|" _
        & "Клапан холодной воды 2" & "|" _
        & "Клапан МС 1" & "|" _
        & "Клапан МС 2" & "|" _
        & "Клапан МС 3" & "|" _
        & "Клапан МС 4" & "|" _
        & "Клапан МС 5" & "|" _
        & "Клапан МС 6" & "|" _
        & "Клапан МС 7" & "|" _
        & "Клапан МС 8" & "|" _
        & "Клапан МС 9" & "|" _
        & "Замок люка 1" & "|" _
        & "Замок люка 2" & "|" _
        & "Слив 1" & "|" _
        & "Слив 2" & "|" _
        & "Нагрев" & "|" _
        & "Мотор"
    
    StepsView.FormatString = s
    
    ' "Тушим" все ячейки таблицы
    For row% = 1 To StepsView.Rows - 1
        For col% = 1 To StepsView.Cols - 1
            StepsView.col = col%
            StepsView.row = row%
            StepsView.CellBackColor = &H8000000F
        Next col%
    Next row%
    
    'col% = StepsView.LeftCol
    
    StepsView.col = 1
    StepsView.row = 1
    
    ImageChecked.Width = StepsView.CellWidth
    ImageChecked.Height = StepsView.CellHeight
    
    ImageUnchecked.Width = StepsView.CellWidth
    ImageUnchecked.Height = StepsView.CellHeight
    
    ImageGrayed.Width = StepsView.CellWidth
    ImageGrayed.Height = StepsView.CellHeight
    
    For col% = 1 To MAX_NUMBER_OF_STEPS
        FuncN% = Manager.GetFunctionType(Manager.ProgramIndex + 1, col%)
        
        If FuncN% > 0 And FuncN% < 11 Then
            For row% = 1 To StepsView.Rows - 1
                StepsView.col = col%
                StepsView.row = row%
                StepsView.CellAlignment = flexAlignCenterCenter
                
                If StepsViewMode = TEXT_VIEW Then StepsView.Text = Mid$(FunctionsStrings(FuncN%), row%, 1)
                
                If GetLoadingsFromFuncN(FuncN%) And (2 ^ (row% - 1)) Then
                    Select Case StepsViewMode
                        Case TEXT_VIEW:
                            If ValveEnabled(col%, row%) Then
                                StepsView.CellBackColor = &HC000&
                            Else
                                StepsView.CellBackColor = &H8000000F
                            End If
                            StepsView.CellPictureAlignment = flexAlignCenterCenter
                        
                        Case CHECKS_VIEW:
                            StepsView.CellBackColor = &HFFFFFF
                            If ValveEnabled(col%, row%) Then
                                Set StepsView.CellPicture = ImageChecked.Picture
                            Else
                                Set StepsView.CellPicture = ImageUnchecked.Picture
                            End If
                            StepsView.CellPictureAlignment = flexAlignCenterCenter
                        
                    End Select
                Else
                    Select Case StepsViewMode
                        Case TEXT_VIEW: StepsView.CellBackColor = &H80000005
                    
                        Case CHECKS_VIEW:
                            StepsView.CellBackColor = &H80000005
                            Set StepsView.CellPicture = ImageGrayed.Picture
                            StepsView.CellPictureAlignment = flexAlignCenterCenter
                        
                        End Select
                End If
            Next row%
        Else
            For row% = 1 To StepsView.Rows - 1
                StepsView.col = col%
                StepsView.row = row%
                StepsView.Text = ""
                StepsView.CellBackColor = &H8000000F
            Next row%
        End If
        
     Next col%
    
    StepsView.col = X%
    StepsView.row = Y%
    
    ' После сделанных изменений можно визуализировать компонент
    StepsView.Visible = True
End Sub

Private Sub RefreshProperties()
    Dim ParamStr As String

    ' Выходим из процедуры, если программы не загружены или отсутствуют
    If Not Manager.FileLoaded Then
        FrameRight.Caption = "Свойства"
        PropertyTable.Visible = False
        PropertyTable.Rows = 1
        PropertyTable.Clear
        ParamStr = "<Параметр|Значение"
        PropertyTable.FormatString = ParamStr
        PropertyTable.Visible = True
        FrameRight.Enabled = False
        Exit Sub
    End If
    
    FrameRight.Enabled = True

    FrameRight.Caption = "Свойства - [" & ListPrograms.Text & ".Шаг" & Manager.StepIndex + 1 & "]"
    
    ' Узнаём номер функции текущего шага
    Dim FuncN%, row%
    
    FuncN% = Manager.GetFunctionType(Manager.ProgramIndex + 1, Manager.StepIndex + 1)
    
    PropertyTable.Visible = False
    PropertyTable.Rows = 2
    PropertyTable.Clear
    ParamStr = "<Параметр|Значение"
    PropertyTable.FormatString = ParamStr
        
    If FuncN% < 12 Then
        Select Case FuncN%
            Case WPC_OPERATION_IDLE ' пропуск
                ModuleIdle.ShowPropertyTableForIdle Me
        
            Case WPC_OPERATION_FILL ' Налив
                ModuleFill.ShowPropertyTableForFill Me
            
            Case WPC_OPERATION_DTRG ' моющие
                ModuleDTRG.ShowPropertyTableForDTRG Me
            
            Case WPC_OPERATION_HEAT ' нагрев
                ModuleHeat.ShowPropertyTableForHeat Me
                
            ' стирка, полоскание, расстряска
            Case WPC_OPERATION_WASH, WPC_OPERATION_RINS, WPC_OPERATION_JOLT, WPC_OPERATION_PAUS
                ModuleWashOrRinsOrJolt.ShowPropertyTableForWashOrRinsOrJolt Me
                
'            Case WPC_OPERATION_PAUS ' пауза
'                ModulePause.ShowPropertyTableForPause Me

            Case WPC_OPERATION_DRAIN ' слив
                ModuleDrain.ShowPropertyTableForDrain Me
                
            Case WPC_OPERATION_SPIN ' отжим
                ModuleSpin.ShowPropertyTableForSpin Me
            
            Case WPC_OPERATION_COOL ' охлаждение
                ModuleCool.ShowPropertyTableForCool Me
                
            Case WPC_OPERATION_TRIN ' тех.полоскание
                ModuleTrin.ShowPropertyTableForTrin Me
        
            Case Else

        End Select
    End If
    
    PropertyTable.Visible = True
End Sub

Public Sub LoadLimits(FileName As String)
    LimitsLoaded = DoesFileExist(FileName)
    
    If Not LimitsLoaded Then Exit Sub
    
    Dim LimitsFile As New TIniFiles
    
    LimitsFile.Create FileName
    
    ' Настройки заголовка
    EndSound.DefaultValue = LimitsFile.ReadInteger(TITLE_SECTION_NAME, "EndSound.Default", ENDSOUND_DEFAULT) > 0
    DoorUnlock.DefaultValue = LimitsFile.ReadInteger(TITLE_SECTION_NAME, "DoorUnlock.Default", DOORUNLOCK_DEFAULT) > 0
    
    Set LimitsFile = Nothing
End Sub

Private Sub func_SetDefaultProgramTitle(N As Integer, ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE)
    
    Dim StepPointer As Long
    
    StepPointer = Manager.DataPointer + (N - 1) * PROGRAM_SIZE_IN_BYTES
    'CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
    PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer
    
    Select Case EndSound.DefaultValue
        Case False: RecordTitle.LowBits = RecordTitle.LowBits And &HFFFE
        Case True: RecordTitle.LowBits = RecordTitle.LowBits Or &H1
    End Select

    Select Case DoorUnlock.DefaultValue
        Case False: RecordTitle.LowBits = RecordTitle.LowBits And &HFFFD
        Case True: RecordTitle.LowBits = RecordTitle.LowBits Or &H2
    End Select
    
    ' Сохраняем изменения
    'CopyMemory ByVal StepPointer, RecordTitle, HEADER_SIZE_IN_BYTES
End Sub
    
Private Sub SetDefaultProgramTitle(N As Integer)
    Dim RecordTitle As TYPE_WPC_TITLE
    
    func_SetDefaultProgramTitle N, 0&, RecordTitle
End Sub

