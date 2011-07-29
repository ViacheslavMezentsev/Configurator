VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FormMain 
   Caption         =   "Конфигуратор УП"
   ClientHeight    =   7128
   ClientLeft      =   2532
   ClientTop       =   1944
   ClientWidth     =   8916
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7128
   ScaleWidth      =   8916
   Begin VB.Frame FrameLog 
      BorderStyle     =   0  'None
      Height          =   1092
      Left            =   5640
      TabIndex        =   20
      Top             =   5640
      Visible         =   0   'False
      Width           =   3252
      Begin VB.TextBox TextLog 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.2
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   480
         Width           =   3012
      End
      Begin VB.Label LabelLogCaption 
         BackColor       =   &H8000000D&
         Caption         =   " Журнал"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   252
         Left            =   120
         MousePointer    =   7  'Size N S
         TabIndex        =   21
         Top             =   120
         Width           =   3000
      End
   End
   Begin VB.Timer TimerAutoUpdate 
      Interval        =   60000
      Left            =   3720
      Top             =   5640
   End
   Begin VB.Timer TimerLogAnimate 
      Interval        =   2000
      Left            =   5160
      Top             =   6240
   End
   Begin VB.Frame FrameLogSplitter 
      BorderStyle     =   0  'None
      Height          =   60
      Left            =   4320
      MousePointer    =   7  'Size N S
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   4572
   End
   Begin MSComctlLib.ImageList ImageListMainToolbar_32x32 
      Left            =   960
      Top             =   5640
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   12632256
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":6432
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":68C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":6E65
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":73AB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   2520
      Top             =   5640
   End
   Begin VB.Frame SplitterLeft 
      BorderStyle     =   0  'None
      Height          =   4932
      Left            =   2400
      MousePointer    =   9  'Size W E
      TabIndex        =   14
      Top             =   600
      Width           =   60
   End
   Begin MSComDlg.CommonDialog SaveFileDialog 
      Left            =   1920
      Top             =   5640
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog OpenFileDialog 
      Left            =   1560
      Top             =   5640
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      DefaultExt      =   "*.bin"
      DialogTitle     =   "Открыть"
      Filter          =   "Файлы проекта (*.bin)|*.bin|Конфигуратор УП 1.2 (*.js)|*.js"
      FilterIndex     =   1
   End
   Begin VB.Frame SplitterRight 
      BorderStyle     =   0  'None
      Height          =   4812
      Left            =   6516
      MousePointer    =   9  'Size W E
      TabIndex        =   13
      Top             =   600
      Width           =   60
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   312
      Left            =   0
      TabIndex        =   12
      Top             =   6816
      Width           =   8916
      _ExtentX        =   15727
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10809
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
      TabIndex        =   11
      Top             =   480
      Width           =   3972
      Begin VB.Frame FrameCodeView 
         BorderStyle     =   0  'None
         Height          =   2172
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Visible         =   0   'False
         Width           =   3732
         Begin VB.TextBox TextByte 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   7.8
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   120
            MaxLength       =   3
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1800
            Visible         =   0   'False
            Width           =   612
         End
         Begin MSFlexGridLib.MSFlexGrid CodeView 
            Height          =   1572
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   2772
            _ExtentX        =   4890
            _ExtentY        =   2773
            _Version        =   393216
            Cols            =   17
            HighLight       =   0
            GridLines       =   0
            BorderStyle     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   204
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
         TabIndex        =   15
         Top             =   240
         Width           =   3732
         Begin MSFlexGridLib.MSFlexGrid StepsView 
            Height          =   1452
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   2292
            _ExtentX        =   4043
            _ExtentY        =   2561
            _Version        =   393216
            Rows            =   16
            Cols            =   81
            AllowBigSelection=   0   'False
            HighLight       =   0
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
      TabIndex        =   10
      Top             =   480
      Width           =   2172
      Begin VB.TextBox TextName 
         BorderStyle     =   0  'None
         Height          =   288
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   4680
         Visible         =   0   'False
         Width           =   732
      End
      Begin MSFlexGridLib.MSFlexGrid ListPrograms 
         Height          =   4332
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1932
         _ExtentX        =   3408
         _ExtentY        =   7641
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         HighLight       =   0
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
      TabIndex        =   9
      Top             =   480
      Width           =   2292
      Begin VB.ComboBox ComboCell 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   4560
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.TextBox TextCell 
         BorderStyle     =   0  'None
         Height          =   288
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   4200
         Visible         =   0   'False
         Width           =   732
      End
      Begin MSFlexGridLib.MSFlexGrid PropertyTable 
         Height          =   3852
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2052
         _ExtentX        =   3620
         _ExtentY        =   6795
         _Version        =   393216
         AllowBigSelection=   0   'False
         HighLight       =   0
         AllowUserResizing=   1
         BorderStyle     =   0
      End
      Begin VB.Shape ShapeDescription 
         Height          =   564
         Left            =   1200
         Top             =   4332
         Width           =   960
      End
      Begin VB.Label LabelDescription 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   17
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
      Top             =   5640
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":78B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":7C05
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":7F59
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMain.frx":82AD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListSquares 
      Left            =   480
      Top             =   5640
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
            Picture         =   "FormMain.frx":8601
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   8916
      _ExtentX        =   15727
      _ExtentY        =   847
      ButtonWidth     =   826
      ButtonHeight    =   804
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageListMainToolbar_32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
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
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Настройки"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Image ImageGrayed 
      Height          =   192
      Left            =   3480
      Picture         =   "FormMain.frx":8955
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image ImageChecked 
      Height          =   192
      Left            =   3240
      Picture         =   "FormMain.frx":8CB5
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image ImageUnchecked 
      Appearance      =   0  'Flat
      Height          =   192
      Left            =   3000
      Picture         =   "FormMain.frx":9027
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
      Begin VB.Menu Separator1 
         Caption         =   "-"
      End
      Begin VB.Menu ImportMainMenuItem 
         Caption         =   "&Импорт..."
      End
      Begin VB.Menu ExportMainMenuItem 
         Caption         =   "&Экспорт..."
      End
      Begin VB.Menu Separator2 
         Caption         =   "-"
      End
      Begin VB.Menu MRUListMenu 
         Caption         =   "История"
         Begin VB.Menu MRUItems 
            Caption         =   ""
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu MRUItems 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu MRUItems 
            Caption         =   ""
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu MRUItems 
            Caption         =   ""
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu MRUItems 
            Caption         =   ""
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu MRUItems 
            Caption         =   ""
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu MRUItems 
            Caption         =   ""
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu MRUItems 
            Caption         =   ""
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu MRUItems 
            Caption         =   ""
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu MRUItems 
            Caption         =   ""
            Index           =   10
            Visible         =   0   'False
         End
         Begin VB.Menu Separator3 
            Caption         =   "-"
         End
         Begin VB.Menu ClearHistoryMenuItem 
            Caption         =   "&Очистить"
         End
      End
      Begin VB.Menu Separator4 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMainMenuItem 
         Caption         =   "&Выход"
      End
   End
   Begin VB.Menu ViewMainMenuItem 
      Caption         =   "&Вид"
      Begin VB.Menu MenuItemShowHideLog 
         Caption         =   "&Журнал"
         Shortcut        =   ^L
      End
      Begin VB.Menu Separator5 
         Caption         =   "-"
      End
      Begin VB.Menu OptionsMainMenuItem 
         Caption         =   "&Настройки..."
      End
   End
   Begin VB.Menu PopupMenuPrograms 
      Caption         =   "П&рограмма"
      Begin VB.Menu GotoMenuItem 
         Caption         =   "&Перейти..."
         Shortcut        =   ^G
      End
      Begin VB.Menu PopupMenuListClear 
         Caption         =   "&Очистить"
      End
      Begin VB.Menu CopyMainMenuItem 
         Caption         =   "&Копировать..."
      End
      Begin VB.Menu Separator6 
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
      Begin VB.Menu MenuItemDoUpdate 
         Caption         =   "О&бновить"
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

'**
'@rem Режим отображения средней панели
Private ViewMode As Integer
'**
'@rem Режимы отображения таблицы шагов
Private StepsViewMode As Integer
'**
'@rem
Private FileName As String

'**
'@rem
Public ModuleIdle As CModuleIdle
'**
'@rem
Public ModuleFill As CModuleFill
'**
'@rem
Public ModuleDTRG As CModuleDTRG
'**
'@rem
Public ModuleHeat As CModuleHeat
'**
'@rem
Public ModuleWashOrRinsOrJolt As CModuleWashOrRinsOrJolt

'<Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:16:52
'Причина: Модуль аналогичен по функционалу с ModuleWashOrRinsOrJolt>
'Public ModulePause As TModulePause
'</Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:16:52>

'**
'@rem
Public ModuleDrain As CModuleDrain
'**
'@rem
Public ModuleSpin As CModuleSpin
'**
'@rem
Public ModuleCool As CModuleCool
'**
'@rem
Public ModuleTrin As CModuleTrin

'**
'@rem
Private WithEvents Kachalka As clsKachalka
Attribute Kachalka.VB_VarHelpID = -1

Dim SplitterRightMoving As Boolean
Dim SplitterLeftMoving As Boolean
Dim BegX As Integer, BegY As Integer

'**
'@see
'@rem Сохранение внешнего вида интерфейса.
Private Sub SavePlacement()
    '<EhHeader>
    On Error GoTo SavePlacement_Err
    '</EhHeader>

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

    '<EhFooter>
    Exit Sub

SavePlacement_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.SavePlacement]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

'**
'@see
'@rem Загружаем настройки внешнего вида интерфейса.
Private Sub LoadPlacement()
    '<EhHeader>
    On Error GoTo LoadPlacement_Err
    '</EhHeader>
    
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
    Dim Result As Integer
    
    Path = String$(255, 0)
    Result = GetModuleFileName(0, Path, 254)
    Path = MiscExtractPathName(Path, True)
    
    CurrentDir = IniFile.ReadString("Settings", "CurrentDir", Path)
    
    '<EhFooter>
    Exit Sub

LoadPlacement_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.LoadPlacement]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Public Sub RefreshComponents(ByVal FramesOnly As Boolean)
    '<EhHeader>
    On Error GoTo RefreshComponents_Err
    '</EhHeader>

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
    RefrefhFrameLog
    RefreshStatusBar
    
    '<EhFooter>
    Exit Sub

RefreshComponents_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.RefreshComponents]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub RefreshFrameLeft()
    '<EhHeader>
    On Error GoTo RefreshFrameLeft_Err
    '</EhHeader>
    
    ListPrograms.Left = 120
    ListPrograms.Width = FrameLeft.Width - ListPrograms.Left - 120
    ListPrograms.Height = FrameLeft.Height - ListPrograms.Top - 120
    
    ListPrograms.ColWidth(0) = ListPrograms.Width
    FrameLeft.FontSize = Settings.StepsViewFontSize
    
    '<EhFooter>
    Exit Sub

RefreshFrameLeft_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.RefreshFrameLeft]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub RefreshFrameMain()
    '<EhHeader>
    On Error GoTo RefreshFrameMain_Err
    '</EhHeader>

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
            ' Поэтому нужно делать обновление после изменения размеров
            RefreshCodeView
            
            FrameCodeView.Visible = True
            
    End Select
    
    FrameMain.Enabled = Manager.FileLoaded
    FrameMain.FontSize = Settings.StepsViewFontSize
    
    '<EhFooter>
    Exit Sub

RefreshFrameMain_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.RefreshFrameMain]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub RefreshFrameRight()
    '<EhHeader>
    On Error GoTo RefreshFrameRight_Err
    '</EhHeader>
    
    PropertyTable.Left = 120
    PropertyTable.Top = 240
    PropertyTable.Width = FrameRight.Width - PropertyTable.Left - 120
    
    If LabelDescription.Visible Then
    
        PropertyTable.Height = FrameRight.Height - PropertyTable.Top - LabelDescription.Height - 120
        LabelDescription.Top = PropertyTable.Top + PropertyTable.Height
        LabelDescription.Width = PropertyTable.Width
        ShapeDescription.Top = LabelDescription.Top
        ShapeDescription.Width = PropertyTable.Width
        
    Else
        PropertyTable.Height = FrameRight.Height - PropertyTable.Top - 120
    End If
    
    FrameRight.FontSize = Settings.StepsViewFontSize
    
    If PropertyTable.Width > PropertyTable.ColWidth(0) Then
    
        PropertyTable.ColWidth(1) = PropertyTable.Width - PropertyTable.ColWidth(0)
    End If
    
    TextCell.Width = PropertyTable.ColWidth(1)
    ComboCell.Width = PropertyTable.ColWidth(1)
    
    '<EhFooter>
    Exit Sub

RefreshFrameRight_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.RefreshFrameRight]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub RefreshForm()
    '<EhHeader>
    On Error GoTo RefreshForm_Err
    '</EhHeader>
    
    SetCaption Manager.FileName
    
    TextCell.Font.Name = Settings.StepsViewFontName
    TextCell.Font.Size = Settings.StepsViewFontSize
    
    TextName.Font.Name = Settings.StepsViewFontName
    TextName.Font.Size = Settings.StepsViewFontSize
    
    TextByte.Font.Size = Settings.StepsViewFontSize

    LabelDescription.Font.Name = Settings.StepsViewFontName
    LabelDescription.Font.Size = Settings.StepsViewFontSize
    
    ComboCell.Font.Name = Settings.StepsViewFontName
    ComboCell.Font.Size = Settings.StepsViewFontSize
    
    '<EhFooter>
    Exit Sub

RefreshForm_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.RefreshForm]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub RefreshMainMenu()
    '<EhHeader>
    On Error GoTo RefreshMainMenu_Err
    '</EhHeader>
    
    PopupMenuPrograms.Visible = Manager.FileLoaded
    
    Select Case ViewMode
        Case STEPS_VIEW: GotoMenuItem.Visible = False
        Case CODE_VIEW: GotoMenuItem.Visible = True
    End Select
    
    StepMainMenuItem.Visible = Manager.FileLoaded And (ViewMode = STEPS_VIEW)
    
    '<EhFooter>
    Exit Sub

RefreshMainMenu_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.RefreshMainMenu]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub RefrefhFrameLog()
    '<EhHeader>
    On Error GoTo RefrefhFrameLog_Err
    '</EhHeader>

    FrameLog.Left = FormMain.ScaleLeft
    FrameLog.Height = 1440
    FrameLog.Top = StatusBar.Top - FrameLog.Height
    FrameLog.Width = FormMain.ScaleWidth
    
    LabelLogCaption.Top = 0
    LabelLogCaption.Left = 0
    LabelLogCaption.Width = FrameLog.Width
    
    TextLog.Left = LabelLogCaption.Left
    TextLog.Top = LabelLogCaption.Height
    TextLog.Width = FrameLog.Width
    TextLog.Height = FrameLog.Height - TextLog.Top
    
    '<EhFooter>
    Exit Sub

RefrefhFrameLog_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormMain.RefrefhFrameLog]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub RefreshStatusBar()
    '<EhHeader>
    On Error GoTo RefreshStatusBar_Err
    '</EhHeader>

    If Modified Then
        StatusBar.Panels(2).Text = "Изменён"
    Else
        StatusBar.Panels(2).Text = ""
    End If
    
    '<EhFooter>
    Exit Sub

RefreshStatusBar_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.RefreshStatusBar]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub AboutMainMenuItem_Click()
    '<EhHeader>
    On Error GoTo AboutMainMenuItem_Click_Err
    '</EhHeader>
    
    FormAbout.Show (vbModal)
    
    '<EhFooter>
    Exit Sub

AboutMainMenuItem_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.AboutMainMenuItem_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub ClearHistoryMenuItem_Click()
    '<EhHeader>
    On Error GoTo ClearHistoryMenuItem_Click_Err
    '</EhHeader>

    MRUFileList.ClearHistory
    DisplayMRU
    
    '<EhFooter>
    Exit Sub

ClearHistoryMenuItem_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.ClearHistoryMenuItem_Click]: " _
        & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub CloseMainMenuItem_Click()
    '<EhHeader>
    On Error GoTo CloseMainMenuItem_Click_Err
    '</EhHeader>

    If Modified = True Then
    
        Dim vbRes As Integer
        
        vbRes = MsgBox("Сохранить изменения в файле:" & _
           VBA.Constants.vbCrLf & VBA.Constants.vbCrLf & """" & Manager.FileName & """?", _
           vbYesNoCancel + vbQuestion, APP_NAME)
        
        Select Case vbRes
        
            Case vbYes
            
                SaveMainMenuItem_Click
                Manager.CloseFile
                SetModified False
                RefreshComponents False
                
            Case vbNo
            
                Manager.CloseFile
                SetModified False
                RefreshComponents False
                
            Case vbCancel
        
        End Select
        
    Else
    
        Manager.CloseFile
        RefreshComponents False
        
    End If
    
    '<EhFooter>
    Exit Sub

CloseMainMenuItem_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.CloseMainMenuItem_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub CodeView_Click()
    '<EhHeader>
    On Error GoTo CodeView_Click_Err
    '</EhHeader>
    
    Dim x As Integer, Y As Integer
    Dim col As Integer, row As Integer
    
    CodeView.Visible = False
    
    x = CodeView.col
    Y = CodeView.row

    For col = 1 To CodeView.Cols - 2
        CodeView.col = col
        CodeView.row = 0
        CodeView.CellFontBold = False
    Next
    
    row = CodeView.TopRow
    
    Do While CodeView.RowIsVisible(row)
        CodeView.col = 0
        CodeView.row = row
    
        CodeView.CellFontBold = False
        row = row + 1

        If row > CodeView.rows - 1 Then Exit Do
    Loop
    
    CodeView.row = 0
    CodeView.col = x
    CodeView.CellFontBold = True
    
    CodeView.row = Y
    CodeView.col = 0
    CodeView.CellFontBold = True
    
    CodeView.col = x
    CodeView.row = Y
    
    CodeView.Visible = True
    CodeView.SetFocus
    
    '<EhFooter>
    Exit Sub

CodeView_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.CodeView_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub CodeView_DblClick()
    '<EhHeader>
    On Error GoTo CodeView_DblClick_Err
    '</EhHeader>
    
    CodeView_KeyDown VBRUN.KeyCodeConstants.vbKeyReturn, 0
    
    '<EhFooter>
    Exit Sub

CodeView_DblClick_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.CodeView_DblClick]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub CodeView_KeyDown(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo CodeView_KeyDown_Err
    '</EhHeader>

    ' При нажатии Enter в ячейке даём возможность редактировать её содержимое
    'If Not (KeyCode = VBRUN.KeyCodeConstants.vbKeyReturn) Then Exit Sub
    
    ' Фильтруем не нужные клавиши
    Select Case KeyCode
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
        
        Case Else: Exit Sub
            
    End Select
    
    Dim col As Integer, row As Integer
    
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
    
    '<EhFooter>
    Exit Sub

CodeView_KeyDown_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.CodeView_KeyDown]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub CodeView_Scroll()
    '<EhHeader>
    On Error GoTo CodeView_Scroll_Err
    '</EhHeader>
    
    RefreshCodeView
    
    '<EhFooter>
    Exit Sub

CodeView_Scroll_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.CodeView_Scroll]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub ComboCell_KeyDown(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo ComboCell_KeyDown_Err
    '</EhHeader>

    If KeyCode = VBRUN.KeyCodeConstants.vbKeyEscape Then
        ComboCell.Visible = False
        LabelDescription.Visible = False
        ShapeDescription.Visible = False
        RefreshFrameRight
        PropertyTable.SetFocus
    End If
    
    If KeyCode = VBRUN.KeyCodeConstants.vbKeyReturn Then
    
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
                    
'<Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:18:47
'Причина: Модуль аналогичен по функционалу с ModuleWashOrRinsOrJolt>
'                Case WPC_OPERATION_PAUS ' пауза
'                    ModulePause.SetComboPropertyForPause Me
'</Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:18:47>
    
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
            Dim CRC8Value As Byte
            Dim row As Integer
            Dim Address As Long
            Dim Size As Long
            
            Address = Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
            Size = PROGRAM_SIZE_IN_BYTES - 1
            
            CRC8Value = Manager.CalculateCRC8(Address + 1, Size)
            Manager.SetByte Address, CRC8Value

            ComboCell.Visible = False
            LabelDescription.Visible = False
            ShapeDescription.Visible = False
            
            row = PropertyTable.row
            RefreshComponents False

            If row < PropertyTable.rows - 1 Then PropertyTable.row = row
            
            PropertyTable.SetFocus
            
        End If
    End If
    
    '<EhFooter>
    Exit Sub

ComboCell_KeyDown_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.ComboCell_KeyDown]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub ComboCell_LostFocus()
    '<EhHeader>
    On Error GoTo ComboCell_LostFocus_Err
    '</EhHeader>
    
    ComboCell.Visible = False
    LabelDescription.Visible = False
    ShapeDescription.Visible = False
    RefreshFrameRight
    
    '<EhFooter>
    Exit Sub

ComboCell_LostFocus_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.ComboCell_LostFocus]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub CopyMainMenuItem_Click()
    '<EhHeader>
    On Error GoTo CopyMainMenuItem_Click_Err
    '</EhHeader>
    
    Dim I As Integer
    
    FormCopy.List1.Clear
    FormCopy.List2.Clear
    
    For I = 1 To Manager.ProgramsCount
        FormCopy.List1.AddItem ListPrograms.TextArray(GetCellIndex(ListPrograms, I, 0))
        FormCopy.List2.AddItem ListPrograms.TextArray(GetCellIndex(ListPrograms, I, 0))
    Next
    
    FormCopy.List1.ListIndex = 0
    FormCopy.List2.ListIndex = 0
    
    FormCopy.Show (vbModal)
    
    RefreshComponents False
    
    '<EhFooter>
    Exit Sub

CopyMainMenuItem_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.CopyMainMenuItem_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub DeleteStepMenuItem_Click()
    '<EhHeader>
    On Error GoTo DeleteStepMenuItem_Click_Err
    '</EhHeader>

    ' Удаляем текущий шаг
    Manager.DeleteStep
    
    ' Пересчитываем CRC поле записи программы
    Dim CRC8Value As Byte
    Dim Address As Long
    Dim Size As Long
    
    Address = Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
    Size = PROGRAM_SIZE_IN_BYTES - 1
    
    CRC8Value = Manager.CalculateCRC8(Address + 1, Size)
    Manager.SetByte Address, CRC8Value
    
    SetModified True
    RefreshDataComponents
    
    StepsView.SetFocus
    
    '<EhFooter>
    Exit Sub

DeleteStepMenuItem_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.DeleteStepMenuItem_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub ExitMainMenuItem_Click()
    '<EhHeader>
    On Error GoTo ExitMainMenuItem_Click_Err
    '</EhHeader>
    
    ' Выходим из программы
    Unload Me
    
    '<EhFooter>
    Exit Sub

ExitMainMenuItem_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.ExitMainMenuItem_Click]: " _
        & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub ExportMainMenuItem_Click()
    '<EhHeader>
    On Error GoTo ExportMainMenuItem_Click_Err
    '</EhHeader>
    
    Dim FName As String

    ' Если у имени файла есть расширение, то меняем его, иначе - добавляем
    If InStrRev(Manager.FileName, ".") = 0 Then
        SaveFileDialog.FileName = Manager.FileName & "." & "json"
    Else
        SaveFileDialog.FileName = Left$(Manager.FileName, InStrRev(Manager.FileName, ".")) & "json"
    End If
    
    SaveFileDialog.DialogTitle = "Экспорт файла..."
    SaveFileDialog.DefaultExt = ".json"
    SaveFileDialog.Filter = "Конфигуратор УП 1.x (*.json)|*.json"
    SaveFileDialog.FilterIndex = 1
    SaveFileDialog.MaxFileSize = 32767
    SaveFileDialog.InitDir = CurrentDir
    SaveFileDialog.CancelError = True
    
    SaveFileDialog.ShowSave

    FName = SaveFileDialog.FileName
    
    If FName <> "" Then
    
        Manager.ExportToJSON FName
        
    End If

    '<EhFooter>
    Exit Sub

ExportMainMenuItem_Click_Err:
    If Err.Number = cdlCancel Then
    Else
        App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.ExportMainMenuItem_Click]: " _
           & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    End If
    
    Resume Next
    '</EhFooter>
End Sub

Private Sub FileMainMenuItem_Click()
    '<EhHeader>
    On Error GoTo FileMainMenuItem_Click_Err
    '</EhHeader>

    ' Настраиваем доступность пунктов меню "Файл":
    
    ' "Файл\Сохранить"
    SaveMainMenuItem.Enabled = Modified
    
    ' "Файл\Сохранить как..."
    SaveAsMainMenuItem.Enabled = Manager.FileLoaded
    
    ' "Файл\Экспорт..."
    ExportMainMenuItem.Enabled = Manager.FileLoaded
    
    ' "Файл\Закрыть"
    CloseMainMenuItem.Enabled = Manager.FileLoaded
    
    '<EhFooter>
    Exit Sub

FileMainMenuItem_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.FileMainMenuItem_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo Form_KeyDown_Err
    '</EhHeader>

    Dim col As Integer, row As Integer
    
    If KeyCode = VBRUN.KeyCodeConstants.vbKeyF3 And Shift = 0 Then

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
    
    If KeyCode = VBRUN.KeyCodeConstants.vbKeyF4 And Shift = 0 Then

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
    
    '<EhFooter>
    Exit Sub

Form_KeyDown_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.Form_KeyDown]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    
    Resume Next
    '</EhFooter>
End Sub

Private Sub DisplayMRU()
    '<EhHeader>
    On Error GoTo DisplayMRU_Err
    '</EhHeader>

    Dim iFile As Long
    
    ' Here I am assuming the MRU is held in a menu array
    ' called mnuFile, to start at Index 1:

    For iFile = 1 To MRUFileList.FileCount

        If (MRUFileList.FileExists(iFile)) Then
            MRUItems(iFile).Visible = True
            MRUItems(iFile).Caption = MRUFileList.MenuCaption(iFile, Settings.FilesHistoryLimitPaths)
            MRUItems(iFile).Tag = CStr(iFile)
        End If
    Next
     
    MRUListMenu.Enabled = (MRUFileList.FileCount > 0)
    
    '<EhFooter>
    Exit Sub

DisplayMRU_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.DisplayMRU]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub
 
Private Sub Form_Load()

    On Local Error GoTo Form_Load_Err
    
    Dim itm As ListItem
    Dim sitm As ListSubItem
    Dim IniFilePath As String
    Dim Result As Integer

    Debug.Print "----------------------------------------------------------------------------"
    Debug.Print Date & " " & Time & ": " & "Версия: " & App.Major & "." & App.Minor & "." & App.Revision
    
    KeyPreview = True
    
    ' Среда разработки часто "вылетает" из-за кода внутри
    ' Поэтому его тестирование нужно проводить только на
    ' откомпилированном приложении
    
    Dim WE_ARE_IN_IDE As Boolean
    
    Debug.Assert MakeTrue(WE_ARE_IN_IDE)
    
    If WE_ARE_IN_IDE Then
    
        ' Код, выполняемый в runtime среды разработки
        DesignMode = True
        Debug.Print Date & " " & Time & " [cop.FormOptions.Form_Load]: " & "Режим разработки."
        
    Else
    
        ' Код, который будет в скомпилированном файле
        DesignMode = False
        
        Timer1.Enabled = True
        Timer1.Interval = 0
    
        HookKeyboard Timer1
        
    End If
    
    ' Создаём контейнер для работы с ошибками
    Set ВекторОшибок = New JVector
    
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

    Debug.Print Date & " " & Time & " [cop.FormOptions.Form_Load]: " & "Текущий путь: " & CurrentDir
    Debug.Print Date & " " & Time & " [cop.FormOptions.Form_Load]: " & "Файл настроек: " & IniFilePath
    
    ' Создаём экземпляр объекта
    Set IniFile = New CIniFiles
    IniFile.Create (IniFilePath)
    
    ' Настройки программы
    Set Settings = New CSettings
    Settings.LoadSettings
    
    Debug.Print Date & " " & Time & " [cop.FormOptions.Form_Load]: " & "Файл лога: " & Settings.LogFilePath
    
    ' При загрузке выставляем флаг необходимости обновления
    ' Он будет действовать до срабатывания таймера автообновления
    If Settings.AutoUpdateEnabled Then AutoUpdateState = AUS_NOT_UPDATED
   
    'TODO: Проверить корректность всех файловых путей
    ' VBRUN.LogModeConstants.vbLogOverwrite не работает по невыясненной причине
    If Settings.RewriteLogFile Then
    
        Debug.Print Date & " " & Time & " [cop.FormOptions.Form_Load]: " & "Файл лога удалён."
        DeleteFile Settings.LogFilePath
        
    End If
    
    App.StartLogging Settings.LogFilePath, VBRUN.LogModeConstants.vbLogToFile
    
    ' Версия программы
    Dim MAX_PATH As Long
    Dim Length As Long
    Dim strFile As String
    Dim szCurrDir As String, szUserName As String
    
    Dim udtFileInfo As FILEINFO
        
    strFile = String(255, 0)
    GetModuleFileName 0, strFile, 255

    If GetFileVersionInformation(strFile, udtFileInfo) = eNoVersion Then
        
        udtFileInfo.FileVersion = "Версия " & App.Major & "." & App.Minor & "." & App.Revision
    
    Else
        
        udtFileInfo.FileVersion = "Версия " & udtFileInfo.FileVersion
        
    End If
    
    MAX_PATH = 255
    szCurrDir = Space(255)
    Length = GetCurrentDirectory(MAX_PATH, szCurrDir)
    szCurrDir = Left$(szCurrDir, Length)
    
    szUserName = Space(255)
    GetUserName szUserName, Length
    szUserName = Left$(szUserName, Length - 1)
    
    App.LogEvent VBA.Constants.vbCrLf & VBA.Constants.vbCrLf _
       & "-----------------------------------------------------------------------" & vbCrLf _
       & "Конфигуратор управляющих программ" & vbCrLf _
       & udtFileInfo.FileVersion & vbCrLf _
       & "Уникальный идентификатор (GUID): " & ProgramGUID & vbCrLf _
       & "Дата запуска: " & Date & " г. в " & Time & vbCrLf _
       & "Операционная система: " & GetOSVersion & vbCrLf _
       & "Имя пользователя: " & szUserName & vbCrLf _
       & "Текущая папка: " & szCurrDir & vbCrLf _
       & "-----------------------------------------------------------------------" & vbCrLf, _
       VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
       
    ' Создаём экземпляр объекта
    Set Manager = New CProgramManager
    
    ' Создаём экземпляр объекта
    Set MRUFileList = New cMRUFileList
    
    App.HelpFile = CurrentDir & "\cop.chm"
    
    Debug.Print Date & " " & Time & " [cop.FormOptions.Form_Load]: " & "Файл справки: " & App.HelpFile
    
    ' Начальные пути для диалоговых окон
    OpenFileDialog.InitDir = CurrentDir
    SaveFileDialog.InitDir = CurrentDir
    
    Set ModuleIdle = New CModuleIdle
    Set ModuleFill = New CModuleFill
    Set ModuleDTRG = New CModuleDTRG
    Set ModuleHeat = New CModuleHeat
    Set ModuleWashOrRinsOrJolt = New CModuleWashOrRinsOrJolt
'<Удалил: Мезенцев Вячеслав, 24.06.2011 г. в 1:31:58
'Причина: Модуль аналогичен по функционалу с ModuleWashOrRinsOrJolt>
'    Set ModulePause = New TModulePause
'</Удалил: Мезенцев Вячеслав, 24.06.2011 г. в 1:31:58>
    Set ModuleDrain = New CModuleDrain
    Set ModuleSpin = New CModuleSpin
    Set ModuleCool = New CModuleCool
    Set ModuleTrin = New CModuleTrin
        
    IniFilePath = CurrentDir & "\limits.ini"
    
    Debug.Print Date & " " & Time & " [cop.FormOptions.Form_Load]: " & "Файл уставок: " & IniFilePath
    
    LoadLimits IniFilePath
    
    If LimitsLoaded Then Debug.Print Date & " " & Time & _
        " [cop.FormOptions.Form_Load]: " & "Уставки загружены."
    
    ModuleIdle.LoadLimits IniFilePath
    ModuleFill.LoadLimits IniFilePath
    ModuleDTRG.LoadLimits IniFilePath
    ModuleHeat.LoadLimits IniFilePath
    ModuleWashOrRinsOrJolt.LoadLimits IniFilePath
'<Удалил: Мезенцев Вячеслав, 24.06.2011 г. в 1:34:29
'Причина: Модуль аналогичен по функционалу с ModuleWashOrRinsOrJolt>
'    ModulePause.LoadLimits IniFilePath
'</Удалил: Мезенцев Вячеслав, 24.06.2011 г. в 1:34:29>
    ModuleDrain.LoadLimits IniFilePath
    ModuleSpin.LoadLimits IniFilePath
    ModuleCool.LoadLimits IniFilePath
    ModuleTrin.LoadLimits IniFilePath
    
    SetModified False
    
    ' Восстанавливаем положение формы и компонентов
    LoadPlacement
    
    ' Восстанавливаем список используемых файлов
    MRUFileList.Load IniFile
    DisplayMRU
    
    Dim s As String
    Dim col As Integer, row
    
    StepsView.Redraw = False
    
    StepsView.Font.Bold = Settings.StepsViewFontBold
    StepsView.Font.Italic = Settings.StepsViewFontItalic
    StepsView.Font.Name = Settings.StepsViewFontName
    StepsView.Font.Size = Settings.StepsViewFontSize
        
    StepsView.Cols = MAX_NUMBER_OF_STEPS + 1
    
    s = "<   |"

    For col = 1 To StepsView.Cols - 1
        StepsView.ColWidth(col) = 250

        If col < StepsView.Cols - 1 Then
            s = s & col & "|"
        Else
            s = s & col
        End If
        StepsView.col = col
        StepsView.row = 0
        StepsView.CellAlignment = flexAlignCenterCenter
    Next
    
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

    For row = 1 To StepsView.rows - 1
        StepsView.RowHeight(row) = Settings.RowHeight
        
        For col = 1 To MAX_NUMBER_OF_STEPS
            StepsView.ColWidth(col) = Settings.StepColWidth
            StepsView.col = col
            StepsView.row = row
            StepsView.CellBackColor = &H8000000F
        Next
    Next
    
    StepsView.col = 1
    StepsView.row = 1
    
    StepsView.Redraw = True
    
    ' Инициализируем окно кода
    s = "<   |"

    For col = 1 To CodeView.Cols - 1
        CodeView.ColWidth(col) = Settings.StepColWidth

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
    Next
    
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

    ' Создаём образ "качалки"
    Set Kachalka = New clsKachalka

    ' Обновляем вид
    RefreshComponents False
    
    ' Симулируем изменение размером формы для вызова Resize()
    Move Left, Top, Width, Height
    
    Exit Sub
    
Form_Load_Err:
    ' Обновляем вид
    RefreshComponents False
    
    ' Симулируем изменение размером формы для вызова Resize()
    Move Left, Top, Width, Height
    
End Sub

Private Sub Form_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    
    RefreshComponents True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '<EhHeader>
    On Error GoTo Form_Unload_Err
    '</EhHeader>

    If Modified = True Then
    
        Dim vbRes As Integer
        
        vbRes = MsgBox("Сохранить изменения в файле:" & _
           VBA.Constants.vbCrLf & VBA.Constants.vbCrLf & _
           """" & Manager.FileName & """?", _
           vbYesNoCancel + vbQuestion, APP_NAME)
           
        Select Case vbRes
        
            Case vbYes: SaveMainMenuItem_Click
            
            Case vbNo:
            
            Case vbCancel
                Cancel = 1
                Exit Sub
                
        End Select
        
    End If
    
    Settings.SaveSettings
    
    ' Сохраняем настройки интерфейса
    SavePlacement
    
    ' Сохраняем список используемых файлов
    MRUFileList.Save IniFile
    
    Unload Me
    Unload FormDownload
    
    Set FormDownload = Nothing
    Set FormMain = Nothing
    
    UnHookKeyboard
    
    ВекторОшибок.removeAllElements
    Set ВекторОшибок = Nothing
    
    '<EhFooter>
    Exit Sub

Form_Unload_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormMain.Form_Unload]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub GotoMenuItem_Click()
    '<EhHeader>
    On Error GoTo GotoMenuItem_Click_Err
    '</EhHeader>
    
    FormGoto.Show (vbModal)
    
    '<EhFooter>
    Exit Sub

GotoMenuItem_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.GotoMenuItem_Click]: " _
        & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub HelpMainMenuSubItem_Click()
    '<EhHeader>
    On Error GoTo HelpMainMenuSubItem_Click_Err
    '</EhHeader>

    If DoesFileExist(App.HelpFile) Then
        Shell ("hh " & App.HelpFile), vbNormalFocus
    End If
    
    '<EhFooter>
    Exit Sub

HelpMainMenuSubItem_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.HelpMainMenuSubItem_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub ImportMainMenuItem_Click()
    '<EhHeader>
    On Error GoTo ImportMainMenuItem_Click_Err
    '</EhHeader>

    ' Если файл загружен, то спрашиваем о действии
    If Manager.FileLoaded Then

        CloseMainMenuItem_Click

        ' Если пользователь нажал "Отмена" в диалоговом окне,
        ' то файл остаётся открытым. В этом случае ничего не делаем
        If Manager.FileLoaded Then Exit Sub

    End If

    ' Теперь можно импортировать файл
    OpenFileDialog.DialogTitle = "Импорт файла..."
    OpenFileDialog.DefaultExt = ".json"
    OpenFileDialog.FileName = ""
    OpenFileDialog.Filter = "Конфигуратор УП 1.x (*.json)|*.json"
    OpenFileDialog.FilterIndex = 1
    OpenFileDialog.MaxFileSize = 32767
    OpenFileDialog.InitDir = CurrentDir
    OpenFileDialog.CancelError = True
    OpenFileDialog.Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly
    
    OpenFileDialog.ShowOpen

    FileName = OpenFileDialog.FileName
    
    Dim Succes As Boolean
    
    If FileName <> "" Then

        If Manager.FileLoaded Then
            CloseMainMenuItem_Click
        End If
        
        Succes = Manager.ImportFromJSON(FileName)
        
        If Succes = True Then
            SetCaption (Manager.FileName)
            SetModified True
        
            ViewMode = STEPS_VIEW
            RefreshComponents False
            'RefreshFrameRight
        Else
            
        End If
    End If

    '<EhFooter>
    Exit Sub

ImportMainMenuItem_Click_Err:
    If Err.Number = cdlCancel Then
    Else
        App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.ImportMainMenuItem_Click]: " _
           & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    End If
    
    Resume Next
    '</EhFooter>
End Sub

Private Sub InsertStepMenuItem_Click()
    '<EhHeader>
    On Error GoTo InsertStepMenuItem_Click_Err
    '</EhHeader>

    Manager.InsertStep
            
    ' Пересчитываем CRC поле записи программы
    Dim CRC8Value As Byte
    Dim Address As Long
    Dim Size As Long
    
    Address = Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
    Size = PROGRAM_SIZE_IN_BYTES - 1
    
    CRC8Value = Manager.CalculateCRC8(Address + 1, Size)
    Manager.SetByte Address, CRC8Value
            
    SetModified True
    RefreshDataComponents

    StepsView.SetFocus
    
    '<EhFooter>
    Exit Sub

InsertStepMenuItem_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.InsertStepMenuItem_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub Kachalka_Complete(ByVal Status As kach_tlb.BINDSTATUS, ByVal StatusText As String)
    '<EhHeader>
    On Error GoTo Kachalka_Complete_Err
    '</EhHeader>
    
    TextLog.Text = TextLog.Text & vbCrLf & Date & " " & Time & ": " & "Загрузка завершена"
    FormDownload.Caption = "Загрузка завершена"
    
    MenuItemDoUpdate.Enabled = True

    '<EhFooter>
    Exit Sub

Kachalka_Complete_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormMain.Kachalka_Complete]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub Kachalka_DataAvailable(ByVal EventType As kach_tlb.BSCF, ByVal Data As String, ByVal DataFormat As Long)

    ' Пропускаем
    
End Sub

Private Sub Kachalka_Progress(ByVal Progress As Long, ByVal ProgressMax As Long, ByVal Status As Long, ByVal StatusText As String, Cancel As Boolean)
    '<EhHeader>
    On Error GoTo Kachalka_Progress_Err
    '</EhHeader>

    Dim sProgress As String
    
    If ProgressMax Then
    
        sProgress = Format(Progress / ProgressMax, "00.00%")
        
    Else
    
        sProgress = "???"
        
    End If
     
    If FormDownload.Visible = True Then
    
        FormDownload.ProgressBar.Value = CInt((100 * Progress) / ProgressMax)
        FormDownload.Caption = "Загрузка: " & sProgress
        
    End If
    
    ' Обновляем интерфейс
    DoEvents
    
    Cancel = SetCancel

    '<EhFooter>
    Exit Sub

Kachalka_Progress_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormMain.Kachalka_Progress]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Function LookupStatus(ByVal ulStatusCode As kach_tlb.BINDSTATUS) As String
    
    If ulStatusCode <= 0 Then
    
        LookupStatus = Hex(ulStatusCode)
        
    Else
    
        LookupStatus = Choose(ulStatusCode, _
            "BINDSTATUS_FINDINGRESOURCE", "BINDSTATUS_CONNECTING", _
            "BINDSTATUS_REDIRECTING", "BINDSTATUS_BEGINDOWNLOADDATA", _
            "BINDSTATUS_DOWNLOADINGDATA", "BINDSTATUS_ENDDOWNLOADDATA", _
            "BINDSTATUS_BEGINDOWNLOADCOMPONENTS", "BINDSTATUS_INSTALLINGCOMPONENTS", _
            "BINDSTATUS_ENDDOWNLOADCOMPONENTS", "BINDSTATUS_USINGCACHEDCOPY", _
            "BINDSTATUS_SENDINGREQUEST", "BINDSTATUS_CLASSIDAVAILABLE", _
            "BINDSTATUS_MIMETYPEAVAILABLE", "BINDSTATUS_CACHEFILENAMEAVAILABLE", _
            "BINDSTATUS_BEGINSYNCOPERATION", "BINDSTATUS_ENDSYNCOPERATION", _
            "BINDSTATUS_BEGINUPLOADDATA", "BINDSTATUS_UPLOADINGDATA", _
            "BINDSTATUS_ENDUPLOADDATA", "BINDSTATUS_PROTOCOLCLASSID", _
            "BINDSTATUS_ENCODING", "BINDSTATUS_VERIFIEDMIMETYPEAVAILABLE", _
            "BINDSTATUS_CLASSINSTALLLOCATION", "BINDSTATUS_DECODING", _
            "BINDSTATUS_LOADINGMIMEHANDLER", "BINDSTATUS_CONTENTDISPOSITIONATTACH", _
            "BINDSTATUS_FILTERREPORTMIMETYPE", "BINDSTATUS_CLSIDCANINSTANTIATE", _
            "BINDSTATUS_IUNKNOWNAVAILABLE", "BINDSTATUS_DIRECTBIND", _
            "BINDSTATUS_RAWMIMETYPE", "BINDSTATUS_PROXYDETECTING", _
            "BINDSTATUS_ACCEPTRANGES", "BINDSTATUS_COOKIE_SENT", _
            "BINDSTATUS_COMPACT_POLICY_RECEIVED", "BINDSTATUS_COOKIE_SUPPRESSED", _
            "BINDSTATUS_COOKIE_STATE_UNKNOWN", "BINDSTATUS_COOKIE_STATE_ACCEPT", _
            "BINDSTATUS_COOKIE_STATE_REJECT", "BINDSTATUS_COOKIE_STATE_PROMPT", _
            "BINDSTATUS_COOKIE_STATE_LEASH", "BINDSTATUS_COOKIE_STATE_DOWNGRADE", _
            "BINDSTATUS_POLICY_HREF", "BINDSTATUS_P3P_HEADER", _
            "BINDSTATUS_SESSION_COOKIE_RECEIVED", "BINDSTATUS_PERSISTENT_COOKIE_RECEIVED", _
            "BINDSTATUS_SESSION_COOKIES_ALLOWED")
    
    End If

End Function

Public Sub ListPrograms_Click()
    '<EhHeader>
    On Error GoTo ListPrograms_Click_Err
    '</EhHeader>

    Dim CRC8Value As Byte
    Dim OldSelectedRow As Integer
    Dim CurrentSelectedRow As Integer
    Dim CellBackColor As Long
    
    OldSelectedRow = Manager.ProgramIndex + 1
    CurrentSelectedRow = ListPrograms.row
    
    ListPrograms.Redraw = False

    ListPrograms.row = OldSelectedRow
    ListPrograms.CellForeColor = &H80000008

    ' Вычисляем признак пустой программы
    CRC8Value = Manager.CalculateCRC8(Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES, PROGRAM_SIZE_IN_BYTES)

    If CRC8Value = CRC8_FOR_DEFAULT_PROGRAM Then

        ListPrograms.CellBackColor = &H8000000F

    Else

        Dim N As Integer
        Dim StepPointer As Long, Value As Long
        Dim s As String
        Dim RecordTitle As TYPE_WPC_TITLE

        StepPointer = Manager.DataPointer + Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES

        Value = 0

        For N = 1 To PROG_NAME_LENGTH - 1
            Value = Value + CLng(RecordTitle.ProgName(N))
        Next

        If Value = 0 Then
            ListPrograms.CellBackColor = &HC0FFFF
        Else
            ListPrograms.CellBackColor = &H80000005
        End If

    End If

    ListPrograms.row = CurrentSelectedRow
    ListPrograms.CellForeColor = &H8000000E
    ListPrograms.CellBackColor = &H8000000D
    
    ListPrograms.Redraw = True
    
    Manager.ProgramIndex = CurrentSelectedRow - 1
    
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
    
    '<EhFooter>
    Exit Sub

ListPrograms_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.ListPrograms_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub ListPrograms_DblClick()
    '<EhHeader>
    On Error GoTo ListPrograms_DblClick_Err
    '</EhHeader>
    
    ListPrograms_KeyDown VBRUN.KeyCodeConstants.vbKeyReturn, 0
    
    '<EhFooter>
    Exit Sub

ListPrograms_DblClick_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.ListPrograms_DblClick]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub ListPrograms_KeyDown(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo ListPrograms_KeyDown_Err
    '</EhHeader>

    If KeyCode = VBRUN.KeyCodeConstants.vbKeyUp Or _
       KeyCode = VBRUN.KeyCodeConstants.vbKeyDown Then
        
        ListPrograms_Click
    End If
    
    If KeyCode = VBRUN.KeyCodeConstants.vbKeyReturn Then
        TextName.Left = ListPrograms.Left + ListPrograms.CellLeft
        TextName.Top = ListPrograms.Top + ListPrograms.CellTop
        TextName.Width = ListPrograms.CellWidth
        TextName.Height = ListPrograms.CellHeight
        
        TextName.Text = ListPrograms.Text
        TextName.SelStart = 0
        TextName.SelLength = Len(TextName.Text)
        TextName.Visible = True
        TextName.SetFocus
    End If
    
    '<EhFooter>
    Exit Sub

ListPrograms_KeyDown_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.ListPrograms_KeyDown]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub ListPrograms_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    '<EhHeader>
    On Error GoTo ListPrograms_MouseDown_Err
    '</EhHeader>
    
    'проверка, нажата ли правая клавиша мыши
    If Button And vbRightButton Then PopupMenu PopupMenuPrograms
    
    '<EhFooter>
    Exit Sub

ListPrograms_MouseDown_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.ListPrograms_MouseDown]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub MenuItemDoUpdate_Click()
    '<EhHeader>
    On Error GoTo MenuItemDoUpdate_Click_Err
    '</EhHeader>

    ' Проверяем интернет-соединение
    Dim InternetConnected As Boolean
    Dim Result As Boolean
    Dim dwConnectionTypes As Long

    StatusBar.Panels(1).Text = "Проверяю доступ к сети"
    TextLog.Text = TextLog.Text & vbCrLf & Date & " " & Time & ": " & "Проверяю доступ к сети"
    
    dwConnectionTypes = INTERNET_CONNECTION_MODEM + INTERNET_CONNECTION_LAN + _
            INTERNET_CONNECTION_PROXY
    
    InternetConnected = InternetGetConnectedState(dwConnectionTypes, 0)

    ' TODO: Отображать процесс автообновления в статус строке
    ' Если имеется подключение к Интернет, то проверяем доступность сервера и
    ' файла автообновления
    If InternetConnected = True Then

        StatusBar.Panels(1).Text = "Проверяю наличие обновлений"
        TextLog.Text = TextLog.Text & vbCrLf & Date & " " & Time & ": " & "Проверяю наличие обновлений"
        
        ' Пытаемся обновиться
        Result = DoAutoUpdate(Settings.AutoUpdateLink)

        If Result = True Then

            StatusBar.Panels(1).Text = "Проверка проведена"
            TextLog.Text = TextLog.Text & vbCrLf & Date & " " & Time & ": " & "Проверка проведена"
            
            AutoUpdateState = AUS_UPDATED
            
            ' Останавливаем таймер и выходим
            TimerAutoUpdate.Interval = 0
            
            Exit Sub

        End If
        
        ' TODO: Подумать что тут написать пользователю
        StatusBar.Panels(1).Text = ""
        
    End If
    
    '<EhFooter>
    Exit Sub

MenuItemDoUpdate_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormMain.MenuItemDoUpdate_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub MenuItemShowHideLog_Click()
    '<EhHeader>
    On Error GoTo MenuItemShowHideLog_Click_Err
    '</EhHeader>

    FrameLog.Visible = Not FrameLog.Visible
    
    '<EhFooter>
    Exit Sub

MenuItemShowHideLog_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormMain.MenuItemShowHideLog_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub MRUItems_Click(Index As Integer)
    '<EhHeader>
    On Error GoTo MRUItems_Click_Err
    '</EhHeader>

    ' Открываем файл из списка
    Dim FName As String
        
    If MRUFileList.FileExists(Index) Then
        FName = MRUFileList.file(Index)
        
        If Manager.FileLoaded Then
            CloseMainMenuItem_Click
        End If
        
        FileName = MiscExtractPathName(FName, False)
        Manager.LoadFromFile (FName)
        SetCaption (Manager.FileName)
        
        ViewMode = STEPS_VIEW
        RefreshComponents False
        'RefreshFrameRight
        
        SetModified False
        CurrentDir = MiscExtractPathName(FName, True)
        
        MRUFileList.AddFile FName
        DisplayMRU
    End If
    
    '<EhFooter>
    Exit Sub

MRUItems_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.MRUItems_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub NewMainMenuItem_Click()
    '<EhHeader>
    On Error GoTo NewMainMenuItem_Click_Err
    '</EhHeader>

    ' Если файл загружен, то спрашиваем о действии
    If Manager.FileLoaded Then

        CloseMainMenuItem_Click

        ' Если пользователь нажал "Отмена" в диалоговом окне,
        ' то файл остаётся открытым. В этом случае ничего не делаем
        If Manager.FileLoaded Then Exit Sub

    End If

    ' Теперь можно создавать новый файл
    Manager.CreateNewFile (DEFAULT_FILE_NAME)
    
    ' Очистить все программы из образа
    PopupMenuListClearAll_Click
    
    '<EhFooter>
    Exit Sub

NewMainMenuItem_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.NewMainMenuItem_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub OptionsMainMenuItem_Click()
    '<EhHeader>
    On Error GoTo OptionsMainMenuItem_Click_Err
    '</EhHeader>
    
    FormOptions.Show vbModal, Me
    
    '<EhFooter>
    Exit Sub

OptionsMainMenuItem_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.OptionsMainMenuItem_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub PopupMenuListClear_Click()
    '<EhHeader>
    On Error GoTo PopupMenuListClear_Click_Err
    '</EhHeader>
    
    Dim StepPointer As Long
    
    ' Очищаем текущую программу
    Manager.ClearProgramN (ListPrograms.row)
    
    ' Устанавливаем заголовок по умолчанию
    If LimitsLoaded Then Manager.SetDefaultProgramHeader Manager.ProgramIndex + 1
            
    ' Пересчитываем CRC поле записи программы
    Dim CRC8Value As Byte
    Dim Address As Long
    Dim Size As Long
    
    Address = (ListPrograms.row - 1) * PROGRAM_SIZE_IN_BYTES
    Size = PROGRAM_SIZE_IN_BYTES - 1
    
    CRC8Value = Manager.CalculateCRC8(Address + 1, Size)
    Manager.SetByte Address, CRC8Value
            
    SetModified True
    
    RefreshDataComponents
    
    '<EhFooter>
    Exit Sub

PopupMenuListClear_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.PopupMenuListClear_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub PopupMenuListClearAll_Click()
    '<EhHeader>
    On Error GoTo PopupMenuListClearAll_Click_Err
    '</EhHeader>
    
    Manager.ClearAll
    
    If LimitsLoaded Then
        Dim CRC8Value As Byte
        Dim I As Integer
        Dim Address As Long
        Dim Size As Long
            
        For I = 1 To Manager.ProgramsCount
            ' Установка заголовка программы по умолчанию
            Manager.SetDefaultProgramHeader I
        
            ' Пересчитываем CRC поле записи программы
            Address = (I - 1) * PROGRAM_SIZE_IN_BYTES
            Size = PROGRAM_SIZE_IN_BYTES - 1
            
            CRC8Value = Manager.CalculateCRC8(Address + 1, Size)
            Manager.SetByte Address, CRC8Value
        Next
    End If
    
    SetModified True
    
    RefreshComponents False
    
    '<EhFooter>
    Exit Sub

PopupMenuListClearAll_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.PopupMenuListClearAll_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub PopupMenuPrograms_Click()
    '<EhHeader>
    On Error GoTo PopupMenuPrograms_Click_Err
    '</EhHeader>
    
    ' Настраиваем доступность пунктов меню "Программа":
    
    ' "Программа\Перейти"
    Select Case ViewMode
    
        Case STEPS_VIEW: GotoMenuItem.Enabled = False
        
        Case CODE_VIEW: GotoMenuItem.Enabled = True
        
    End Select
    
    '<EhFooter>
    Exit Sub

PopupMenuPrograms_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.PopupMenuPrograms_Click]: " _
        & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub PropertyTable_DblClick()
    '<EhHeader>
    On Error GoTo PropertyTable_DblClick_Err
    '</EhHeader>
    
    PropertyTable_KeyDown VBRUN.KeyCodeConstants.vbKeyReturn, 0
    
    '<EhFooter>
    Exit Sub

PropertyTable_DblClick_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.PropertyTable_DblClick]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub PropertyTable_KeyDown(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo PropertyTable_KeyDown_Err
    '</EhHeader>
    
    ' При нажатии Enter в ячейке даём возможность редактировать её содержимое
    If Not (KeyCode = VBRUN.KeyCodeConstants.vbKeyReturn) Then Exit Sub
    
    Dim col As Integer, row As Integer, FuncN As Integer
    
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
                
'<Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:19:28
'Причина: Модуль аналогичен по функционалу с ModuleWashOrRinsOrJolt>
'            Case WPC_OPERATION_PAUS ' пауза
'                ModulePause.EditPropertyForPause Me
'</Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:19:28>

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
    
    '<EhFooter>
    Exit Sub

PropertyTable_KeyDown_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.PropertyTable_KeyDown]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub StepMainMenuItem_Click()
    '<EhHeader>
    On Error GoTo StepMainMenuItem_Click_Err
    '</EhHeader>
    
    InsertStepMenuItem = ActiveControl Is StepsView
    DeleteStepMenuItem.Enabled = ActiveControl Is StepsView
    
    '<EhFooter>
    Exit Sub

StepMainMenuItem_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.StepMainMenuItem_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub StepsView_Click()
    '<EhHeader>
    On Error GoTo StepsView_Click_Err
    '</EhHeader>

    Dim x As Integer, Y As Integer
    Dim col As Integer, row As Integer
    
    StepsView.Redraw = False
    
    x = StepsView.col
    Y = StepsView.row
    
    For col = 1 To StepsView.Cols - 2
        StepsView.col = col
        StepsView.row = 0
        StepsView.CellFontBold = False
    Next
    
    row = StepsView.TopRow
    
    Do While StepsView.RowIsVisible(row)
        StepsView.col = 0
        StepsView.row = row
    
        StepsView.CellFontBold = False
        row = row + 1

        If row > StepsView.rows - 1 Then Exit Do
    Loop
       
    StepsView.row = 0
    StepsView.col = x
    StepsView.CellFontBold = True
    
    StepsView.row = Y
    StepsView.col = 0
    StepsView.CellFontBold = True
    
    StepsView.col = x
    StepsView.row = Y
    
    Manager.StepIndex = x - 1
    
    CodeView.TopRow = (PROGRAM_SIZE_IN_BYTES * Manager.ProgramIndex + _
       HEADER_SIZE_IN_BYTES + STEP_SIZE_IN_BYTES * Manager.StepIndex) / 16 + 1
    
    StepsView.Redraw = True
    StepsView.SetFocus
    
    ' Обновляем зависимые компоненты
    RefreshProperties
    RefreshFrameMain
    RefreshFrameRight
    RefreshCodeView
    
    '<EhFooter>
    Exit Sub

StepsView_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.StepsView_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub OpenMainMenuItem_Click()
    '<EhHeader>
    On Error GoTo OpenMainMenuItem_Click_Err
    '</EhHeader>

    ' Если файл загружен, то спрашиваем о действии
    If Manager.FileLoaded Then

        CloseMainMenuItem_Click

        ' Если пользователь нажал "Отмена" в диалоговом окне,
        ' то файл остаётся открытым. В этом случае ничего не делаем
        If Manager.FileLoaded Then Exit Sub

    End If
    
    ' Теперь можно открывать новый файл
    Dim FileName As String

    OpenFileDialog.DialogTitle = "Открыть файл..."
    OpenFileDialog.DefaultExt = ".bin"
    OpenFileDialog.FileName = ""
    OpenFileDialog.Filter = "Файлы проекта (*.bin)|*.bin"
    OpenFileDialog.FilterIndex = 1
    OpenFileDialog.MaxFileSize = 32767 ' Размер буфера под имя файла
    OpenFileDialog.InitDir = CurrentDir
    OpenFileDialog.CancelError = True
    OpenFileDialog.Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly
    
    OpenFileDialog.ShowOpen

    FileName = OpenFileDialog.FileName
        
    If FileName <> "" Then
              
        FileName = MiscExtractPathName(OpenFileDialog.FileName, False)
        Manager.LoadFromFile (OpenFileDialog.FileName)
        SetCaption (Manager.FileName)
        
        ViewMode = STEPS_VIEW
        RefreshComponents False
        'RefreshFrameRight
        
        SetModified False
        CurrentDir = MiscExtractPathName(OpenFileDialog.FileName, True)
        
        MRUFileList.AddFile OpenFileDialog.FileName
        DisplayMRU
    End If

    '<EhFooter>
    Exit Sub

OpenMainMenuItem_Click_Err:
    If Err.Number = cdlCancel Then
    Else
        App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.OpenMainMenuItem_Click]: " _
           & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    End If
    
    Resume Next
    '</EhFooter>
End Sub

Private Sub SetCaption(FileName As String)
    '<EhHeader>
    On Error GoTo SetCaption_Err
    '</EhHeader>

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
    
    '<EhFooter>
    Exit Sub

SetCaption_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.SetCaption]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub SaveAsMainMenuItem_Click()
    '<EhHeader>
    On Error GoTo SaveAsMainMenuItem_Click_Err
    '</EhHeader>

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

    '<EhFooter>
    Exit Sub

SaveAsMainMenuItem_Click_Err:
    If Err.Number = cdlCancel Then
    Else
        App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.SaveAsMainMenuItem_Click]: " _
           & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    End If
    
    Resume Next
    '</EhFooter>
End Sub

Private Sub SaveMainMenuItem_Click()
    '<EhHeader>
    On Error GoTo SaveMainMenuItem_Click_Err
    '</EhHeader>

    If Modified Then

        If DoesFileExist(Manager.FileName) Then
        
            Manager.SaveToFile (Manager.FileName)
            SetModified False
            RefreshDataComponents
            
        Else
        
            SaveAsMainMenuItem_Click
            
        End If
        
        MRUFileList.AddFile Manager.FileName
        DisplayMRU
        
    End If
    
    '<EhFooter>
    Exit Sub

SaveMainMenuItem_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.SaveMainMenuItem_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub SplitterLeft_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    '<EhHeader>
    On Error GoTo SplitterLeft_MouseDown_Err
    '</EhHeader>
    
    ' Показываем разделительную линию
    SplitterLeft.BackColor = &H80000010
    
    BegX = x
    BegY = Y
    
    SplitterLeftMoving = True
    
    '<EhFooter>
    Exit Sub

SplitterLeft_MouseDown_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.SplitterLeft_MouseDown]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub SplitterLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    '<EhHeader>
    On Error GoTo SplitterLeft_MouseMove_Err
    '</EhHeader>

    If SplitterLeftMoving = True Then
        SplitterLeft.Left = SplitterLeft.Left + x - BegX
        FrameLeft.Width = SplitterLeft.Left
        
        FrameMain.Left = SplitterLeft.Left + SplitterLeft.Width
        FrameMain.Width = SplitterRight.Left - FrameMain.Left
        
        RefreshFrameLeft
        RefreshFrameMain
        Refresh
    End If
    
    '<EhFooter>
    Exit Sub

SplitterLeft_MouseMove_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.SplitterLeft_MouseMove]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub SplitterLeft_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    '<EhHeader>
    On Error GoTo SplitterLeft_MouseUp_Err
    '</EhHeader>
    
    SplitterLeft.BackColor = &H8000000F
    SplitterLeftMoving = False
    
    '<EhFooter>
    Exit Sub

SplitterLeft_MouseUp_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.SplitterLeft_MouseUp]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub SplitterRight_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Показываем разделительную линию
    '<EhHeader>
    On Error GoTo SplitterRight_MouseDown_Err
    '</EhHeader>
    
    SplitterRight.BackColor = &H80000010
    
    BegX = x
    BegY = Y
    
    SplitterRightMoving = True
    
    '<EhFooter>
    Exit Sub

SplitterRight_MouseDown_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.SplitterRight_MouseDown]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub SplitterRight_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    '<EhHeader>
    On Error GoTo SplitterRight_MouseMove_Err
    '</EhHeader>

    If SplitterRightMoving = True Then
        SplitterRight.Left = SplitterRight.Left + x - BegX
        
        FrameRight.Left = SplitterRight.Left + SplitterRight.Width
        FrameRight.Width = Me.ScaleWidth - FrameRight.Left
        
        FrameMain.Width = SplitterRight.Left - FrameMain.Left
        
        RefreshFrameRight
        RefreshFrameMain
        Refresh
    End If
    
    '<EhFooter>
    Exit Sub

SplitterRight_MouseMove_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.SplitterRight_MouseMove]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub SplitterRight_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    '<EhHeader>
    On Error GoTo SplitterRight_MouseUp_Err
    '</EhHeader>
    
    SplitterRight.BackColor = &H8000000F
    SplitterRightMoving = False
    
    '<EhFooter>
    Exit Sub

SplitterRight_MouseUp_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.SplitterRight_MouseUp]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub StepsView_DblClick()

    Dim FuncN As Integer
    
    FuncN = Manager.GetFunctionType(Manager.ProgramIndex + 1, Manager.StepIndex + 1)
    
    ' Сохраняем изменённое значение

    If FuncN < 12 Then

        Select Case FuncN
            Case WPC_OPERATION_IDLE ' пропуск
                ModuleIdle.SetCheckBoxForIdle Me, StepsView.row
        
            Case WPC_OPERATION_FILL ' Налив
                ModuleFill.SetCheckBoxForFill Me, StepsView.row
            
            Case WPC_OPERATION_DTRG ' моющие
                ModuleDTRG.SetCheckBoxForDTRG Me, StepsView.row
            
            Case WPC_OPERATION_HEAT ' нагрев
                ModuleHeat.SetCheckBoxForHeat Me, StepsView.row
                
                ' стирка, полоскание, расстряска
            Case WPC_OPERATION_WASH, WPC_OPERATION_RINS, WPC_OPERATION_JOLT, WPC_OPERATION_PAUS
                ModuleWashOrRinsOrJolt.SetCheckBoxForWashOrRinsOrJolt Me, StepsView.row
                
            Case WPC_OPERATION_DRAIN ' слив
                ModuleDrain.SetCheckBoxForDrain Me, StepsView.row
                
            Case WPC_OPERATION_SPIN ' отжим
                ModuleSpin.SetCheckBoxForSpin Me, StepsView.row
            
            Case WPC_OPERATION_COOL ' охлаждение
                ModuleCool.SetCheckBoxForCool Me, StepsView.row
                
            Case WPC_OPERATION_TRIN ' тех.полоскание
                ModuleTrin.SetCheckBoxForTrin Me, StepsView.row
        
            Case Else

        End Select
            
        ' Пересчитываем CRC поле записи программы
        Dim CRC8Value As Byte
        Dim Address As Long
        Dim Size As Long
        
        Address = Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        Size = PROGRAM_SIZE_IN_BYTES - 1
        
        CRC8Value = Manager.CalculateCRC8(Address + 1, Size)
        Manager.SetByte Address, CRC8Value

        RefreshComponents False
    End If

End Sub

Private Sub StepsView_KeyDown(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo StepsView_KeyDown_Err
    '</EhHeader>

    If KeyCode = VBRUN.KeyCodeConstants.vbKeyInsert Then
        InsertStepMenuItem_Click
        Exit Sub
    End If
        
    If KeyCode = VBRUN.KeyCodeConstants.vbKeyDelete Then
        DeleteStepMenuItem_Click
        Exit Sub
    End If
        
    If KeyCode = VBRUN.KeyCodeConstants.vbKeyLeft Or _
       KeyCode = VBRUN.KeyCodeConstants.vbKeyRight Then
        
        StepsView_Click
    End If
    
    '<EhFooter>
    Exit Sub

StepsView_KeyDown_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.StepsView_KeyDown]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub StepsView_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    '<EhHeader>
    On Error GoTo StepsView_MouseDown_Err
    '</EhHeader>
    
    'проверка, нажата ли правая клавиша мыши

    If Button And vbRightButton Then PopupMenu StepMainMenuItem
    
    '<EhFooter>
    Exit Sub

StepsView_MouseDown_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.StepsView_MouseDown]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub TextByte_Change()
    '<EhHeader>
    On Error GoTo TextByte_Change_Err
    '</EhHeader>
    
    TextByte.Text = Mid$(TextByte.Text, 1, 2)
    
    '<EhFooter>
    Exit Sub

TextByte_Change_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.TextByte_Change]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub TextByte_KeyDown(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo TextByte_KeyDown_Err
    '</EhHeader>

    If KeyCode = VBRUN.KeyCodeConstants.vbKeyEscape Then
        TextByte.Visible = False
        CodeView.SetFocus
    End If
    
    If KeyCode = VBRUN.KeyCodeConstants.vbKeyReturn Then
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
        Dim CRC8Value As Byte
        Dim Address As Long
        Dim Size As Long
        
        Address = Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        Size = PROGRAM_SIZE_IN_BYTES - 1
        
        CRC8Value = Manager.CalculateCRC8(Address + 1, Size)
        Manager.SetByte Address, CRC8Value

        SetModified True
        
        TextByte.Visible = False
        TopRow = CodeView.TopRow
        RefreshComponents False
        CodeView.TopRow = TopRow
        
        CodeView.row = row
        CodeView.col = col
        CodeView.SetFocus
    End If
    
    '<EhFooter>
    Exit Sub

TextByte_KeyDown_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.TextByte_KeyDown]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub TextByte_KeyPress(KeyAscii As Integer)
    '<EhHeader>
    On Error GoTo TextByte_KeyPress_Err
    '</EhHeader>
    
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
    
    '<EhFooter>
    Exit Sub

TextByte_KeyPress_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.TextByte_KeyPress]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub TextByte_LostFocus()
    '<EhHeader>
    On Error GoTo TextByte_LostFocus_Err
    '</EhHeader>
    
    TextByte.Visible = False
    
    '<EhFooter>
    Exit Sub

TextByte_LostFocus_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.TextByte_LostFocus]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub TextCell_KeyDown(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo TextCell_KeyDown_Err
    '</EhHeader>

    If KeyCode = VBRUN.KeyCodeConstants.vbKeyEscape Then
        TextCell.Visible = False
        LabelDescription.Visible = False
        ShapeDescription.Visible = False
        RefreshFrameRight
        PropertyTable.SetFocus
    End If
    
    If KeyCode = VBRUN.KeyCodeConstants.vbKeyReturn Then
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
                    
'<Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:20:10
'Причина: Модуль аналогичен по функционалу с ModuleWashOrRinsOrJolt>
'                Case WPC_OPERATION_PAUS ' пауза
'                    ModulePause.SetComboPropertyForPause Me
'</Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:20:10>
    
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
            Dim CRC8Value As Byte
            Dim Address As Long
            Dim Size As Long
            
            Address = Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
            Size = PROGRAM_SIZE_IN_BYTES - 1
            
            CRC8Value = Manager.CalculateCRC8(Address + 1, Size)
            Manager.SetByte Address, CRC8Value
            
            TextCell.Visible = False
            LabelDescription.Visible = False
            ShapeDescription.Visible = False
            Dim row As Integer
            row = PropertyTable.row
            RefreshComponents False

            If row < PropertyTable.rows - 1 Then PropertyTable.row = row
            PropertyTable.SetFocus
        End If
    End If
    
    '<EhFooter>
    Exit Sub

TextCell_KeyDown_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.TextCell_KeyDown]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub TextCell_KeyPress(KeyAscii As Integer)
    '<EhHeader>
    On Error GoTo TextCell_KeyPress_Err
    '</EhHeader>

    If KeyAscii = VBRUN.KeyCodeConstants.vbKeyReturn Then KeyAscii = 0
    
    '<EhFooter>
    Exit Sub

TextCell_KeyPress_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.TextCell_KeyPress]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub TextCell_LostFocus()
    '<EhHeader>
    On Error GoTo TextCell_LostFocus_Err
    '</EhHeader>
    
    TextCell.Visible = False
    LabelDescription.Visible = False
    ShapeDescription.Visible = False
    RefreshFrameRight
    
    '<EhFooter>
    Exit Sub

TextCell_LostFocus_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.TextCell_LostFocus]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub TextName_KeyDown(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo TextName_KeyDown_Err
    '</EhHeader>
    
    Dim I As Integer
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    
    If KeyCode = VBRUN.KeyCodeConstants.vbKeyEscape Then
        TextName.Visible = False
        ListPrograms.SetFocus
    End If
    
    If KeyCode = VBRUN.KeyCodeConstants.vbKeyReturn Then
        StepPointer = Manager.DataPointer + Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        
        For I = 1 To PROG_NAME_LENGTH - 1

            If I <= Len(TextName.Text) Then
                RecordTitle.ProgName(I) = Asc(Mid$(TextName.Text, I, 1))
            Else
                RecordTitle.ProgName(I) = 0
            End If
        Next
        
        RecordTitle.ProgName(PROG_NAME_LENGTH) = 0
        ' Сохраняем изменения
        CopyMemory ByVal StepPointer, RecordTitle, HEADER_SIZE_IN_BYTES
        SetModified True
        
        ' Пересчитываем CRC поле записи программы
        Dim CRC8Value As Byte
        Dim Address As Long
        Dim Size As Long
        
        Address = Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        Size = PROGRAM_SIZE_IN_BYTES - 1
        
        CRC8Value = Manager.CalculateCRC8(Address + 1, Size)
        Manager.SetByte Address, CRC8Value
        
        TextName.Visible = False
        Dim row As Integer
        row = ListPrograms.row
        RefreshComponents False
        
        If row < ListPrograms.rows - 1 Then ListPrograms.row = row
        ListPrograms.SetFocus
    End If
    
    '<EhFooter>
    Exit Sub

TextName_KeyDown_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.TextName_KeyDown]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub TextName_KeyPress(KeyAscii As Integer)
    '<EhHeader>
    On Error GoTo TextName_KeyPress_Err
    '</EhHeader>

    If KeyAscii = VBRUN.KeyCodeConstants.vbKeyReturn Then KeyAscii = 0
    
    '<EhFooter>
    Exit Sub

TextName_KeyPress_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.TextName_KeyPress]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub TextName_LostFocus()
    '<EhHeader>
    On Error GoTo TextName_LostFocus_Err
    '</EhHeader>
    
    TextName.Visible = False
    
    '<EhFooter>
    Exit Sub

TextName_LostFocus_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.TextName_LostFocus]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub Timer1_Timer()
    '<EhHeader>
    On Error GoTo Timer1_Timer_Err
    '</EhHeader>
    
    ' Особый случай: мы ловим клавиши при помощи хуков
    ' Когда файлы не загружены, нет активного элемента управления
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
    
    '<EhFooter>
    Exit Sub

Timer1_Timer_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.TextName_LostFocus]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Function DoAutoUpdate(UpdateFileLink As String) As Boolean
    '<EhHeader>
    On Error GoTo DoAutoUpdate_Err
    '</EhHeader>

    Dim Result As Boolean
       
    Result = False
    
    ' Создаём временный файл
    Dim szBuffer As String, szTempFileName As String
    Dim MAX_PATH As Long
    Dim Length As Integer
    
    MAX_PATH = 255
    szBuffer = Space(255)
    
    ' Получаем путь к временной папке
    Length = GetTempPath(MAX_PATH, szBuffer)

    ' Формируем путь к временному файлу
    szTempFileName = Space(255)
    GetTempFileName szBuffer, "cop", 0, szTempFileName
       
    TextLog.Text = TextLog.Text & vbCrLf & Date & " " & Time & ": " & "Загрузка файла обновления"
    
    ' Пытаемся скачать файл описания с сервера
    Kachalka.DownloadToFile UpdateFileLink, szTempFileName
    
    ' Обрабатываем скачанный файл
    Dim I As Integer
    Dim CurrMajor As Integer, CurrMinor As Integer, CurrRevision As Integer, CurrBuild As Integer
    Dim Major As Integer, Minor As Integer, Revision As Integer, Build As Integer
    Dim sInputJson As String
    Dim Version As String
    Dim strFile As String
    Dim DownloadLink As String
    Dim VerArr() As String
    
    Dim udtFileInfo As FILEINFO
    Dim p As Object
    
    strFile = String(255, 0)
    GetModuleFileName 0, strFile, 255
    
    ' На время отладки задаём отладочный входной файл
    If DesignMode Then szTempFileName = "D:\Projects\vbasic\Projects\Configurator\update"
    
    ' Считываем файл и декодируем его
    sInputJson = FromUTF8(LoadFromJSONFile(szTempFileName))

    ' Производим разбор данных из файла
    Set p = JSON.parse(sInputJson)
   
    TextLog.Text = TextLog.Text & vbCrLf & Date & " " & Time & ": " & "Поиск информации об обновлении"
    ' Ищем запись, имеющую необходимый GUID в поле ProgID
    For I = 1 To p.Count
    
        If (ProgramGUID = p.Item(I).Item("ProgID")) Then
        
            TextLog.Text = TextLog.Text & vbCrLf & Date & " " & Time & ": " & "Сверка версий"
            
            ' Считываем информацию о версии
            Major = p.Item(I).Item("Major")
            Minor = p.Item(I).Item("Minor")
            Revision = p.Item(I).Item("Revision")
            Build = p.Item(I).Item("Build")
            DownloadLink = p.Item(I).Item("DownloadLink")
                       
            ' Узнаём свою текущую версию
            If GetFileVersionInformation(strFile, udtFileInfo) = eNoVersion Then
                
                CurrMajor = Major
                CurrMinor = Minor
                CurrRevision = Revision
                CurrBuild = Build
                           
            Else
                
                VerArr = Split(udtFileInfo.FileVersion, ".")
                
                ' Косвенно проверяем формат своей версии
                If (UBound(VerArr) = 3) Then
                
                    CurrMajor = CInt(VerArr(0))
                    CurrMinor = CInt(VerArr(1))
                    CurrRevision = CInt(VerArr(2))
                    CurrBuild = CInt(VerArr(3))
                    
                Else
                                  
                    CurrMajor = Major
                    CurrMinor = Minor
                    CurrRevision = Revision
                    CurrBuild = Build
                
                End If
                
            End If
            
            Dim NeedUpdate As Boolean
            
            NeedUpdate = False
            
            ' Если текущая версия устарела, то выводим сообщение об этом
            If CurrMajor >= Major Then
                
                If CurrMinor >= Minor Then
                    
                    If CurrRevision >= Revision Then
                        
                        If CurrBuild >= Build Then
                            
                        Else
                            
                            NeedUpdate = True
                            
                        End If
                    
                    Else
                        
                        NeedUpdate = True
                        
                    End If
                    
                Else
                    
                    NeedUpdate = True
                    
                End If
            
            Else
            
                NeedUpdate = True
                
            End If
            
            ' Спрашиваем и качаем дистрибутив
            If NeedUpdate = True Then
            
                Dim vbRes As Integer
                
                vbRes = MsgBox("Доступно обновление:" & _
                    vbCrLf & vbCrLf & _
                    "Новая версия: " & CStr(Major) & "." & CStr(Minor) & "." & CStr(Revision) & "." & CStr(Build) & vbCrLf & _
                    "Текущая версия: " & CStr(CurrMajor) & "." & CStr(CurrMinor) & "." & CStr(CurrRevision) & "." & CStr(CurrBuild) & _
                    vbCrLf & vbCrLf & "Загрузить обновление?", _
                    vbYesNo + vbQuestion, APP_NAME)
                
                Select Case vbRes
                
                    Case vbYes
                    
                        Dim FileName As String
                    
                        SaveFileDialog.FileName = MiscExtractPathName(DownloadLink, False, "/")
                        SaveFileDialog.DialogTitle = "Сохранить файл..."
                        SaveFileDialog.DefaultExt = ""
                        SaveFileDialog.Filter = "Все файлы (*.*)|*.*"
                        SaveFileDialog.FilterIndex = 1
                        SaveFileDialog.MaxFileSize = 32767
                        SaveFileDialog.InitDir = CurrentDir
                        SaveFileDialog.CancelError = True
                        
                        SaveFileDialog.ShowSave
                    
                        FileName = SaveFileDialog.FileName
                        
                        If FileName <> "" Then
                                           
                            TextLog.Text = TextLog.Text & vbCrLf & Date & " " & Time & ": " & _
                                "Загрузка файла:" & vbCrLf & FileName
                            
                            FormDownload.LabelFrom = "Откуда: " & DownloadLink
                            FormDownload.LabelTo = "Куда: " & FileName
                            
                            ' Показываем форму загрузки
                            FormDownload.Show
                            
                            ' Пытаемся скачать файл
                            Kachalka.DownloadToFile DownloadLink, FileName
                            
                            ' Устанавливаем признак успешной загрузки
                            AutoUpdateState = AUS_UPDATED
                            
                            ' Запоминаем дату
                            Settings.AutoUpdateLastDate = CStr(Date)
                    
                        End If
                        
                    Case vbNo
                
                End Select
                
            End If
            
            Result = True
        
            ' Выходим из цикла
            Exit For
            
        End If
        
    Next
    
    ' Удаляем временный файл
    If DoesFileExist(szTempFileName) Then DeleteFile szTempFileName
       
    Set p = Nothing
       
    DoAutoUpdate = Result
    
    '<EhFooter>
    Exit Function

DoAutoUpdate_Err:
    If Err.Number = cdlCancel Then
        
        ' Удаляем временный файл
        If DoesFileExist(szTempFileName) Then DeleteFile szTempFileName
        
        Set p = Nothing
        
        DoAutoUpdate = False
    
    Else
    
        App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
                " [INFO] [cop.FormMain.DoAutoUpdate]: " & GetErrorMessageById( _
                Err.Number, Err.Description), _
                VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
                
        Resume Next
        
    End If

    '</EhFooter>
End Function

Private Sub TimerAutoUpdate_Timer()
    '<EhHeader>
    On Error GoTo TimerAutoUpdate_Timer_Err
    '</EhHeader>

    ' Если пользователь отменил автообновление
    If Settings.AutoUpdateEnabled = False Then

        ' Останавливаем таймер
        TimerAutoUpdate.Interval = 0
        Exit Sub

    End If

    ' Если обновились в текущем сеансе, то выходим
    If AutoUpdateState = AUS_UPDATED Then

        ' Останавливаем таймер
        TimerAutoUpdate.Interval = 0
        Exit Sub

    End If

    Dim Days As Long
    
    ' В режиме отладки проверка обновления будет происходить сразу
    If DesignMode = True Then
        
        Days = DateDiff("d", Now, CDate(Settings.AutoUpdateLastDate))
        
        Debug.Print "Разница в днях: " & Days
        
    Else
    
        Days = Abs(DateDiff("d", CDate(Settings.AutoUpdateLastDate), Now))
               
        Select Case Settings.AutoUpdatePeriod
        
            Case AUP_EVERY_DAY:
                
                If Days < 1 Then
                
                    ' Останавливаем таймер
                    TimerAutoUpdate.Interval = 0
                    Exit Sub
                    
                End If
                
            Case AUP_ONES_PER_WEEK:
            
                If Days < 7 Then
                
                    ' Останавливаем таймер
                    TimerAutoUpdate.Interval = 0
                    Exit Sub
                    
                End If
                
            Case AUP_ONES_PER_MONTH:
            
                If Days < 30 Then
                
                    ' Останавливаем таймер
                    TimerAutoUpdate.Interval = 0
                    Exit Sub
                    
                End If
                
            Case Else:
            
                    TimerAutoUpdate.Interval = 0
                    Exit Sub
                    
        End Select
        
    End If

    ' Выполняем обновление
    MenuItemDoUpdate_Click

    '<EhFooter>
    Exit Sub

TimerAutoUpdate_Timer_Err:
    
    TextLog.Text = TextLog.Text & vbCrLf & Date & " " & Time & ": " & "Ошибка (см. лог)"
    
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormMain.TimerAutoUpdate_Timer]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    '<EhHeader>
    On Error GoTo Toolbar1_ButtonClick_Err
    '</EhHeader>

    If (Button.Index = 1) Then NewMainMenuItem_Click

    If (Button.Index = 2) Then OpenMainMenuItem_Click

    If (Button.Index = 3) Then SaveMainMenuItem_Click

    If (Button.Index = 5) Then OptionsMainMenuItem_Click
    
    '<EhFooter>
    Exit Sub

Toolbar1_ButtonClick_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.Toolbar1_ButtonClick]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Public Sub RefreshDataComponents()
    '<EhHeader>
    On Error GoTo RefreshDataComponents_Err
    '</EhHeader>
    
    RefreshForm
    RefreshMainMenu
    RefreshList
    RefreshCodeView
    RefreshStepsView
    RefreshProperties
    RefreshFrameRight
    RefreshStatusBar
    
    '<EhFooter>
    Exit Sub

RefreshDataComponents_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.RefreshDataComponents]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub RefreshListSelection()
    '<EhHeader>
    On Error GoTo RefreshListSelection_Err
    '</EhHeader>

    Dim CRC8Value As Byte
    Dim Cnt As Integer
  
    ' Проверяем CRC для каждой из управляющих программ

    For Cnt = 0 To Manager.ProgramsCount - 1
        
        ListPrograms.row = Cnt + 1
        
        ' Вычисляем признак пустой программы
        CRC8Value = Manager.CalculateCRC8(Cnt * PROGRAM_SIZE_IN_BYTES, PROGRAM_SIZE_IN_BYTES)
        
        If CRC8Value = CRC8_FOR_DEFAULT_PROGRAM Then
            ListPrograms.CellBackColor = &H8000000F
        End If
        
        CRC8Value = Manager.CalculateCRC8(Cnt * PROGRAM_SIZE_IN_BYTES + 1, _
           PROGRAM_SIZE_IN_BYTES - 1)

        If Not CRC8Value = Manager.GetByte(Cnt * PROGRAM_SIZE_IN_BYTES) Then
            ListPrograms.CellBackColor = &H8080FF
        End If
        
        If ListPrograms.row = Manager.ProgramIndex + 1 Then
            ListPrograms.CellForeColor = &H80000005
            ListPrograms.CellBackColor = &H8000000D
        End If
    Next

    '<EhFooter>
    Exit Sub

RefreshListSelection_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.RefreshListSelection]: " _
        & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub RefreshList()
    '<EhHeader>
    On Error GoTo RefreshList_Err
    '</EhHeader>

    If Not Manager.FileLoaded Then
        FrameLeft.Caption = "Программы"
        ListPrograms.Clear
        ListPrograms.Font.Bold = Settings.StepsViewFontBold
        ListPrograms.Font.Italic = Settings.StepsViewFontItalic
        ListPrograms.Font.Name = Settings.StepsViewFontName
        ListPrograms.Font.Size = Settings.StepsViewFontSize
        ListPrograms.FormatString = "<Список"
        ListPrograms.rows = 1
        FrameLeft.Enabled = False
        Exit Sub
    End If
    
    FrameLeft.Enabled = True
    FrameLeft.Caption = "Программы [" & Manager.ProgramsCount & "]"
    
    ListPrograms.Redraw = False
    ListPrograms.Clear
    ListPrograms.rows = 1
    
    ListPrograms.Font.Bold = Settings.StepsViewFontBold
    ListPrograms.Font.Italic = Settings.StepsViewFontItalic
    ListPrograms.Font.Name = Settings.StepsViewFontName
    ListPrograms.Font.Size = Settings.StepsViewFontSize
    
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
            Next
            
            If Value = 0 Then
                ListPrograms.AddItem "Программа" & Cnt
                ListPrograms.row = Cnt
                ListPrograms.CellBackColor = &HC0FFFF
            Else
                s = ""

                For N = 1 To PROG_NAME_LENGTH - 1
                    s = s & Chr$(RecordTitle.ProgName(N))
                Next
                ListPrograms.AddItem s
            End If
        Next
    
        ' Обновляем задний фон ячеек списка в зависимости от состояния
        RefreshListSelection
        
        ListPrograms.row = Manager.ProgramIndex + 1
    End If

    ListPrograms.ColWidth(0) = ListPrograms.Width
    ListPrograms.Redraw = True
    
    '<EhFooter>
    Exit Sub

RefreshList_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.RefreshList]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub RefreshCodeView()
    '<EhHeader>
    On Error GoTo RefreshCodeView_Err
    '</EhHeader>
    
    Dim col As Integer, row As Integer
    Dim s As String
    
    ' Если файл не загружен, то выводить нечего,
    ' поэтому отображаем вид без данных

    If (Not Manager.FileLoaded) Then
        FrameMain.Caption = "Код"
        CodeView.Visible = False
        CodeView.Clear

        CodeView.Font.Bold = Settings.StepsViewFontBold
        CodeView.Font.Italic = Settings.StepsViewFontItalic
        CodeView.Font.Name = Settings.StepsViewFontName
        CodeView.Font.Size = Settings.StepsViewFontSize
        
        CodeView.rows = 2
        ' Инициализируем окно кода
        s = "<   |"

        For col = 1 To CodeView.Cols - 2
            CodeView.ColWidth(col) = Settings.StepColWidth

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
        Next
        
        CodeView.FormatString = s
        CodeView.Visible = True
        
        FrameMain.Enabled = False
        Exit Sub
    End If
       
    CodeView.Visible = False
    CodeView.Clear
    
    CodeView.Font.Bold = Settings.StepsViewFontBold
    CodeView.Font.Italic = Settings.StepsViewFontItalic
    CodeView.Font.Name = Settings.StepsViewFontName
    CodeView.Font.Size = Settings.StepsViewFontSize
        
    ' Формируем заголовки столбцов
    s = "<   |"

    For col = 1 To CodeView.Cols - 2
        CodeView.ColWidth(col) = Settings.StepColWidth

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
    Next
    
    CodeView.FormatString = s
    
    ' Формируем заголовки строк
    Dim HexValue As Long
   
    CodeView.ColWidth(0) = 2 * Settings.StepColWidth
    CodeView.rows = Manager.ImageSize / 16
    
    For row = 1 To CodeView.rows - 1
        CodeView.col = 0
        CodeView.row = row
        
        HexValue = (row - 1) * 16
        
        If HexValue < &H10 Then
            CodeView.Text = "0000"
        Else

            If HexValue < &H100 Then
                CodeView.Text = "00" & Hex$(HexValue)
            Else

                If HexValue < &H1000 Then
                    CodeView.Text = "0" & Hex$(HexValue)
                Else

                    If HexValue < &H10000 Then
                        CodeView.Text = "" & Hex$(HexValue)
                    End If
                End If
            End If
        End If
        
        CodeView.CellAlignment = flexAlignRightCenter
    Next
    
    ' Выводим данные
    row = CodeView.TopRow
    
    Do While CodeView.RowIsVisible(row)

        CodeView.RowHeight(row) = Settings.RowHeight
        
        For col = 1 To CodeView.Cols - 2
            CodeView.col = col
            CodeView.row = row
            CodeView.CellAlignment = flexAlignCenterCenter
            
            If ((16 * (row - 1) + col - 1) < Manager.ImageSize) Then
                HexValue = Manager.GetByte(16 * (row - 1) + col - 1)
            
                If (HexValue < &H10) Then
                    CodeView.Text = "0" & Hex$(HexValue)
                Else
                    CodeView.Text = "" & Hex$(HexValue)
                End If
            Else
                CodeView.Text = ""
            End If
        Next
        
        row = row + 1

        If (row > CodeView.rows - 1) Then Exit Do
    Loop
    
    ' После сделанных изменений можно визуализировать компонент
    CodeView.Visible = True
    
    '<EhFooter>
    Exit Sub

RefreshCodeView_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.RefreshCodeView]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Function ValveEnabled(col As Integer, row As Integer) As Boolean
    '<EhHeader>
    On Error GoTo ValveEnabled_Err
    '</EhHeader>
    
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
                
                ' стирка, полоскание, расстряска, пауза
            Case WPC_OPERATION_WASH, WPC_OPERATION_RINS, WPC_OPERATION_JOLT, WPC_OPERATION_PAUS
                ValveEnabled = ModuleWashOrRinsOrJolt.ValveEnabled(Me, col - 1, row)
                Exit Function
                
'<Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:20:26
'Причина: Модуль аналогичен по функционалу с ModuleWashOrRinsOrJolt>
'            Case WPC_OPERATION_PAUS ' пауза
'               ValveEnabled = ModulePause.SetComboPropertyForPause(Me, col - 1, row)
'               Exit Function
'</Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:20:26>

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
    
    '<EhFooter>
    Exit Function

ValveEnabled_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.ValveEnabled]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Function

Private Sub RefreshStepsView()
    '<EhHeader>
    On Error GoTo RefreshStepsView_Err
    '</EhHeader>
    
    Dim col As Integer, row As Integer, x As Integer, Y As Integer, FuncN As Integer
    Dim s As String
    
    ' Выходим из процедуры, если программы не загружены или отсутствуют

    If Not Manager.FileLoaded Then
        FrameMain.Caption = "Шаги"
        StepsView.Redraw = False
        StepsView.Clear
        
        StepsView.Font.Bold = Settings.StepsViewFontBold
        StepsView.Font.Italic = Settings.StepsViewFontItalic
        StepsView.Font.Name = Settings.StepsViewFontName
        StepsView.Font.Size = Settings.StepsViewFontSize
    
        StepsView.Cols = MAX_NUMBER_OF_STEPS + 1
        
        s = "<   |"

        For col = 1 To StepsView.Cols - 1
            StepsView.ColWidth(col) = Settings.StepColWidth
            
            If col < StepsView.Cols - 1 Then

                If col < 10 Then
                    s = s & "0" & col & "|"
                Else
                    s = s & col & "|"
                End If
            Else
                s = s & col
            End If
            StepsView.col = col
            StepsView.row = 0
            StepsView.CellAlignment = flexAlignCenterCenter
        Next
        
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

        For row = 1 To StepsView.rows - 1
            StepsView.RowHeight(row) = Settings.RowHeight
            
            For col = 1 To StepsView.Cols - 1
                StepsView.ColWidth(col) = Settings.StepColWidth
                StepsView.col = col
                StepsView.row = row
                StepsView.CellBackColor = &H8000000F
            Next
        Next
        
        StepsView.col = 1
        StepsView.row = 1
        StepsView.Redraw = True
        FrameMain.Enabled = False
        Exit Sub
    End If
        
    StepsView.Redraw = False
    
    x = StepsView.col
    Y = StepsView.row
    
    StepsView.Clear
    
    StepsView.Font.Bold = Settings.StepsViewFontBold
    StepsView.Font.Italic = Settings.StepsViewFontItalic
    StepsView.Font.Name = Settings.StepsViewFontName
    StepsView.Font.Size = Settings.StepsViewFontSize
        
    StepsView.Cols = MAX_NUMBER_OF_STEPS + 1
    
    s = "<   |"

    For col = 1 To StepsView.Cols - 1
        StepsView.ColWidth(col) = Settings.StepColWidth
        
        If col < StepsView.Cols - 1 Then

            If col < 10 Then
                s = s & "0" & col & "|"
            Else
                s = s & col & "|"
            End If
        Else
            s = s & col
        End If
        StepsView.col = col
        StepsView.row = 0
        StepsView.CellAlignment = flexAlignCenterCenter
    Next
    
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

    For row = 1 To StepsView.rows - 1
        StepsView.RowHeight(row) = Settings.RowHeight
        
        For col = 1 To StepsView.Cols - 1
            StepsView.col = col
            StepsView.row = row
            StepsView.CellBackColor = &H8000000F
        Next
    Next
    
    'col = StepsView.LeftCol
    
    StepsView.col = 1
    StepsView.row = 1
    
    ImageChecked.Width = StepsView.CellWidth
    ImageChecked.Height = StepsView.CellHeight
    
    ImageUnchecked.Width = StepsView.CellWidth
    ImageUnchecked.Height = StepsView.CellHeight
    
    ImageGrayed.Width = StepsView.CellWidth
    ImageGrayed.Height = StepsView.CellHeight
    
    For col = 1 To MAX_NUMBER_OF_STEPS
        FuncN = Manager.GetFunctionType(Manager.ProgramIndex + 1, col)
        
        If FuncN > 0 And FuncN < 11 Then

            For row = 1 To StepsView.rows - 1
                StepsView.col = col
                StepsView.row = row
                StepsView.CellAlignment = flexAlignCenterCenter
                
                If StepsViewMode = TEXT_VIEW Then StepsView.Text = Mid$(FunctionsStrings(FuncN), row, 1)
                
                If GetLoadingsFromFuncN(FuncN) And (2 ^ (row - 1)) Then

                    Select Case StepsViewMode
                        Case TEXT_VIEW:

                            If ValveEnabled(col, row) Then
                                StepsView.CellBackColor = &HC000&
                            Else
                                StepsView.CellBackColor = &H8000000F
                            End If
                            StepsView.CellPictureAlignment = flexAlignCenterCenter
                        
                        Case CHECKS_VIEW:
                            StepsView.CellBackColor = &HFFFFFF

                            If ValveEnabled(col, row) Then
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
            Next
        Else

            For row = 1 To StepsView.rows - 1
                StepsView.col = col
                StepsView.row = row
                StepsView.Text = ""
                StepsView.CellBackColor = &H8000000F
            Next
        End If
        
    Next
    
    StepsView.col = x
    StepsView.row = Y
    
    ' После сделанных изменений можно визуализировать компонент
    StepsView.Redraw = True
    
    '<EhFooter>
    Exit Sub

RefreshStepsView_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.RefreshStepsView]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub RefreshProperties()
    '<EhHeader>
    On Error GoTo RefreshProperties_Err
    '</EhHeader>
    
    Dim ParamStr As String

    ' Выходим из процедуры, если программы не загружены или отсутствуют

    If Not Manager.FileLoaded Then
        FrameRight.Caption = "Свойства"
        PropertyTable.Redraw = False
        
        PropertyTable.Font.Bold = Settings.StepsViewFontBold
        PropertyTable.Font.Italic = Settings.StepsViewFontItalic
        PropertyTable.Font.Name = Settings.StepsViewFontName
        PropertyTable.Font.Size = Settings.StepsViewFontSize
        
        PropertyTable.rows = 1
        PropertyTable.Clear
        ParamStr = "<Параметр|Значение"
        PropertyTable.FormatString = ParamStr
        PropertyTable.Redraw = True
        FrameRight.Enabled = False
        Exit Sub
    End If
    
    FrameRight.Enabled = True

    FrameRight.Caption = "Свойства - [" & ListPrograms.Text & ".Шаг" & Manager.StepIndex + 1 & "]"
    
    ' Узнаём номер функции текущего шага
    Dim FuncN As Integer
    
    FuncN = Manager.GetFunctionType(Manager.ProgramIndex + 1, Manager.StepIndex + 1)
    
    PropertyTable.Redraw = False
    
    PropertyTable.Font.Bold = Settings.StepsViewFontBold
    PropertyTable.Font.Italic = Settings.StepsViewFontItalic
    PropertyTable.Font.Name = Settings.StepsViewFontName
    PropertyTable.Font.Size = Settings.StepsViewFontSize
        
    PropertyTable.rows = 2
    PropertyTable.Clear
    ParamStr = "<Параметр|Значение"
    PropertyTable.FormatString = ParamStr
        
    If FuncN < 12 Then

        Select Case FuncN
            Case WPC_OPERATION_IDLE ' пропуск
                ModuleIdle.ShowPropertyTableForIdle Me
        
            Case WPC_OPERATION_FILL ' Налив
                ModuleFill.ShowPropertyTableForFill Me
            
            Case WPC_OPERATION_DTRG ' моющие
                ModuleDTRG.ShowPropertyTableForDTRG Me
            
            Case WPC_OPERATION_HEAT ' нагрев
                ModuleHeat.ShowPropertyTableForHeat Me
                
                ' стирка, полоскание, расстряска, пауза
            Case WPC_OPERATION_WASH, WPC_OPERATION_RINS, WPC_OPERATION_JOLT, WPC_OPERATION_PAUS
                ModuleWashOrRinsOrJolt.ShowPropertyTableForWashOrRinsOrJolt Me
                
'<Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:19:45
'Причина: Модуль аналогичен по функционалу с ModuleWashOrRinsOrJolt>
'            Case WPC_OPERATION_PAUS ' пауза
'                ModulePause.ShowPropertyTableForPause Me
'</Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:19:45>

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
    
    PropertyTable.Redraw = True
    
    '<EhFooter>
    Exit Sub

RefreshProperties_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.RefreshProperties]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Public Sub LoadLimits(FileName As String)
    '<EhHeader>
    On Error GoTo LoadLimits_Err
    '</EhHeader>
    
    LimitsLoaded = DoesFileExist(FileName)
    
    If Not LimitsLoaded Then Exit Sub
    
    Dim LimitsFile As New CIniFiles
    
    LimitsFile.Create FileName
    
    ' Настройки заголовка
    EndSound.DefaultValue = LimitsFile.ReadBoolean(TITLE_SECTION_NAME, "EndSound.Default", ENDSOUND_DEFAULT)
    DoorUnlock.DefaultValue = LimitsFile.ReadBoolean(TITLE_SECTION_NAME, "DoorUnlock.Default", DOORUNLOCK_DEFAULT)
    
    Set LimitsFile = Nothing
    
    '<EhFooter>
    Exit Sub

LoadLimits_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormMain.LoadLimits]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub ViewMainMenuItem_Click()
    '<EhHeader>
    On Error GoTo ViewMainMenuItem_Click_Err
    '</EhHeader>

    MenuItemShowHideLog.Checked = FrameLog.Visible

    '<EhFooter>
    Exit Sub

ViewMainMenuItem_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormMain.ViewMainMenuItem_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub
