VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FormOptions 
   Caption         =   "Настройки"
   ClientHeight    =   7248
   ClientLeft      =   2580
   ClientTop       =   1512
   ClientWidth     =   6984
   Icon            =   "FormOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7248
   ScaleWidth      =   6984
   Begin VB.CommandButton CommandCancel 
      Caption         =   "&Отмена"
      Height          =   360
      Left            =   5760
      TabIndex        =   22
      Top             =   6783
      Width           =   990
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6612
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6972
      _ExtentX        =   12298
      _ExtentY        =   11663
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   8
      TabHeight       =   420
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Приложение "
      TabPicture(0)   =   "FormOptions.frx":23D2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ShapeSSTabFrame"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameSettings"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameSplitterUpDown"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameDescription"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Образец "
      TabPicture(1)   =   "FormOptions.frx":23EE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameExample"
      Tab(1).ControlCount=   1
      Begin VB.Frame FrameExample 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6012
         Left            =   -74880
         TabIndex        =   6
         Top             =   360
         Width           =   6732
         Begin VB.PictureBox PictureHSelRight 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   252
            Left            =   5400
            ScaleHeight     =   252
            ScaleWidth      =   24
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   720
            Visible         =   0   'False
            Width           =   24
         End
         Begin VB.PictureBox PictureHSelBottom 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   20
            Left            =   120
            ScaleHeight     =   24
            ScaleWidth      =   5316
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   960
            Visible         =   0   'False
            Width           =   5316
         End
         Begin VB.PictureBox PictureHSelLeft 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   252
            Left            =   120
            ScaleHeight     =   252
            ScaleWidth      =   24
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   720
            Visible         =   0   'False
            Width           =   24
         End
         Begin VB.PictureBox PictureHSelTop 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   20
            Left            =   120
            ScaleHeight     =   24
            ScaleWidth      =   5316
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   720
            Visible         =   0   'False
            Width           =   5316
         End
         Begin VB.PictureBox PictureVSelBottom 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   20
            Left            =   1320
            ScaleHeight     =   24
            ScaleWidth      =   396
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1800
            Visible         =   0   'False
            Width           =   400
         End
         Begin VB.PictureBox PictureVSelTop 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   20
            Left            =   1320
            ScaleHeight     =   24
            ScaleWidth      =   396
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   480
            Visible         =   0   'False
            Width           =   400
         End
         Begin VB.PictureBox PictureVSelRight 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   1332
            Left            =   1680
            ScaleHeight     =   1332
            ScaleWidth      =   24
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   480
            Visible         =   0   'False
            Width           =   20
         End
         Begin VB.PictureBox PictureVSelLeft 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   1332
            Left            =   1320
            ScaleHeight     =   1332
            ScaleWidth      =   24
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   480
            Visible         =   0   'False
            Width           =   20
         End
         Begin MSFlexGridLib.MSFlexGrid StepsView 
            Height          =   4212
            Left            =   120
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   480
            Width           =   6252
            _ExtentX        =   11028
            _ExtentY        =   7430
            _Version        =   393216
            Rows            =   16
            Cols            =   81
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            BorderStyle     =   0
            Appearance      =   0
         End
         Begin VB.Label LabelFrameMain 
            AutoSize        =   -1  'True
            BackColor       =   &H00F4C0C0&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00FFFFFF&
            Height          =   192
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   1020
         End
         Begin VB.Shape ShapeFrameMainCaption 
            BackColor       =   &H00F4C0C0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FF8080&
            Height          =   252
            Left            =   0
            Top             =   0
            Width           =   6612
         End
         Begin VB.Label LabelFont 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LabelFont"
            Height          =   192
            Left            =   120
            TabIndex        =   17
            Top             =   5640
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Shape ShapeFrameMain 
            BackColor       =   &H00F4E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FF8080&
            Height          =   5916
            Left            =   0
            Top             =   0
            Width           =   6648
         End
      End
      Begin VB.Frame FrameDescription 
         BackColor       =   &H00F4E0E0&
         BorderStyle     =   0  'None
         Height          =   816
         Left            =   240
         TabIndex        =   5
         Top             =   5640
         Width           =   6672
         Begin VB.Label LabelDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "Пояснение"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.2
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   492
            Left            =   120
            TabIndex        =   20
            Top             =   120
            Width           =   6420
            WordWrap        =   -1  'True
         End
         Begin VB.Shape ShapeMessageBorderLight 
            BackColor       =   &H00F4E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   792
            Left            =   12
            Shape           =   4  'Rounded Rectangle
            Top             =   12
            Width           =   6648
         End
         Begin VB.Shape ShapeMessageBorderDark 
            BorderColor     =   &H80000010&
            Height          =   816
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   6672
         End
      End
      Begin VB.Frame FrameSplitterUpDown 
         BackColor       =   &H00F4C0C0&
         BorderStyle     =   0  'None
         Height          =   40
         Left            =   120
         MousePointer    =   7  'Size N S
         TabIndex        =   4
         Top             =   5520
         Width           =   6768
      End
      Begin VB.Frame FrameSettings 
         BorderStyle     =   0  'None
         Height          =   5052
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   6060
         Begin VB.CommandButton CommandBrowse 
            Caption         =   "..."
            Height          =   240
            Left            =   1800
            TabIndex        =   19
            Top             =   4440
            Visible         =   0   'False
            Width           =   252
         End
         Begin VB.ComboBox ComboCell 
            Height          =   288
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   4440
            Visible         =   0   'False
            Width           =   732
         End
         Begin VB.TextBox TextCell 
            BorderStyle     =   0  'None
            Height          =   288
            Left            =   120
            TabIndex        =   16
            Top             =   4440
            Visible         =   0   'False
            Width           =   732
         End
         Begin MSFlexGridLib.MSFlexGrid MSFGSettings 
            Height          =   972
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   1715
            _Version        =   393216
            FixedCols       =   0
            BackColor       =   16048352
            BackColorFixed  =   16040128
            BackColorBkg    =   16048352
            GridColor       =   16048352
            GridColorFixed  =   16040128
            AllowBigSelection=   0   'False
            HighLight       =   0
            GridLinesFixed  =   1
            ScrollBars      =   2
            AllowUserResizing=   1
            BorderStyle     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Shape ShapeSSTabFrame 
         BackColor       =   &H00F4E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF8080&
         Height          =   6156
         Left            =   120
         Top             =   360
         Width           =   6768
      End
   End
   Begin MSComDlg.CommonDialog SaveFileDialog 
      Left            =   480
      Top             =   6840
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog FontDialog 
      Left            =   0
      Top             =   6840
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      DialogTitle     =   "Шрифт"
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   6768
      Width           =   1095
   End
End
Attribute VB_Name = "FormOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**
'@author <a href="mailto:unihomelab@ya.ru">Мезенцев Вячеслав</a>
'@revision Дата ревизии: 16.06.2011 г., время: 4:19:56
'@rem <h1><b>FormOptions</b></h1>
'<pre>
'--------------------------------------------------------------------------------
' Проект   :       Конфигуратор управляющих программ
' Модуль   :       FormOptions
' Описание :       Диалоговая форма редактирования настроек программы
' Автор    :       Мезенцев Вячеслав
' Изменён  :       16.06.2011 г., время: 4:19:56
'--------------------------------------------------------------------------------
'</pre>
Option Explicit

Dim SplitterMoving As Boolean
Dim BegX As Integer, BegY As Integer

'**
'@see
'@rem Загружаем настройки внешнего вида интерфейса.
Private Sub LoadPlacement()
    '<EhHeader>
    On Error GoTo LoadPlacement_Err
    '</EhHeader>
    
    ' Размеры формы
    Left = IniFile.ReadInteger("FormOptions", "Left", 2532)
    Top = IniFile.ReadInteger("FormOptions", "Top", 1176)
    Width = IniFile.ReadInteger("FormOptions", "Width", 7080)
    Height = IniFile.ReadInteger("FormOptions", "Height", 7632)
    
    ' Размеры и положение компонентов
    ' TabControl
    SSTab.Left = IniFile.ReadInteger("FormOptions", "SSTab.Left", 0)
    SSTab.Top = IniFile.ReadInteger("FormOptions", "SSTab.Top", 0)
    SSTab.Width = IniFile.ReadInteger("FormOptions", "SSTab.Width", ScaleWidth)
    SSTab.Height = IniFile.ReadInteger("FormOptions", "SSTab.Height", 6612)

    ' Вкладка "Приложение"
    FrameSettings.Left = IniFile.ReadInteger("FormOptions", "FrameSettings.Left", 0)
    FrameSettings.Top = IniFile.ReadInteger("FormOptions", "FrameSettings.Top", SSTab.TabHeight)
    FrameSettings.Width = IniFile.ReadInteger("FormOptions", "FrameSettings.Width", SSTab.Width)
    FrameSettings.Height = IniFile.ReadInteger("FormOptions", "FrameSettings.Height", 5172)
    
    FrameSplitterUpDown.Left = IniFile.ReadInteger("FormOptions", "FrameSplitterUpDown.Left", FrameSettings.Left)
    FrameSplitterUpDown.Top = IniFile.ReadInteger("FormOptions", "FrameSplitterUpDown.Top", FrameSettings.Top + FrameSettings.Height)
    FrameSplitterUpDown.Width = IniFile.ReadInteger("FormOptions", "FrameSplitterUpDown.Width", FrameSettings.Width)
    FrameSplitterUpDown.Height = IniFile.ReadInteger("FormOptions", "FrameSplitterUpDown.Height", Settings.SplittersThickness)
    
    FrameDescription.Left = IniFile.ReadInteger("FormOptions", "FrameDescription.Left", FrameSettings.Left)
    FrameDescription.Top = IniFile.ReadInteger("FormOptions", "FrameDescription.Top", FrameSplitterUpDown.Top + FrameSplitterUpDown.Height)
    FrameDescription.Width = IniFile.ReadInteger("FormOptions", "FrameDescription.Width", FrameSettings.Width)
    FrameDescription.Height = IniFile.ReadInteger("FormOptions", "FrameDescription.Height", SSTab.Height - SSTab.TabHeight - FrameDescription.Top)
    
    MSFGSettings.Left = IniFile.ReadInteger("FormOptions", "MSFGSettings.Left", 0)
    MSFGSettings.Top = IniFile.ReadInteger("FormOptions", "MSFGSettings.Top", 0)
    MSFGSettings.Width = IniFile.ReadInteger("FormOptions", "MSFGSettings.Width", FrameSettings.Width)
    MSFGSettings.Height = IniFile.ReadInteger("FormOptions", "MSFGSettings.Height", FrameSettings.Height)
    
    MSFGSettings.ColWidth(0) = IniFile.ReadInteger("FormOptions", "MSFGSettings.ColWidth0", MSFGSettings.Width / 2)
    MSFGSettings.ColWidth(1) = IniFile.ReadInteger("FormOptions", "MSFGSettings.ColWidth1", MSFGSettings.Width / 2)
    
    '<EhFooter>
    Exit Sub

LoadPlacement_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.LoadPlacement]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

'**
'@see
'@rem Сохранение внешнего вида интерфейса.
Private Sub SavePlacement()
    '<EhHeader>
    On Error GoTo SavePlacement_Err
    '</EhHeader>

    ' Размеры формы
    IniFile.WriteInteger "FormOptions", "Left", Left
    IniFile.WriteInteger "FormOptions", "Top", Top
    IniFile.WriteInteger "FormOptions", "Width", Width
    IniFile.WriteInteger "FormOptions", "Height", Height
    
    ' Размеры и положение компонентов
    ' TabControl
    IniFile.WriteInteger "FormOptions", "SSTab.Left", SSTab.Left
    IniFile.WriteInteger "FormOptions", "SSTab.Top", SSTab.Top
    IniFile.WriteInteger "FormOptions", "SSTab.Width", SSTab.Width
    IniFile.WriteInteger "FormOptions", "SSTab.Height", SSTab.Height
    
    ' Вкладка "Приложение"
    IniFile.WriteInteger "FormOptions", "FrameSettings.Left", FrameSettings.Left
    IniFile.WriteInteger "FormOptions", "FrameSettings.Top", FrameSettings.Top
    IniFile.WriteInteger "FormOptions", "FrameSettings.Width", FrameSettings.Width
    IniFile.WriteInteger "FormOptions", "FrameSettings.Height", FrameSettings.Height
    
    IniFile.WriteInteger "FormOptions", "FrameSplitterUpDown.Left", FrameSplitterUpDown.Left
    IniFile.WriteInteger "FormOptions", "FrameSplitterUpDown.Top", FrameSplitterUpDown.Top
    IniFile.WriteInteger "FormOptions", "FrameSplitterUpDown.Width", FrameSplitterUpDown.Width
    IniFile.WriteInteger "FormOptions", "FrameSplitterUpDown.Height", FrameSplitterUpDown.Height
       
    IniFile.WriteInteger "FormOptions", "FrameDescription.Left", FrameDescription.Left
    IniFile.WriteInteger "FormOptions", "FrameDescription.Top", FrameDescription.Top
    IniFile.WriteInteger "FormOptions", "FrameDescription.Width", FrameDescription.Width
    IniFile.WriteInteger "FormOptions", "FrameDescription.Height", FrameDescription.Height

    IniFile.WriteInteger "FormOptions", "MSFGSettings.Left", MSFGSettings.Left
    IniFile.WriteInteger "FormOptions", "MSFGSettings.Top", MSFGSettings.Top
    IniFile.WriteInteger "FormOptions", "MSFGSettings.Width", MSFGSettings.Width
    IniFile.WriteInteger "FormOptions", "MSFGSettings.Height", MSFGSettings.Height
    
    IniFile.WriteInteger "FormOptions", "MSFGSettings.ColWidth0", MSFGSettings.ColWidth(0)
    IniFile.WriteInteger "FormOptions", "MSFGSettings.ColWidth1", MSFGSettings.ColWidth(1)
    
    '<EhFooter>
    Exit Sub

SavePlacement_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.SavePlacement]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub ShowHorizontalSelector()
    '<EhHeader>
    On Error GoTo ShowHorizontalSelector_Err
    '</EhHeader>

    Dim SelectorWidth As Integer
    
    SelectorWidth = Settings.StepsSelectorWidth
    
    ' Отображаем горизонтальный селектор
    If Settings.StepsHSelectorEnabled Then
    
        PictureHSelLeft.Top = StepsView.Top + StepsView.RowPos(StepsView.RowSel) - SelectorWidth / 2
        PictureHSelLeft.Left = StepsView.Left
        PictureHSelLeft.Width = SelectorWidth
        PictureHSelLeft.Height = StepsView.RowHeight(StepsView.RowSel)
        
        PictureHSelRight.Top = PictureHSelLeft.Top
        PictureHSelRight.Left = PictureHSelLeft.Left + StepsView.ColWidth(0) + StepsView.ColWidth(1) * (StepsView.Cols - 1)
        PictureHSelRight.Height = PictureHSelLeft.Height
        PictureHSelRight.Width = SelectorWidth
        
        PictureHSelTop.Left = PictureHSelLeft.Left
        PictureHSelTop.Top = PictureHSelLeft.Top
        PictureHSelTop.Height = SelectorWidth
        PictureHSelTop.Width = StepsView.ColWidth(0) + StepsView.ColWidth(1) * (StepsView.Cols - 1) + SelectorWidth
        
        PictureHSelBottom.Left = PictureHSelLeft.Left
        PictureHSelBottom.Top = PictureHSelLeft.Top + PictureHSelLeft.Height
        PictureHSelBottom.Height = SelectorWidth
        PictureHSelBottom.Width = StepsView.ColWidth(0) + StepsView.ColWidth(1) * (StepsView.Cols - 1) + SelectorWidth
        
        If StepsView.RowIsVisible(StepsView.RowSel) Then
        
            PictureHSelLeft.Visible = True
            PictureHSelRight.Visible = True
            PictureHSelTop.Visible = True
            PictureHSelBottom.Visible = True
        
        Else
        
            PictureHSelLeft.Visible = False
            PictureHSelRight.Visible = False
            PictureHSelTop.Visible = False
            PictureHSelBottom.Visible = False
        
        End If

    Else
        
        PictureHSelLeft.Visible = False
        PictureHSelRight.Visible = False
        PictureHSelTop.Visible = False
        PictureHSelBottom.Visible = False
            
    End If

    '<EhFooter>
    Exit Sub

ShowHorizontalSelector_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.ShowHorizontalSelector]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub ShowVerticalSelector()
    '<EhHeader>
    On Error GoTo ShowVerticalSelector_Err
    '</EhHeader>

    Dim SelectorWidth As Integer
    
    SelectorWidth = Settings.StepsSelectorWidth
    
    ' Отображаем вертикальный селектор
    If Settings.StepsVSelectorEnabled Then
            
        PictureVSelLeft.Top = StepsView.Top
        PictureVSelLeft.Left = StepsView.Left + StepsView.ColPos(StepsView.ColSel) - SelectorWidth / 2
        PictureVSelLeft.Width = SelectorWidth
        PictureVSelLeft.Height = StepsView.RowHeight(0) + StepsView.RowHeight(1) * (StepsView.rows - 1)
        
        PictureVSelRight.Top = PictureVSelLeft.Top
        PictureVSelRight.Left = PictureVSelLeft.Left + StepsView.ColWidth(StepsView.ColSel)
        PictureVSelRight.Height = PictureVSelLeft.Height
        PictureVSelRight.Width = SelectorWidth
        
        PictureVSelTop.Left = PictureVSelLeft.Left
        PictureVSelTop.Top = StepsView.Top
        PictureVSelTop.Height = SelectorWidth
        PictureVSelTop.Width = StepsView.ColWidth(StepsView.ColSel) + SelectorWidth / 2
        
        PictureVSelBottom.Left = PictureVSelLeft.Left
        PictureVSelBottom.Top = PictureVSelLeft.Top + StepsView.RowHeight(0) + StepsView.RowHeight(1) * (StepsView.rows - 1)
        PictureVSelBottom.Height = SelectorWidth
        PictureVSelBottom.Width = PictureVSelTop.Width + SelectorWidth / 2
        
        If StepsView.ColIsVisible(StepsView.ColSel) Then
        
            PictureVSelLeft.Visible = True
            PictureVSelRight.Visible = True
            PictureVSelTop.Visible = True
            PictureVSelBottom.Visible = True
        
        Else
        
            PictureVSelLeft.Visible = False
            PictureVSelRight.Visible = False
            PictureVSelTop.Visible = False
            PictureVSelBottom.Visible = False
        
        End If
        
    Else
    
        PictureVSelLeft.Visible = False
        PictureVSelRight.Visible = False
        PictureVSelTop.Visible = False
        PictureVSelBottom.Visible = False
        
    End If

    '<EhFooter>
    Exit Sub

ShowVerticalSelector_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.ShowVerticalSelector]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub cmdOK_Click()
    '<EhHeader>
    On Error GoTo cmdOK_Click_Err
    '</EhHeader>

    ' Применяем параметры к интерфейсу программы
    FormMain.RefreshComponents False
       
    Unload Me

    '<EhFooter>
    Exit Sub

cmdOK_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormOptions.cmdOK_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub RefreshSettingsView()

    Dim I As Integer
    
    ' ------
    ' Загружаем установки в таблицу
    ' Отключаем отображение
    With MSFGSettings
    
        ' Обновляем значения параметров
        .Redraw = False
        
        .Col = 1
        
        For I = 1 To .rows - 1
        
            Select Case .RowData(I)
            
                Case SETTINGS_STEPS_COL_WIDTH:
                    
                    .row = I
                    .Text = CStr(Settings.StepsColWidth)
                    .CellBackColor = &HFFFFFF
                    
                Case SETTINGS_STEPS_ROW_HEIGHT:

                    .row = I
                    .Text = CStr(Settings.StepsRowHeight)
                    .CellBackColor = &HFFFFFF

                Case SETTINGS_STEPSVIEW_FONT:

                    .row = I
                    .CellAlignment = flexAlignRightCenter
                    .Text = Settings.StepsViewFontName & ", " & CStr(Settings.StepsViewFontSize)
                    .CellFontBold = Settings.StepsViewFontBold
                    .CellFontItalic = Settings.StepsViewFontItalic
                    .CellFontName = Settings.StepsViewFontName
                    .CellFontSize = Settings.StepsViewFontSize
                    .CellBackColor = &HFFFFFF

                Case SETTINGS_STEPS_SELECTOR_WIDTH:

                    .row = I
                    .Text = CStr(Settings.StepsSelectorWidth)
                    .CellBackColor = &HFFFFFF

                Case SETTINGS_STEPS_VSELECTOR_ENABLED:

                    .row = I
                    .CellAlignment = flexAlignRightCenter

                    Select Case Settings.StepsVSelectorEnabled
                        Case False: .Text = STRING_NO
                        Case True: .Text = STRING_YES
                    End Select

                    .CellBackColor = &HFFFFFF

                Case SETTINGS_STEPS_HSELECTOR_ENABLED:

                    .row = I
                    .CellAlignment = flexAlignRightCenter

                    Select Case Settings.StepsHSelectorEnabled
                        Case False: .Text = STRING_NO
                        Case True: .Text = STRING_YES
                    End Select

                    .CellBackColor = &HFFFFFF

                Case SETTINGS_REWRITE_LOGFILE:

                    .row = I
                    .CellAlignment = flexAlignRightCenter

                    Select Case Settings.RewriteLogFile
                        Case False: .Text = STRING_NO
                        Case True: .Text = STRING_YES
                    End Select

                    .CellBackColor = &HFFFFFF

                Case SETTINGS_LOG_FILEPATH:

                    .row = I
                    .CellAlignment = flexAlignRightCenter
                    .Text = Settings.LogFilePath
                    .CellBackColor = &HFFFFFF

                Case SETTINGS_FILES_HISTORY_SIZE:

                    .row = I
                    .CellAlignment = flexAlignRightCenter
                    .Text = CStr(MRUFileList.MaxFileCount)
                    .CellBackColor = &HFFFFFF

                Case SETTINGS_FILES_HISTORY_LIMIT_PATHS:

                    .row = I
                    .CellAlignment = flexAlignRightCenter

                    Select Case Settings.FilesHistoryLimitPaths
                        Case False: .Text = STRING_NO
                        Case True: .Text = STRING_YES
                    End Select

                    .CellBackColor = &HFFFFFF

                Case SETTINGS_AUTOUPDATE_ENABLED:

                    .row = I
                    .CellAlignment = flexAlignRightCenter

                    Select Case Settings.AutoUpdateEnabled
                        Case False: .Text = STRING_NO
                        Case True: .Text = STRING_YES
                    End Select

                    .CellBackColor = &HFFFFFF

                Case SETTINGS_AUTOUPDATE_PERIOD:

                    .row = I
                    .CellAlignment = flexAlignRightCenter

                    Select Case Settings.AutoUpdatePeriod

                        Case AUP_EVERY_DAY: .Text = "каждый день"

                        Case AUP_ONES_PER_WEEK: .Text = "раз в неделю"

                        Case AUP_ONES_PER_MONTH: .Text = "раз в месяц"

                    End Select

                    .CellBackColor = &HFFFFFF

                Case SETTINGS_AUTOUPDATE_LAST_DATE:
                    
                    .row = I
                    .Text = CStr(Settings.AutoUpdateLastDate)
                    .CellBackColor = &HFFFFFF
                    
                Case SETTINGS_IMPORT_JSON_CODEPAGE:

                    .row = I
                    .CellAlignment = flexAlignRightCenter
                    .Text = "UTF-8"
                    .CellBackColor = &HFFFFFF

                Case SETTINGS_EXPORT_JSON_CODEPAGE:

                    .row = I
                    .CellAlignment = flexAlignRightCenter
                    .Text = "UTF-8"
                    .CellBackColor = &HFFFFFF

            End Select
            
        Next
    
        .Redraw = True
    
    End With
    
End Sub

Private Sub RefreshStepsView()
    '<EhHeader>
    On Error GoTo RefreshStepsView_Err
    '</EhHeader>

    Dim S As String
    Dim Col As Integer, row As Integer

    StepsView.Redraw = False

    StepsView.Font.Bold = LabelFont.FontBold
    StepsView.Font.Italic = LabelFont.FontItalic
    StepsView.Font.Name = LabelFont.FontName
    StepsView.Font.Size = LabelFont.FontSize

    StepsView.Cols = 10 + 1

    S = "<   |"

    For Col = 1 To StepsView.Cols - 1

        If Col < StepsView.Cols - 1 Then

            If Col < 10 Then
            
                S = S & "0" & Col & "|"
                
            Else
            
                S = S & Col & "|"
                
            End If
            
        Else
        
            S = S & Col
            
        End If
        
        StepsView.Col = Col
        StepsView.row = 0
        StepsView.CellAlignment = flexAlignCenterCenter
        
    Next

    StepsView.FormatString = S

    S = ";|" _
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

    StepsView.FormatString = S

    ' "Тушим" все ячейки таблицы

    For row = 1 To StepsView.rows - 1
    
        StepsView.RowHeight(row) = Settings.StepsRowHeight

        For Col = 1 To 10
        
            StepsView.ColWidth(Col) = Settings.StepsColWidth
            StepsView.Col = Col
            StepsView.row = row
            StepsView.CellBackColor = &HC8D0D4
            
        Next
        
    Next

    StepsView.Col = 1
    StepsView.row = 1

    StepsView.Redraw = True
    
    ' Отображаем горизонтальный селектор
    ShowHorizontalSelector

    ' Отображаем вертикальный селектор
    ShowVerticalSelector

    '<EhFooter>
    Exit Sub

RefreshStepsView_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormOptions.RefreshStepsView]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub ComboCell_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = VBRUN.KeyCodeConstants.vbKeyEscape Then
    
        ComboCell.Visible = False
        RefreshTabControl
        MSFGSettings.SetFocus
        
    End If
    
    If KeyCode = VBRUN.KeyCodeConstants.vbKeyReturn Then
     
        With MSFGSettings
        
            ' Если первая колонка (имена параметров), то ничего не делаем
            If .Col = 0 Then Exit Sub
        
            Select Case .RowData(.row)
            
                Case SETTINGS_STEPS_COL_WIDTH:
                    
                Case SETTINGS_STEPS_ROW_HEIGHT:
                
                Case SETTINGS_STEPSVIEW_FONT:

                Case SETTINGS_STEPS_SELECTOR_WIDTH:
                
                Case SETTINGS_STEPS_VSELECTOR_ENABLED:
                
                    Select Case ComboCell.ListIndex
                        Case 0: Settings.StepsVSelectorEnabled = False
                        Case 1: Settings.StepsVSelectorEnabled = True
                    End Select
                
                Case SETTINGS_STEPS_HSELECTOR_ENABLED:

                    Select Case ComboCell.ListIndex
                        Case 0: Settings.StepsHSelectorEnabled = False
                        Case 1: Settings.StepsHSelectorEnabled = True
                    End Select
                    
                Case SETTINGS_REWRITE_LOGFILE:

                    Select Case ComboCell.ListIndex
                        Case 0: Settings.RewriteLogFile = False
                        Case 1: Settings.RewriteLogFile = True
                    End Select
                    
                Case SETTINGS_LOG_FILEPATH:

                Case SETTINGS_FILES_HISTORY_SIZE:

                Case SETTINGS_FILES_HISTORY_LIMIT_PATHS:

                    Select Case ComboCell.ListIndex
                        Case 0: Settings.FilesHistoryLimitPaths = False
                        Case 1: Settings.FilesHistoryLimitPaths = True
                    End Select
                    
                Case SETTINGS_AUTOUPDATE_ENABLED:
                    
                    Select Case ComboCell.ListIndex
                        Case 0: Settings.AutoUpdateEnabled = False
                        Case 1: Settings.AutoUpdateEnabled = True
                    End Select
                    
                Case SETTINGS_AUTOUPDATE_PERIOD:
                    
                    Select Case ComboCell.ListIndex
                        Case 0: Settings.AutoUpdatePeriod = AUP_EVERY_DAY
                        Case 1: Settings.AutoUpdatePeriod = AUP_ONES_PER_WEEK
                        Case 2: Settings.AutoUpdatePeriod = AUP_ONES_PER_MONTH
                    End Select
                    
                Case SETTINGS_IMPORT_JSON_CODEPAGE:

                Case SETTINGS_EXPORT_JSON_CODEPAGE:

            End Select
        
            ComboCell.Visible = False
            
            Dim row As Integer
            
            row = .row
            
            RefreshSettingsView
            RefreshStepsView
            
            If row < .rows - 1 Then .row = row
            
            .SetFocus
        
        End With
            
    End If

End Sub

Private Sub ComboCell_LostFocus()
    
    ComboCell.Visible = False
    
End Sub

Private Sub CommandBrowse_Click()
    '<EhHeader>
    On Error GoTo CommandBrowse_Click_Err
    '</EhHeader>

    SaveFileDialog.FileName = Settings.LogFilePath
    SaveFileDialog.DialogTitle = "Обзор..."
    SaveFileDialog.DefaultExt = ".log"
    SaveFileDialog.Filter = "Файл лога (*.log)|*.log|Все файлы (*.*)|(*.*)"
    SaveFileDialog.FilterIndex = 1
    SaveFileDialog.MaxFileSize = 32767
    SaveFileDialog.InitDir = CurrentDir
    SaveFileDialog.CancelError = True

    SaveFileDialog.ShowSave

    TextCell.Text = SaveFileDialog.FileName
    TextCell.SelStart = Len(TextCell.Text)
    TextCell.SelLength = 0
    TextCell.SetFocus

    '<EhFooter>
    Exit Sub

CommandBrowse_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.CommandBrowse_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub RefreshTabControl()

    SSTab.Top = 0
    SSTab.Left = 0
    SSTab.Width = ScaleWidth
    SSTab.Height = ScaleHeight - 576
    
    Select Case SSTab.Tab
    
        Case 0:
        
            
            FrameExample.Visible = False
            
            ShapeSSTabFrame.Visible = True
            FrameDescription.Visible = True
            FrameSplitterUpDown.Visible = True
            FrameSettings.Visible = True
            
            ShapeSSTabFrame.Top = SSTab.Top + SSTab.TabHeight
            ShapeSSTabFrame.Left = SSTab.Left
            ShapeSSTabFrame.Width = SSTab.Width
            ShapeSSTabFrame.Height = SSTab.Height - SSTab.TabHeight
            
            FrameDescription.Top = ShapeSSTabFrame.Top + ShapeSSTabFrame.Height - FrameDescription.Height - 40
            FrameDescription.Left = ShapeSSTabFrame.Left + 40
            FrameDescription.Width = ShapeSSTabFrame.Width - 80
            
            FrameSplitterUpDown.Height = Settings.SplittersThickness
            FrameSplitterUpDown.Top = FrameDescription.Top - FrameSplitterUpDown.Height
            FrameSplitterUpDown.Left = FrameDescription.Left
            FrameSplitterUpDown.Width = FrameDescription.Width
            
            FrameSettings.Top = ShapeSSTabFrame.Top + 40
            FrameSettings.Left = ShapeSSTabFrame.Left + 40
            FrameSettings.Width = ShapeSSTabFrame.Width - 80
            FrameSettings.Height = FrameSplitterUpDown.Top - FrameSettings.Top
        
            ShapeMessageBorderDark.Top = 40
            ShapeMessageBorderDark.Left = 0
            ShapeMessageBorderDark.Width = FrameDescription.Width
            ShapeMessageBorderDark.Height = FrameDescription.Height - 40
            
            ShapeMessageBorderLight.Top = ShapeMessageBorderDark.Top + 12
            ShapeMessageBorderLight.Left = ShapeMessageBorderDark.Left + 12
            ShapeMessageBorderLight.Width = ShapeMessageBorderDark.Width - 24
            ShapeMessageBorderLight.Height = ShapeMessageBorderDark.Height - 24
        
            MSFGSettings.Top = 0
            MSFGSettings.Left = 0
            MSFGSettings.Width = FrameSettings.Width - MSFGSettings.Left
            MSFGSettings.Height = FrameSettings.Height
              
            ' Если строки не умещаются во фрейме, то появляется вертикальная полоска прокрутки
            ' Корректируем ширину столбцов для этого случая
            Dim ScrollWidth As Long
            
            ScrollWidth = Screen.TwipsPerPixelX * GetSystemMetrics(SM_CXVSCROLL)
            
            If MSFGSettings.rows * MSFGSettings.RowHeight(0) > MSFGSettings.Height Then
    
                If MSFGSettings.Width > (MSFGSettings.ColWidth(0) + ScrollWidth) Then
                    
                    MSFGSettings.ColWidth(1) = MSFGSettings.Width - MSFGSettings.ColWidth(0) - ScrollWidth
                
                End If
            
            Else
                
                If MSFGSettings.Width > MSFGSettings.ColWidth(0) Then
                    
                    MSFGSettings.ColWidth(1) = MSFGSettings.Width - MSFGSettings.ColWidth(0)
                
                End If

            End If
        
        Case 1:
    
            FrameDescription.Visible = False
            FrameSplitterUpDown.Visible = False
            FrameSettings.Visible = False
            ShapeSSTabFrame.Visible = False
            
            FrameExample.Top = SSTab.Top + SSTab.TabHeight
            FrameExample.Left = SSTab.Left
            FrameExample.Width = SSTab.Width
            FrameExample.Height = SSTab.Height - SSTab.TabHeight
            
            ShapeFrameMain.Top = 0
            ShapeFrameMain.Left = 0
            ShapeFrameMain.Width = FrameExample.Width
            ShapeFrameMain.Height = FrameExample.Height
            
            ShapeFrameMainCaption.Top = 0
            ShapeFrameMainCaption.Left = 0
            ShapeFrameMainCaption.Width = ShapeFrameMain.Width
    
            LabelFrameMain.Top = ShapeFrameMainCaption.Top
            LabelFrameMain.Left = ShapeFrameMainCaption.Left + 120
            LabelFrameMain.Width = ShapeFrameMainCaption.Width - 240
            
            LabelFrameMain.FontSize = Settings.StepsViewFontSize
            
            LabelFont.FontBold = Settings.StepsViewFontBold
            LabelFont.FontItalic = Settings.StepsViewFontItalic
            LabelFont.FontName = Settings.StepsViewFontName
            LabelFont.FontSize = Settings.StepsViewFontSize
            
            StepsView.Top = ShapeFrameMainCaption.Top + ShapeFrameMainCaption.Height + 120
            StepsView.Left = 120
            StepsView.Width = FrameExample.Width - StepsView.Left - 120
            StepsView.Height = FrameExample.Height - StepsView.Top - 120
            
            FrameExample.Visible = True
            
            RefreshStepsView
    
    End Select
    
End Sub

Private Sub RefreshButtons()

    cmdOk.Top = ScaleHeight - 456
    cmdOk.Left = ScaleWidth - 2424
    
    CommandCancel.Top = ScaleHeight - 456
    CommandCancel.Left = ScaleWidth - 1224
    
End Sub

Private Sub RefreshComponents(Optional RefreshWithData As Boolean = False)

    If RefreshWithData = True Then RefreshSettingsView

    RefreshTabControl
    RefreshButtons
    
End Sub

Private Sub CommandCancel_Click()
    '<EhHeader>
    On Error GoTo CommandCancel_Click_Err
    '</EhHeader>

    Unload Me
    
    '<EhFooter>
    Exit Sub

CommandCancel_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.CommandCancel_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub Form_Load()
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>
   
    ' Восстанавливаем положение формы и компонентов
    LoadPlacement
    
    ' ------
    ' Загружаем установки в таблицу
    ' Отключаем отображение
    With MSFGSettings
    
        .Redraw = False
    
        .rows = 21
    
        ' Очищаем таблицу установок
        .Clear
    
        Dim row As Long
    
        .FormatString = "<Параметр|Значение"
        .row = 0
        .CellFontBold = True
        .CellForeColor = &HFFFFFF
        .Col = 1
        .CellAlignment = flexAlignRightCenter
        .CellFontBold = True
        .CellForeColor = &HFFFFFF
    
        row = .row
    
        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .Text = "Вид"
        .CellFontBold = True
        
        
        .row = Inc(row)
        .RowData(.row) = SETTINGS_STEPS_COL_WIDTH
        .Text = "Ширина столбца шага"
        .CellBackColor = &HFFFFFF
    
        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .RowData(.row) = SETTINGS_STEPS_ROW_HEIGHT
        .Text = "Высота строк таблиц"
        .CellBackColor = &HFFFFFF
    
        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .RowData(.row) = SETTINGS_STEPSVIEW_FONT
        .Text = "Шрифт"
        .CellBackColor = &HFFFFFF
       
        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .RowData(.row) = SETTINGS_STEPS_SELECTOR_WIDTH
        .Text = "Толщина селектора"
        .CellBackColor = &HFFFFFF
    
        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .RowData(.row) = SETTINGS_STEPS_VSELECTOR_ENABLED
        .Text = "Вертикальный селектор"
        .CellBackColor = &HFFFFFF
    
        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .RowData(.row) = SETTINGS_STEPS_HSELECTOR_ENABLED
        .Text = "Горизонтальный селектор"
        .CellBackColor = &HFFFFFF
    
        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .Text = "Лог"
        .CellFontBold = True
    
        .row = Inc(row)
        .RowData(.row) = SETTINGS_REWRITE_LOGFILE
        .Text = "Перезаписывать файл лога при запуске"
        .CellBackColor = &HFFFFFF
    
        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .RowData(.row) = SETTINGS_LOG_FILEPATH
        .Text = "Путь к файлу"
        .CellBackColor = &HFFFFFF
    
        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .Text = "История файлов"
        .CellFontBold = True
    
        .row = Inc(row)
        .RowData(.row) = SETTINGS_FILES_HISTORY_SIZE
        .Text = "Помнить не более (файлов)"
        .CellBackColor = &HFFFFFF
    
        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .RowData(.row) = SETTINGS_FILES_HISTORY_LIMIT_PATHS
        .Text = "Ограничивать длину пути"
        .CellBackColor = &HFFFFFF

        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .Text = "Обновление"
        .CellFontBold = True
    
        .row = Inc(row)
        .RowData(.row) = SETTINGS_AUTOUPDATE_ENABLED
        .Text = "Автоматическое обновление"
        .CellBackColor = &HFFFFFF

        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .RowData(.row) = SETTINGS_AUTOUPDATE_PERIOD
        .Text = "Период автообновления"
        .CellBackColor = &HFFFFFF

        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .RowData(.row) = SETTINGS_AUTOUPDATE_LAST_DATE
        .Text = "Последнее обновление"
        .CellBackColor = &HFFFFFF
        
        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .Text = "Импорт/экспорт (JSON)"
        .CellFontBold = True
    
        .row = Inc(row)
        .RowData(.row) = SETTINGS_IMPORT_JSON_CODEPAGE
        .Text = "Кодировка импорта"
        .CellBackColor = &HFFFFFF

        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .RowData(.row) = SETTINGS_EXPORT_JSON_CODEPAGE
        .Text = "Кодировка экспорта"
        .CellBackColor = &HFFFFFF

        .Redraw = True
    
    End With
        
    RefreshStepsView

    StepsView.row = 1
    StepsView.Col = 1

    ShowHorizontalSelector
    ShowVerticalSelector
    
    RefreshComponents True
       
    ' Симулируем изменение размером формы для вызова Resize()
    Move Left, Top, Width, Height

    '<EhFooter>
    Exit Sub

Form_Load_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormOptions.Form_Load]: " _
        & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

    RefreshComponents

End Sub

Private Sub Form_Unload(Cancel As Integer)
    '<EhHeader>
    On Error GoTo Form_Unload_Err
    '</EhHeader>

    SavePlacement

    '<EhFooter>
    Exit Sub

Form_Unload_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.Form_Unload]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub FrameSplitterUpDown_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    SplitterMoving = True
    BegX = x
    BegY = y
    
End Sub

Private Sub FrameSplitterUpDown_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If SplitterMoving Then
    
        FrameDescription.Height = FrameDescription.Height - y + BegY
        
        RefreshTabControl
    
    End If
    
End Sub

Private Sub FrameSplitterUpDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    SplitterMoving = False
    
End Sub

Private Sub MSFGSettings_Click()
    '<EhHeader>
    On Error GoTo MSFGSettings_Click_Err
    '</EhHeader>

        With MSFGSettings
        
            Select Case .RowData(.row)
            
                Case SETTINGS_STEPS_COL_WIDTH:
                
                    LabelDescription.Caption = "Изменение ширины всех колонок шагов"
                    Exit Sub
                
                Case SETTINGS_STEPS_ROW_HEIGHT:
                
                    LabelDescription.Caption = "Изменение высоты всех строк таблицы шагов"
                    Exit Sub
                
                Case SETTINGS_STEPSVIEW_FONT:

                    LabelDescription.Caption = "Настройки шрифта для таблицы шагов"
                    Exit Sub
                    
                Case SETTINGS_STEPS_SELECTOR_WIDTH:
                
                    LabelDescription.Caption = "Толщина рамки селектора"
                    Exit Sub
                    
                Case SETTINGS_STEPS_VSELECTOR_ENABLED:

                    LabelDescription.Caption = "Показать/скрыть вертикальный селектор"
                    Exit Sub
                    
                Case SETTINGS_STEPS_HSELECTOR_ENABLED:

                    LabelDescription.Caption = "Показать/скрыть горизонтальный селектор"
                    Exit Sub
                    
                Case SETTINGS_REWRITE_LOGFILE:

                    LabelDescription.Caption = "Перезаписывать файл лога при запуске программы"
                    Exit Sub
                    
                Case SETTINGS_LOG_FILEPATH:
                    
                    LabelDescription.Caption = "Путь к файлу лога программы"
                    Exit Sub
                    
                Case SETTINGS_FILES_HISTORY_SIZE:
                    
                    LabelDescription.Caption = "Количество файлов в истории (максимально 10)"
                    Exit Sub
                    
                Case SETTINGS_FILES_HISTORY_LIMIT_PATHS:

                    LabelDescription.Caption = "Ограничение длины пути файла в истории"
                    Exit Sub
                    
                Case SETTINGS_AUTOUPDATE_ENABLED:

                    LabelDescription.Caption = "Настройка режима обновления: автомат или ручное"
                    Exit Sub
                    
                Case SETTINGS_AUTOUPDATE_PERIOD:
                    
                    LabelDescription.Caption = "Настройка интервала обновления в автомате"
                    Exit Sub
                    
                Case SETTINGS_AUTOUPDATE_LAST_DATE:
                    
                    LabelDescription.Caption = "Дата последнего обновления"
                    Exit Sub
                    
                Case SETTINGS_IMPORT_JSON_CODEPAGE:

                    LabelDescription.Caption = "Тип кодировки при импорте из JSON-формата"
                    Exit Sub

                Case SETTINGS_EXPORT_JSON_CODEPAGE:

                    LabelDescription.Caption = "Тип кодировки при экспорте в JSON-формат"
                    Exit Sub

            End Select

            LabelDescription.Caption = ""
            
        End With
            
    '<EhFooter>
    Exit Sub

MSFGSettings_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.MSFGSettings_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub MSFGSettings_DblClick()
    '<EhHeader>
    On Error GoTo MSFGSettings_DblClick_Err
    '</EhHeader>

    MSFGSettings_KeyDown VBRUN.KeyCodeConstants.vbKeyReturn, 0

    '<EhFooter>
    Exit Sub

MSFGSettings_DblClick_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.MSFGSettings_DblClick]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub MSFGSettings_KeyDown(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo MSFGSettings_KeyDown_Err
    '</EhHeader>

    If KeyCode = VBRUN.vbKeyReturn Then
    
        With MSFGSettings
        
            ' Если первая колонка (имена параметров), то ничего не делаем
            If .Col = 0 Or .RowData(.row) = 0 Then
            
                Exit Sub
                
            Else
                
                TextCell.FontName = .CellFontName
                TextCell.FontSize = .CellFontSize
        
                TextCell.Top = .Top + .CellTop
                TextCell.Left = .Left + .CellLeft
                TextCell.Width = .CellWidth
                TextCell.Height = .CellHeight
                
                ComboCell.FontName = .CellFontName
                ComboCell.FontSize = .CellFontSize
                
                ComboCell.Left = .Left + .CellLeft
                ComboCell.Top = .Top + .CellTop
                ComboCell.Width = .CellWidth
                ComboCell.Clear
                        
                Select Case .RowData(.row)
                
                    Case SETTINGS_STEPS_COL_WIDTH, SETTINGS_STEPS_ROW_HEIGHT, _
                        SETTINGS_STEPS_SELECTOR_WIDTH, SETTINGS_FILES_HISTORY_SIZE:
                    
                        TextCell.Text = .Text
                        TextCell.SelStart = 0
                        TextCell.SelLength = Len(TextCell.Text)
                        TextCell.Visible = True
                        TextCell.SetFocus
                    
                    Case SETTINGS_STEPS_VSELECTOR_ENABLED:
                    
                        ComboCell.AddItem STRING_NO
                        ComboCell.AddItem STRING_YES
                        
                        Select Case Settings.StepsVSelectorEnabled
                            Case False: ComboCell.ListIndex = 0
                            Case True: ComboCell.ListIndex = 1
                        End Select
                        
                        ComboCell.Visible = True
                        ComboCell.SetFocus
                    
                    Case SETTINGS_STEPS_HSELECTOR_ENABLED:
                    
                        ComboCell.AddItem STRING_NO
                        ComboCell.AddItem STRING_YES
                        
                        Select Case Settings.StepsHSelectorEnabled
                            Case False: ComboCell.ListIndex = 0
                            Case True: ComboCell.ListIndex = 1
                        End Select
                        
                        ComboCell.Visible = True
                        ComboCell.SetFocus
                    
                    Case SETTINGS_REWRITE_LOGFILE:
                    
                        ComboCell.AddItem STRING_NO
                        ComboCell.AddItem STRING_YES
                        
                        Select Case Settings.RewriteLogFile
                            Case False: ComboCell.ListIndex = 0
                            Case True: ComboCell.ListIndex = 1
                        End Select
                        
                        ComboCell.Visible = True
                        ComboCell.SetFocus
                        
                    Case SETTINGS_FILES_HISTORY_LIMIT_PATHS:
                    
                        ComboCell.AddItem STRING_NO
                        ComboCell.AddItem STRING_YES
                        
                        Select Case Settings.FilesHistoryLimitPaths
                            Case False: ComboCell.ListIndex = 0
                            Case True: ComboCell.ListIndex = 1
                        End Select
                        
                        ComboCell.Visible = True
                        ComboCell.SetFocus
                        
                    Case SETTINGS_AUTOUPDATE_ENABLED:
                        
                        ComboCell.AddItem STRING_NO
                        ComboCell.AddItem STRING_YES
                        
                        Select Case Settings.AutoUpdateEnabled
                            Case False: ComboCell.ListIndex = 0
                            Case True: ComboCell.ListIndex = 1
                        End Select
                        
                        ComboCell.Visible = True
                        ComboCell.SetFocus
    
                    Case SETTINGS_STEPSVIEW_FONT:

                        FontDialog.FontBold = Settings.StepsViewFontBold
                        FontDialog.FontItalic = Settings.StepsViewFontItalic
                        FontDialog.FontName = Settings.StepsViewFontName
                        FontDialog.FontSize = Settings.StepsViewFontSize
                        FontDialog.Flags = cdlCFBoth
                    
                        FontDialog.ShowFont
                        
                        Settings.StepsViewFontBold = FontDialog.FontBold
                        Settings.StepsViewFontItalic = FontDialog.FontItalic
                        Settings.StepsViewFontName = FontDialog.FontName
                        Settings.StepsViewFontSize = FontDialog.FontSize
                        
                        LabelFont.FontBold = FontDialog.FontBold
                        LabelFont.FontItalic = FontDialog.FontItalic
                        LabelFont.FontName = FontDialog.FontName
                        LabelFont.FontSize = FontDialog.FontSize
                                   
                        Dim row As Integer
                        
                        row = .row
                        
                        RefreshSettingsView
                        RefreshStepsView
                        
                        If row < .rows - 1 Then .row = row
                        
                        .SetFocus

                    Case SETTINGS_LOG_FILEPATH:
                        
                        TextCell.Width = TextCell.Width - CommandBrowse.Width
                        TextCell.Text = .Text
                        TextCell.SelStart = Len(TextCell.Text)
                        TextCell.SelLength = 0
                        CommandBrowse.Left = TextCell.Left + TextCell.Width
                        CommandBrowse.Top = TextCell.Top
                        CommandBrowse.Height = TextCell.Height
                        
                        TextCell.Visible = True
                        CommandBrowse.Visible = True
                        CommandBrowse.SetFocus
    
                    Case SETTINGS_AUTOUPDATE_PERIOD:
                        
                        ComboCell.AddItem "каждый день"
                        ComboCell.AddItem "раз в неделю"
                        ComboCell.AddItem "раз в месяц"
                        
                        Select Case Settings.AutoUpdatePeriod
                            Case AUP_EVERY_DAY: ComboCell.ListIndex = 0
                            Case AUP_ONES_PER_WEEK: ComboCell.ListIndex = 1
                            Case AUP_ONES_PER_MONTH: ComboCell.ListIndex = 2
                        End Select
                        
                        ComboCell.Visible = True
                        ComboCell.SetFocus
                        
                    Case SETTINGS_IMPORT_JSON_CODEPAGE:
    
                    Case SETTINGS_EXPORT_JSON_CODEPAGE:
    
                End Select
                
            End If
        
        End With
    
    End If

    '<EhFooter>
    Exit Sub

MSFGSettings_KeyDown_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.MSFGSettings_KeyDown]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
    '<EhHeader>
    On Error GoTo SSTab_Click_Err
    '</EhHeader>

    RefreshTabControl

    '<EhFooter>
    Exit Sub

SSTab_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.SSTab_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub TextCell_KeyDown(KeyCode As Integer, Shift As Integer)
    '<EhHeader>
    On Error GoTo TextCell_KeyDown_Err
    '</EhHeader>

    If KeyCode = VBRUN.KeyCodeConstants.vbKeyEscape Then
    
        TextCell.Visible = False
        RefreshTabControl
        MSFGSettings.SetFocus
        
    End If
    
    If KeyCode = VBRUN.KeyCodeConstants.vbKeyReturn Then
     
        With MSFGSettings
        
            ' Если первая колонка (имена параметров), то ничего не делаем
            If .Col = 0 Then Exit Sub
        
            Select Case .RowData(.row)
            
                Case SETTINGS_STEPS_COL_WIDTH:
                
                    Settings.StepsColWidth = CLng(TextCell.Text)
                    
                Case SETTINGS_STEPS_ROW_HEIGHT:
                
                    Settings.StepsRowHeight = CLng(TextCell.Text)

                Case SETTINGS_STEPSVIEW_FONT:

                Case SETTINGS_STEPS_SELECTOR_WIDTH:
                
                    Settings.StepsSelectorWidth = CLng(TextCell.Text)

                Case SETTINGS_STEPS_VSELECTOR_ENABLED:

                Case SETTINGS_STEPS_HSELECTOR_ENABLED:

                Case SETTINGS_REWRITE_LOGFILE:

                Case SETTINGS_LOG_FILEPATH:
                
                    Settings.LogFilePath = TextCell.Text

                Case SETTINGS_FILES_HISTORY_SIZE:
                
                    MRUFileList.MaxFileCount = CLng(TextCell.Text)

                Case SETTINGS_FILES_HISTORY_LIMIT_PATHS:

                Case SETTINGS_AUTOUPDATE_ENABLED:

                Case SETTINGS_AUTOUPDATE_PERIOD:

                Case SETTINGS_IMPORT_JSON_CODEPAGE:

                Case SETTINGS_EXPORT_JSON_CODEPAGE:

            End Select
        
            TextCell.Visible = False
            
            Dim row As Integer
            
            row = .row
            
            RefreshSettingsView
            RefreshStepsView
            
            If row < .rows - 1 Then .row = row
            
            .SetFocus
        
        End With
            
    End If
    
    '<EhFooter>
    Exit Sub

TextCell_KeyDown_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.TextCell_KeyDown]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

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
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormOptions.TextCell_KeyPress]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub TextCell_LostFocus()

    TextCell.Visible = False
    CommandBrowse.Visible = False
    
End Sub
