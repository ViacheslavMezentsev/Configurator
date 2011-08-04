VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FormDownload 
   Caption         =   "Загрузка"
   ClientHeight    =   3516
   ClientLeft      =   2772
   ClientTop       =   3768
   ClientWidth     =   6588
   Icon            =   "FormDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3516
   ScaleWidth      =   6588
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   2772
      Left            =   0
      TabIndex        =   2
      Top             =   -120
      Width           =   6492
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   252
         Left            =   0
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   2124
         _ExtentX        =   3747
         _ExtentY        =   445
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Frame FrameDescription 
         BorderStyle     =   0  'None
         Height          =   1212
         Left            =   0
         TabIndex        =   5
         Top             =   1560
         Width           =   6252
         Begin VB.Label LabelAttMessage 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Пожалуйста, дождитесь окончания загрузки, затем нажмите кнопку ""Закрыть""."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.2
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   552
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   5988
            WordWrap        =   -1  'True
         End
         Begin VB.Label LabelAttention 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Внимание"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.2
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   1800
            TabIndex        =   7
            Top             =   120
            Width           =   2136
         End
         Begin VB.Shape ShapeMessageBoderLight 
            BackColor       =   &H00F4E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   1092
            Left            =   12
            Shape           =   4  'Rounded Rectangle
            Top             =   12
            Width           =   6120
         End
         Begin VB.Shape ShapeMessageBoderDark 
            BorderColor     =   &H80000010&
            Height          =   1116
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   6144
         End
      End
      Begin VB.Frame FrameSplitterUpDown 
         BackColor       =   &H00F4C0C0&
         BorderStyle     =   0  'None
         Height          =   40
         Left            =   0
         MousePointer    =   7  'Size N S
         TabIndex        =   3
         Top             =   1440
         Width           =   6288
      End
      Begin MSFlexGridLib.MSFlexGrid MSFGSettings 
         Height          =   1332
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   6540
         _ExtentX        =   11536
         _ExtentY        =   2350
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16048352
         BackColorFixed  =   16040128
         BackColorBkg    =   16048352
         GridColor       =   13160660
         GridColorFixed  =   13160660
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
   Begin VB.CommandButton CommanClose 
      Caption         =   "&Закрыть"
      Height          =   360
      Left            =   5400
      TabIndex        =   1
      Top             =   3000
      Width           =   1092
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Отмена"
      Height          =   372
      Left            =   4080
      TabIndex        =   0
      Top             =   3000
      Width           =   1212
   End
End
Attribute VB_Name = "FormDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SplitterMoving As Boolean
Dim BegX As Integer, BegY As Integer

Public Sub ShowFromText(Text As String)

    MSFGSettings.col = 1
    MSFGSettings.row = 2
    MSFGSettings.Text = Text
    
End Sub

Public Sub ShowToText(Text As String)

    MSFGSettings.col = 1
    MSFGSettings.row = 3
    MSFGSettings.Text = Text
    
End Sub

Public Sub ShowStateText(Text As String)

    MSFGSettings.col = 1
    MSFGSettings.row = 5
    MSFGSettings.Text = Text
    
End Sub

Public Sub SetProgress(Value As Integer)

    ProgressBar.Value = Value

End Sub

Private Sub RefreshFrame()

    Frame.Top = 0
    Frame.Left = 0
    Frame.Width = ScaleWidth
    Frame.Height = ScaleHeight - 600

    FrameDescription.Top = Frame.Height - FrameDescription.Height
    FrameDescription.Left = 0
    FrameDescription.Width = Frame.Width
    
    FrameSplitterUpDown.Height = Settings.SplittersThickness
    FrameSplitterUpDown.Top = FrameDescription.Top - FrameSplitterUpDown.Height
    FrameSplitterUpDown.Left = FrameDescription.Left
    FrameSplitterUpDown.Width = FrameDescription.Width
    
    MSFGSettings.Top = 0
    MSFGSettings.Left = 0
    MSFGSettings.Width = Frame.Width
    MSFGSettings.Height = FrameSplitterUpDown.Top
    
    If MSFGSettings.Width > MSFGSettings.ColWidth(0) Then
    
        MSFGSettings.ColWidth(1) = MSFGSettings.Width - MSFGSettings.ColWidth(0)
        
    End If
            
    MSFGSettings.col = 1
    MSFGSettings.row = 4
    
    ProgressBar.Top = MSFGSettings.CellTop
    ProgressBar.Left = MSFGSettings.CellLeft
    ProgressBar.Width = MSFGSettings.CellWidth
    ProgressBar.Height = MSFGSettings.CellHeight
    
    ShapeMessageBoderDark.Top = 0
    ShapeMessageBoderDark.Left = 0
    ShapeMessageBoderDark.Width = FrameDescription.Width
    ShapeMessageBoderDark.Height = FrameDescription.Height

    ShapeMessageBoderLight.Top = ShapeMessageBoderDark.Top + 12
    ShapeMessageBoderLight.Left = ShapeMessageBoderDark.Left + 12
    ShapeMessageBoderLight.Width = ShapeMessageBoderDark.Width - ShapeMessageBoderLight.Left - 12
    ShapeMessageBoderLight.Height = ShapeMessageBoderDark.Height - ShapeMessageBoderLight.Top - 12

End Sub

Private Sub RefreshComponents()

    RefreshFrame
    
    CommanClose.Left = ScaleWidth - CommanClose.Width - 120
    CommanClose.Top = ScaleHeight - CommanClose.Height - 120
    
    CancelButton.Left = CommanClose.Left - CancelButton.Width - 120
    CancelButton.Top = ScaleHeight - CancelButton.Height - 120

End Sub

Private Sub CancelButton_Click()
    '<EhHeader>
    On Error GoTo CancelButton_Click_Err
    '</EhHeader>

    SetCancel = True
    
    '<EhFooter>
    Exit Sub

CancelButton_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormDownload.CancelButton_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub CommanClose_Click()
    '<EhHeader>
    On Error GoTo CommanClose_Click_Err
    '</EhHeader>

    Unload Me
    
    '<EhFooter>
    Exit Sub

CommanClose_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormDownload.CommanClose_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub Form_Load()
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>

    ProgressBar.Value = 0
    SetCancel = False

    ' ------
    ' Загружаем установки в таблицу
    ' Отключаем отображение
    MSFGSettings.Redraw = False
    
    MSFGSettings.rows = 6
    
    ' Очищаем таблицу установок
    MSFGSettings.Clear
    
    MSFGSettings.FormatString = "<Параметр|Значение"
    MSFGSettings.col = 1
    MSFGSettings.row = 0
    MSFGSettings.CellAlignment = flexAlignRightCenter
    MSFGSettings.ColWidth(0) = 1500
    
    ' -----------------------------------------------
    MSFGSettings.col = 0
    MSFGSettings.row = 1
    MSFGSettings.Text = "Загрузка"
    MSFGSettings.CellFontBold = True
    
    MSFGSettings.row = 2
    MSFGSettings.Text = "Откуда"
    MSFGSettings.CellBackColor = &HFFFFFF
    
    MSFGSettings.col = 1
    MSFGSettings.Text = ""
    MSFGSettings.CellBackColor = &HFFFFFF
    
    ' -----------------------------------------------
    MSFGSettings.col = 0
    MSFGSettings.row = 3
    MSFGSettings.Text = "Куда"
    MSFGSettings.CellBackColor = &HFFFFFF
    
    MSFGSettings.col = 1
    MSFGSettings.Text = ""
    MSFGSettings.CellBackColor = &HFFFFFF
    
    ' -----------------------------------------------
    MSFGSettings.col = 0
    MSFGSettings.row = 4
    MSFGSettings.Text = "Прогресс"
    MSFGSettings.CellBackColor = &HFFFFFF
    
    MSFGSettings.col = 1
    MSFGSettings.Text = ""
    MSFGSettings.CellBackColor = &HFFFFFF
    
    ProgressBar.Top = MSFGSettings.CellTop
    ProgressBar.Left = MSFGSettings.CellLeft
    ProgressBar.Width = MSFGSettings.CellWidth
    ProgressBar.Height = MSFGSettings.CellHeight
    ProgressBar.Visible = True
    
    ' -----------------------------------------------
    MSFGSettings.col = 0
    MSFGSettings.row = 5
    MSFGSettings.Text = "Статус"
    MSFGSettings.CellBackColor = &HFFFFFF
    
    MSFGSettings.col = 1
    MSFGSettings.Text = ""
    MSFGSettings.CellBackColor = &HFFFFFF
    
    MSFGSettings.Redraw = True
    
    RefreshComponents
       
    ' Симулируем изменение размеров формы для вызова Resize()
    Move Left, Top, Width, Height
    
    '<EhFooter>
    Exit Sub

Form_Load_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormDownload.Form_Load]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Private Sub Form_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    
    RefreshComponents

End Sub

Private Sub FrameSplitterUpDown_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    SplitterMoving = True
    BegX = x
    BegY = y
    
End Sub

Private Sub FrameSplitterUpDown_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If SplitterMoving Then
    
        FrameDescription.Height = FrameDescription.Height - y + BegY
        
        RefreshFrame
    
    End If
    
End Sub

Private Sub FrameSplitterUpDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    SplitterMoving = False
    
End Sub

