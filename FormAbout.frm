VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "О программе"
   ClientHeight    =   3708
   ClientLeft      =   2340
   ClientTop       =   1932
   ClientWidth     =   5412
   ClipControls    =   0   'False
   Icon            =   "FormAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3708
   ScaleWidth      =   5412
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PictureLogoAT 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   768
      Left            =   204
      MouseIcon       =   "FormAbout.frx":6432
      MousePointer    =   99  'Custom
      Picture         =   "FormAbout.frx":6CFC
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "ООО НПФ ""Авторские технологии"""
      Top             =   2352
      Width           =   768
   End
   Begin VB.PictureBox PictureLogoVyazma 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   768
      Left            =   4440
      Picture         =   "FormAbout.frx":9D40
      ScaleHeight     =   768
      ScaleWidth      =   768
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2352
      Width           =   768
   End
   Begin MSComctlLib.ImageList ImageListPhotos 
      Left            =   720
      Top             =   3240
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormAbout.frx":CD84
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormAbout.frx":FDD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormAbout.frx":12E2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer 
      Left            =   240
      Top             =   3360
   End
   Begin VB.PictureBox picIcon 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   768
      Left            =   228
      MouseIcon       =   "FormAbout.frx":15E7E
      MousePointer    =   99  'Custom
      Picture         =   "FormAbout.frx":1C488
      ScaleHeight     =   526.236
      ScaleMode       =   0  'User
      ScaleWidth      =   526.236
      TabIndex        =   1
      Top             =   348
      Width           =   768
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   348
      Left            =   2076
      TabIndex        =   0
      Top             =   3276
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Мезенцев Вячеслав [unihomelab@yandex.ru]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Left            =   1320
      MouseIcon       =   "FormAbout.frx":1D35B
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "Щёлкните, чтобы написать письмо автору"
      Top             =   1848
      Width           =   3684
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Зыков Василий [vassily@at.ur.ru]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Left            =   1320
      MouseIcon       =   "FormAbout.frx":33375
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "Щёлкните, чтобы написать письмо автору"
      Top             =   1560
      Width           =   2688
   End
   Begin VB.Label lblDisclaimer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Авторы: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   264
      TabIndex        =   3
      Top             =   1524
      Width           =   828
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Конфигуратор управляющих программ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   1320
      TabIndex        =   4
      Top             =   204
      Width           =   3636
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Версия"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1320
      MouseIcon       =   "FormAbout.frx":4938F
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "Щёлкните для копирования в буфер"
      Top             =   480
      Width           =   660
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Предназначен для создания или изменения управляющих программ контроллера MCU-401"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   444
      Left            =   1320
      TabIndex        =   2
      Top             =   780
      Width           =   3888
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Екатеринбург, 2011 г."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1680
      TabIndex        =   6
      Top             =   2760
      Width           =   2052
   End
   Begin VB.Label LabelCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ООО НПФ ""Авторские технологии"""
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Left            =   1320
      MouseIcon       =   "FormAbout.frx":49C59
      MousePointer    =   99  'Custom
      TabIndex        =   9
      ToolTipText     =   "Щёлкните, чтобы перейти на сайт"
      Top             =   2400
      Width           =   2868
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H80000010&
      Height          =   936
      Left            =   108
      Shape           =   4  'Rounded Rectangle
      Top             =   2268
      Width           =   5196
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000010&
      Height          =   756
      Left            =   108
      Shape           =   4  'Rounded Rectangle
      Top             =   1416
      Width           =   5196
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000010&
      Height          =   1236
      Left            =   1190
      Shape           =   4  'Rounded Rectangle
      Top             =   110
      Width           =   4112
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000010&
      Height          =   1232
      Left            =   110
      Shape           =   4  'Rounded Rectangle
      Top             =   110
      Width           =   992
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   912
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   5172
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   732
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1428
      Width           =   5172
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1212
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4092
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1212
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   972
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const AS_START = 0
Private Const AS_COLLAPS = 1
Private Const AS_EXPAND = 2
Private Const AS_FINISH = 3

Private Const FRAMES_COUNT = 6

Dim AnimateState As Byte
Dim AnimateCounter As Long, ImageCounter As Long
Dim DX As Long, PicTop As Long, PicHeight As Long

Private Sub cmdOK_Click()
    '<EhHeader>
    On Error GoTo cmdOK_Click_Err
    '</EhHeader>

    Unload Me
    
    '<EhFooter>
    Exit Sub

cmdOK_Click_Err:
    Logger.Info "[cop.FormAbout.cmdOK_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description)

    Resume Next

    '</EhFooter>
End Sub

Private Sub Form_Load()
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>

    
    Me.Caption = "О программе"
    
    ' Версия программы
    Dim strFile As String
    Dim udtFileInfo As FILEINFO
    
    strFile = String(255, 0)
    GetModuleFileName 0, strFile, 255

    If GetFileVersionInformation(strFile, udtFileInfo) = eNoVersion Then
        
        udtFileInfo.FileVersion = "Версия " & App.Major & "." & App.Minor & "." & App.Revision
    
    Else
        
        udtFileInfo.FileVersion = "Версия " & udtFileInfo.FileVersion
        
    End If
    
    lblVersion.Caption = udtFileInfo.FileVersion

    lblTitle.Caption = "Конфигуратор управляющих программ"

    AnimateState = AS_FINISH
    ImageCounter = 1
    picIcon.Picture = ImageListPhotos.ListImages.Item(ImageCounter).Picture
    
    '<EhFooter>
    Exit Sub

Form_Load_Err:
    Logger.Info "[cop.FormAbout.Form_Load]: " & GetErrorMessageById( _
            Err.Number, Err.Description)

    Resume Next

    '</EhFooter>
End Sub

Private Sub Label2_Click()
    '<EhHeader>
    On Error GoTo Label2_Click_Err
    '</EhHeader>
    
    Dim Success As Integer
    
    ' Вызываем почтовую программу по умолчанию
    Success = ShellExecute(Me.hWnd, vbNullString, "mailto: vassily@at.ur.ru", vbNullString, vbNullString, SW_SHOWNORMAL)
    
    '<EhFooter>
    Exit Sub

Label2_Click_Err:
    Logger.Info "[cop.FormAbout.Label2_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description)

    Resume Next

    '</EhFooter>
End Sub

Private Sub Label3_Click()
    '<EhHeader>
    On Error GoTo Label3_Click_Err
    '</EhHeader>
    
    Dim Success As Integer
    
    ' Вызываем почтовую программу по умолчанию
    Success = ShellExecute(Me.hWnd, vbNullString, "mailto: unihomelab@ya.ru", vbNullString, vbNullString, SW_SHOWNORMAL)
    
    '<EhFooter>
    Exit Sub

Label3_Click_Err:
    Logger.Info "[cop.FormAbout.Label3_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description)

    Resume Next

    '</EhFooter>
End Sub

Private Sub LabelCompany_Click()
    '<EhHeader>
    On Error GoTo LabelCompany_Click_Err
    '</EhHeader>

    Dim Success As Integer
    
    ' Переходим на страничку
    Success = ShellExecute(Me.hWnd, vbNullString, "http://www.at.ur.ru", vbNullString, vbNullString, SW_SHOWNORMAL)

    '<EhFooter>
    Exit Sub

LabelCompany_Click_Err:
    Logger.Info "[cop.FormAbout.LabelCompany_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description)

    Resume Next

    '</EhFooter>
End Sub

Private Sub lblVersion_Click()
    '<EhHeader>
    On Error GoTo lblVersion_Click_Err
    '</EhHeader>

    Clipboard.SetText lblVersion.Caption

    '<EhFooter>
    Exit Sub

lblVersion_Click_Err:
    Logger.Info "[cop.FormAbout.lblVersion_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description)

    Resume Next

    '</EhFooter>
End Sub

Private Sub picIcon_Click()
    '<EhHeader>
    On Error GoTo picIcon_Click_Err
    '</EhHeader>
    
    Timer.Interval = 30
    Timer.Enabled = True
    AnimateState = AS_START
    
    '<EhFooter>
    Exit Sub

picIcon_Click_Err:
    Logger.Info "[cop.FormAbout.picIcon_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description)

    Resume Next

    '</EhFooter>
End Sub

Private Sub PictureLogoAT_Click()
    '<EhHeader>
    On Error GoTo PictureLogoAT_Click_Err
    '</EhHeader>

    LabelCompany_Click

    '<EhFooter>
    Exit Sub

PictureLogoAT_Click_Err:
    Logger.Info "[cop.FormAbout.PictureLogoAT_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description)

    Resume Next

    '</EhFooter>
End Sub

' Простой конечный автомат состояний
Private Sub Timer_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    
    Select Case AnimateState
    
        Case AS_START
        
            AnimateState = AS_COLLAPS
            AnimateCounter = FRAMES_COUNT
            DX = Shape5.Height / (2 * (AnimateCounter + 1))
            picIcon.AutoRedraw = False
            PicTop = picIcon.Top
            PicHeight = picIcon.Height
        
        Case AS_COLLAPS
        
            Shape5.Top = Shape5.Top + DX
            Shape1.Top = Shape1.Top + DX
            Shape5.Height = Shape5.Height - 2 * DX
            Shape1.Height = Shape1.Height - 2 * DX
            
            If Shape5.Top > picIcon.Top Then picIcon.Top = Shape5.Top
            If Shape5.Height < picIcon.Height Then picIcon.Height = Shape5.Height
            
            Dec AnimateCounter
            
            If AnimateCounter = 0 Then
            
                ' Меняем картинку и возвращаемся
                Inc ImageCounter
                ImageCounter = ((ImageCounter - 1) Mod ImageListPhotos.ListImages.Count) + 1
                picIcon.Picture = ImageListPhotos.ListImages.Item(ImageCounter).Picture
                AnimateState = AS_EXPAND
            End If
        
        Case AS_EXPAND
        
            Shape1.Top = Shape1.Top - DX
            Shape5.Top = Shape5.Top - DX
            Shape1.Height = Shape1.Height + 2 * DX
            Shape5.Height = Shape5.Height + 2 * DX
            
            If Shape5.Height > PicHeight Then
                picIcon.Top = PicTop
                picIcon.Height = PicHeight
            Else
                picIcon.Top = Shape5.Top
                picIcon.Height = Shape5.Height
            End If
            
            Inc AnimateCounter
            
            If AnimateCounter = FRAMES_COUNT Then AnimateState = AS_FINISH
            
        Case AS_FINISH
        
            Timer.Interval = 0
            picIcon.AutoRedraw = True
        
    End Select
    
End Sub
