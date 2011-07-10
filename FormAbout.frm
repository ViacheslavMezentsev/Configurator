VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "О программе"
   ClientHeight    =   3192
   ClientLeft      =   2340
   ClientTop       =   1932
   ClientWidth     =   5412
   ClipControls    =   0   'False
   Icon            =   "FormAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3192
   ScaleWidth      =   5412
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageListPhotos 
      Left            =   600
      Top             =   2640
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
            Picture         =   "FormAbout.frx":6432
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormAbout.frx":9484
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormAbout.frx":C4D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer 
      Left            =   120
      Top             =   2760
   End
   Begin VB.PictureBox picIcon 
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   768
      Left            =   228
      MouseIcon       =   "FormAbout.frx":F52C
      MousePointer    =   99  'Custom
      Picture         =   "FormAbout.frx":15B36
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
      Left            =   2040
      TabIndex        =   0
      Top             =   2760
      Width           =   1260
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H80000010&
      Height          =   392
      Left            =   110
      Shape           =   4  'Rounded Rectangle
      Top             =   2270
      Width           =   5192
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000010&
      Height          =   752
      Left            =   110
      Shape           =   4  'Rounded Rectangle
      Top             =   1430
      Width           =   5192
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
      BorderColor     =   &H00FFFFFF&
      Height          =   372
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   5172
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   732
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   5172
   End
   Begin VB.Shape Shape2 
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Мезенцев Вячеслав [unihomelab@yandex.ru]"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   1320
      MouseIcon       =   "FormAbout.frx":16A09
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1848
      Width           =   3768
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Зыков Василий [vassily@at.ur.ru]"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   1320
      MouseIcon       =   "FormAbout.frx":2CA23
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1560
      Width           =   2880
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Екатеринбург, 2011 г."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1656
      TabIndex        =   6
      Top             =   2316
      Width           =   2100
   End
   Begin VB.Label lblDescription 
      Caption         =   "Конфигуратор предназначен для создания или изменения управляющих программ"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Конфигуратор управляющих программ"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   288
      Left            =   1320
      TabIndex        =   4
      Top             =   204
      Width           =   3888
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Версия"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   10.2
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1320
      TabIndex        =   5
      Top             =   480
      Width           =   708
   End
   Begin VB.Label lblDisclaimer 
      AutoSize        =   -1  'True
      Caption         =   "Авторы: "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   288
      Left            =   264
      TabIndex        =   3
      Top             =   1524
      Width           =   852
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
    Unload Me
End Sub

Private Sub Form_Load()
    'Me.Caption = "About " & App.Title
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
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormAbout.Form_Load]: " & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub Label2_Click()
    '<EhHeader>
    On Error GoTo Label2_Click_Err
    '</EhHeader>
    
    Dim Success As Integer
    
    ' Вызываем почтовую программу по умолчанию
    Success = ShellExecute(Me.hwnd, vbNullString, "mailto: vassily@at.ur.ru", vbNullString, vbNullString, SW_SHOWNORMAL)
    
    '<EhFooter>
    Exit Sub

Label2_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormAbout.Label2_Click]: " _
        & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

Private Sub Label3_Click()
    '<EhHeader>
    On Error GoTo Label3_Click_Err
    '</EhHeader>
    
    Dim Success As Integer
    
    ' Вызываем почтовую программу по умолчанию
    Success = ShellExecute(Me.hwnd, vbNullString, "mailto: unihomelab@ya.ru", vbNullString, vbNullString, SW_SHOWNORMAL)
    
    '<EhFooter>
    
    Exit Sub

Label3_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormAbout.Label3_Click]: " _
        & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
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
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormAbout.picIcon_Click]: " _
        & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
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
