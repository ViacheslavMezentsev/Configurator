VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormDownload 
   Caption         =   "Загрузка..."
   ClientHeight    =   1536
   ClientLeft      =   2772
   ClientTop       =   3768
   ClientWidth     =   5328
   Icon            =   "FormDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1536
   ScaleWidth      =   5328
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CommanClose 
      Caption         =   "&Закрыть"
      Height          =   360
      Left            =   4080
      TabIndex        =   4
      Top             =   1080
      Width           =   1092
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Отмена"
      Height          =   372
      Left            =   2760
      TabIndex        =   2
      Top             =   1080
      Width           =   1212
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   5120
      _ExtentX        =   9038
      _ExtentY        =   445
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label LabelTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Куда:"
      Height          =   192
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   420
   End
   Begin VB.Label LabelFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Откуда:"
      Height          =   192
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   612
   End
End
Attribute VB_Name = "FormDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

    ProgressBar.Width = ScaleWidth - ProgressBar.Left - 120
    
    CommanClose.Left = ScaleWidth - CommanClose.Width - 120
    CommanClose.Top = ScaleHeight - CommanClose.Height - 120
    
    CancelButton.Left = CommanClose.Left - CancelButton.Width - 120
    CancelButton.Top = ScaleHeight - CancelButton.Height - 120

End Sub
