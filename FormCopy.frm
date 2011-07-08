VERSION 5.00
Begin VB.Form FormCopy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Копирование"
   ClientHeight    =   3708
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   3984
   Icon            =   "FormCopy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3708
   ScaleWidth      =   3984
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CommandCopy 
      Caption         =   "Копировать"
      Height          =   372
      Left            =   1680
      TabIndex        =   5
      Top             =   3240
      Width           =   1092
   End
   Begin VB.CommandButton CommandCancel 
      Caption         =   "Отмена"
      Height          =   372
      Left            =   2880
      TabIndex        =   4
      Top             =   3240
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Height          =   3132
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3972
      Begin VB.ListBox List2 
         Height          =   2736
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   1692
      End
      Begin VB.ListBox List1 
         Height          =   2736
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1692
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1800
         TabIndex        =   3
         Top             =   1560
         Width           =   372
      End
   End
End
Attribute VB_Name = "FormCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandCancel_Click()
    Unload Me
End Sub

Private Sub CommandCopy_Click()
    '<EhHeader>
    On Error GoTo CommandCopy_Click_Err
    '</EhHeader>

    If List1.SelCount > 0 And List2.SelCount > 0 Then
    
        Manager.CopyProgram List1.ListIndex, List2.ListIndex
        SetModified True
        FormMain.RefreshDataComponents
    End If
    
    '<EhFooter>
    Exit Sub

CommandCopy_Click_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormCopy.CommandCopy_Click]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

