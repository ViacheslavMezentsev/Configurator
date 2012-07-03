VERSION 5.00
Begin VB.Form FormGoto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Перейти"
   ClientHeight    =   1608
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   2844
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1608
   ScaleWidth      =   2844
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   372
      Left            =   1560
      TabIndex        =   5
      Top             =   1116
      Width           =   1092
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Перейти"
      Height          =   372
      Left            =   240
      TabIndex        =   4
      Top             =   1116
      Width           =   1092
   End
   Begin VB.ComboBox ComboSteps 
      Height          =   288
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1572
   End
   Begin VB.ComboBox ComboProgramNames 
      Height          =   288
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   192
      Width           =   1572
   End
   Begin VB.Label LabelSteps 
      AutoSize        =   -1  'True
      Caption         =   "Шаг:"
      Height          =   192
      Left            =   708
      TabIndex        =   3
      Top             =   648
      Width           =   336
   End
   Begin VB.Label LabelProgramNames 
      AutoSize        =   -1  'True
      Caption         =   "Программа:"
      Height          =   192
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   924
   End
End
Attribute VB_Name = "FormGoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    '<EhHeader>
    On Error GoTo cmdCancel_Click_Err
    '</EhHeader>
    
    Unload Me
    
    '<EhFooter>
    Exit Sub

cmdCancel_Click_Err:
    Logger.Info "[cop.FormGoto.cmdCancel_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description)

    Resume Next

    '</EhFooter>
End Sub

Private Sub cmdOK_Click()
    '<EhHeader>
    On Error GoTo cmdOK_Click_Err
    '</EhHeader>

    FormMain.ListPrograms.row = ComboProgramNames.ListIndex + 1
    FormMain.ListPrograms_Click
    
    Unload Me
    
    '<EhFooter>
    Exit Sub

cmdOK_Click_Err:
    Logger.Info "[cop.FormGoto.cmdOK_Click]: " & GetErrorMessageById( _
            Err.Number, Err.Description)

    Resume Next

    '</EhFooter>
End Sub

Private Sub Form_Load()
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>

    Dim I As Integer
    
    For I = FormMain.ListPrograms.FixedRows To FormMain.ListPrograms.rows - 1
    
        ComboProgramNames.AddItem FormMain.ListPrograms.TextMatrix(I, 0)
    Next
    
    For I = 1 To MAX_NUMBER_OF_STEPS
    
        ComboSteps.AddItem "Шаг " & CStr(I)
    Next
    
    ComboProgramNames.ListIndex = FormMain.ListPrograms.row - 1
    ComboSteps.ListIndex = FormMain.StepsView.Col - 1
    
    '<EhFooter>
    Exit Sub

Form_Load_Err:
    Logger.Info "[cop.FormGoto.Form_Load]: " & GetErrorMessageById( _
            Err.Number, Err.Description)

    Resume Next

    '</EhFooter>
End Sub
