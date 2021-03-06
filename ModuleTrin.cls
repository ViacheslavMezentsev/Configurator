VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CModuleTrin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const TRIN_PARAMETERS_COUNT = IDLE_PARAMETERS_COUNT + 0

' �������� ������
Private Const TRIN_SECTION_NAME = "Trin"

' ��������� �� ���������
' ���������� ����
Private Const PAUSE_DEFAULT = False

' ��������� ������� ����
' ���������� ����
Private Pause As TYPE_BOOL_DESCRIPTION

' ���������� ����

Public LimitsLoaded As Boolean

' �����������
Private Sub Class_Initialize()
    LimitsLoaded = False
End Sub

Public Sub LoadLimits(FileName As String)
    LimitsLoaded = DoesFileExist(FileName)

    If Not LimitsLoaded Then Exit Sub

    Dim LimitsFile As New CIniFile

    LimitsFile.Create FileName

    ' ��������� �������
    ' ���������� ����
    Pause.DefaultValue = LimitsFile.ReadBoolean(TRIN_SECTION_NAME, "Pause.Default", PAUSE_DEFAULT)

    ' ���������� ����

    Set LimitsFile = Nothing
End Sub

' �������-���������
Private Sub func_SetDefaults(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordStep As TYPE_WPC_STEP)

    Dim StepPointer As Long

    StepPointer = Manager.DataPointer + _
       Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
       HEADER_SIZE_IN_BYTES + _
       Manager.StepIndex * STEP_SIZE_IN_BYTES

    PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

    ' ���������� ����

    Select Case Pause.DefaultValue
        Case False: RecordStep.Bits = RecordStep.Bits And &HFFEF
        Case True: RecordStep.Bits = RecordStep.Bits Or &H10
    End Select

    ' ���������� ����

End Sub

Public Sub SetDefaults(frm As FormMain)
    Dim RecordStep As TYPE_WPC_STEP

    func_SetDefaults frm, 0&, RecordStep
End Sub

' �������-���������
Private Sub func_ShowPropertyTableForTrin(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordStep As TYPE_WPC_STEP)

    Dim I As Integer
    Dim j As Integer
    Dim ParamStr As String
    Dim S As String
    Dim StepPointer As Long

    With frm
        ' ��������� ������� � ���������� ����������
        ParamStr = ";��������|"

        For I = 1 To TRIN_PARAMETERS_COUNT

            Select Case I
                    ' ����� ���������
                Case IDLE_ENDSOUND_FIELD:
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_ENDSOUND

                Case IDLE_DOORUNLOCK_FIELD:
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_DOORUNLOCK

                Case IDLE_PROGNAME_FIELD:
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_PROGNAME

                Case IDLE_STEP_FIELD:
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_STEP

                Case IDLE_FUNCTION_FIELD:
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_FUNCTION

                    ' ����������� ���������

                Case Else
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_UNKNOWN
            End Select

            If (I < TRIN_PARAMETERS_COUNT) Then ParamStr = ParamStr & "|"
        Next

        ' ��� ������ ������������� �������� ���������� �����
        .PropertyTable.FormatString = ParamStr

        StepPointer = Manager.DataPointer + Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

        StepPointer = Manager.DataPointer + _
           Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
           HEADER_SIZE_IN_BYTES + _
           Manager.StepIndex * STEP_SIZE_IN_BYTES

        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer

        .PropertyTable.Col = 1
        .PropertyTable.row = 0
        .PropertyTable.CellAlignment = flexAlignRightCenter
        
        For I = 1 To TRIN_PARAMETERS_COUNT
            .PropertyTable.row = I
            .PropertyTable.CellAlignment = flexAlignRightCenter

            Select Case I
                    ' ����� ���������
                Case IDLE_ENDSOUND_FIELD:

                    If (RecordTitle.LowBits And &H1) Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                    .PropertyTable.CellBackColor = &HE0E0E0

                Case IDLE_DOORUNLOCK_FIELD:

                    If (RecordTitle.LowBits And &H2) / &H2 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                    .PropertyTable.CellBackColor = &HE0E0E0

                Case IDLE_PROGNAME_FIELD:
                    S = ""

                    For j = 1 To PROG_NAME_LENGTH - 1
                        S = S & Chr$(CLng(RecordTitle.ProgName(j)))
                    Next

                    .PropertyTable.Text = S
                    .PropertyTable.CellBackColor = &HE0E0E0

                Case IDLE_STEP_FIELD:
                    .PropertyTable.CellFontBold = True
                    .PropertyTable.Text = "" & Manager.StepIndex + 1

                Case IDLE_FUNCTION_FIELD:
                    .PropertyTable.CellFontBold = True
                    .PropertyTable.Text = FunctionsStrings(RecordStep.Bits And &HF)

                    ' ����������� ���������

                Case Else
                    .PropertyTable.CellBackColor = &H8000000F

            End Select
        Next
    End With
End Sub

Public Sub ShowPropertyTableForTrin(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP

    func_ShowPropertyTableForTrin frm, 0&, RecordTitle, RecordStep
End Sub

' �������-���������
Private Sub func_EditPropertyForTrin(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordStep As TYPE_WPC_STEP)

    Dim I As Integer
    Dim StepPointer As Long
    Dim S As String

    With frm
        StepPointer = Manager.DataPointer + Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

        StepPointer = Manager.DataPointer + _
           Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
           HEADER_SIZE_IN_BYTES + _
           Manager.StepIndex * STEP_SIZE_IN_BYTES

        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer

        .ComboCell.Left = .PropertyTable.Left + .PropertyTable.CellLeft
        .ComboCell.Top = .PropertyTable.Top + .PropertyTable.CellTop
        .ComboCell.Width = .PropertyTable.CellWidth
        .ComboCell.Clear

        .TextCell.Left = .PropertyTable.Left + .PropertyTable.CellLeft
        .TextCell.Top = .PropertyTable.Top + .PropertyTable.CellTop
        .TextCell.Width = .PropertyTable.CellWidth
        .TextCell.Height = .PropertyTable.CellHeight

        Select Case .PropertyTable.row
                ' ����� ��������� ��������� � ����
            Case IDLE_ENDSOUND_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = RecordTitle.LowBits And &H1
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case IDLE_DOORUNLOCK_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordTitle.LowBits And &H2) / &H2
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case IDLE_PROGNAME_FIELD:
                S = ""

                For I = 1 To PROG_NAME_LENGTH - 1
                    S = S & Chr$(CLng(RecordTitle.ProgName(I)))
                Next
                .TextCell.Text = S
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case IDLE_STEP_FIELD:

                For I = 1 To MAX_NUMBER_OF_STEPS
                    .ComboCell.AddItem ("��� " & I)
                Next
                .ComboCell.ListIndex = Manager.StepIndex
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case IDLE_FUNCTION_FIELD:

                For I = 1 To NUMBER_OF_FUNCS
                    .ComboCell.AddItem (FunctionsStrings(I - 1))
                Next

                .ComboCell.ListIndex = RecordStep.Bits And &HF
                .ComboCell.Visible = True
                .ComboCell.SetFocus

                ' ����������� ��������� ����

            Case Else

        End Select

        .TextCell.SelStart = 0
        .TextCell.SelLength = Len(.TextCell.Text)
    End With
End Sub

Public Sub EditPropertyForTrin(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP

    func_EditPropertyForTrin frm, 0&, RecordTitle, RecordStep
End Sub

' �������-���������
Private Sub func_SetCheckBoxForTrin(frm As FormMain, _
       ByVal ValveNumber As Integer, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordStep As TYPE_WPC_STEP)

    Dim StepPointer As Long
    
    StepPointer = Manager.DataPointer + _
       Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
       HEADER_SIZE_IN_BYTES + _
       Manager.StepIndex * STEP_SIZE_IN_BYTES
    
    PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer

    Select Case ValveNumber
    
        Case 1:
        Case 2:
        Case 3:
        Case 4:
        Case 5:
        Case 6:
        Case 7:
        Case 8:
        Case 9:
        Case 10:
        Case 11:
        Case 12:
        Case 13:
        Case 14:
        Case 15:
        Case 16:
        Case 17:
        Case 18:
    
    End Select

End Sub

Public Sub SetCheckBoxForTrin(frm As FormMain, ValveNumber As Integer)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP

    func_SetCheckBoxForTrin frm, ValveNumber, 0&, RecordTitle, RecordStep
End Sub

' �������-���������
Private Sub func_SetComboPropertyForTrin(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordStep As TYPE_WPC_STEP)

    On Error GoTo ErrorHandler

    Dim I As Integer
    Dim StepPointer As Long

    With frm
        StepPointer = Manager.DataPointer + Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

        StepPointer = Manager.DataPointer + _
           Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
           HEADER_SIZE_IN_BYTES + _
           Manager.StepIndex * STEP_SIZE_IN_BYTES

        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer

        Select Case .PropertyTable.row
                ' ����� ��������� ��������� � ����
            Case IDLE_ENDSOUND_FIELD:

                Select Case .ComboCell.ListIndex
                    Case 0: RecordTitle.LowBits = RecordTitle.LowBits And &HFFFE
                    Case 1: RecordTitle.LowBits = RecordTitle.LowBits Or &H1
                End Select

            Case IDLE_DOORUNLOCK_FIELD:

                Select Case .ComboCell.ListIndex
                    Case 0: RecordTitle.LowBits = RecordTitle.LowBits And &HFFFD
                    Case 1: RecordTitle.LowBits = RecordTitle.LowBits Or &H2
                End Select

            Case IDLE_PROGNAME_FIELD:

                For I = 1 To PROG_NAME_LENGTH - 1

                    If I <= Len(.TextCell.Text) Then
                        RecordTitle.ProgName(I) = Asc(Mid$(.TextCell.Text, I, 1))
                    Else
                        RecordTitle.ProgName(I) = 0
                    End If
                Next
                RecordTitle.ProgName(PROG_NAME_LENGTH) = 0

            Case IDLE_STEP_FIELD:
                Manager.StepIndex = .ComboCell.ListIndex
                Exit Sub

            Case IDLE_FUNCTION_FIELD:
                ' ���� ������� �� ��, �� ������ �� ������

                If (RecordStep.Bits And &HF) = (.ComboCell.ListIndex And &HF) Then Exit Sub

                ' ��� ������ ����� ������� ��� ���� �� ������ �������� ���������,
                ZeroMemory RecordStep, STEP_SIZE_IN_BYTES

                ' ���������� �������� �� ��������� � ����������� �� �������,

                If LimitsLoaded Then

                    Select Case .ComboCell.ListIndex And &HF
                        Case WPC_OPERATION_IDLE ' �������
                            .ModuleIdle.SetDefaults frm

                        Case WPC_OPERATION_FILL ' �����
                            .ModuleFill.SetDefaults frm

                        Case WPC_OPERATION_DTRG ' ������
                            .ModuleDTRG.SetDefaults frm

                        Case WPC_OPERATION_HEAT ' ������
                            .ModuleHeat.SetDefaults frm

                            ' ������, ����������, ����������
                        Case WPC_OPERATION_WASH, WPC_OPERATION_RINS, WPC_OPERATION_JOLT, WPC_OPERATION_PAUS
                            .ModuleWashOrRinsOrJolt.SetDefaults frm

'<������: �������� ��������, 17.06.2011 �. � 17:23:38
'�������: ������ ���������� �� ����������� � ModuleWashOrRinsOrJolt>
'                        Case WPC_OPERATION_PAUS ' �����
'</������: �������� ��������, 17.06.2011 �. � 17:23:38>

                        Case WPC_OPERATION_DRAIN ' ����
                            .ModuleDrain.SetDefaults frm

                        Case WPC_OPERATION_SPIN ' �����
                            .ModuleSpin.SetDefaults frm

                        Case WPC_OPERATION_COOL ' ����������
                            .ModuleCool.SetDefaults frm

                        Case WPC_OPERATION_TRIN ' ���.����������
                            .ModuleTrin.SetDefaults frm

                        Case Else

                    End Select
                End If

                ' �������� ����� �������� � ���� ���� �������
                RecordStep.Bits = RecordStep.Bits Or (.ComboCell.ListIndex And &HF)

                ' ����������� ��������� ����

            Case Else

        End Select

        SetModified True
    End With

    Exit Sub

ErrorHandler:
    Err.Clear
End Sub

Public Sub SetComboPropertyForTrin(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP

    func_SetComboPropertyForTrin frm, 0&, RecordTitle, RecordStep
End Sub

' �������-���������
Private Function func_ValveEnabled(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordStep As TYPE_WPC_STEP, _
       ByVal StepIndex As Integer, _
       ByVal Num As Integer) As Boolean

    Dim StepPointer As Long

    With frm
        StepPointer = Manager.DataPointer + Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

        StepPointer = Manager.DataPointer + _
           Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
           HEADER_SIZE_IN_BYTES + _
           StepIndex * STEP_SIZE_IN_BYTES

        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer

        Select Case Num
            Case 1:
            Case 2:
            Case 3:
            Case 4:
            Case 5:
            Case 6:
            Case 7:
            Case 8:
            Case 9:
            Case 10:
            Case 11:
            Case 12:
            Case 13:
            Case 14:
            Case 15:
            Case 16:
            Case 17:
            Case 18:

            Case Else
                func_ValveEnabled = False
        End Select
    End With
End Function

Public Function ValveEnabled(frm As FormMain, StepIndex As Integer, Num As Integer) As Boolean
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP

    ValveEnabled = func_ValveEnabled(frm, 0&, RecordTitle, RecordStep, StepIndex, Num)
End Function

' �������-���������
Private Function func_ShowStepTableForTrin(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordStep As TYPE_WPC_STEP, _
       ByRef StepGrid As MSFlexGrid) As Boolean

    Dim I As Integer, j As Integer
    Dim StepPointer As Long
    Dim S As String
    
    ' ------
    ' ��������� ��������� � �������
    With StepGrid
    
        ' ��������� �����������
        .Redraw = False
    
        .rows = 7
    
        ' ������� ������� ���������
        .Clear
    
        Dim row As Long
    
        .FormatString = "<��������|��������"
        .ColWidth(0) = 2 * .Width / 3
        .ColWidth(1) = .Width / 3
        
        .row = 0
        .Col = 0
        '.CellFontBold = True
        .CellForeColor = &HFFFFFF
        
        .Col = 1
        .CellAlignment = flexAlignRightCenter
        '.CellFontBold = True
        .CellForeColor = &HFFFFFF
        .RowHeight(.row) = Settings.StepsRowHeight
    
        row = .row
    
        ' ����� ��������� [���������]
        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .Text = "���������"
        .CellFontBold = True
        .RowHeight(.row) = Settings.StepsRowHeight
        
        .Col = 1
        .Text = Manager.ProgramIndex + 1
        .CellFontBold = True
        
        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .RowData(.row) = IDLE_ENDSOUND_FIELD
        .Text = IDLE_PARAMETER_DESCR_ENDSOUND
        .CellBackColor = &HFFFFFF
        .RowHeight(.row) = Settings.StepsRowHeight
        
        ' -----------------------------------------------
        .row = Inc(row)
        .RowData(.row) = IDLE_DOORUNLOCK_FIELD
        .Text = IDLE_PARAMETER_DESCR_DOORUNLOCK
        .CellBackColor = &HFFFFFF
        .RowHeight(.row) = Settings.StepsRowHeight
        
        ' -----------------------------------------------
        .row = Inc(row)
        .RowData(.row) = IDLE_PROGNAME_FIELD
        .Text = IDLE_PARAMETER_DESCR_PROGNAME
        .CellBackColor = &HFFFFFF
        .RowHeight(.row) = Settings.StepsRowHeight
        
        ' [���]
        ' -----------------------------------------------
        .row = Inc(row)
        .RowData(.row) = IDLE_STEP_FIELD
        .Text = "���"
        .CellFontBold = True
        .RowHeight(.row) = Settings.StepsRowHeight
        
        ' -----------------------------------------------
        .row = Inc(row)
        .RowData(.row) = IDLE_FUNCTION_FIELD
        .Text = IDLE_PARAMETER_DESCR_FUNCTION
        .CellBackColor = &HFFFFFF
        .RowHeight(.row) = Settings.StepsRowHeight
        
        ' ����������� ���������
        ' -----------------------------------------------
        
        .Height = Settings.StepsRowHeight * .rows

        StepPointer = Manager.DataPointer + Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

        StepPointer = Manager.DataPointer + _
           Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
           HEADER_SIZE_IN_BYTES + _
           (.Tag - 1) * STEP_SIZE_IN_BYTES

        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer
        
        .Col = 1
        
        For I = 1 To .rows - 1
        
            .row = I
            .CellAlignment = flexAlignRightCenter
            .CellBackColor = &HFFFFFF

            Select Case .RowData(I)
            
                    ' ����� ���������
                Case IDLE_ENDSOUND_FIELD:

                    If (RecordTitle.LowBits And &H1) Then
                        .Text = STRING_YES
                    Else
                        .Text = STRING_NO
                    End If

                Case IDLE_DOORUNLOCK_FIELD:

                    If (RecordTitle.LowBits And &H2) / &H2 Then
                        .Text = STRING_YES
                    Else
                        .Text = STRING_NO
                    End If

                Case IDLE_PROGNAME_FIELD:
                    
                    S = ""

                    For j = 1 To PROG_NAME_LENGTH - 1
                        S = S & Chr$(CLng(RecordTitle.ProgName(j)))
                    Next

                    .Text = S
                    

                Case IDLE_STEP_FIELD:
                    
                    .CellFontBold = True
                    .CellBackColor = &HF4E0E0
                    .Text = "" & .Tag

                Case IDLE_FUNCTION_FIELD:
                    
                    .CellFontBold = True
                    .Text = FunctionsStrings(RecordStep.Bits And &HF)

                    ' ����������� ���������

                Case Else
                
                    .CellBackColor = &HF4E0E0

            End Select
            
        Next

        .Redraw = True

    End With
        
End Function

Public Function ShowStepTableForTrin(ByRef frm As FormMain, ByRef StepGrid As MSFlexGrid) As Boolean
    
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP

    ShowStepTableForTrin = func_ShowStepTableForTrin(frm, 0&, RecordTitle, RecordStep, StepGrid)
    
End Function

