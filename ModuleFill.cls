VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CModuleFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const FILL_PAUSE_FIELD = IDLE_FUNCTION_FIELD + 1
Private Const FILL_COLDWATERGATE_FIELD = IDLE_FUNCTION_FIELD + 2
Private Const FILL_HOTWATERGATE_FIELD = IDLE_FUNCTION_FIELD + 3
Private Const FILL_RECYCLEDWATERGATE_FIELD = IDLE_FUNCTION_FIELD + 4
Private Const FILL_ROTATION_FIELD = IDLE_FUNCTION_FIELD + 5
Private Const FILL_LEVEL_FIELD = IDLE_FUNCTION_FIELD + 6
Private Const FILL_ROTATIONTIME_FIELD = IDLE_FUNCTION_FIELD + 7
Private Const FILL_PAUSETIME_FIELD = IDLE_FUNCTION_FIELD + 8
Private Const FILL_DRUMSPEED_FIELD = IDLE_FUNCTION_FIELD + 9

Private Const FILL_PARAMETERS_COUNT = IDLE_PARAMETERS_COUNT + 9

' ���������� ���������
Private Const FILL_PARAMETER_DESCR_PAUSE = "����. �����"
Private Const FILL_PARAMETER_DESCR_COLDWATERGATE = "����. ��1"
Private Const FILL_PARAMETER_DESCR_HOTWATERGATE = "����. ��"
Private Const FILL_PARAMETER_DESCR_RECYCLEDWATERGATE = "����. ��2"
Private Const FILL_PARAMETER_DESCR_ROTATION = "����. ����."

' ���������� ���������
Private Const FILL_PARAMETER_DESCR_LEVEL = "������� ����. �����."
Private Const FILL_PARAMETER_DESCR_ROTATIONTIME = "����� ����. ������"
Private Const FILL_PARAMETER_DESCR_PAUSETIME = "����� ����� ����. ���."
Private Const FILL_PARAMETER_DESCR_DRUMSPEED = "�������� ����. ��������"

' �������� ������
Private Const FILL_SECTION_NAME = "Fill"

' ��������� �� ���������
' ���������� ����
Private Const PAUSE_DEFAULT = False
Private Const COLDWATERGATE_DEFAULT = False
Private Const HOTWATERGATE_DEFAULT = False
Private Const RECYCLEDWATERGATE_DEFAULT = False
Private Const ROTATION_DEFAULT = True

' ���������� ����
Private Const LEVEL_MIN = 0
Private Const LEVEL_MAX = 99
Private Const LEVEL_DEFAULT = 15
Private Const LEVEL_DIMENSION = "��"

Private Const ROTATIONTIME_MIN = 1
Private Const ROTATIONTIME_MAX = 250
Private Const ROTATIONTIME_DEFAULT = 6
Private Const ROTATIONTIME_DIMENSION = "���"

Private Const PAUSETIME_MIN = 1
Private Const PAUSETIME_MAX = 250
Private Const PAUSETIME_DEFAULT = 12
Private Const PAUSETIME_DIMENSION = "���"

Private Const DRUMSPEED_MIN = 30
Private Const DRUMSPEED_MAX = 70
Private Const DRUMSPEED_DEFAULT = 50
Private Const DRUMSPEED_DIMENSION = "��/���"

' ��������� ������� ����
' ���������� ����
Private Pause As TYPE_BOOL_DESCRIPTION
Private ColdWaterGate As TYPE_BOOL_DESCRIPTION
Private HotWaterGate As TYPE_BOOL_DESCRIPTION
Private RecycledWaterGate As TYPE_BOOL_DESCRIPTION
Private Rotation As TYPE_BOOL_DESCRIPTION

Private Level As TYPE_BYTE_DESCRIPTION
Private RotationTime As TYPE_BYTE_DESCRIPTION
Private PauseTime As TYPE_BYTE_DESCRIPTION
Private DrumSpeed As TYPE_BYTE_DESCRIPTION

' ���������� ����
Public LimitsLoaded As Boolean

' �����������
Private Sub Class_Initialize()
    LimitsLoaded = False
End Sub

Public Sub LoadLimits(FileName As String)
    LimitsLoaded = DoesFileExist(FileName)

    If Not LimitsLoaded Then Exit Sub

    Dim LimitsFile As New CIniFiles

    LimitsFile.Create FileName

    ' ��������� �������
    ' ���������� ����
    Pause.DefaultValue = LimitsFile.ReadBoolean(FILL_SECTION_NAME, "Pause.Default", PAUSE_DEFAULT)
    ColdWaterGate.DefaultValue = LimitsFile.ReadBoolean(FILL_SECTION_NAME, "ColdWaterGate.Default", COLDWATERGATE_DEFAULT)
    HotWaterGate.DefaultValue = LimitsFile.ReadBoolean(FILL_SECTION_NAME, "HotWaterGate.Default", HOTWATERGATE_DEFAULT)
    RecycledWaterGate.DefaultValue = LimitsFile.ReadBoolean(FILL_SECTION_NAME, "RecycledWaterGate.Default", RECYCLEDWATERGATE_DEFAULT)
    Rotation.DefaultValue = LimitsFile.ReadBoolean(FILL_SECTION_NAME, "Rotation.Default", ROTATION_DEFAULT)

    ' ���������� ����

    With Level
        .MinValue = LimitsFile.ReadInteger(FILL_SECTION_NAME, "Level.Min", LEVEL_MIN)
        .MaxValue = LimitsFile.ReadInteger(FILL_SECTION_NAME, "Level.Max", LEVEL_MAX)
        .DefaultValue = LimitsFile.ReadInteger(FILL_SECTION_NAME, "Level.Default", LEVEL_DEFAULT)
        .Dimension = LimitsFile.ReadString(FILL_SECTION_NAME, "Level.Dimension", LEVEL_DIMENSION)
    End With

    With RotationTime
        .MinValue = LimitsFile.ReadInteger(FILL_SECTION_NAME, "RotationTime.Min", ROTATIONTIME_MIN)
        .MaxValue = LimitsFile.ReadInteger(FILL_SECTION_NAME, "RotationTime.Max", ROTATIONTIME_MAX)
        .DefaultValue = LimitsFile.ReadInteger(FILL_SECTION_NAME, "RotationTime.Default", ROTATIONTIME_DEFAULT)
        .Dimension = LimitsFile.ReadString(FILL_SECTION_NAME, "RotationTime.Dimension", ROTATIONTIME_DIMENSION)
    End With

    With PauseTime
        .MinValue = LimitsFile.ReadInteger(FILL_SECTION_NAME, "PauseTime.Min", PAUSETIME_MIN)
        .MaxValue = LimitsFile.ReadInteger(FILL_SECTION_NAME, "PauseTime.Max", PAUSETIME_MAX)
        .DefaultValue = LimitsFile.ReadInteger(FILL_SECTION_NAME, "PauseTime.Default", PAUSETIME_DEFAULT)
        .Dimension = LimitsFile.ReadString(FILL_SECTION_NAME, "PauseTime.Dimension", PAUSETIME_DIMENSION)
    End With

    With DrumSpeed
        .MinValue = LimitsFile.ReadInteger(FILL_SECTION_NAME, "DrumSpeed.Min", DRUMSPEED_MIN)
        .MaxValue = LimitsFile.ReadInteger(FILL_SECTION_NAME, "DrumSpeed.Max", DRUMSPEED_MAX)
        .DefaultValue = LimitsFile.ReadInteger(FILL_SECTION_NAME, "DrumSpeed.Default", DRUMSPEED_DEFAULT)
        .Dimension = LimitsFile.ReadString(FILL_SECTION_NAME, "DrumSpeed.Dimension", DRUMSPEED_DIMENSION)
    End With

    Set LimitsFile = Nothing
End Sub

' �������-���������
Private Sub func_SetDefaults(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordFill As TYPE_WPC_FILL)

    Dim StepPointer As Long

    StepPointer = Manager.DataPointer + _
       Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
       HEADER_SIZE_IN_BYTES + _
       Manager.StepIndex * STEP_SIZE_IN_BYTES

    PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

    ' ���������� ����

    Select Case Pause.DefaultValue
        Case False: RecordFill.Bits = RecordFill.Bits And &HFFEF
        Case True: RecordFill.Bits = RecordFill.Bits Or &H10
    End Select

    Select Case ColdWaterGate.DefaultValue
        Case False: RecordFill.Bits = RecordFill.Bits And &HFFDF
        Case True: RecordFill.Bits = RecordFill.Bits Or &H20
    End Select

    Select Case HotWaterGate.DefaultValue
        Case False: RecordFill.Bits = RecordFill.Bits And &HFFBF
        Case True: RecordFill.Bits = RecordFill.Bits Or &H40
    End Select

    Select Case RecycledWaterGate.DefaultValue
        Case False: RecordFill.Bits = RecordFill.Bits And &HFF7F
        Case True: RecordFill.Bits = RecordFill.Bits Or &H80
    End Select

    Select Case Rotation.DefaultValue
        Case False: RecordFill.Bits = RecordFill.Bits And &HFEFF
        Case True: RecordFill.Bits = RecordFill.Bits Or &H100
    End Select

    ' ���������� ����
    RecordFill.Level = Level.DefaultValue
    RecordFill.RotationTime = RotationTime.DefaultValue
    RecordFill.PauseTime = PauseTime.DefaultValue
    RecordFill.DrumSpeed = DrumSpeed.DefaultValue

End Sub

Public Sub SetDefaults(frm As FormMain)
    Dim RecordFill As TYPE_WPC_FILL

    func_SetDefaults frm, 0&, RecordFill
End Sub

' �������-���������
Private Sub func_ShowPropertyTableForFill(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordFill As TYPE_WPC_FILL)

    Dim I As Integer, J As Integer
    Dim StepPointer As Long
    Dim ParamStr As String, s As String

    With frm
        ' ��������� ������� � ���������� ����������
        ParamStr = ";��������|"

        For I = 1 To FILL_PARAMETERS_COUNT

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
                Case FILL_PAUSE_FIELD:
                    ParamStr = ParamStr & FILL_PARAMETER_DESCR_PAUSE

                Case FILL_COLDWATERGATE_FIELD:
                    ParamStr = ParamStr & FILL_PARAMETER_DESCR_COLDWATERGATE

                Case FILL_HOTWATERGATE_FIELD:
                    ParamStr = ParamStr & FILL_PARAMETER_DESCR_HOTWATERGATE

                Case FILL_RECYCLEDWATERGATE_FIELD:
                    ParamStr = ParamStr & FILL_PARAMETER_DESCR_RECYCLEDWATERGATE

                Case FILL_ROTATION_FIELD:
                    ParamStr = ParamStr & FILL_PARAMETER_DESCR_ROTATION

                Case FILL_LEVEL_FIELD:
                    ParamStr = ParamStr & FILL_PARAMETER_DESCR_LEVEL

                Case FILL_ROTATIONTIME_FIELD:
                    ParamStr = ParamStr & FILL_PARAMETER_DESCR_ROTATIONTIME

                Case FILL_PAUSETIME_FIELD:
                    ParamStr = ParamStr & FILL_PARAMETER_DESCR_PAUSETIME

                Case FILL_DRUMSPEED_FIELD:
                    ParamStr = ParamStr & FILL_PARAMETER_DESCR_DRUMSPEED

                Case Else
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_UNKNOWN
            End Select

            If (I < FILL_PARAMETERS_COUNT) Then ParamStr = ParamStr & "|"
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

        .PropertyTable.col = 1

        For I = 1 To FILL_PARAMETERS_COUNT
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
                    s = ""

                    For J = 1 To PROG_NAME_LENGTH - 1
                        s = s & Chr$(CLng(RecordTitle.ProgName(J)))
                    Next

                    .PropertyTable.Text = s
                    .PropertyTable.CellBackColor = &HE0E0E0

                Case IDLE_STEP_FIELD:
                    .PropertyTable.CellFontBold = True
                    .PropertyTable.Text = "" & Manager.StepIndex + 1

                Case IDLE_FUNCTION_FIELD:
                    .PropertyTable.CellFontBold = True
                    .PropertyTable.Text = FunctionsStrings(RecordFill.Bits And &HF)

                    ' ����������� ���������
                Case FILL_PAUSE_FIELD:

                    If (RecordFill.Bits And &H10) / &H10 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                Case FILL_COLDWATERGATE_FIELD:

                    If (RecordFill.Bits And &H20) / &H20 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                Case FILL_HOTWATERGATE_FIELD:

                    If (RecordFill.Bits And &H40) / &H40 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                Case FILL_RECYCLEDWATERGATE_FIELD:

                    If (RecordFill.Bits And &H80) / &H80 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                Case FILL_ROTATION_FIELD:

                    If (RecordFill.Bits And &H100) / &H100 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                Case FILL_LEVEL_FIELD:

                    If LimitsLoaded Then

                        If RecordFill.Level < Level.MinValue Or RecordFill.Level > Level.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordFill.Level

                Case FILL_ROTATIONTIME_FIELD:

                    If LimitsLoaded Then

                        If RecordFill.RotationTime < RotationTime.MinValue Or RecordFill.RotationTime > RotationTime.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordFill.RotationTime

                Case FILL_PAUSETIME_FIELD:

                    If LimitsLoaded Then

                        If RecordFill.PauseTime < PauseTime.MinValue Or RecordFill.PauseTime > PauseTime.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordFill.PauseTime

                Case FILL_DRUMSPEED_FIELD:

                    If LimitsLoaded Then

                        If RecordFill.DrumSpeed < DrumSpeed.MinValue Or RecordFill.DrumSpeed > DrumSpeed.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordFill.DrumSpeed

                Case Else
                    .PropertyTable.CellBackColor = &H8000000F

            End Select
        Next
    End With
End Sub

Public Sub ShowPropertyTableForFill(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordFill As TYPE_WPC_FILL

    func_ShowPropertyTableForFill frm, 0&, RecordTitle, RecordFill
End Sub

' �������-���������
Private Sub func_EditPropertyForFill(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordFill As TYPE_WPC_FILL)

    Dim I As Integer
    Dim StepPointer As Long
    Dim s As String

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

        .LabelDescription.Left = .PropertyTable.Left
        .LabelDescription.Width = .PropertyTable.Width
        .LabelDescription.Height = 5 * .PropertyTable.RowHeight(0)
        .LabelDescription.Top = .PropertyTable.Top + .PropertyTable.Height - .LabelDescription.Height

        .ShapeDescription.Left = .LabelDescription.Left
        .ShapeDescription.Width = .LabelDescription.Width
        .ShapeDescription.Height = .LabelDescription.Height
        .ShapeDescription.Top = .LabelDescription.Top

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
                .ComboCell.ListIndex = (RecordTitle.LowBits And &H2) / 2 ^ 1
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case IDLE_PROGNAME_FIELD:
                s = ""

                For I = 1 To PROG_NAME_LENGTH - 1
                    s = s & Chr$(CLng(RecordTitle.ProgName(I)))
                Next
                .TextCell.Text = s
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
                .ComboCell.ListIndex = RecordFill.Bits And &HF
                .ComboCell.Visible = True
                .ComboCell.SetFocus

                ' ����������� ��������� ����
            Case FILL_PAUSE_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordFill.Bits And &H10) / &H10
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case FILL_COLDWATERGATE_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordFill.Bits And &H20) / &H20
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case FILL_HOTWATERGATE_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordFill.Bits And &H40) / &H40
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case FILL_RECYCLEDWATERGATE_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordFill.Bits And &H80) / &H80
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case FILL_ROTATION_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordFill.Bits And &H100) / 2 ^ 8
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case FILL_LEVEL_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & Level.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & Level.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & Level.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & Level.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordFill.Level
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case FILL_ROTATIONTIME_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & RotationTime.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & RotationTime.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & RotationTime.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & RotationTime.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordFill.RotationTime
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case FILL_PAUSETIME_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & PauseTime.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & PauseTime.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & PauseTime.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & PauseTime.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordFill.PauseTime
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case FILL_DRUMSPEED_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & DrumSpeed.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & DrumSpeed.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & DrumSpeed.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & DrumSpeed.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordFill.DrumSpeed
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case Else

        End Select

        .TextCell.SelStart = 0
        .TextCell.SelLength = Len(.TextCell.Text)
    End With
End Sub

Public Sub EditPropertyForFill(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordFill As TYPE_WPC_FILL

    func_EditPropertyForFill frm, 0&, RecordTitle, RecordFill
End Sub

' �������-���������
Private Sub func_SetCheckBoxForFill(frm As FormMain, _
       ByVal ValveNumber As Integer, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordFill As TYPE_WPC_FILL)
       
    Dim StepPointer As Long
    
    StepPointer = Manager.DataPointer + _
       Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
       HEADER_SIZE_IN_BYTES + _
       Manager.StepIndex * STEP_SIZE_IN_BYTES
    
    PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer

    Select Case ValveNumber
    
        Case 1: RecordFill.Bits = RecordFill.Bits Xor &H40 ' "����. ��"
        Case 2: RecordFill.Bits = RecordFill.Bits Xor &H20 ' "����. ��1"
        Case 3: RecordFill.Bits = RecordFill.Bits Xor &H80 ' "����. ��2"
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
        Case 18: RecordFill.Bits = RecordFill.Bits Xor &H100 ' "����. ����."
    
    End Select
    
    SetModified True
End Sub

Public Sub SetCheckBoxForFill(frm As FormMain, ValveNumber As Integer)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordFill As TYPE_WPC_FILL

    func_SetCheckBoxForFill frm, ValveNumber, 0&, RecordTitle, RecordFill
End Sub

' �������-���������
Private Sub func_SetComboPropertyForFill(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordFill As TYPE_WPC_FILL)

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

                If (RecordFill.Bits And &HF) = (.ComboCell.ListIndex And &HF) Then Exit Sub

                ' ��� ������ ����� ������� ��� ���� �� ������ �������� ���������,
                ZeroMemory RecordFill, STEP_SIZE_IN_BYTES

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

'<������: �������� ��������, 17.06.2011 �. � 17:22:29
'�������: ������ ���������� �� ����������� � ModuleWashOrRinsOrJolt>
'                        Case WPC_OPERATION_PAUS ' �����
'</������: �������� ��������, 17.06.2011 �. � 17:22:29>

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
                RecordFill.Bits = RecordFill.Bits Or (.ComboCell.ListIndex And &HF)

                ' ����������� ��������� ����
            Case FILL_PAUSE_FIELD:

                Select Case .ComboCell.ListIndex
                    Case 0: RecordFill.Bits = RecordFill.Bits And &HFFEF
                    Case 1: RecordFill.Bits = RecordFill.Bits Or &H10
                End Select

            Case FILL_COLDWATERGATE_FIELD:

                Select Case .ComboCell.ListIndex
                    Case 0: RecordFill.Bits = RecordFill.Bits And &HFFDF
                    Case 1: RecordFill.Bits = RecordFill.Bits Or &H20
                End Select

            Case FILL_HOTWATERGATE_FIELD:

                Select Case .ComboCell.ListIndex
                    Case 0: RecordFill.Bits = RecordFill.Bits And &HFFBF
                    Case 1: RecordFill.Bits = RecordFill.Bits Or &H40
                End Select

            Case FILL_RECYCLEDWATERGATE_FIELD:

                Select Case .ComboCell.ListIndex
                    Case 0: RecordFill.Bits = RecordFill.Bits And &HFF7F
                    Case 1: RecordFill.Bits = RecordFill.Bits Or &H80
                End Select

            Case FILL_ROTATION_FIELD:

                Select Case .ComboCell.ListIndex
                    Case 0: RecordFill.Bits = RecordFill.Bits And &HFEFF
                    Case 1: RecordFill.Bits = RecordFill.Bits Or &H100
                End Select

            Case FILL_LEVEL_FIELD:
                RecordFill.Level = Val(.TextCell.Text)

            Case FILL_ROTATIONTIME_FIELD:
                RecordFill.RotationTime = Val(.TextCell.Text)

            Case FILL_PAUSETIME_FIELD:
                RecordFill.PauseTime = Val(.TextCell.Text)

            Case FILL_DRUMSPEED_FIELD:
                RecordFill.DrumSpeed = Val(.TextCell.Text)

        End Select

        SetModified True

    End With

    Exit Sub

ErrorHandler:
    Err.Clear
End Sub

Public Sub SetComboPropertyForFill(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordFill As TYPE_WPC_FILL

    func_SetComboPropertyForFill frm, 0&, RecordTitle, RecordFill
End Sub

' �������-���������
Private Function func_ValveEnabled(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordFill As TYPE_WPC_FILL, _
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
            Case 1: func_ValveEnabled = (RecordFill.Bits And &H40) > 0
            Case 2: func_ValveEnabled = (RecordFill.Bits And &H20) > 0
            Case 3: func_ValveEnabled = (RecordFill.Bits And &H80) > 0
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
            Case 18: func_ValveEnabled = (RecordFill.Bits And &H100) > 0

            Case Else
                func_ValveEnabled = False
        End Select
    End With
End Function

Public Function ValveEnabled(frm As FormMain, StepIndex As Integer, Num As Integer) As Boolean
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordFill As TYPE_WPC_FILL

    ValveEnabled = func_ValveEnabled(frm, 0&, RecordTitle, RecordFill, StepIndex, Num)
End Function