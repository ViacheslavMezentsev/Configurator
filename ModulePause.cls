VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CModulePause"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const PAUSE_PAUSE_FIELD = IDLE_FUNCTION_FIELD + 1
Private Const PAUSE_ROTATION_FIELD = IDLE_FUNCTION_FIELD + 2

Private Const PAUSE_TIME_FIELD = IDLE_FUNCTION_FIELD + 3
Private Const PAUSE_RotationTime_FIELD = IDLE_FUNCTION_FIELD + 4
Private Const PAUSE_PAUSETIME_FIELD = IDLE_FUNCTION_FIELD + 5
Private Const PAUSE_DRUMSPEED_FIELD = IDLE_FUNCTION_FIELD + 6

Private Const PAUSE_PARAMETERS_COUNT = IDLE_PARAMETERS_COUNT + 6

' Дискретные параметры
Private Const PAUSE_PARAMETER_DESCR_PAUSE = "Разр. паузы"
Private Const PAUSE_PARAMETER_DESCR_ROTATION = "Разр. вращ."

' Аналоговые параметры
Private Const PAUSE_PARAMETER_DESCR_TIME = "Время"
Private Const PAUSE_PARAMETER_DESCR_RotationTime = "Время вращ. мотора"
Private Const PAUSE_PARAMETER_DESCR_PAUSETIME = "Время паузы вращ. мот."
Private Const PAUSE_PARAMETER_DESCR_DRUMSPEED = "Скорость вращ. барабана"

' Функция-посредник
Private Sub func_ShowPropertyTableForPause(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordStep As TYPE_WPC_STEP)

    Dim I, J, row As Integer
    Dim ParamStr, s As String
    Dim StepPointer As Long

    With frm
        ' Формируем столбик с названиями параметров
        ParamStr = ";Параметр|"

        For I = 1 To PAUSE_PARAMETERS_COUNT
            Select Case I
                ' Общие параметры
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

                ' Специальные параметры

                Case Else
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_UNKNOWN
            End Select

            If (I < PAUSE_PARAMETERS_COUNT) Then ParamStr = ParamStr & "|"
        Next I

        ' Эта строка автоматически изменяет количество строк
        .PropertyTable.FormatString = ParamStr

        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        'CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES

        'CopyMemory RecordStep, ByVal StepPointer, STEP_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer

        .PropertyTable.col = 1

        For I = 1 To PAUSE_PARAMETERS_COUNT
            .PropertyTable.row = I
            .PropertyTable.CellAlignment = flexAlignRightCenter

            Select Case I
                ' Общие параметры
                Case IDLE_ENDSOUND_FIELD:
                    If (RecordTitle.LowBits And &H1) Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                Case IDLE_DOORUNLOCK_FIELD:
                    If (RecordTitle.LowBits And &H2) / &H2 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                Case IDLE_PROGNAME_FIELD:
                    s = ""
                    For J = 1 To PROG_NAME_LENGTH - 1
                        s = s & Chr(CLng(RecordTitle.ProgName(J)))
                    Next J
                    .PropertyTable.Text = s

                Case IDLE_STEP_FIELD:
                    .PropertyTable.CellFontBold = True
                    .PropertyTable.Text = "" & .Manager.StepIndex + 1

                Case IDLE_FUNCTION_FIELD:
                    .PropertyTable.CellFontBold = True
                    .PropertyTable.Text = FunctionsStrings(RecordStep.Bits And &HF)

                ' Специальные параметры

                Case Else
                    .PropertyTable.CellBackColor = &H8000000F

            End Select
        Next I
    End With
End Sub

Public Sub ShowPropertyTableForPause(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP

    func_ShowPropertyTableForPause frm, 0&, RecordTitle, RecordStep
End Sub

' Функция-посредник
Private Sub func_EditPropertyForPause(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordStep As TYPE_WPC_STEP)

    Dim I As Integer
    Dim StepPointer As Long
    Dim s As String

    With frm
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        'CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES

        'CopyMemory RecordStep, ByVal StepPointer, STEP_SIZE_IN_BYTES
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
            ' Общие параметры программы и шага
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
                s = ""
                For I = 1 To PROG_NAME_LENGTH - 1
                    s = s & Chr(CLng(RecordTitle.ProgName(I)))
                Next I
                .TextCell.Text = s
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case IDLE_STEP_FIELD:
                For I = 1 To MAX_NUMBER_OF_STEPS
                    .ComboCell.AddItem ("Шаг " & I)
                Next I
                .ComboCell.ListIndex = .Manager.StepIndex
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case IDLE_FUNCTION_FIELD:
                For I = 0 To 11
                    .ComboCell.AddItem (FunctionsStrings(I))
                Next I
                .ComboCell.ListIndex = RecordStep.Bits And &HF
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            ' Специальные параметры шага

            Case Else

        End Select

        .TextCell.SelStart = 0
        .TextCell.SelLength = Len(.TextCell.Text)
    End With
End Sub

Public Sub EditPropertyForPause(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP

    func_EditPropertyForPause frm, 0&, RecordTitle, RecordStep
End Sub

' Функция-посредник
Private Sub func_SetComboPropertyForPause(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordStep As TYPE_WPC_STEP)

    On Error GoTo ErrorHandler

    Dim I As Integer
    Dim StepPointer As Long

    With frm
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        'CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES

        'CopyMemory RecordStep, ByVal StepPointer, STEP_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer

        Select Case .PropertyTable.row
            ' Общие параметры программы и шага
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
                        RecordTitle.ProgName(I) = Asc(Mid(.TextCell.Text, I, 1))
                    Else
                        RecordTitle.ProgName(I) = 0
                    End If
                Next I
                RecordTitle.ProgName(PROG_NAME_LENGTH) = 0

            Case IDLE_STEP_FIELD:
                .Manager.StepIndex = .ComboCell.ListIndex
                Exit Sub

            Case IDLE_FUNCTION_FIELD:
                ' При выборе новой функции для шага мы должны обнулить структуру
                ZeroMemory RecordStep, STEP_SIZE_IN_BYTES
                ' и записать новое значение в поле типа функции
                RecordStep.Bits = .ComboCell.ListIndex And &HF

            ' Специальные параметры шага

            Case Else

        End Select

        .SetModified True

        ' Сохраняем изменения
        'CopyMemory ByVal StepPointer, RecordStep, STEP_SIZE_IN_BYTES

        'StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        'CopyMemory ByVal StepPointer, RecordTitle, HEADER_SIZE_IN_BYTES
    End With
    Exit Sub

ErrorHandler:
    Err.Clear
End Sub

Public Sub SetComboPropertyForPause(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP

    func_SetComboPropertyForPause frm, 0&, RecordTitle, RecordStep
End Sub
