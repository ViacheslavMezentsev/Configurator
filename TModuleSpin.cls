VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TModuleSpin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const SPIN_PAUSE_FIELD = IDLE_FUNCTION_FIELD + 1
Private Const SPIN_DRAINGATE1_FIELD = IDLE_FUNCTION_FIELD + 2
Private Const SPIN_DRAINGATE2_FIELD = IDLE_FUNCTION_FIELD + 3

Private Const SPIN_DRUMSPEED_FIELD = IDLE_FUNCTION_FIELD + 4
Private Const SPIN_TIME_FIELD = IDLE_FUNCTION_FIELD + 5

Private Const SPIN_PARAMETERS_COUNT = IDLE_PARAMETERS_COUNT + 5

' Дискретные параметры
Private Const SPIN_PARAMETER_DESCR_PAUSE = "Разр. паузы"
Private Const SPIN_PARAMETER_DESCR_DRAINGATE1 = "Разр. кл. слива 1"
Private Const SPIN_PARAMETER_DESCR_DRAINGATE2 = "Разр. кл. слива 2"

' Аналоговые параметры
Private Const SPIN_PARAMETER_DESCR_DRUMSPEED = "Скорость вращ. барабана"
Private Const SPIN_PARAMETER_DESCR_TIME = "Время отжима"

' Функция-посредник
Private Sub func_ShowPropertyTableForSpin(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordSpin As TYPE_WPC_SPIN)
    
    Dim I, J, row As Integer
    Dim ParamStr, s As String
    Dim StepPointer As Long
    
    With frm
        ' Формируем столбик с названиями параметров
        ParamStr = ";Параметр|"
        
        For I = 1 To SPIN_PARAMETERS_COUNT
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
                Case SPIN_PAUSE_FIELD:
                    ParamStr = ParamStr & SPIN_PARAMETER_DESCR_PAUSE
                    
                Case SPIN_DRAINGATE1_FIELD:
                    ParamStr = ParamStr & SPIN_PARAMETER_DESCR_DRAINGATE1
                    
                Case SPIN_DRAINGATE2_FIELD:
                    ParamStr = ParamStr & SPIN_PARAMETER_DESCR_DRAINGATE2
                    
                Case SPIN_DRUMSPEED_FIELD:
                    ParamStr = ParamStr & SPIN_PARAMETER_DESCR_DRUMSPEED
                
                Case SPIN_TIME_FIELD:
                    ParamStr = ParamStr & SPIN_PARAMETER_DESCR_TIME
                    
                Case Else
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_UNKNOWN
            End Select
            
            If (I < SPIN_PARAMETERS_COUNT) Then ParamStr = ParamStr & "|"
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
        
        'CopyMemory RecordSpin, ByVal StepPointer, STEP_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer
        
        .PropertyTable.col = 1
        
        For I = 1 To SPIN_PARAMETERS_COUNT
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
                    
                ' TODO: Доделать отображение имени программы
                Case IDLE_PROGNAME_FIELD:
                    s = ""
                    For J = 1 To PROG_NAME_LENGTH - 1
                        'if Chr(RecordStep.Reserved(J))
                        s = s & "*"
                    Next J
                    .PropertyTable.Text = s
                
                Case IDLE_STEP_FIELD:
                    .PropertyTable.CellFontBold = True
                    .PropertyTable.Text = "" & .Manager.StepIndex + 1
                
                Case IDLE_FUNCTION_FIELD:
                    .PropertyTable.CellFontBold = True
                    .PropertyTable.Text = FunctionsStrings(RecordSpin.Bits And &HF)
                    
                ' Специальные параметры
                Case SPIN_PAUSE_FIELD:
                    If (RecordSpin.Bits And &H10) / &H10 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If
                    
                Case SPIN_DRAINGATE1_FIELD:
                    If (RecordSpin.Bits And &H20) / &H20 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If
                    
                Case SPIN_DRAINGATE2_FIELD:
                    If (RecordSpin.Bits And &H40) / &H40 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If
                    
                Case SPIN_DRUMSPEED_FIELD:
                    .PropertyTable.Text = "" & RecordSpin.DrumSpeed
                
                Case SPIN_TIME_FIELD:
                    .PropertyTable.Text = "" & RecordSpin.Time
                    
                Case Else
                    .PropertyTable.CellBackColor = &H8000000F
                    
            End Select
        Next I
    End With
End Sub

Public Sub ShowPropertyTableForSpin(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordSpin As TYPE_WPC_SPIN
    
    func_ShowPropertyTableForSpin frm, 0&, RecordTitle, RecordSpin
End Sub

' Функция-посредник
Private Sub func_EditPropertyForSpin(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordSpin As TYPE_WPC_SPIN)
    
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
        
        'CopyMemory RecordSpin, ByVal StepPointer, STEP_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer
        
        .ComboCell(0).Left = .PropertyTable.Left + .PropertyTable.CellLeft
        .ComboCell(0).Top = .PropertyTable.Top + .PropertyTable.CellTop
        .ComboCell(0).Width = .PropertyTable.CellWidth
        .ComboCell(0).Clear
        
        .TextCell(0).Left = .PropertyTable.Left + .PropertyTable.CellLeft
        .TextCell(0).Top = .PropertyTable.Top + .PropertyTable.CellTop
        .TextCell(0).Width = .PropertyTable.CellWidth
        .TextCell(0).Height = .PropertyTable.CellHeight
        
        Select Case .PropertyTable.row
            ' Общие параметры программы и шага
            Case IDLE_ENDSOUND_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = RecordTitle.LowBits And &H1
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
            
            Case IDLE_DOORUNLOCK_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordTitle.LowBits And &H2) / &H2
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case IDLE_PROGNAME_FIELD:
                .TextCell(0).Text = .PropertyTable.Text
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case IDLE_STEP_FIELD:
                For I = 1 To MAX_NUMBER_OF_STEPS
                    .ComboCell(0).AddItem ("Шаг " & I)
                Next I
                .ComboCell(0).ListIndex = .Manager.StepIndex
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case IDLE_FUNCTION_FIELD:
                For I = 0 To 11
                    .ComboCell(0).AddItem (FunctionsStrings(I))
                Next I
                .ComboCell(0).ListIndex = RecordSpin.Bits And &HF
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            ' Специальные параметры шага
            Case SPIN_PAUSE_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordSpin.Bits And &H10) / &H10
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case SPIN_DRAINGATE1_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordSpin.Bits And &H20) / &H20
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case SPIN_DRAINGATE2_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordSpin.Bits And &H40) / &H40
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case SPIN_DRUMSPEED_FIELD:
                .TextCell(0).Text = "" & RecordSpin.DrumSpeed
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case SPIN_TIME_FIELD:
                .TextCell(0).Text = "" & RecordSpin.Time
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case Else

        End Select
    
        .TextCell(0).SelStart = 0
        .TextCell(0).SelLength = Len(.TextCell(0).Text)
    End With
End Sub

Public Sub EditPropertyForSpin(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordSpin As TYPE_WPC_SPIN
    
    func_EditPropertyForSpin frm, 0&, RecordTitle, RecordSpin
End Sub

' Функция-посредник
Private Sub func_SetComboPropertyForSpin(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordSpin As TYPE_WPC_SPIN)
    
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
        
        'CopyMemory RecordSpin, ByVal StepPointer, STEP_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer
    
        Select Case .PropertyTable.row
            ' Общие параметры программы и шага
            Case IDLE_ENDSOUND_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordTitle.LowBits = RecordTitle.LowBits And &HFFFE
                    Case 1: RecordTitle.LowBits = RecordTitle.LowBits Or &H1
                End Select
            
            Case IDLE_DOORUNLOCK_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordTitle.LowBits = RecordTitle.LowBits And &HFFFD
                    Case 1: RecordTitle.LowBits = RecordTitle.LowBits Or &H2
                End Select
                
            Case IDLE_PROGNAME_FIELD:
                For I = 1 To PROG_NAME_LENGTH - 1
                    If I < Len(.TextCell(0).Text) Then
                        RecordTitle.ProgName(I) = Asc(Mid(.TextCell(0).Text, I, 1))
                    Else
                        RecordTitle.ProgName(I) = 0
                    End If
                Next I
                RecordTitle.ProgName(PROG_NAME_LENGTH) = 0
            
            Case IDLE_STEP_FIELD:
                .Manager.StepIndex = .ComboCell(0).ListIndex
                Exit Sub
                
            Case IDLE_FUNCTION_FIELD:
                ' При выборе новой функции для шага мы должны обнулить структуру
                ZeroMemory RecordSpin, STEP_SIZE_IN_BYTES
                ' и записать новое значение в поле типа функции
                RecordSpin.Bits = .ComboCell(0).ListIndex And &HF
                
            ' Специальные параметры шага
            Case SPIN_PAUSE_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordSpin.Bits = RecordSpin.Bits And &HFFEF
                    Case 1: RecordSpin.Bits = RecordSpin.Bits Or &H10
                End Select
                
            Case SPIN_DRAINGATE1_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordSpin.Bits = RecordSpin.Bits And &HFFDF
                    Case 1: RecordSpin.Bits = RecordSpin.Bits Or &H20
                End Select
                
            Case SPIN_DRAINGATE2_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordSpin.Bits = RecordSpin.Bits And &HFFBF
                    Case 1: RecordSpin.Bits = RecordSpin.Bits Or &H40
                End Select
                
            Case SPIN_DRUMSPEED_FIELD:
                RecordSpin.DrumSpeed = Val(.TextCell(0).Text)
                
            Case SPIN_TIME_FIELD:
                RecordSpin.Time = Val(.TextCell(0).Text)
                
            Case Else
    
        End Select
        
        .SetModified True
        
        ' Сохраняем изменения
        'CopyMemory ByVal StepPointer, RecordSpin, STEP_SIZE_IN_BYTES
        
        'StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        'CopyMemory ByVal StepPointer, RecordTitle, HEADER_SIZE_IN_BYTES
    End With
End Sub

Public Sub SetComboPropertyForSpin(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordSpin As TYPE_WPC_SPIN
    
    func_SetComboPropertyForSpin frm, 0&, RecordTitle, RecordSpin
End Sub


