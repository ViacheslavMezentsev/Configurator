VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TModuleDrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const DRAIN_PAUSE_FIELD = IDLE_FUNCTION_FIELD + 1
Private Const DRAIN_DRAINGATE1_FIELD = IDLE_FUNCTION_FIELD + 2
Private Const DRAIN_DRAINGATE2_FIELD = IDLE_FUNCTION_FIELD + 3
Private Const DRAIN_ROTATION_FIELD = IDLE_FUNCTION_FIELD + 4

Private Const DRAIN_LEVEL_FIELD = IDLE_FUNCTION_FIELD + 5
Private Const DRAIN_ROTTIME_FIELD = IDLE_FUNCTION_FIELD + 6
Private Const DRAIN_PAUSETIME_FIELD = IDLE_FUNCTION_FIELD + 7
Private Const DRAIN_DRUMSPEED1_FIELD = IDLE_FUNCTION_FIELD + 8
Private Const DRAIN_TIME2_FIELD = IDLE_FUNCTION_FIELD + 9
Private Const DRAIN_DRUMSPEED2_FIELD = IDLE_FUNCTION_FIELD + 10

Private Const DRAIN_PARAMETERS_COUNT = IDLE_PARAMETERS_COUNT + 10

' Дискретные параметры
Private Const DRAIN_PARAMETER_DESCR_PAUSE = "Разр. паузы"
Private Const DRAIN_PARAMETER_DESCR_DRAINGATE1 = "Разр. кл. слива 1"
Private Const DRAIN_PARAMETER_DESCR_DRAINGATE2 = "Разр. кл. слива 2"
Private Const DRAIN_PARAMETER_DESCR_ROTATION = "Разр. вращ."

' Аналоговые параметры
Private Const DRAIN_PARAMETER_DESCR_LEVEL = "Уровень моющ. раств."
Private Const DRAIN_PARAMETER_DESCR_ROTTIME = "Время вращ. мотора"
Private Const DRAIN_PARAMETER_DESCR_PAUSETIME = "Время паузы вращ. мот."
Private Const DRAIN_PARAMETER_DESCR_DRUMSPEED1 = "Скорость вращ. бар. (реверс)"
Private Const DRAIN_PARAMETER_DESCR_TIME2 = "Время раскладки"
Private Const DRAIN_PARAMETER_DESCR_DRUMSPEED2 = "Скорость вращ. бар. (раскл.)"

Public Sub ShowPropertyTableForDrain(frm As FormMain)
    Dim I, J, row As Integer
    Dim ParamStr, s As String
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordDrain As TYPE_WPC_DRAIN

    With frm
        ' Формируем столбик с названиями параметров
        ParamStr = ";Параметр|"
        
        For I = 1 To DRAIN_PARAMETERS_COUNT
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
                Case DRAIN_PAUSE_FIELD:
                    ParamStr = ParamStr & DRAIN_PARAMETER_DESCR_PAUSE
                    
                Case DRAIN_DRAINGATE1_FIELD:
                    ParamStr = ParamStr & DRAIN_PARAMETER_DESCR_DRAINGATE1
                    
                Case DRAIN_DRAINGATE2_FIELD:
                    ParamStr = ParamStr & DRAIN_PARAMETER_DESCR_DRAINGATE2
                    
                Case DRAIN_ROTATION_FIELD:
                    ParamStr = ParamStr & DRAIN_PARAMETER_DESCR_ROTATION
                
                Case DRAIN_LEVEL_FIELD:
                    ParamStr = ParamStr & DRAIN_PARAMETER_DESCR_LEVEL
                
                Case DRAIN_ROTTIME_FIELD:
                    ParamStr = ParamStr & DRAIN_PARAMETER_DESCR_ROTTIME
                
                Case DRAIN_PAUSETIME_FIELD:
                    ParamStr = ParamStr & DRAIN_PARAMETER_DESCR_PAUSETIME
                
                Case DRAIN_DRUMSPEED1_FIELD:
                    ParamStr = ParamStr & DRAIN_PARAMETER_DESCR_DRUMSPEED1
                    
                Case DRAIN_TIME2_FIELD:
                    ParamStr = ParamStr & DRAIN_PARAMETER_DESCR_TIME2
                    
                Case DRAIN_DRUMSPEED2_FIELD:
                    ParamStr = ParamStr & DRAIN_PARAMETER_DESCR_DRUMSPEED2
                    
                Case Else
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_UNKNOWN
            End Select
            
            If (I < DRAIN_PARAMETERS_COUNT) Then ParamStr = ParamStr & "|"
        Next I
                
        ' Эта строка автоматически изменяет количество строк
        .PropertyTable.FormatString = ParamStr
        
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES

        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        CopyMemory RecordDrain, ByVal StepPointer, STEP_SIZE_IN_BYTES
        
        .PropertyTable.col = 1
        
        For I = 1 To DRAIN_PARAMETERS_COUNT
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
                    .PropertyTable.Text = FunctionsStrings(RecordDrain.Bits And &HF)
                    
                ' Специальные параметры
                Case DRAIN_PAUSE_FIELD:
                    If (RecordDrain.Bits And &H10) / &H10 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                Case DRAIN_DRAINGATE1_FIELD:
                    If (RecordDrain.Bits And &H20) / &H20 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If
                    
                Case DRAIN_DRAINGATE2_FIELD:
                    If (RecordDrain.Bits And &H40) / &H40 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If
                    
                Case DRAIN_ROTATION_FIELD:
                    If (RecordDrain.Bits And &H80) / &H80 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If
                    
                Case DRAIN_LEVEL_FIELD:
                    .PropertyTable.Text = "" & RecordDrain.Level
                    
                Case DRAIN_ROTTIME_FIELD:
                    .PropertyTable.Text = "" & RecordDrain.RotTime
                
                Case DRAIN_PAUSETIME_FIELD:
                    .PropertyTable.Text = "" & RecordDrain.PauseTime
                
                Case DRAIN_DRUMSPEED1_FIELD:
                    .PropertyTable.Text = "" & RecordDrain.DrumSpeed1
                
                Case DRAIN_TIME2_FIELD:
                    .PropertyTable.Text = "" & RecordDrain.Time2
                
                Case DRAIN_DRUMSPEED2_FIELD:
                    .PropertyTable.Text = "" & _
                        (RecordDrain.HiByteDrumSpeed2 * &H100 + _
                        RecordDrain.LoByteDrumSpeed2)
                
                Case Else
                    .PropertyTable.CellBackColor = &H8000000F
                    
            End Select
        Next I
    End With
End Sub

Public Sub EditPropertyForDrain(frm As FormMain)
    Dim I As Integer
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordDrain As TYPE_WPC_DRAIN
    
    With frm
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        CopyMemory RecordDrain, ByVal StepPointer, STEP_SIZE_IN_BYTES
        
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
                .ComboCell(0).ListIndex = RecordDrain.Bits And &HF
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            ' Специальные параметры шага
            Case DRAIN_PAUSE_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordDrain.Bits And &H10) / &H10
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus

            Case DRAIN_DRAINGATE1_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordDrain.Bits And &H20) / &H20
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case DRAIN_DRAINGATE2_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordDrain.Bits And &H40) / &H40
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case DRAIN_ROTATION_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordDrain.Bits And &H80) / &H80
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case DRAIN_LEVEL_FIELD:
                .TextCell(0).Text = "" & RecordDrain.Level
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case DRAIN_ROTTIME_FIELD:
                .TextCell(0).Text = "" & RecordDrain.RotTime
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case DRAIN_PAUSETIME_FIELD:
                .TextCell(0).Text = "" & RecordDrain.PauseTime
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case DRAIN_DRUMSPEED1_FIELD:
                .TextCell(0).Text = "" & RecordDrain.DrumSpeed1
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case DRAIN_TIME2_FIELD:
                .TextCell(0).Text = "" & RecordDrain.Time2
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case DRAIN_DRUMSPEED2_FIELD:
                .TextCell(0).Text = "" & _
                    (RecordDrain.HiByteDrumSpeed2 * &H100 + _
                    RecordDrain.LoByteDrumSpeed2)
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case Else

        End Select
    
        .TextCell(0).SelStart = 0
        .TextCell(0).SelLength = Len(.TextCell(0).Text)
    End With
End Sub

Public Sub SetComboPropertyForDrain(frm As FormMain)
    Dim I As Integer
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordDrain As TYPE_WPC_DRAIN
    
    With frm
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        CopyMemory RecordDrain, ByVal StepPointer, STEP_SIZE_IN_BYTES
    
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
                ZeroMemory RecordDrain, STEP_SIZE_IN_BYTES
                ' и записать новое значение в поле типа функции
                RecordDrain.Bits = .ComboCell(0).ListIndex And &HF
                
            ' Специальные параметры шага
            Case DRAIN_PAUSE_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordDrain.Bits = RecordDrain.Bits And &HFFEF
                    Case 1: RecordDrain.Bits = RecordDrain.Bits Or &H10
                End Select
    
            Case DRAIN_DRAINGATE1_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordDrain.Bits = RecordDrain.Bits And &HFFDF
                    Case 1: RecordDrain.Bits = RecordDrain.Bits Or &H20
                End Select
            
            Case DRAIN_DRAINGATE2_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordDrain.Bits = RecordDrain.Bits And &HFFBF
                    Case 1: RecordDrain.Bits = RecordDrain.Bits Or &H40
                End Select
            
            Case DRAIN_ROTATION_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordDrain.Bits = RecordDrain.Bits And &HFF7F
                    Case 1: RecordDrain.Bits = RecordDrain.Bits Or &H80
                End Select

            Case DRAIN_LEVEL_FIELD:
                RecordDrain.Level = Val(.TextCell(0).Text)
    
            Case DRAIN_ROTTIME_FIELD:
                RecordDrain.RotTime = Val(.TextCell(0).Text)
            
            Case DRAIN_PAUSETIME_FIELD:
                RecordDrain.PauseTime = Val(.TextCell(0).Text)
            
            Case DRAIN_DRUMSPEED1_FIELD:
                RecordDrain.DrumSpeed1 = Val(.TextCell(0).Text)

            Case DRAIN_TIME2_FIELD:
                RecordDrain.Time2 = Val(.TextCell(0).Text)

            Case DRAIN_DRUMSPEED2_FIELD:
                Dim v As Integer
                v = CInt(Val(.TextCell(0).Text))
                RecordDrain.LoByteDrumSpeed2 = (v And &HFF)
                RecordDrain.HiByteDrumSpeed2 = (v And &HFF00) / &H100
            
            Case Else
    
        End Select
        
        .SetModified True
        
        ' Сохраняем изменения
        CopyMemory ByVal StepPointer, RecordDrain, STEP_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory ByVal StepPointer, RecordTitle, HEADER_SIZE_IN_BYTES
    End With
End Sub


