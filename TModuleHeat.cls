VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TModuleHeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const HEAT_PAUSE_FIELD = IDLE_FUNCTION_FIELD + 1
Private Const HEAT_ROTATION_FIELD = IDLE_FUNCTION_FIELD + 2

Private Const HEAT_TEMPERATURE_FIELD = IDLE_FUNCTION_FIELD + 3
Private Const HEAT_ROTTIME_FIELD = IDLE_FUNCTION_FIELD + 4
Private Const HEAT_PAUSETIME_FIELD = IDLE_FUNCTION_FIELD + 5
Private Const HEAT_DRUMSPEED_FIELD = IDLE_FUNCTION_FIELD + 6

Private Const HEAT_PARAMETERS_COUNT = IDLE_PARAMETERS_COUNT + 6

' Дискретные параметры
Private Const HEAT_PARAMETER_DESCR_PAUSE = "Разр. паузы"
Private Const HEAT_PARAMETER_DESCR_ROTATION = "Разр. вращ."

' Аналоговые параметры
Private Const HEAT_PARAMETER_DESCR_TEMPERATURE = "Температура"
Private Const HEAT_PARAMETER_DESCR_ROTTIME = "Время вращ. мотора"
Private Const HEAT_PARAMETER_DESCR_PAUSETIME = "Время паузы вращ. мот."
Private Const HEAT_PARAMETER_DESCR_DRUMSPEED = "Скорость вращ. барабана"

Public Sub ShowPropertyTableForHeat(frm As FormMain)
    Dim I, J, row As Integer
    Dim ParamStr, s As String
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordHeat As TYPE_WPC_HEAT
    
    With frm
        ' Формируем столбик с названиями параметров
        ParamStr = ";Параметр|"
        
        For I = 1 To HEAT_PARAMETERS_COUNT
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
                Case HEAT_PAUSE_FIELD:
                    ParamStr = ParamStr & HEAT_PARAMETER_DESCR_PAUSE
                    
                Case HEAT_ROTATION_FIELD:
                    ParamStr = ParamStr & HEAT_PARAMETER_DESCR_ROTATION
                
                Case HEAT_TEMPERATURE_FIELD:
                    ParamStr = ParamStr & HEAT_PARAMETER_DESCR_TEMPERATURE
                
                Case HEAT_ROTTIME_FIELD:
                    ParamStr = ParamStr & HEAT_PARAMETER_DESCR_ROTTIME
                
                Case HEAT_PAUSETIME_FIELD:
                    ParamStr = ParamStr & HEAT_PARAMETER_DESCR_PAUSETIME
                
                Case HEAT_DRUMSPEED_FIELD:
                    ParamStr = ParamStr & HEAT_PARAMETER_DESCR_DRUMSPEED
                    
                Case Else
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_UNKNOWN
            End Select
            
            If (I < HEAT_PARAMETERS_COUNT) Then ParamStr = ParamStr & "|"
        Next I
                
        ' Эта строка автоматически изменяет количество строк
        .PropertyTable.FormatString = ParamStr
        
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        CopyMemory RecordHeat, ByVal StepPointer, STEP_SIZE_IN_BYTES
        
        .PropertyTable.col = 1
        
        For I = 1 To HEAT_PARAMETERS_COUNT
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
                    .PropertyTable.Text = FunctionsStrings(RecordHeat.Bits And &HF)
                    
                ' Специальные параметры
                Case HEAT_PAUSE_FIELD:
                    If (RecordHeat.Bits And &H10) / &H10 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                Case HEAT_ROTATION_FIELD:
                    If (RecordHeat.Bits And &H20) / &H20 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If
                    
                Case HEAT_TEMPERATURE_FIELD:
                    .PropertyTable.Text = "" & RecordHeat.Temperature
                    
                Case HEAT_ROTTIME_FIELD:
                    .PropertyTable.Text = "" & RecordHeat.RotTime
                
                Case HEAT_PAUSETIME_FIELD:
                    .PropertyTable.Text = "" & RecordHeat.PauseTime
                
                Case HEAT_DRUMSPEED_FIELD:
                    .PropertyTable.Text = "" & RecordHeat.DrumSpeed
                
                Case Else
                    .PropertyTable.CellBackColor = &H8000000F
                    
            End Select
        Next I
    End With
End Sub

Public Sub EditPropertyForHeat(frm As FormMain)
    Dim I As Integer
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordHeat As TYPE_WPC_HEAT
    
    With frm
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        CopyMemory RecordHeat, ByVal StepPointer, STEP_SIZE_IN_BYTES
        
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
                .ComboCell(0).ListIndex = RecordHeat.Bits And &HF
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            ' Специальные параметры шага
            Case HEAT_PAUSE_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordHeat.Bits And &H10) / &H10
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus

            Case HEAT_ROTATION_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordHeat.Bits And &H20) / &H20
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus

            Case HEAT_TEMPERATURE_FIELD:
                .TextCell(0).Text = "" & RecordHeat.Temperature
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus

            Case HEAT_ROTTIME_FIELD:
                .TextCell(0).Text = "" & RecordHeat.RotTime
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case HEAT_PAUSETIME_FIELD:
                .TextCell(0).Text = "" & RecordHeat.PauseTime
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case HEAT_DRUMSPEED_FIELD:
                .TextCell(0).Text = "" & RecordHeat.DrumSpeed
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case Else

        End Select
    
        .TextCell(0).SelStart = 0
        .TextCell(0).SelLength = Len(.TextCell(0).Text)
    End With
End Sub

Public Sub SetComboPropertyForHeat(frm As FormMain)
    Dim I As Integer
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordHeat As TYPE_WPC_HEAT
    
    With frm
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        CopyMemory RecordHeat, ByVal StepPointer, STEP_SIZE_IN_BYTES
    
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
                ZeroMemory RecordHeat, STEP_SIZE_IN_BYTES
                ' и записать новое значение в поле типа функции
                RecordHeat.Bits = .ComboCell(0).ListIndex And &HF
                
            ' Специальные параметры шага
            Case HEAT_PAUSE_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordHeat.Bits = RecordHeat.Bits And &HFFEF
                    Case 1: RecordHeat.Bits = RecordHeat.Bits Or &H10
                End Select
    
            Case HEAT_ROTATION_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordHeat.Bits = RecordHeat.Bits And &HFFDF
                    Case 1: RecordHeat.Bits = RecordHeat.Bits Or &H20
                End Select

            Case HEAT_TEMPERATURE_FIELD:
                RecordHeat.Temperature = Val(.TextCell(0).Text)
    
            Case HEAT_ROTTIME_FIELD:
                RecordHeat.RotTime = Val(.TextCell(0).Text)
            
            Case HEAT_PAUSETIME_FIELD:
                RecordHeat.PauseTime = Val(.TextCell(0).Text)
            
            Case HEAT_DRUMSPEED_FIELD:
                RecordHeat.DrumSpeed = Val(.TextCell(0).Text)
            
            Case Else
    
        End Select
        
        .SetModified True
        
        ' Сохраняем изменения
        CopyMemory ByVal StepPointer, RecordHeat, STEP_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory ByVal StepPointer, RecordTitle, HEADER_SIZE_IN_BYTES
    End With
End Sub


