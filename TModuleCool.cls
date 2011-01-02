VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TModuleCool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const COOL_PAUSE_FIELD = IDLE_FUNCTION_FIELD + 1
Private Const COOL_FAST_FIELD = IDLE_FUNCTION_FIELD + 2
Private Const COOL_ROTATION_FIELD = IDLE_FUNCTION_FIELD + 3

Private Const COOL_TEMPERATURE_FIELD = IDLE_FUNCTION_FIELD + 4
Private Const COOL_COLDWATERTIME_FIELD = IDLE_FUNCTION_FIELD + 5
Private Const COOL_ROTTIME_FIELD = IDLE_FUNCTION_FIELD + 6
Private Const COOL_PAUSETIME_FIELD = IDLE_FUNCTION_FIELD + 7
Private Const COOL_DRUMSPEED_FIELD = IDLE_FUNCTION_FIELD + 8

Private Const COOL_PARAMETERS_COUNT = IDLE_PARAMETERS_COUNT + 8

' Дискретные параметры
Private Const COOL_PARAMETER_DESCR_PAUSE = "Разр. паузы"
Private Const COOL_PARAMETER_DESCR_FAST = "Разр. быстр. охл."
Private Const COOL_PARAMETER_DESCR_ROTATION = "Разр. вращ."

' Аналоговые параметры
Private Const COOL_PARAMETER_DESCR_TEMPERATURE = "Температура"
Private Const COOL_PARAMETER_DESCR_COLDWATERTIME = "Время откр. кл. ХВ"
Private Const COOL_PARAMETER_DESCR_ROTTIME = "Время вращ. мотора"
Private Const COOL_PARAMETER_DESCR_PAUSETIME = "Время паузы вращ. мот."
Private Const COOL_PARAMETER_DESCR_DRUMSPEED = "Скорость вращ. барабана"

Public Sub ShowPropertyTableForCool(frm As FormMain)
    Dim I, J, row As Integer
    Dim ParamStr, s As String
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordCool As TYPE_WPC_COOL
    
    With frm
        ' Формируем столбик с названиями параметров
        ParamStr = ";Параметр|"
        
        For I = 1 To COOL_PARAMETERS_COUNT
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
                Case COOL_PAUSE_FIELD:
                    ParamStr = ParamStr & COOL_PARAMETER_DESCR_PAUSE
                    
                Case COOL_FAST_FIELD:
                    ParamStr = ParamStr & COOL_PARAMETER_DESCR_FAST
                
                Case COOL_ROTATION_FIELD:
                    ParamStr = ParamStr & COOL_PARAMETER_DESCR_ROTATION
                
                Case COOL_TEMPERATURE_FIELD:
                    ParamStr = ParamStr & COOL_PARAMETER_DESCR_TEMPERATURE
                
                Case COOL_COLDWATERTIME_FIELD:
                    ParamStr = ParamStr & COOL_PARAMETER_DESCR_COLDWATERTIME
                
                Case COOL_ROTTIME_FIELD:
                    ParamStr = ParamStr & COOL_PARAMETER_DESCR_ROTTIME
                
                Case COOL_PAUSETIME_FIELD:
                    ParamStr = ParamStr & COOL_PARAMETER_DESCR_PAUSETIME
                
                Case COOL_DRUMSPEED_FIELD:
                    ParamStr = ParamStr & COOL_PARAMETER_DESCR_DRUMSPEED
                    
                Case Else
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_UNKNOWN
            End Select
            
            If (I < COOL_PARAMETERS_COUNT) Then ParamStr = ParamStr & "|"
        Next I
                
        ' Эта строка автоматически изменяет количество строк
        .PropertyTable.FormatString = ParamStr
        
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        CopyMemory RecordCool, ByVal StepPointer, STEP_SIZE_IN_BYTES
        
        .PropertyTable.col = 1
        For I = 1 To COOL_PARAMETERS_COUNT
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
                    .PropertyTable.Text = FunctionsStrings(RecordCool.Bits And &HF)
                    
                ' Специальные параметры
                Case COOL_PAUSE_FIELD:
                    If (RecordCool.Bits And &H10) / &H10 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                Case COOL_FAST_FIELD:
                    If (RecordCool.Bits And &H20) / &H20 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If
                    
                Case COOL_ROTATION_FIELD:
                    If (RecordCool.Bits And &H40) / &H40 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If
                    
                Case COOL_TEMPERATURE_FIELD:
                    .PropertyTable.Text = "" & RecordCool.Temperature
                    
                Case COOL_COLDWATERTIME_FIELD:
                    .PropertyTable.Text = "" & RecordCool.ColdWaterTime
                    
                Case COOL_ROTTIME_FIELD:
                    .PropertyTable.Text = "" & RecordCool.RotTime
                
                Case COOL_PAUSETIME_FIELD:
                    .PropertyTable.Text = "" & RecordCool.PauseTime
                
                Case COOL_DRUMSPEED_FIELD:
                    .PropertyTable.Text = "" & RecordCool.DrumSpeed
                
                Case Else
                    .PropertyTable.CellBackColor = &H8000000F
                    
            End Select
        Next I
    End With
End Sub

Public Sub EditPropertyForCool(frm As FormMain)
    Dim I As Integer
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordCool As TYPE_WPC_COOL
    
    With frm
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        CopyMemory RecordCool, ByVal StepPointer, STEP_SIZE_IN_BYTES
        
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
                .ComboCell(0).ListIndex = RecordCool.Bits And &HF
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            ' Специальные параметры шага
            Case COOL_PAUSE_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordCool.Bits And &H10) / &H10
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case COOL_FAST_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordCool.Bits And &H20) / &H20
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case COOL_ROTATION_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordCool.Bits And &H40) / &H40
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case COOL_TEMPERATURE_FIELD:
                .TextCell(0).Text = "" & RecordCool.Temperature
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus

            Case COOL_COLDWATERTIME_FIELD:
                .TextCell(0).Text = "" & RecordCool.ColdWaterTime
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
                
            Case COOL_ROTTIME_FIELD:
                .TextCell(0).Text = "" & RecordCool.RotTime
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
                
            Case COOL_PAUSETIME_FIELD:
                .TextCell(0).Text = "" & RecordCool.PauseTime
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
                
            Case COOL_DRUMSPEED_FIELD:
                .TextCell(0).Text = "" & RecordCool.DrumSpeed
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
                
            Case Else

        End Select
    
        .TextCell(0).SelStart = 0
        .TextCell(0).SelLength = Len(.TextCell(0).Text)
    End With
End Sub

Public Sub SetComboPropertyForCool(frm As FormMain)
    Dim I As Integer
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordCool As TYPE_WPC_COOL
    
    With frm
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        CopyMemory RecordCool, ByVal StepPointer, STEP_SIZE_IN_BYTES
    
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
                ZeroMemory RecordCool, STEP_SIZE_IN_BYTES
                ' и записать новое значение в поле типа функции
                RecordCool.Bits = .ComboCell(0).ListIndex And &HF
                
            ' Специальные параметры шага
            Case COOL_PAUSE_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordCool.Bits = RecordCool.Bits And &HFFEF
                    Case 1: RecordCool.Bits = RecordCool.Bits Or &H10
                End Select
                
            Case COOL_FAST_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordCool.Bits = RecordCool.Bits And &HFFDF
                    Case 1: RecordCool.Bits = RecordCool.Bits Or &H20
                End Select
                
            Case COOL_ROTATION_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordCool.Bits = RecordCool.Bits And &HFFBF
                    Case 1: RecordCool.Bits = RecordCool.Bits Or &H40
                End Select
                
            Case COOL_TEMPERATURE_FIELD:
                RecordCool.Temperature = Val(.TextCell(0).Text)
            
            Case COOL_COLDWATERTIME_FIELD:
                RecordCool.ColdWaterTime = Val(.TextCell(0).Text)
    
            Case COOL_ROTTIME_FIELD:
                RecordCool.RotTime = Val(.TextCell(0).Text)
            
            Case COOL_PAUSETIME_FIELD:
                RecordCool.PauseTime = Val(.TextCell(0).Text)
            
            Case COOL_DRUMSPEED_FIELD:
                RecordCool.DrumSpeed = Val(.TextCell(0).Text)
            
            Case Else
    
        End Select
        
        .SetModified True
        
        ' Сохраняем изменения
        CopyMemory ByVal StepPointer, RecordCool, STEP_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory ByVal StepPointer, RecordTitle, HEADER_SIZE_IN_BYTES
    End With
End Sub


