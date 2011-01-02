VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TModuleFill"
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
Private Const FILL_ROTTIME_FIELD = IDLE_FUNCTION_FIELD + 7
Private Const FILL_PAUSETIME_FIELD = IDLE_FUNCTION_FIELD + 8
Private Const FILL_DRUMSPEED_FIELD = IDLE_FUNCTION_FIELD + 9

Private Const FILL_PARAMETERS_COUNT = IDLE_PARAMETERS_COUNT + 9

' Дискретные параметры
Private Const FILL_PARAMETER_DESCR_PAUSE = "Разр. паузы"
Private Const FILL_PARAMETER_DESCR_COLDWATERGATE = "Разр. ХВ1"
Private Const FILL_PARAMETER_DESCR_HOTWATERGATE = "Разр. ГВ"
Private Const FILL_PARAMETER_DESCR_RECYCLEDWATERGATE = "Разр. ХВ2"
Private Const FILL_PARAMETER_DESCR_ROTATION = "Разр. вращ."

' Аналоговые параметры
Private Const FILL_PARAMETER_DESCR_LEVEL = "Уровень моющ. раств."
Private Const FILL_PARAMETER_DESCR_ROTTIME = "Время вращ. мотора"
Private Const FILL_PARAMETER_DESCR_PAUSETIME = "Время паузы вращ. мот."
Private Const FILL_PARAMETER_DESCR_DRUMSPEED = "Скорость вращ. барабана"

Public Sub ShowPropertyTableForFill(frm As FormMain)
    Dim I, J, row As Integer
    Dim ParamStr, s As String
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordFill As TYPE_WPC_FILL
    
    With frm
        ' Формируем столбик с названиями параметров
        ParamStr = ";Параметр|"
        
        For I = 1 To FILL_PARAMETERS_COUNT
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
                
                Case FILL_ROTTIME_FIELD:
                    ParamStr = ParamStr & FILL_PARAMETER_DESCR_ROTTIME
                
                Case FILL_PAUSETIME_FIELD:
                    ParamStr = ParamStr & FILL_PARAMETER_DESCR_PAUSETIME
                
                Case FILL_DRUMSPEED_FIELD:
                    ParamStr = ParamStr & FILL_PARAMETER_DESCR_DRUMSPEED
                    
                Case Else
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_UNKNOWN
            End Select
            
            If (I < FILL_PARAMETERS_COUNT) Then ParamStr = ParamStr & "|"
        Next I
                
        ' Эта строка автоматически изменяет количество строк
        .PropertyTable.FormatString = ParamStr
        
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        CopyMemory RecordFill, ByVal StepPointer, STEP_SIZE_IN_BYTES
        
        .PropertyTable.col = 1
        
        For I = 1 To FILL_PARAMETERS_COUNT
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
                    .PropertyTable.Text = FunctionsStrings(RecordFill.Bits And &HF)
                    
                ' Специальные параметры
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
                    .PropertyTable.Text = "" & RecordFill.Level
                    
                Case FILL_ROTTIME_FIELD:
                    .PropertyTable.Text = "" & RecordFill.RotTime
                
                Case FILL_PAUSETIME_FIELD:
                    .PropertyTable.Text = "" & RecordFill.PauseTime
                
                Case FILL_DRUMSPEED_FIELD:
                    .PropertyTable.Text = "" & RecordFill.DrumSpeed
                
                Case Else
                    .PropertyTable.CellBackColor = &H8000000F
                    
            End Select
        Next I
    End With
End Sub

Public Sub EditPropertyForFill(frm As FormMain)
    Dim I As Integer
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordFill As TYPE_WPC_FILL
    
    With frm
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        CopyMemory RecordFill, ByVal StepPointer, STEP_SIZE_IN_BYTES
        
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
                .ComboCell(0).ListIndex = (RecordTitle.LowBits And &H2) / 2 ^ 1
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
                .ComboCell(0).ListIndex = RecordFill.Bits And &HF
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            ' Специальные параметры шага
            Case FILL_PAUSE_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordFill.Bits And &H10) / 2 ^ 4
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case FILL_COLDWATERGATE_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordFill.Bits And &H20) / 2 ^ 5
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case FILL_HOTWATERGATE_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordFill.Bits And &H40) / 2 ^ 6
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case FILL_RECYCLEDWATERGATE_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordFill.Bits And &H80) / 2 ^ 7
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case FILL_ROTATION_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordFill.Bits And &H100) / 2 ^ 8
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case FILL_LEVEL_FIELD:
                .TextCell(0).Text = "" & RecordFill.Level
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
                
            Case FILL_ROTTIME_FIELD:
                .TextCell(0).Text = "" & RecordFill.RotTime
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
                
            Case FILL_PAUSETIME_FIELD:
                .TextCell(0).Text = "" & RecordFill.PauseTime
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
                
            Case FILL_DRUMSPEED_FIELD:
                .TextCell(0).Text = "" & RecordFill.DrumSpeed
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
                
            Case Else
            
        End Select
        
        .TextCell(0).SelStart = 0
        .TextCell(0).SelLength = Len(.TextCell(0).Text)
    End With
End Sub

Public Sub SetComboPropertyForFill(frm As FormMain)
    Dim I As Integer
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordFill As TYPE_WPC_FILL
    
    With frm
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        CopyMemory RecordFill, ByVal StepPointer, STEP_SIZE_IN_BYTES
    
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
                ZeroMemory RecordFill, STEP_SIZE_IN_BYTES
                ' и записать новое значение в поле типа функции
                RecordFill.Bits = .ComboCell(0).ListIndex And &HF
                
                ' Специальные параметры шага
            Case FILL_PAUSE_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordFill.Bits = RecordFill.Bits And &HFFEF
                    Case 1: RecordFill.Bits = RecordFill.Bits Or &H10
                End Select
                
            Case FILL_COLDWATERGATE_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordFill.Bits = RecordFill.Bits And &HFFDF
                    Case 1: RecordFill.Bits = RecordFill.Bits Or &H20
                End Select
                
            Case FILL_HOTWATERGATE_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordFill.Bits = RecordFill.Bits And &HFFBF
                    Case 1: RecordFill.Bits = RecordFill.Bits Or &H40
                End Select
                
            Case FILL_RECYCLEDWATERGATE_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordFill.Bits = RecordFill.Bits And &HFF7F
                    Case 1: RecordFill.Bits = RecordFill.Bits Or &H80
                End Select
                
            Case FILL_ROTATION_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordFill.Bits = RecordFill.Bits And &HFEFF
                    Case 1: RecordFill.Bits = RecordFill.Bits Or &H100
                End Select
                
            Case FILL_LEVEL_FIELD:
                RecordFill.Level = Val(.TextCell(0).Text)
                
            Case FILL_ROTTIME_FIELD:
                RecordFill.RotTime = Val(.TextCell(0).Text)
                
            Case FILL_PAUSETIME_FIELD:
                RecordFill.PauseTime = Val(.TextCell(0).Text)
                
            Case FILL_DRUMSPEED_FIELD:
                RecordFill.DrumSpeed = Val(.TextCell(0).Text)

        End Select
        
        .SetModified True
        
        ' Сохраняем изменения
        CopyMemory ByVal StepPointer, RecordFill, STEP_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory ByVal StepPointer, RecordTitle, HEADER_SIZE_IN_BYTES
    End With
End Sub


