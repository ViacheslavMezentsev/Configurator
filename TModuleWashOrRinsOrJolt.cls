VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TModuleWashOrRinsOrJolt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const WASH_PAUSE_FIELD = IDLE_FUNCTION_FIELD + 1
Private Const WASH_ROTATION_FIELD = IDLE_FUNCTION_FIELD + 2

Private Const WASH_TIME_FIELD = IDLE_FUNCTION_FIELD + 3
Private Const WASH_ROTTIME_FIELD = IDLE_FUNCTION_FIELD + 4
Private Const WASH_PAUSETIME_FIELD = IDLE_FUNCTION_FIELD + 5
Private Const WASH_DRUMSPEED_FIELD = IDLE_FUNCTION_FIELD + 6

Private Const WASH_PARAMETERS_COUNT = IDLE_PARAMETERS_COUNT + 6

' Дискретные параметры
Private Const WASH_PARAMETER_DESCR_PAUSE = "Разр. паузы"
Private Const WASH_PARAMETER_DESCR_ROTATION = "Разр. вращ."

' Аналоговые параметры
Private Const WASH_PARAMETER_DESCR_TIME = "Время"
Private Const WASH_PARAMETER_DESCR_ROTTIME = "Время вращ. мотора"
Private Const WASH_PARAMETER_DESCR_PAUSETIME = "Время паузы вращ. мот."
Private Const WASH_PARAMETER_DESCR_DRUMSPEED = "Скорость вращ. барабана"

' Функция-посредник
Private Sub func_ShowPropertyTableForWashOrRinsOrJolt(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordWash As TYPE_WPC_WASH)
    
    Dim I, J, row As Integer
    Dim ParamStr, s As String
    Dim StepPointer As Long
    
    With frm
        ' Формируем столбик с названиями параметров
        ParamStr = ";Параметр|"
        
        For I = 1 To WASH_PARAMETERS_COUNT
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
                Case WASH_PAUSE_FIELD:
                    ParamStr = ParamStr & WASH_PARAMETER_DESCR_PAUSE
                    
                Case WASH_ROTATION_FIELD:
                    ParamStr = ParamStr & WASH_PARAMETER_DESCR_ROTATION
                
                Case WASH_TIME_FIELD:
                    ParamStr = ParamStr & WASH_PARAMETER_DESCR_TIME
                
                Case WASH_ROTTIME_FIELD:
                    ParamStr = ParamStr & WASH_PARAMETER_DESCR_ROTTIME
                
                Case WASH_PAUSETIME_FIELD:
                    ParamStr = ParamStr & WASH_PARAMETER_DESCR_PAUSETIME
                
                Case WASH_DRUMSPEED_FIELD:
                    ParamStr = ParamStr & WASH_PARAMETER_DESCR_DRUMSPEED
                    
                Case Else
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_UNKNOWN
            End Select
            
            If (I < WASH_PARAMETERS_COUNT) Then ParamStr = ParamStr & "|"
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
        
        'CopyMemory RecordWash, ByVal StepPointer, STEP_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer
        
        .PropertyTable.col = 1
        
        For I = 1 To WASH_PARAMETERS_COUNT
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
                    .PropertyTable.Text = FunctionsStrings(RecordWash.Bits And &HF)
                    
                ' Специальные параметры
                Case WASH_PAUSE_FIELD:
                    If (RecordWash.Bits And &H10) / &H10 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                Case WASH_ROTATION_FIELD:
                    If (RecordWash.Bits And &H20) / &H20 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If
                    
                Case WASH_TIME_FIELD:
                    .PropertyTable.Text = "" & RecordWash.Time
                    
                Case WASH_ROTTIME_FIELD:
                    .PropertyTable.Text = "" & RecordWash.RotTime
                
                Case WASH_PAUSETIME_FIELD:
                    .PropertyTable.Text = "" & RecordWash.PauseTime
                
                Case WASH_DRUMSPEED_FIELD:
                    .PropertyTable.Text = "" & RecordWash.DrumSpeed
                
                Case Else
                    .PropertyTable.CellBackColor = &H8000000F
                    
            End Select
        Next I
    End With
End Sub

Public Sub ShowPropertyTableForWashOrRinsOrJolt(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordWash As TYPE_WPC_WASH
    
    func_ShowPropertyTableForWashOrRinsOrJolt frm, 0&, RecordTitle, RecordWash
End Sub

' Функция-посредник
Private Sub func_EditPropertyForWashOrRinsOrJolt(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordWash As TYPE_WPC_WASH)
    
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
        
        'CopyMemory RecordWash, ByVal StepPointer, STEP_SIZE_IN_BYTES
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
                .ComboCell(0).ListIndex = RecordWash.Bits And &HF
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            ' Специальные параметры шага
            Case WASH_PAUSE_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordWash.Bits And &H10) / &H10
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus

            Case WASH_ROTATION_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordWash.Bits And &H20) / &H20
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus

            Case WASH_TIME_FIELD:
                .TextCell(0).Text = "" & RecordWash.Time
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus

            Case WASH_ROTTIME_FIELD:
                .TextCell(0).Text = "" & RecordWash.RotTime
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case WASH_PAUSETIME_FIELD:
                .TextCell(0).Text = "" & RecordWash.PauseTime
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case WASH_DRUMSPEED_FIELD:
                .TextCell(0).Text = "" & RecordWash.DrumSpeed
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case Else

        End Select
    
        .TextCell(0).SelStart = 0
        .TextCell(0).SelLength = Len(.TextCell(0).Text)
    End With
End Sub

Public Sub EditPropertyForWashOrRinsOrJolt(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordWash As TYPE_WPC_WASH
    
    func_EditPropertyForWashOrRinsOrJolt frm, 0&, RecordTitle, RecordWash
End Sub

' Функция-посредник
Private Sub func_SetComboPropertyForWashOrRinsOrJolt(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordWash As TYPE_WPC_WASH)
    
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
        
        'CopyMemory RecordWash, ByVal StepPointer, STEP_SIZE_IN_BYTES
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
                ZeroMemory RecordWash, STEP_SIZE_IN_BYTES
                ' и записать новое значение в поле типа функции
                RecordWash.Bits = .ComboCell(0).ListIndex And &HF
                
            ' Специальные параметры шага
            Case WASH_PAUSE_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordWash.Bits = RecordWash.Bits And &HFFEF
                    Case 1: RecordWash.Bits = RecordWash.Bits Or &H10
                End Select
    
            Case WASH_ROTATION_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordWash.Bits = RecordWash.Bits And &HFFDF
                    Case 1: RecordWash.Bits = RecordWash.Bits Or &H20
                End Select

            Case WASH_TIME_FIELD:
                RecordWash.Time = Val(.TextCell(0).Text)
    
            Case WASH_ROTTIME_FIELD:
                RecordWash.RotTime = Val(.TextCell(0).Text)
            
            Case WASH_PAUSETIME_FIELD:
                RecordWash.PauseTime = Val(.TextCell(0).Text)
            
            Case WASH_DRUMSPEED_FIELD:
                RecordWash.DrumSpeed = Val(.TextCell(0).Text)
            
            Case Else
    
        End Select
        
        .SetModified True
        
        ' Сохраняем изменения
        'CopyMemory ByVal StepPointer, RecordWash, STEP_SIZE_IN_BYTES
        
        'StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        'CopyMemory ByVal StepPointer, RecordTitle, HEADER_SIZE_IN_BYTES
    End With
End Sub

Public Sub SetComboPropertyForWashOrRinsOrJolt(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordWash As TYPE_WPC_WASH
    
    func_SetComboPropertyForWashOrRinsOrJolt frm, 0&, RecordTitle, RecordWash
End Sub

