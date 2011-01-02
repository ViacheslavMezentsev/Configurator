VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TModuleDTRG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const DTRG_PAUSE_FIELD = IDLE_FUNCTION_FIELD + 1
Private Const DTRG_ROTATION_FIELD = IDLE_FUNCTION_FIELD + 2

Private Const DTRG_DETERGENT1_FIELD = IDLE_FUNCTION_FIELD + 3
Private Const DTRG_DETERGENT2_FIELD = IDLE_FUNCTION_FIELD + 4
Private Const DTRG_DETERGENT3_FIELD = IDLE_FUNCTION_FIELD + 5
Private Const DTRG_DETERGENT4_FIELD = IDLE_FUNCTION_FIELD + 6
Private Const DTRG_DETERGENT5_FIELD = IDLE_FUNCTION_FIELD + 7
Private Const DTRG_DETERGENT6_FIELD = IDLE_FUNCTION_FIELD + 8
Private Const DTRG_DETERGENT7_FIELD = IDLE_FUNCTION_FIELD + 9
Private Const DTRG_DETERGENT8_FIELD = IDLE_FUNCTION_FIELD + 10
Private Const DTRG_DETERGENT9_FIELD = IDLE_FUNCTION_FIELD + 11

Private Const DTRG_ROTTIME_FIELD = IDLE_FUNCTION_FIELD + 12
Private Const DTRG_PAUSETIME_FIELD = IDLE_FUNCTION_FIELD + 13
Private Const DTRG_DRUMSPEED_FIELD = IDLE_FUNCTION_FIELD + 14

Private Const DTRG_PARAMETERS_COUNT = IDLE_PARAMETERS_COUNT + 14

' Дискретные параметры
Private Const DTRG_PARAMETER_DESCR_PAUSE = "Разр. паузы"
Private Const DTRG_PARAMETER_DESCR_ROTATION = "Разр. вращ."

' Аналоговые параметры
Private Const DTRG_PARAMETER_DESCR_DETERGENT1 = "Время подачи моющ. 1"
Private Const DTRG_PARAMETER_DESCR_DETERGENT2 = "Время подачи моющ. 2"
Private Const DTRG_PARAMETER_DESCR_DETERGENT3 = "Время подачи моющ. 3"
Private Const DTRG_PARAMETER_DESCR_DETERGENT4 = "Время подачи моющ. 4"
Private Const DTRG_PARAMETER_DESCR_DETERGENT5 = "Время подачи моющ. 5"
Private Const DTRG_PARAMETER_DESCR_DETERGENT6 = "Время подачи моющ. 6"
Private Const DTRG_PARAMETER_DESCR_DETERGENT7 = "Время подачи моющ. 7"
Private Const DTRG_PARAMETER_DESCR_DETERGENT8 = "Время подачи моющ. 8"
Private Const DTRG_PARAMETER_DESCR_DETERGENT9 = "Время подачи моющ. 9"

Private Const DTRG_PARAMETER_DESCR_ROTTIME = "Время вращ. мотора"
Private Const DTRG_PARAMETER_DESCR_PAUSETIME = "Время паузы вращ. мот."
Private Const DTRG_PARAMETER_DESCR_DRUMSPEED = "Скорость вращ. барабана"

Public Sub ShowPropertyTableForDTRG(frm As FormMain)
    Dim I, J, row As Integer
    Dim ParamStr, s As String
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordDetergent As TYPE_WPC_DETERGENT
    
    With frm
        ' Формируем столбик с названиями параметров
        ParamStr = ";Параметр|"
        For I = 1 To DTRG_PARAMETERS_COUNT
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
                Case DTRG_PAUSE_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_PAUSE
                    
                Case DTRG_ROTATION_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_ROTATION
                    
                Case DTRG_DETERGENT1_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_DETERGENT1
                    
                Case DTRG_DETERGENT2_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_DETERGENT2
                
                Case DTRG_DETERGENT3_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_DETERGENT3
                
                Case DTRG_DETERGENT4_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_DETERGENT4
                
                Case DTRG_DETERGENT5_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_DETERGENT5
                
                Case DTRG_DETERGENT6_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_DETERGENT6
                
                Case DTRG_DETERGENT7_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_DETERGENT7
                    
                Case DTRG_DETERGENT8_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_DETERGENT8
                    
                Case DTRG_DETERGENT9_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_DETERGENT9
                
                Case DTRG_ROTTIME_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_ROTTIME
                    
                Case DTRG_PAUSETIME_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_PAUSETIME
                    
                Case DTRG_DRUMSPEED_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_DRUMSPEED
                    
                Case Else
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_UNKNOWN
            End Select
            
            If (I < DTRG_PARAMETERS_COUNT) Then ParamStr = ParamStr & "|"
        Next I
                
        ' Эта строка автоматически изменяет количество строк
        .PropertyTable.FormatString = ParamStr
        
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        CopyMemory RecordDetergent, ByVal StepPointer, STEP_SIZE_IN_BYTES
        
        .PropertyTable.col = 1
               
        For I = 1 To DTRG_PARAMETERS_COUNT
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
                    .PropertyTable.Text = FunctionsStrings(RecordDetergent.Bits And &HF)
                    
                ' Специальные параметры
                Case DTRG_PAUSE_FIELD:
                    If (RecordDetergent.Bits And &H10) / &H10 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If
                    
                Case DTRG_ROTATION_FIELD:
                    If (RecordDetergent.Bits And &H20) / &H20 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If
                    
                Case DTRG_DETERGENT1_FIELD:
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_1_Time
                    
                Case DTRG_DETERGENT2_FIELD:
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_2_Time
                
                Case DTRG_DETERGENT3_FIELD:
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_3_Time
                
                Case DTRG_DETERGENT4_FIELD:
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_4_Time
                
                Case DTRG_DETERGENT5_FIELD:
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_5_Time
                
                Case DTRG_DETERGENT6_FIELD:
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_6_Time
                
                Case DTRG_DETERGENT7_FIELD:
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_7_Time
                    
                Case DTRG_DETERGENT8_FIELD:
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_8_Time
                    
                Case DTRG_DETERGENT9_FIELD:
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_9_Time
                
                Case DTRG_ROTTIME_FIELD:
                    .PropertyTable.Text = "" & RecordDetergent.RotTime
                    
                Case DTRG_PAUSETIME_FIELD:
                    .PropertyTable.Text = "" & RecordDetergent.PauseTime
                    
                Case DTRG_DRUMSPEED_FIELD:
                    .PropertyTable.Text = "" & RecordDetergent.DrumSpeed
                    
                Case Else
                    .PropertyTable.CellBackColor = &H8000000F
                
            End Select
        Next I
    End With
End Sub

Public Sub EditPropertyForDTRG(frm As FormMain)
    Dim I As Integer
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordDetergent As TYPE_WPC_DETERGENT
    
    With frm
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        CopyMemory RecordDetergent, ByVal StepPointer, STEP_SIZE_IN_BYTES
        
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
                .ComboCell(0).ListIndex = RecordDetergent.Bits And &HF
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            ' Специальные параметры шага
            Case DTRG_PAUSE_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordDetergent.Bits And &H10) / &H10
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case DTRG_ROTATION_FIELD:
                .ComboCell(0).AddItem STRING_NO
                .ComboCell(0).AddItem STRING_YES
                .ComboCell(0).ListIndex = (RecordDetergent.Bits And &H20) / &H20
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case DTRG_DETERGENT1_FIELD:
                .TextCell(0).Text = "" & RecordDetergent.Detergent_1_Time
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
                
            Case DTRG_DETERGENT2_FIELD:
                .TextCell(0).Text = "" & RecordDetergent.Detergent_2_Time
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case DTRG_DETERGENT3_FIELD:
                .TextCell(0).Text = "" & RecordDetergent.Detergent_3_Time
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case DTRG_DETERGENT4_FIELD:
                .TextCell(0).Text = "" & RecordDetergent.Detergent_4_Time
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case DTRG_DETERGENT5_FIELD:
                .TextCell(0).Text = "" & RecordDetergent.Detergent_5_Time
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case DTRG_DETERGENT6_FIELD:
                .TextCell(0).Text = "" & RecordDetergent.Detergent_6_Time
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case DTRG_DETERGENT7_FIELD:
                .TextCell(0).Text = "" & RecordDetergent.Detergent_7_Time
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
                
            Case DTRG_DETERGENT8_FIELD:
                .TextCell(0).Text = "" & RecordDetergent.Detergent_8_Time
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
                
            Case DTRG_DETERGENT9_FIELD:
                .TextCell(0).Text = "" & RecordDetergent.Detergent_9_Time
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
            
            Case DTRG_ROTTIME_FIELD:
                .TextCell(0).Text = "" & RecordDetergent.RotTime
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
                
            Case DTRG_PAUSETIME_FIELD:
                .TextCell(0).Text = "" & RecordDetergent.PauseTime
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
                
            Case DTRG_DRUMSPEED_FIELD:
                .TextCell(0).Text = "" & RecordDetergent.DrumSpeed
                .TextCell(0).Visible = True
                .TextCell(0).SetFocus
                
            Case Else
        
        End Select
        
        .TextCell(0).SelStart = 0
        .TextCell(0).SelLength = Len(.TextCell(0).Text)
    End With
End Sub

Public Sub SetComboPropertyForDTRG(frm As FormMain)
    Dim I As Integer
    Dim StepPointer As Long
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordDetergent As TYPE_WPC_DETERGENT
    
    With frm
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        CopyMemory RecordDetergent, ByVal StepPointer, STEP_SIZE_IN_BYTES
        
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
                ZeroMemory RecordDetergent, STEP_SIZE_IN_BYTES
                ' и записать новое значение в поле типа функции
                RecordDetergent.Bits = .ComboCell(0).ListIndex And &HF
                
            ' Специальные параметры шага
            Case DTRG_PAUSE_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordDetergent.Bits = RecordDetergent.Bits And &HFFEF
                    Case 1: RecordDetergent.Bits = RecordDetergent.Bits Or &H10
                End Select
                
            Case DTRG_ROTATION_FIELD:
                Select Case .ComboCell(0).ListIndex
                    Case 0: RecordDetergent.Bits = RecordDetergent.Bits And &HFFDF
                    Case 1: RecordDetergent.Bits = RecordDetergent.Bits Or &H20
                End Select
                
            Case DTRG_DETERGENT1_FIELD:
                RecordDetergent.Detergent_1_Time = Val(.TextCell(0).Text)
                
            Case DTRG_DETERGENT2_FIELD:
                RecordDetergent.Detergent_2_Time = Val(.TextCell(0).Text)
            
            Case DTRG_DETERGENT3_FIELD:
                RecordDetergent.Detergent_3_Time = Val(.TextCell(0).Text)
            
            Case DTRG_DETERGENT4_FIELD:
                RecordDetergent.Detergent_4_Time = Val(.TextCell(0).Text)
            
            Case DTRG_DETERGENT5_FIELD:
                RecordDetergent.Detergent_5_Time = Val(.TextCell(0).Text)
            
            Case DTRG_DETERGENT6_FIELD:
                RecordDetergent.Detergent_6_Time = Val(.TextCell(0).Text)
            
            Case DTRG_DETERGENT7_FIELD:
                RecordDetergent.Detergent_7_Time = Val(.TextCell(0).Text)
                
            Case DTRG_DETERGENT8_FIELD:
                RecordDetergent.Detergent_8_Time = Val(.TextCell(0).Text)
                
            Case DTRG_DETERGENT9_FIELD:
                RecordDetergent.Detergent_9_Time = Val(.TextCell(0).Text)
            
            Case DTRG_ROTTIME_FIELD:
                RecordDetergent.RotTime = Val(.TextCell(0).Text)
                
            Case DTRG_PAUSETIME_FIELD:
                RecordDetergent.PauseTime = Val(.TextCell(0).Text)
                
            Case DTRG_DRUMSPEED_FIELD:
                RecordDetergent.DrumSpeed = Val(.TextCell(0).Text)
                
            Case Else

        End Select
        
        .SetModified True
        
        ' Сохраняем изменения
        CopyMemory ByVal StepPointer, RecordDetergent, STEP_SIZE_IN_BYTES
        
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        CopyMemory ByVal StepPointer, RecordTitle, HEADER_SIZE_IN_BYTES
    End With
End Sub


