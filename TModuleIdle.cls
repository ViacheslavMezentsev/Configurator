VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TModuleIdle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Функция-посредник
Private Sub func_ShowPropertyTableForIdle(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordStep As TYPE_WPC_STEP)
    
    Dim I, J, row As Integer
    Dim StepPointer As Long
    Dim ParamStr, s As String
    
    With frm
        ' Формируем столбик с названиями параметров
        ParamStr = ";Параметр|"
        
        For I = 1 To IDLE_PARAMETERS_COUNT
            Select Case I
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
                    
                Case Else
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_UNKNOWN
            End Select
            
            If (I < IDLE_PARAMETERS_COUNT) Then ParamStr = ParamStr & "|"
        Next I
        
        .PropertyTable.FormatString = ParamStr
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        'CopyMemory RecordStep, ByVal StepPointer, STEP_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer
        
        .PropertyTable.col = 1
        
        For I = 1 To IDLE_PARAMETERS_COUNT
            .PropertyTable.row = I
            
            Select Case I
                Case IDLE_STEP_FIELD:
                    .PropertyTable.CellFontBold = True
                    .PropertyTable.Text = "" & .Manager.StepIndex + 1
                
                Case IDLE_FUNCTION_FIELD:
                    .PropertyTable.CellFontBold = True
                    .PropertyTable.Text = FunctionsStrings(RecordStep.Bits And &HF)
                    
                Case Else
                    .PropertyTable.CellBackColor = &H8000000F
                    
            End Select
        Next I
    End With
End Sub

Public Sub ShowPropertyTableForIdle(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP

    func_ShowPropertyTableForIdle frm, 0&, RecordTitle, RecordStep
End Sub

' Функция-посредник
Private Sub func_EditPropertyForIdle(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordStep As TYPE_WPC_STEP)
    
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
            Case IDLE_ENDSOUND_FIELD:
            
            Case IDLE_DOORUNLOCK_FIELD:
            
            Case IDLE_PROGNAME_FIELD:
            
            Case IDLE_STEP_FIELD:
                .ComboCell(0).Left = .PropertyTable.Left + .PropertyTable.CellLeft
                .ComboCell(0).Top = .PropertyTable.Top + .PropertyTable.CellTop
                .ComboCell(0).Width = .PropertyTable.CellWidth
                .ComboCell(0).Clear
                For I = 1 To MAX_NUMBER_OF_STEPS
                    .ComboCell(0).AddItem ("Шаг " & I)
                Next I
                .ComboCell(0).ListIndex = .Manager.StepIndex
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
                
            Case IDLE_FUNCTION_FIELD:
                .ComboCell(0).Left = .PropertyTable.Left + .PropertyTable.CellLeft
                .ComboCell(0).Top = .PropertyTable.Top + .PropertyTable.CellTop
                .ComboCell(0).Width = .PropertyTable.CellWidth
                .ComboCell(0).Clear
                For I = 0 To 11
                    .ComboCell(0).AddItem (FunctionsStrings(I))
                Next I
                .ComboCell(0).ListIndex = RecordStep.Bits And &HF
                .ComboCell(0).Visible = True
                .ComboCell(0).SetFocus
            
            Case Else
        End Select
    End With
End Sub

Public Sub EditPropertyForIdle(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP
    
    func_EditPropertyForIdle frm, 0&, RecordTitle, RecordStep
End Sub

' Функция-посредник
Private Sub func_SetComboPropertyForIdle(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordStep As TYPE_WPC_STEP)
    
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
            ' В режиме пропуска доступны для изменения только два поля
            Case IDLE_ENDSOUND_FIELD:
            Case IDLE_DOORUNLOCK_FIELD:
            Case IDLE_PROGNAME_FIELD:
            
            Case IDLE_STEP_FIELD:
                .Manager.StepIndex = .ComboCell(0).ListIndex
                Exit Sub
                
            Case IDLE_FUNCTION_FIELD:
                ' При выборе новой функции для шага мы должны обнулить структуру
                ZeroMemory RecordStep, STEP_SIZE_IN_BYTES
                ' и записать новое значение в поле типа функции
                RecordStep.Bits = .ComboCell(0).ListIndex And &HF
                
            Case Else
                
        End Select
        
        .SetModified True
        
        ' Сохраняем изменения
        'CopyMemory ByVal StepPointer, RecordStep, STEP_SIZE_IN_BYTES
    End With
End Sub

Public Sub SetComboPropertyForIdle(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP
    
    func_SetComboPropertyForIdle frm, 0&, RecordTitle, RecordStep
End Sub


