VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TModuleTrin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const TRIN_PARAMETERS_COUNT = IDLE_PARAMETERS_COUNT + 0

' Название секции
Private Const TRIN_SECTION_NAME = "Trin"

' Настройки по умолчанию
' Дискретные типы
Private Const PAUSE_DEFAULT = 0

' Параметры функции шага
' Дискретные типы
Private Pause As TYPE_BOOL_DESCRIPTION

' Аналоговые типы

Public LimitsLoaded As Boolean

' Конструктор
Private Sub Class_Initialize()
    LimitsLoaded = False
End Sub

' Деструктор
Private Sub Class_Terminate()
End Sub

Public Sub LoadLimits(FileName As String)
    LimitsLoaded = DoesFileExist(FileName)
    
    If Not LimitsLoaded Then Exit Sub
    
    Dim LimitsFile As New TIniFiles
    
    LimitsFile.Create FileName
    
    ' Настройки функции
    ' Дискретные типы
    Pause.DefaultValue = LimitsFile.ReadInteger(TRIN_SECTION_NAME, "Pause.Default", PAUSE_DEFAULT) > 0
    
    ' Аналоговые типы
    
    Set LimitsFile = Nothing
End Sub

' Функция-посредник
Private Sub func_SetDefaults(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordStep As TYPE_WPC_STEP)

    Dim StepPointer As Long
    
    StepPointer = frm.Manager.DataPointer + _
        frm.Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
        HEADER_SIZE_IN_BYTES + _
        frm.Manager.StepIndex * STEP_SIZE_IN_BYTES
    
    'CopyMemory RecordCool, ByVal StepPointer, STEP_SIZE_IN_BYTES
    PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer
    
    ' Дискретные типы
    Select Case Pause.DefaultValue
        Case False: RecordStep.Bits = RecordStep.Bits And &HFFEF
        Case True: RecordStep.Bits = RecordStep.Bits Or &H10
    End Select
    
    ' Аналоговые типы
    
    ' Сохраняем изменения
    'StepPointer = frm.Manager.DataPointer + frm.Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
    'CopyMemory ByVal StepPointer, RecordStep, HEADER_SIZE_IN_BYTES
End Sub

Public Sub SetDefaults(frm As FormMain)
    Dim RecordStep As TYPE_WPC_STEP
    
    func_SetDefaults frm, 0&, RecordStep
End Sub
' Функция-посредник
Private Sub func_ShowPropertyTableForTrin(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordStep As TYPE_WPC_STEP)
    
    Dim I, J, row As Integer
    Dim ParamStr, s As String
    Dim StepPointer As Long
    
    With frm
        ' Формируем столбик с названиями параметров
        ParamStr = ";Параметр|"
        
        For I = 1 To TRIN_PARAMETERS_COUNT
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
            
            If (I < TRIN_PARAMETERS_COUNT) Then ParamStr = ParamStr & "|"
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
        
        For I = 1 To TRIN_PARAMETERS_COUNT
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

Public Sub ShowPropertyTableForTrin(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP
    
    func_ShowPropertyTableForTrin frm, 0&, RecordTitle, RecordStep
End Sub

' Функция-посредник
Private Sub func_EditPropertyForTrin(frm As FormMain, _
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

Public Sub EditPropertyForTrin(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP
    
    func_EditPropertyForTrin frm, 0&, RecordTitle, RecordStep
End Sub

' Функция-посредник
Private Sub func_SetComboPropertyForTrin(frm As FormMain, _
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
                ' Если функция та же, то ничего не делаем
                If (RecordStep.Bits And &HF) = (.ComboCell.ListIndex And &HF) Then Exit Sub
                
                ' При выборе новой функции для шага мы должны обнулить структуру,
                ZeroMemory RecordStep, STEP_SIZE_IN_BYTES
                
                ' установить значения по умолчанию в зависимости от функции,
                If LimitsLoaded Then
                    Select Case .ComboCell.ListIndex And &HF
                        Case WPC_OPERATION_IDLE ' пропуск
                            .ModuleIdle.SetDefaults frm
                            
                        Case WPC_OPERATION_FILL ' Налив
                            .ModuleFill.SetDefaults frm
                        
                        Case WPC_OPERATION_DTRG ' моющие
                            .ModuleDTRG.SetDefaults frm
                            
                        Case WPC_OPERATION_HEAT ' нагрев
                            .ModuleHeat.SetDefaults frm
                            
                        ' стирка, полоскание, расстряска
                        Case WPC_OPERATION_WASH, WPC_OPERATION_RINS, WPC_OPERATION_JOLT, WPC_OPERATION_PAUS
                            .ModuleWashOrRinsOrJolt.SetDefaults frm
                            
                        'Case WPC_OPERATION_PAUS ' пауза
            
                        Case WPC_OPERATION_DRAIN ' слив
                            .ModuleDrain.SetDefaults frm
                            
                        Case WPC_OPERATION_SPIN ' отжим
                            .ModuleSpin.SetDefaults frm
                        
                        Case WPC_OPERATION_COOL ' охлаждение
                            .ModuleCool.SetDefaults frm
                            
                        Case WPC_OPERATION_TRIN ' тех.полоскание
                            .ModuleTrin.SetDefaults frm
                    
                        Case Else
                        
                    End Select
                End If
                
                ' записать новое значение в поле типа функции
                RecordStep.Bits = RecordStep.Bits Or (.ComboCell.ListIndex And &HF)
                
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

Public Sub SetComboPropertyForTrin(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP
    
    func_SetComboPropertyForTrin frm, 0&, RecordTitle, RecordStep
End Sub


