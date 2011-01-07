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
Private Const WASH_ROTATIONTIME_FIELD = IDLE_FUNCTION_FIELD + 4
Private Const WASH_PAUSETIME_FIELD = IDLE_FUNCTION_FIELD + 5
Private Const WASH_DRUMSPEED_FIELD = IDLE_FUNCTION_FIELD + 6

Private Const WASH_PARAMETERS_COUNT = IDLE_PARAMETERS_COUNT + 6

' Дискретные параметры
Private Const WASH_PARAMETER_DESCR_PAUSE = "Разр. паузы"
Private Const WASH_PARAMETER_DESCR_ROTATION = "Разр. вращ."

' Аналоговые параметры
Private Const WASH_PARAMETER_DESCR_TIME = "Время"
Private Const WASH_PARAMETER_DESCR_RotationTime = "Время вращ. мотора"
Private Const WASH_PARAMETER_DESCR_PAUSETIME = "Время паузы вращ. мот."
Private Const WASH_PARAMETER_DESCR_DRUMSPEED = "Скорость вращ. барабана"

' Название секции
Private Const WASH_SECTION_NAME = "Wash"

' Настройки по умолчанию
' Дискретные типы
Private Const PAUSE_DEFAULT = 0
Private Const ROTATION_DEFAULT = 1

' Аналоговые типы
Private Const TIME_MIN = 0
Private Const TIME_MAX = 255
Private Const TIME_DEFAULT = 30
Private Const TIME_DIMENSION = "мин"

Private Const ROTATIONTIME_MIN = 1
Private Const ROTATIONTIME_MAX = 250
Private Const ROTATIONTIME_DEFAULT = 6
Private Const ROTATIONTIME_DIMENSION = "сек"

Private Const PAUSETIME_MIN = 1
Private Const PAUSETIME_MAX = 250
Private Const PAUSETIME_DEFAULT = 12
Private Const PAUSETIME_DIMENSION = "сек"

Private Const DRUMSPEED_MIN = 30
Private Const DRUMSPEED_MAX = 70
Private Const DRUMSPEED_DEFAULT = 50
Private Const DRUMSPEED_DIMENSION = "об/мин"

' Параметры функции шага
' Дискретные типы
Private Pause As TYPE_BOOL_DESCRIPTION
Private Rotation As TYPE_BOOL_DESCRIPTION

' Аналоговые типы
Private Time As TYPE_BYTE_DESCRIPTION
Private RotationTime As TYPE_BYTE_DESCRIPTION
Private PauseTime As TYPE_BYTE_DESCRIPTION
Private DrumSpeed As TYPE_BYTE_DESCRIPTION

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
    Pause.DefaultValue = LimitsFile.ReadInteger(WASH_SECTION_NAME, "Pause.Default", PAUSE_DEFAULT) > 0
    Rotation.DefaultValue = LimitsFile.ReadInteger(WASH_SECTION_NAME, "Rotation.Default", ROTATION_DEFAULT) > 0
    
    ' Аналоговые типы
    With Time
        .MinValue = LimitsFile.ReadInteger(WASH_SECTION_NAME, "Time.Min", TIME_MIN)
        .MaxValue = LimitsFile.ReadInteger(WASH_SECTION_NAME, "Time.Max", TIME_MAX)
        .DefaultValue = LimitsFile.ReadInteger(WASH_SECTION_NAME, "Time.Default", TIME_DEFAULT)
        .Dimension = LimitsFile.ReadString(WASH_SECTION_NAME, "Time.Dimension", TIME_DIMENSION)
    End With
        
    With RotationTime
        .MinValue = LimitsFile.ReadInteger(WASH_SECTION_NAME, "RotationTime.Min", ROTATIONTIME_MIN)
        .MaxValue = LimitsFile.ReadInteger(WASH_SECTION_NAME, "RotationTime.Max", ROTATIONTIME_MAX)
        .DefaultValue = LimitsFile.ReadInteger(WASH_SECTION_NAME, "RotationTime.Default", ROTATIONTIME_DEFAULT)
        .Dimension = LimitsFile.ReadString(WASH_SECTION_NAME, "RotationTime.Dimension", ROTATIONTIME_DIMENSION)
    End With
    
    With PauseTime
        .MinValue = LimitsFile.ReadInteger(WASH_SECTION_NAME, "PauseTime.Min", PAUSETIME_MIN)
        .MaxValue = LimitsFile.ReadInteger(WASH_SECTION_NAME, "PauseTime.Max", PAUSETIME_MAX)
        .DefaultValue = LimitsFile.ReadInteger(WASH_SECTION_NAME, "PauseTime.Default", PAUSETIME_DEFAULT)
        .Dimension = LimitsFile.ReadString(WASH_SECTION_NAME, "PauseTime.Dimension", PAUSETIME_DIMENSION)
    End With
    
    With DrumSpeed
        .MinValue = LimitsFile.ReadInteger(WASH_SECTION_NAME, "DrumSpeed.Min", DRUMSPEED_MIN)
        .MaxValue = LimitsFile.ReadInteger(WASH_SECTION_NAME, "DrumSpeed.Max", DRUMSPEED_MAX)
        .DefaultValue = LimitsFile.ReadInteger(WASH_SECTION_NAME, "DrumSpeed.Default", DRUMSPEED_DEFAULT)
        .Dimension = LimitsFile.ReadString(WASH_SECTION_NAME, "DrumSpeed.Dimension", DRUMSPEED_DIMENSION)
    End With
    
    Set LimitsFile = Nothing
End Sub

' Функция-посредник
Private Sub func_SetDefaults(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordWash As TYPE_WPC_WASH)

    Dim StepPointer As Long
    
    StepPointer = frm.Manager.DataPointer + _
        frm.Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
        HEADER_SIZE_IN_BYTES + _
        frm.Manager.StepIndex * STEP_SIZE_IN_BYTES
    
    'CopyMemory RecordWash, ByVal StepPointer, STEP_SIZE_IN_BYTES
    PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer
    
    ' Дискретные типы
    Select Case Pause.DefaultValue
        Case False: RecordWash.Bits = RecordWash.Bits And &HFFEF
        Case True: RecordWash.Bits = RecordWash.Bits Or &H10
    End Select

    Select Case Rotation.DefaultValue
        Case False: RecordWash.Bits = RecordWash.Bits And &HFFDF
        Case True: RecordWash.Bits = RecordWash.Bits Or &H20
    End Select
    
    ' Аналоговые типы
    RecordWash.Time = Time.DefaultValue
    RecordWash.RotationTime = RotationTime.DefaultValue
    RecordWash.PauseTime = PauseTime.DefaultValue
    RecordWash.DrumSpeed = DrumSpeed.DefaultValue
    
    ' Сохраняем изменения
    'StepPointer = frm.Manager.DataPointer + frm.Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
    'CopyMemory ByVal StepPointer, RecordWash, HEADER_SIZE_IN_BYTES
End Sub

Public Sub SetDefaults(frm As FormMain)
    Dim RecordWash As TYPE_WPC_WASH
    
    func_SetDefaults frm, 0&, RecordWash
End Sub

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
                
                Case WASH_ROTATIONTIME_FIELD:
                    ParamStr = ParamStr & WASH_PARAMETER_DESCR_RotationTime
                
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
                    If LimitsLoaded Then
                        If RecordWash.Time < Time.MinValue Or RecordWash.Time > Time.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordWash.Time
                    
                Case WASH_ROTATIONTIME_FIELD:
                    If LimitsLoaded Then
                        If RecordWash.RotationTime < RotationTime.MinValue Or RecordWash.RotationTime > RotationTime.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordWash.RotationTime
                
                Case WASH_PAUSETIME_FIELD:
                    If LimitsLoaded Then
                        If RecordWash.PauseTime < PauseTime.MinValue Or RecordWash.PauseTime > PauseTime.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordWash.PauseTime
                
                Case WASH_DRUMSPEED_FIELD:
                    If LimitsLoaded Then
                        If RecordWash.DrumSpeed < DrumSpeed.MinValue Or RecordWash.DrumSpeed > DrumSpeed.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
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
    Dim s As String
    
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
        
        .ComboCell.Left = .PropertyTable.Left + .PropertyTable.CellLeft
        .ComboCell.Top = .PropertyTable.Top + .PropertyTable.CellTop
        .ComboCell.Width = .PropertyTable.CellWidth
        .ComboCell.Clear
        
        .TextCell.Left = .PropertyTable.Left + .PropertyTable.CellLeft
        .TextCell.Top = .PropertyTable.Top + .PropertyTable.CellTop
        .TextCell.Width = .PropertyTable.CellWidth
        .TextCell.Height = .PropertyTable.CellHeight
        
        .LabelDescription.Left = .PropertyTable.Left
        .LabelDescription.Width = .PropertyTable.Width
        .LabelDescription.Height = 5 * .PropertyTable.RowHeight(0)
        .LabelDescription.Top = .PropertyTable.Top + .PropertyTable.Height - .LabelDescription.Height
        
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
                .ComboCell.ListIndex = RecordWash.Bits And &HF
                .ComboCell.Visible = True
                .ComboCell.SetFocus
                
            ' Специальные параметры шага
            Case WASH_PAUSE_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordWash.Bits And &H10) / &H10
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case WASH_ROTATION_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordWash.Bits And &H20) / &H20
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case WASH_TIME_FIELD:
                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = Chr(13) & Chr(10) & _
                                                "Минимальное значение: " & Time.MinValue & Chr(13) & Chr(10) & _
                                                "Максимальное значение: " & Time.MaxValue & Chr(13) & Chr(10) & _
                                                "Значение по умолчанию: " & Time.DefaultValue & Chr(13) & Chr(10) & _
                                                "Размерность: [" & Time.Dimension & "]"
                    .LabelDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordWash.Time
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case WASH_ROTATIONTIME_FIELD:
                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = Chr(13) & Chr(10) & _
                                                "Минимальное значение: " & RotationTime.MinValue & Chr(13) & Chr(10) & _
                                                "Максимальное значение: " & RotationTime.MaxValue & Chr(13) & Chr(10) & _
                                                "Значение по умолчанию: " & RotationTime.DefaultValue & Chr(13) & Chr(10) & _
                                                "Размерность: [" & RotationTime.Dimension & "]"
                    .LabelDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordWash.RotationTime
                .TextCell.Visible = True
                .TextCell.SetFocus
            
            Case WASH_PAUSETIME_FIELD:
                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = Chr(13) & Chr(10) & _
                                                "Минимальное значение: " & PauseTime.MinValue & Chr(13) & Chr(10) & _
                                                "Максимальное значение: " & PauseTime.MaxValue & Chr(13) & Chr(10) & _
                                                "Значение по умолчанию: " & PauseTime.DefaultValue & Chr(13) & Chr(10) & _
                                                "Размерность: [" & PauseTime.Dimension & "]"
                    .LabelDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordWash.PauseTime
                .TextCell.Visible = True
                .TextCell.SetFocus
            
            Case WASH_DRUMSPEED_FIELD:
                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = Chr(13) & Chr(10) & _
                                                "Минимальное значение: " & DrumSpeed.MinValue & Chr(13) & Chr(10) & _
                                                "Максимальное значение: " & DrumSpeed.MaxValue & Chr(13) & Chr(10) & _
                                                "Значение по умолчанию: " & DrumSpeed.DefaultValue & Chr(13) & Chr(10) & _
                                                "Размерность: [" & DrumSpeed.Dimension & "]"
                    .LabelDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordWash.DrumSpeed
                .TextCell.Visible = True
                .TextCell.SetFocus
            
            Case Else

        End Select
    
        .TextCell.SelStart = 0
        .TextCell.SelLength = Len(.TextCell.Text)
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
        
        'CopyMemory RecordWash, ByVal StepPointer, STEP_SIZE_IN_BYTES
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
                If (RecordWash.Bits And &HF) = (.ComboCell.ListIndex And &HF) Then Exit Sub

                ' При выборе новой функции для шага мы должны обнулить структуру,
                ZeroMemory RecordWash, STEP_SIZE_IN_BYTES
                
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
                RecordWash.Bits = RecordWash.Bits Or (.ComboCell.ListIndex And &HF)
                
            ' Специальные параметры шага
            Case WASH_PAUSE_FIELD:
                Select Case .ComboCell.ListIndex
                    Case 0: RecordWash.Bits = RecordWash.Bits And &HFFEF
                    Case 1: RecordWash.Bits = RecordWash.Bits Or &H10
                End Select
    
            Case WASH_ROTATION_FIELD:
                Select Case .ComboCell.ListIndex
                    Case 0: RecordWash.Bits = RecordWash.Bits And &HFFDF
                    Case 1: RecordWash.Bits = RecordWash.Bits Or &H20
                End Select

            Case WASH_TIME_FIELD:
                RecordWash.Time = Val(.TextCell.Text)
    
            Case WASH_ROTATIONTIME_FIELD:
                RecordWash.RotationTime = Val(.TextCell.Text)
            
            Case WASH_PAUSETIME_FIELD:
                RecordWash.PauseTime = Val(.TextCell.Text)
            
            Case WASH_DRUMSPEED_FIELD:
                RecordWash.DrumSpeed = Val(.TextCell.Text)
            
            Case Else
    
        End Select
        
        .SetModified True
        
        ' Сохраняем изменения
        'CopyMemory ByVal StepPointer, RecordWash, STEP_SIZE_IN_BYTES
        
        'StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        'CopyMemory ByVal StepPointer, RecordTitle, HEADER_SIZE_IN_BYTES
    End With
    Exit Sub

ErrorHandler:
    Err.Clear
End Sub

Public Sub SetComboPropertyForWashOrRinsOrJolt(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordWash As TYPE_WPC_WASH
    
    func_SetComboPropertyForWashOrRinsOrJolt frm, 0&, RecordTitle, RecordWash
End Sub

