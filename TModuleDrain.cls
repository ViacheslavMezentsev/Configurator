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
Private Const DRAIN_ROTATIONTIME_FIELD = IDLE_FUNCTION_FIELD + 6
Private Const DRAIN_PAUSETIME_FIELD = IDLE_FUNCTION_FIELD + 7
Private Const DRAIN_DRUMSPEED1_FIELD = IDLE_FUNCTION_FIELD + 8
'Private Const DRAIN_DRUMSPEED2_FIELD = IDLE_FUNCTION_FIELD + 9
'Private Const DRAIN_TIME2_FIELD = IDLE_FUNCTION_FIELD + 10

Private Const DRAIN_PARAMETERS_COUNT = IDLE_PARAMETERS_COUNT + 8

' Дискретные параметры
Private Const DRAIN_PARAMETER_DESCR_PAUSE = "Разр. паузы"
Private Const DRAIN_PARAMETER_DESCR_DRAINGATE1 = "Разр. кл. слива 1"
Private Const DRAIN_PARAMETER_DESCR_DRAINGATE2 = "Разр. кл. слива 2"
Private Const DRAIN_PARAMETER_DESCR_ROTATION = "Разр. вращ."

' Аналоговые параметры
Private Const DRAIN_PARAMETER_DESCR_LEVEL = "Уровень моющ. раств."
Private Const DRAIN_PARAMETER_DESCR_ROTATIONTIME = "Время вращ. мотора"
Private Const DRAIN_PARAMETER_DESCR_PAUSETIME = "Время паузы вращ. мот."
Private Const DRAIN_PARAMETER_DESCR_DRUMSPEED1 = "Скорость вращ. бар. (реверс)"
'Private Const DRAIN_PARAMETER_DESCR_TIME2 = "Время раскладки"
'Private Const DRAIN_PARAMETER_DESCR_DRUMSPEED2 = "Скорость вращ. бар. (раскл.)"

' Название секции
Private Const DRAIN_SECTION_NAME = "Drain"

' Настройки по умолчанию
' Дискретные типы
Private Const PAUSE_DEFAULT = 0
Private Const DRAINGATE1_DEFAULT = 1
Private Const DRAINGATE2_DEFAULT = 0
Private Const ROTATION_DEFAULT = 1

' Аналоговые типы
Private Const LEVEL_MIN = 0
Private Const LEVEL_MAX = 99
Private Const LEVEL_DEFAULT = 0
Private Const LEVEL_DIMENSION = "см"

Private Const ROTATIONTIME_MIN = 1
Private Const ROTATIONTIME_MAX = 250
Private Const ROTATIONTIME_DEFAULT = 6
Private Const ROTATIONTIME_DIMENSION = "сек"

Private Const PAUSETIME_MIN = 1
Private Const PAUSETIME_MAX = 250
Private Const PAUSETIME_DEFAULT = 12
Private Const PAUSETIME_DIMENSION = "сек"

Private Const DRUMSPEED1_MIN = 30
Private Const DRUMSPEED1_MAX = 70
Private Const DRUMSPEED1_DEFAULT = 50
Private Const DRUMSPEED1_DIMENSION = "об/мин"

'Private Const DRUMSPEED2_MIN = 60
'Private Const DRUMSPEED2_MAX = 100
'Private Const DRUMSPEED2_DEFAULT = 70
'Private Const DRUMSPEED2_DIMENSION = "об/мин"

'Private Const TIME2_MIN = 1
'Private Const TIME2_MAX = 255
'Private Const TIME2_DEFAULT = 3
'Private Const TIME2_DIMENSION = "мин"

' Параметры функции шага
' Дискретные типы
Private Pause As TYPE_BOOL_DESCRIPTION
Private DrainGate1 As TYPE_BOOL_DESCRIPTION
Private DrainGate2 As TYPE_BOOL_DESCRIPTION
Private Rotation As TYPE_BOOL_DESCRIPTION

' Аналоговые типы
Private Level As TYPE_BYTE_DESCRIPTION
Private RotationTime As TYPE_BYTE_DESCRIPTION
Private PauseTime As TYPE_BYTE_DESCRIPTION
Private DrumSpeed1 As TYPE_BYTE_DESCRIPTION
'Private DrumSpeed2 As TYPE_INT_DESCRIPTION
'Private Time2 As TYPE_BYTE_DESCRIPTION

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
    Pause.DefaultValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "Pause.Default", PAUSE_DEFAULT) > 0
    DrainGate1.DefaultValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "DrainGate1.Default", DRAINGATE1_DEFAULT) > 0
    DrainGate2.DefaultValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "DrainGate2.Default", DRAINGATE2_DEFAULT) > 0
    Rotation.DefaultValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "Rotation.Default", ROTATION_DEFAULT) > 0
    
    ' Аналоговые типы
    With Level
        .MinValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "Level.Min", LEVEL_MIN)
        .MaxValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "Level.Max", LEVEL_MAX)
        .DefaultValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "Level.Default", LEVEL_DEFAULT)
        .Dimension = LimitsFile.ReadString(DRAIN_SECTION_NAME, "Level.Dimension", LEVEL_DIMENSION)
    End With
      
    With RotationTime
        .MinValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "RotationTime.Min", ROTATIONTIME_MIN)
        .MaxValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "RotationTime.Max", ROTATIONTIME_MAX)
        .DefaultValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "RotationTime.Default", ROTATIONTIME_DEFAULT)
        .Dimension = LimitsFile.ReadString(DRAIN_SECTION_NAME, "RotationTime.Dimension", ROTATIONTIME_DIMENSION)
    End With
    
    With PauseTime
        .MinValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "PauseTime.Min", PAUSETIME_MIN)
        .MaxValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "PauseTime.Max", PAUSETIME_MAX)
        .DefaultValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "PauseTime.Default", PAUSETIME_DEFAULT)
        .Dimension = LimitsFile.ReadString(DRAIN_SECTION_NAME, "PauseTime.Dimension", PAUSETIME_DIMENSION)
    End With
    
    With DrumSpeed1
        .MinValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "DrumSpeed1.Min", DRUMSPEED1_MIN)
        .MaxValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "DrumSpeed1.Max", DRUMSPEED1_MAX)
        .DefaultValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "DrumSpeed1.Default", DRUMSPEED1_DEFAULT)
        .Dimension = LimitsFile.ReadString(DRAIN_SECTION_NAME, "DrumSpeed1.Dimension", DRUMSPEED1_DIMENSION)
    End With
    
'    With DrumSpeed2
'        .MinValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "DrumSpeed2.Min", DRUMSPEED2_MIN)
'        .MaxValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "DrumSpeed2.Max", DRUMSPEED2_MAX)
'        .DefaultValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "DrumSpeed2.Default", DRUMSPEED2_DEFAULT)
'        .Dimension = LimitsFile.ReadString(DRAIN_SECTION_NAME, "DrumSpeed2.Dimension", DRUMSPEED2_DIMENSION)
'    End With
'
'    With Time2
'        .MinValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "Time2.Min", TIME2_MIN)
'        .MaxValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "Time2.Max", TIME2_MAX)
'        .DefaultValue = LimitsFile.ReadInteger(DRAIN_SECTION_NAME, "Time2.Default", TIME2_DEFAULT)
'        .Dimension = LimitsFile.ReadString(DRAIN_SECTION_NAME, "Time2.Dimension", TIME2_DIMENSION)
'    End With
        
    Set LimitsFile = Nothing
End Sub

' Функция-посредник
Private Sub func_SetDefaults(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordDrain As TYPE_WPC_DRAIN)

    Dim StepPointer As Long
    
    StepPointer = frm.Manager.DataPointer + _
        frm.Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
        HEADER_SIZE_IN_BYTES + _
        frm.Manager.StepIndex * STEP_SIZE_IN_BYTES
    
    'CopyMemory RecordDrain, ByVal StepPointer, STEP_SIZE_IN_BYTES
    PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer
    
    ' Дискретные типы
    Select Case Pause.DefaultValue
        Case False: RecordDrain.Bits = RecordDrain.Bits And &HFFEF
        Case True: RecordDrain.Bits = RecordDrain.Bits Or &H10
    End Select

    Select Case DrainGate1.DefaultValue
        Case False: RecordDrain.Bits = RecordDrain.Bits And &HFFDF
        Case True: RecordDrain.Bits = RecordDrain.Bits Or &H20
    End Select
    
    Select Case DrainGate2.DefaultValue
        Case False: RecordDrain.Bits = RecordDrain.Bits And &HFFBF
        Case True: RecordDrain.Bits = RecordDrain.Bits Or &H40
    End Select
    
    Select Case Rotation.DefaultValue
        Case False: RecordDrain.Bits = RecordDrain.Bits And &HFF7F
        Case True: RecordDrain.Bits = RecordDrain.Bits Or &H80
    End Select
    
    ' Аналоговые типы
    RecordDrain.Level = Level.DefaultValue
    RecordDrain.RotationTime = RotationTime.DefaultValue
    RecordDrain.PauseTime = PauseTime.DefaultValue
    RecordDrain.DrumSpeed1 = DrumSpeed1.DefaultValue
'    RecordDrain.DrumSpeed2 = DrumSpeed2.DefaultValue
'    RecordDrain.Time2 = Time2.DefaultValue
    
    ' Сохраняем изменения
    'StepPointer = frm.Manager.DataPointer + frm.Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
    'CopyMemory ByVal StepPointer, RecordDrain, HEADER_SIZE_IN_BYTES
End Sub

Public Sub SetDefaults(frm As FormMain)
    Dim RecordDrain As TYPE_WPC_DRAIN
    
    func_SetDefaults frm, 0&, RecordDrain
End Sub

' Функция-посредник
Private Sub func_ShowPropertyTableForDrain(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordDrain As TYPE_WPC_DRAIN)
    
    Dim I, J, row As Integer
    Dim ParamStr, s As String
    Dim StepPointer As Long
    
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
                
                Case DRAIN_ROTATIONTIME_FIELD:
                    ParamStr = ParamStr & DRAIN_PARAMETER_DESCR_ROTATIONTIME
                
                Case DRAIN_PAUSETIME_FIELD:
                    ParamStr = ParamStr & DRAIN_PARAMETER_DESCR_PAUSETIME
                
                Case DRAIN_DRUMSPEED1_FIELD:
                    ParamStr = ParamStr & DRAIN_PARAMETER_DESCR_DRUMSPEED1
                    
'                Case DRAIN_TIME2_FIELD:
'                    ParamStr = ParamStr & DRAIN_PARAMETER_DESCR_TIME2
'
'                Case DRAIN_DRUMSPEED2_FIELD:
'                    ParamStr = ParamStr & DRAIN_PARAMETER_DESCR_DRUMSPEED2
                    
                Case Else
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_UNKNOWN
            End Select
            
            If (I < DRAIN_PARAMETERS_COUNT) Then ParamStr = ParamStr & "|"
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
        
        'CopyMemory RecordDrain, ByVal StepPointer, STEP_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer
        
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
                    If LimitsLoaded Then
                        If RecordDrain.Level < Level.MinValue Or RecordDrain.Level > Level.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordDrain.Level
                    
                Case DRAIN_ROTATIONTIME_FIELD:
                    If LimitsLoaded Then
                        If RecordDrain.RotationTime < RotationTime.MinValue Or RecordDrain.RotationTime > RotationTime.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordDrain.RotationTime
                
                Case DRAIN_PAUSETIME_FIELD:
                    If LimitsLoaded Then
                        If RecordDrain.PauseTime < PauseTime.MinValue Or RecordDrain.PauseTime > PauseTime.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordDrain.PauseTime
                
                Case DRAIN_DRUMSPEED1_FIELD:
                    If LimitsLoaded Then
                        If RecordDrain.DrumSpeed1 < DrumSpeed1.MinValue Or RecordDrain.DrumSpeed1 > DrumSpeed1.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordDrain.DrumSpeed1
                
'                Case DRAIN_TIME2_FIELD:
'                    If LimitsLoaded Then
'                        If RecordDrain.Time2 < Time2.MinValue Or RecordDrain.Time2 > Time2.MaxValue Then
'                            .PropertyTable.CellBackColor = &H8080FF
'                        Else
'                            .PropertyTable.CellBackColor = &H80000005
'                        End If
'                    End If
'                    .PropertyTable.Text = "" & RecordDrain.Time2
'
'                Case DRAIN_DRUMSPEED2_FIELD:
'                    If LimitsLoaded Then
'                        If RecordDrain.DrumSpeed2 < DrumSpeed2.MinValue Or RecordDrain.DrumSpeed2 > DrumSpeed2.MaxValue Then
'                            .PropertyTable.CellBackColor = &H8080FF
'                        Else
'                            .PropertyTable.CellBackColor = &H80000005
'                        End If
'                    End If
'                    .PropertyTable.Text = "" & RecordDrain.DrumSpeed2
                
                Case Else
                    .PropertyTable.CellBackColor = &H8000000F
                    
            End Select
        Next I
    End With
End Sub

Public Sub ShowPropertyTableForDrain(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordDrain As TYPE_WPC_DRAIN
    
    func_ShowPropertyTableForDrain frm, 0&, RecordTitle, RecordDrain
End Sub

' Функция-посредник
Private Sub func_EditPropertyForDrain(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordDrain As TYPE_WPC_DRAIN)
    
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
        
        'CopyMemory RecordDrain, ByVal StepPointer, STEP_SIZE_IN_BYTES
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
                For I = 0 To NUMBER_OF_FUNCS - 1
                    .ComboCell.AddItem (FunctionsStrings(I))
                Next I
                .ComboCell.ListIndex = RecordDrain.Bits And &HF
                .ComboCell.Visible = True
                .ComboCell.SetFocus
                
            ' Специальные параметры шага
            Case DRAIN_PAUSE_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordDrain.Bits And &H10) / &H10
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case DRAIN_DRAINGATE1_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordDrain.Bits And &H20) / &H20
                .ComboCell.Visible = True
                .ComboCell.SetFocus
                
            Case DRAIN_DRAINGATE2_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordDrain.Bits And &H40) / &H40
                .ComboCell.Visible = True
                .ComboCell.SetFocus
                
            Case DRAIN_ROTATION_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordDrain.Bits And &H80) / &H80
                .ComboCell.Visible = True
                .ComboCell.SetFocus
                
            Case DRAIN_LEVEL_FIELD:
                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = Chr(13) & Chr(10) & _
                                                "Минимальное значение: " & Level.MinValue & Chr(13) & Chr(10) & _
                                                "Максимальное значение: " & Level.MaxValue & Chr(13) & Chr(10) & _
                                                "Значение по умолчанию: " & Level.DefaultValue & Chr(13) & Chr(10) & _
                                                "Размерность: [" & Level.Dimension & "]"
                    .LabelDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordDrain.Level
                .TextCell.Visible = True
                .TextCell.SetFocus
            
            Case DRAIN_ROTATIONTIME_FIELD:
                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = Chr(13) & Chr(10) & _
                                                "Минимальное значение: " & RotationTime.MinValue & Chr(13) & Chr(10) & _
                                                "Максимальное значение: " & RotationTime.MaxValue & Chr(13) & Chr(10) & _
                                                "Значение по умолчанию: " & RotationTime.DefaultValue & Chr(13) & Chr(10) & _
                                                "Размерность: [" & RotationTime.Dimension & "]"
                    .LabelDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordDrain.RotationTime
                .TextCell.Visible = True
                .TextCell.SetFocus
            
            Case DRAIN_PAUSETIME_FIELD:
                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = Chr(13) & Chr(10) & _
                                                "Минимальное значение: " & PauseTime.MinValue & Chr(13) & Chr(10) & _
                                                "Максимальное значение: " & PauseTime.MaxValue & Chr(13) & Chr(10) & _
                                                "Значение по умолчанию: " & PauseTime.DefaultValue & Chr(13) & Chr(10) & _
                                                "Размерность: [" & PauseTime.Dimension & "]"
                    .LabelDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordDrain.PauseTime
                .TextCell.Visible = True
                .TextCell.SetFocus
            
            Case DRAIN_DRUMSPEED1_FIELD:
                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = Chr(13) & Chr(10) & _
                                                "Минимальное значение: " & DrumSpeed1.MinValue & Chr(13) & Chr(10) & _
                                                "Максимальное значение: " & DrumSpeed1.MaxValue & Chr(13) & Chr(10) & _
                                                "Значение по умолчанию: " & DrumSpeed1.DefaultValue & Chr(13) & Chr(10) & _
                                                "Размерность: [" & DrumSpeed1.Dimension & "]"
                    .LabelDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordDrain.DrumSpeed1
                .TextCell.Visible = True
                .TextCell.SetFocus
            
'            Case DRAIN_TIME2_FIELD:
'                If LimitsLoaded Then
'                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
'                    .LabelDescription.Caption = Chr(13) & Chr(10) & _
'                                                "Минимальное значение: " & Time2.MinValue & Chr(13) & Chr(10) & _
'                                                "Максимальное значение: " & Time2.MaxValue & Chr(13) & Chr(10) & _
'                                                "Значение по умолчанию: " & Time2.DefaultValue & Chr(13) & Chr(10) & _
'                                                "Размерность: [" & Time2.Dimension & "]"
'                    .LabelDescription.Visible = True
'                End If
'                .TextCell.Text = "" & RecordDrain.Time2
'                .TextCell.Visible = True
'                .TextCell.SetFocus
'
'            Case DRAIN_DRUMSPEED2_FIELD:
'                If LimitsLoaded Then
'                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
'                    .LabelDescription.Caption = Chr(13) & Chr(10) & _
'                                                "Минимальное значение: " & DrumSpeed2.MinValue & Chr(13) & Chr(10) & _
'                                                "Максимальное значение: " & DrumSpeed2.MaxValue & Chr(13) & Chr(10) & _
'                                                "Значение по умолчанию: " & DrumSpeed2.DefaultValue & Chr(13) & Chr(10) & _
'                                                "Размерность: [" & DrumSpeed2.Dimension & "]"
'                    .LabelDescription.Visible = True
'                End If
'                .TextCell.Text = "" & RecordDrain.DrumSpeed2
'                .TextCell.Visible = True
'                .TextCell.SetFocus
            
            Case Else

        End Select
    
        .TextCell.SelStart = 0
        .TextCell.SelLength = Len(.TextCell.Text)
    End With
End Sub

Public Sub EditPropertyForDrain(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordDrain As TYPE_WPC_DRAIN
    
    func_EditPropertyForDrain frm, 0&, RecordTitle, RecordDrain
End Sub

' Функция-посредник
Private Sub func_SetComboPropertyForDrain(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordDrain As TYPE_WPC_DRAIN)
    
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
        
        'CopyMemory RecordDrain, ByVal StepPointer, STEP_SIZE_IN_BYTES
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
                If (RecordDrain.Bits And &HF) = (.ComboCell.ListIndex And &HF) Then Exit Sub

                ' При выборе новой функции для шага мы должны обнулить структуру,
                ZeroMemory RecordDrain, STEP_SIZE_IN_BYTES
                
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
                RecordDrain.Bits = RecordDrain.Bits Or (.ComboCell.ListIndex And &HF)
                
            ' Специальные параметры шага
            Case DRAIN_PAUSE_FIELD:
                Select Case .ComboCell.ListIndex
                    Case 0: RecordDrain.Bits = RecordDrain.Bits And &HFFEF
                    Case 1: RecordDrain.Bits = RecordDrain.Bits Or &H10
                End Select
    
            Case DRAIN_DRAINGATE1_FIELD:
                Select Case .ComboCell.ListIndex
                    Case 0: RecordDrain.Bits = RecordDrain.Bits And &HFFDF
                    Case 1: RecordDrain.Bits = RecordDrain.Bits Or &H20
                End Select
            
            Case DRAIN_DRAINGATE2_FIELD:
                Select Case .ComboCell.ListIndex
                    Case 0: RecordDrain.Bits = RecordDrain.Bits And &HFFBF
                    Case 1: RecordDrain.Bits = RecordDrain.Bits Or &H40
                End Select
            
            Case DRAIN_ROTATION_FIELD:
                Select Case .ComboCell.ListIndex
                    Case 0: RecordDrain.Bits = RecordDrain.Bits And &HFF7F
                    Case 1: RecordDrain.Bits = RecordDrain.Bits Or &H80
                End Select

            Case DRAIN_LEVEL_FIELD:
                RecordDrain.Level = Val(.TextCell.Text)
    
            Case DRAIN_ROTATIONTIME_FIELD:
                RecordDrain.RotationTime = Val(.TextCell.Text)
            
            Case DRAIN_PAUSETIME_FIELD:
                RecordDrain.PauseTime = Val(.TextCell.Text)
            
            Case DRAIN_DRUMSPEED1_FIELD:
                RecordDrain.DrumSpeed1 = Val(.TextCell.Text)

'            Case DRAIN_TIME2_FIELD:
'                RecordDrain.Time2 = Val(.TextCell.Text)
'
'            Case DRAIN_DRUMSPEED2_FIELD:
'                RecordDrain.DrumSpeed2 = Val(.TextCell.Text)
            
            Case Else
    
        End Select
        
        .SetModified True
        
        ' Сохраняем изменения
        'CopyMemory ByVal StepPointer, RecordDrain, STEP_SIZE_IN_BYTES
        
        'StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        'CopyMemory ByVal StepPointer, RecordTitle, HEADER_SIZE_IN_BYTES
    End With
    Exit Sub

ErrorHandler:
    Err.Clear
End Sub

Public Sub SetComboPropertyForDrain(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordDrain As TYPE_WPC_DRAIN
    
    func_SetComboPropertyForDrain frm, 0&, RecordTitle, RecordDrain
End Sub

' Функция-посредник
Private Function func_ValveEnabled(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordDrain As TYPE_WPC_DRAIN, _
    ByVal StepIndex As Integer, _
    ByVal Num As Integer)

    Dim I As Integer
    Dim StepPointer As Long
    
    With frm
        StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        'CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            StepIndex * STEP_SIZE_IN_BYTES
        
        'CopyMemory RecordCool, ByVal StepPointer, STEP_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer
        
        Select Case Num
            Case 1:
            Case 2:
            Case 3:
            Case 4:
            Case 5:
            Case 6:
            Case 7:
            Case 8:
            Case 9:
            Case 10:
            Case 11:
            Case 12:
            Case 13:
            Case 14:
            Case 15: func_ValveEnabled = (RecordDrain.Bits And &H20) > 0
            Case 16: func_ValveEnabled = (RecordDrain.Bits And &H40) > 0
            Case 17:
            Case 18: func_ValveEnabled = (RecordDrain.Bits And &H80) > 0
            
            Case Else
                func_ValveEnabled = False
        End Select
    End With
End Function

Public Function ValveEnabled(frm As FormMain, StepIndex As Integer, Num As Integer)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordDrain As TYPE_WPC_DRAIN
    
    ValveEnabled = func_ValveEnabled(frm, 0&, RecordTitle, RecordDrain, StepIndex, Num)
End Function

