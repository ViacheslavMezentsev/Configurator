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
Private Const COOL_ROTATIONTIME_FIELD = IDLE_FUNCTION_FIELD + 6
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
Private Const COOL_PARAMETER_DESCR_ROTATIONTIME = "Время вращ. мотора"
Private Const COOL_PARAMETER_DESCR_PAUSETIME = "Время паузы вращ. мот."
Private Const COOL_PARAMETER_DESCR_DRUMSPEED = "Скорость вращ. барабана"

' Название секции
Private Const COOL_SECTION_NAME = "Cool"

' Настройки по умолчанию
' Дискретные типы
Private Const PAUSE_DEFAULT = 0
Private Const FAST_DEFAULT = 0
Private Const ROTATION_DEFAULT = 1

' Аналоговые типы
Private Const TEMPERATURE_MIN = 0
Private Const TEMPERATURE_MAX = 90
Private Const TEMPERATURE_DEFAULT = 1
Private Const TEMPERATURE_DIMENSION = "градусы цельсия"

Private Const COLDWATERTIME_MIN = 1
Private Const COLDWATERTIME_MAX = 29
Private Const COLDWATERTIME_DEFAULT = 10
Private Const COLDWATERTIME_DIMENSION = "сек"

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
Private Fast As TYPE_BOOL_DESCRIPTION
Private Rotation As TYPE_BOOL_DESCRIPTION

' Аналоговые типы
Private Temperature As TYPE_BYTE_DESCRIPTION
Private ColdWaterTime As TYPE_BYTE_DESCRIPTION
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
    Pause.DefaultValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "Pause.Default", PAUSE_DEFAULT) > 0
    Fast.DefaultValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "Fast.Default", FAST_DEFAULT) > 0
    Rotation.DefaultValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "Rotation.Default", ROTATION_DEFAULT) > 0
    
    ' Аналоговые типы
    With Temperature
        .MinValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "Temperature.Min", TEMPERATURE_MIN)
        .MaxValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "Temperature.Max", TEMPERATURE_MAX)
        .DefaultValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "Temperature.Default", TEMPERATURE_DEFAULT)
        .Dimension = LimitsFile.ReadString(COOL_SECTION_NAME, "Temperature.Dimension", TEMPERATURE_DIMENSION)
    End With
    
    With ColdWaterTime
        .MinValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "ColdWaterTime.Min", COLDWATERTIME_MIN)
        .MaxValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "ColdWaterTime.Max", COLDWATERTIME_MAX)
        .DefaultValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "ColdWaterTime.Default", COLDWATERTIME_DEFAULT)
        .Dimension = LimitsFile.ReadString(COOL_SECTION_NAME, "ColdWaterTime.Dimension", COLDWATERTIME_DIMENSION)
    End With
    
    With RotationTime
        .MinValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "RotationTime.Min", ROTATIONTIME_MIN)
        .MaxValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "RotationTime.Max", ROTATIONTIME_MAX)
        .DefaultValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "RotationTime.Default", ROTATIONTIME_DEFAULT)
        .Dimension = LimitsFile.ReadString(COOL_SECTION_NAME, "RotationTime.Dimension", ROTATIONTIME_DIMENSION)
    End With
    
    With PauseTime
        .MinValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "PauseTime.Min", PAUSETIME_MIN)
        .MaxValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "PauseTime.Max", PAUSETIME_MAX)
        .DefaultValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "PauseTime.Default", PAUSETIME_DEFAULT)
        .Dimension = LimitsFile.ReadString(COOL_SECTION_NAME, "PauseTime.Dimension", PAUSETIME_DIMENSION)
    End With
    
    With DrumSpeed
        .MinValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "DrumSpeed.Min", DRUMSPEED_MIN)
        .MaxValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "DrumSpeed.Max", DRUMSPEED_MAX)
        .DefaultValue = LimitsFile.ReadInteger(COOL_SECTION_NAME, "DrumSpeed.Default", DRUMSPEED_DEFAULT)
        .Dimension = LimitsFile.ReadString(COOL_SECTION_NAME, "DrumSpeed.Dimension", DRUMSPEED_DIMENSION)
    End With
    
    Set LimitsFile = Nothing
End Sub

' Функция-посредник
Private Sub func_SetDefaults(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordCool As TYPE_WPC_COOL)

    Dim StepPointer As Long
    
    StepPointer = frm.Manager.DataPointer + _
        frm.Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
        HEADER_SIZE_IN_BYTES + _
        frm.Manager.StepIndex * STEP_SIZE_IN_BYTES
    
    'CopyMemory RecordCool, ByVal StepPointer, STEP_SIZE_IN_BYTES
    PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer
    
    ' Дискретные типы
    Select Case Pause.DefaultValue
        Case False: RecordCool.Bits = RecordCool.Bits And &HFFEF
        Case True: RecordCool.Bits = RecordCool.Bits Or &H10
    End Select

    Select Case Fast.DefaultValue
        Case False: RecordCool.Bits = RecordCool.Bits And &HFFDF
        Case True: RecordCool.Bits = RecordCool.Bits Or &H20
    End Select
    
    Select Case Rotation.DefaultValue
        Case False: RecordCool.Bits = RecordCool.Bits And &HFFBF
        Case True: RecordCool.Bits = RecordCool.Bits Or &H40
    End Select
    
    ' Аналоговые типы
    RecordCool.Temperature = Temperature.DefaultValue
    RecordCool.ColdWaterTime = ColdWaterTime.DefaultValue
    RecordCool.RotationTime = RotationTime.DefaultValue
    RecordCool.PauseTime = PauseTime.DefaultValue
    RecordCool.DrumSpeed = DrumSpeed.DefaultValue
    
    ' Сохраняем изменения
    'StepPointer = frm.Manager.DataPointer + frm.Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
    'CopyMemory ByVal StepPointer, RecordCool, HEADER_SIZE_IN_BYTES
End Sub

Public Sub SetDefaults(frm As FormMain)
    Dim RecordCool As TYPE_WPC_COOL
    
    func_SetDefaults frm, 0&, RecordCool
End Sub

' Функция-посредник
Private Sub func_ShowPropertyTableForCool(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordCool As TYPE_WPC_COOL)
    
    Dim I, J, Row As Integer
    Dim ParamStr, s As String
    Dim StepPointer As Long
    
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
                
                Case COOL_ROTATIONTIME_FIELD:
                    ParamStr = ParamStr & COOL_PARAMETER_DESCR_ROTATIONTIME
                
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
        'CopyMemory RecordTitle, ByVal StepPointer, HEADER_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer
        
        StepPointer = .Manager.DataPointer + _
            .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
            HEADER_SIZE_IN_BYTES + _
            .Manager.StepIndex * STEP_SIZE_IN_BYTES
        
        'CopyMemory RecordCool, ByVal StepPointer, STEP_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer
        
        .PropertyTable.Col = 1
        For I = 1 To COOL_PARAMETERS_COUNT
            .PropertyTable.Row = I
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
                    If LimitsLoaded Then
                        If RecordCool.Temperature < Temperature.MinValue Or RecordCool.Temperature > Temperature.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordCool.Temperature
                    
                Case COOL_COLDWATERTIME_FIELD:
                    If LimitsLoaded Then
                        If RecordCool.ColdWaterTime < ColdWaterTime.MinValue Or RecordCool.ColdWaterTime > ColdWaterTime.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordCool.ColdWaterTime
                    
                Case COOL_ROTATIONTIME_FIELD:
                    If LimitsLoaded Then
                        If RecordCool.RotationTime < RotationTime.MinValue Or RecordCool.RotationTime > RotationTime.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordCool.RotationTime
                
                Case COOL_PAUSETIME_FIELD:
                    If LimitsLoaded Then
                        If RecordCool.PauseTime < PauseTime.MinValue Or RecordCool.PauseTime > PauseTime.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordCool.PauseTime
                
                Case COOL_DRUMSPEED_FIELD:
                    If LimitsLoaded Then
                        If RecordCool.DrumSpeed < DrumSpeed.MinValue Or RecordCool.DrumSpeed > DrumSpeed.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordCool.DrumSpeed
                
                Case Else
                    .PropertyTable.CellBackColor = &H8000000F
                    
            End Select
        Next I
    End With
End Sub
    
Public Sub ShowPropertyTableForCool(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordCool As TYPE_WPC_COOL
    
    func_ShowPropertyTableForCool frm, 0&, RecordTitle, RecordCool
End Sub

' Функция-посредник
Private Sub func_EditPropertyForCool(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordCool As TYPE_WPC_COOL)

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
        
        'CopyMemory RecordCool, ByVal StepPointer, STEP_SIZE_IN_BYTES
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
        
        Select Case .PropertyTable.Row
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
                .ComboCell.ListIndex = RecordCool.Bits And &HF
                .ComboCell.Visible = True
                .ComboCell.SetFocus
                
            ' Специальные параметры шага
            Case COOL_PAUSE_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordCool.Bits And &H10) / &H10
                .ComboCell.Visible = True
                .ComboCell.SetFocus
                
            Case COOL_FAST_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordCool.Bits And &H20) / &H20
                .ComboCell.Visible = True
                .ComboCell.SetFocus
                
            Case COOL_ROTATION_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordCool.Bits And &H40) / &H40
                .ComboCell.Visible = True
                .ComboCell.SetFocus
                
            Case COOL_TEMPERATURE_FIELD:
                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = Chr(13) & Chr(10) & _
                                                "Минимальное значение: " & Temperature.MinValue & Chr(13) & Chr(10) & _
                                                "Максимальное значение: " & Temperature.MaxValue & Chr(13) & Chr(10) & _
                                                "Значение по умолчанию: " & Temperature.DefaultValue & Chr(13) & Chr(10) & _
                                                "Размерность: [" & Temperature.Dimension & "]"
                    .LabelDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordCool.Temperature
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case COOL_COLDWATERTIME_FIELD:
                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = Chr(13) & Chr(10) & _
                                                "Минимальное значение: " & ColdWaterTime.MinValue & Chr(13) & Chr(10) & _
                                                "Максимальное значение: " & ColdWaterTime.MaxValue & Chr(13) & Chr(10) & _
                                                "Значение по умолчанию: " & ColdWaterTime.DefaultValue & Chr(13) & Chr(10) & _
                                                "Размерность: [" & ColdWaterTime.Dimension & "]"
                    .LabelDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordCool.ColdWaterTime
                .TextCell.Visible = True
                .TextCell.SetFocus
                
            Case COOL_ROTATIONTIME_FIELD:
                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = Chr(13) & Chr(10) & _
                                                "Минимальное значение: " & RotationTime.MinValue & Chr(13) & Chr(10) & _
                                                "Максимальное значение: " & RotationTime.MaxValue & Chr(13) & Chr(10) & _
                                                "Значение по умолчанию: " & RotationTime.DefaultValue & Chr(13) & Chr(10) & _
                                                "Размерность: [" & RotationTime.Dimension & "]"
                    .LabelDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordCool.RotationTime
                .TextCell.Visible = True
                .TextCell.SetFocus
                
            Case COOL_PAUSETIME_FIELD:
                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = Chr(13) & Chr(10) & _
                                                "Минимальное значение: " & PauseTime.MinValue & Chr(13) & Chr(10) & _
                                                "Максимальное значение: " & PauseTime.MaxValue & Chr(13) & Chr(10) & _
                                                "Значение по умолчанию: " & PauseTime.DefaultValue & Chr(13) & Chr(10) & _
                                                "Размерность: [" & PauseTime.Dimension & "]"
                    .LabelDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordCool.PauseTime
                .TextCell.Visible = True
                .TextCell.SetFocus
                
            Case COOL_DRUMSPEED_FIELD:
                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = Chr(13) & Chr(10) & _
                                                "Минимальное значение: " & DrumSpeed.MinValue & Chr(13) & Chr(10) & _
                                                "Максимальное значение: " & DrumSpeed.MaxValue & Chr(13) & Chr(10) & _
                                                "Значение по умолчанию: " & DrumSpeed.DefaultValue & Chr(13) & Chr(10) & _
                                                "Размерность: [" & DrumSpeed.Dimension & "]"
                    .LabelDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordCool.DrumSpeed
                .TextCell.Visible = True
                .TextCell.SetFocus
                
            Case Else

        End Select
    
        .TextCell.SelStart = 0
        .TextCell.SelLength = Len(.TextCell.Text)
    End With
End Sub

Public Sub EditPropertyForCool(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordCool As TYPE_WPC_COOL
    
    func_EditPropertyForCool frm, 0&, RecordTitle, RecordCool
End Sub

' Функция-посредник
Private Sub func_SetComboPropertyForCool(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordCool As TYPE_WPC_COOL)
    
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
        
        'CopyMemory RecordCool, ByVal StepPointer, STEP_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer
    
        Select Case .PropertyTable.Row
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
                If (RecordCool.Bits And &HF) = (.ComboCell.ListIndex And &HF) Then Exit Sub
                
                ' При выборе новой функции для шага мы должны обнулить структуру,
                ZeroMemory RecordCool, STEP_SIZE_IN_BYTES

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
                RecordCool.Bits = RecordCool.Bits Or (.ComboCell.ListIndex And &HF)
                
            ' Специальные параметры шага
            Case COOL_PAUSE_FIELD:
                Select Case .ComboCell.ListIndex
                    Case 0: RecordCool.Bits = RecordCool.Bits And &HFFEF
                    Case 1: RecordCool.Bits = RecordCool.Bits Or &H10
                End Select
                
            Case COOL_FAST_FIELD:
                Select Case .ComboCell.ListIndex
                    Case 0: RecordCool.Bits = RecordCool.Bits And &HFFDF
                    Case 1: RecordCool.Bits = RecordCool.Bits Or &H20
                End Select
                
            Case COOL_ROTATION_FIELD:
                Select Case .ComboCell.ListIndex
                    Case 0: RecordCool.Bits = RecordCool.Bits And &HFFBF
                    Case 1: RecordCool.Bits = RecordCool.Bits Or &H40
                End Select
                
            Case COOL_TEMPERATURE_FIELD:
                RecordCool.Temperature = Val(.TextCell.Text)
            
            Case COOL_COLDWATERTIME_FIELD:
                RecordCool.ColdWaterTime = Val(.TextCell.Text)
    
            Case COOL_ROTATIONTIME_FIELD:
                RecordCool.RotationTime = Val(.TextCell.Text)
            
            Case COOL_PAUSETIME_FIELD:
                RecordCool.PauseTime = Val(.TextCell.Text)
            
            Case COOL_DRUMSPEED_FIELD:
                RecordCool.DrumSpeed = Val(.TextCell.Text)
            
            Case Else
    
        End Select
        
        .SetModified True
        
        ' Сохраняем изменения
        'CopyMemory ByVal StepPointer, RecordCool, STEP_SIZE_IN_BYTES
        
        'StepPointer = .Manager.DataPointer + .Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        'CopyMemory ByVal StepPointer, RecordTitle, HEADER_SIZE_IN_BYTES
    End With
    Exit Sub

ErrorHandler:
    Err.Clear
End Sub

Public Sub SetComboPropertyForCool(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordCool As TYPE_WPC_COOL
    
    func_SetComboPropertyForCool frm, 0&, RecordTitle, RecordCool
End Sub

' Функция-посредник
Private Function func_ValveEnabled(frm As FormMain, _
    ByVal begin_of_pointers As Long, _
    ByRef RecordTitle As TYPE_WPC_TITLE, _
    ByRef RecordCool As TYPE_WPC_COOL, _
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
            Case 2: func_ValveEnabled = (RecordCool.Bits And &H20) > 0
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
            Case 15:
            Case 16:
            Case 17:
            Case 18: func_ValveEnabled = (RecordCool.Bits And &H40) > 0
            
            Case Else
                func_ValveEnabled = False
        End Select
    End With
End Function

Public Function ValveEnabled(frm As FormMain, StepIndex As Integer, Num As Integer)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordCool As TYPE_WPC_COOL
    
    ValveEnabled = func_ValveEnabled(frm, 0&, RecordTitle, RecordCool, StepIndex, Num)
End Function


