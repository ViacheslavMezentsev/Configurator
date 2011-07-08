VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CModuleDTRG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Const DTRG_PAUSE_FIELD = IDLE_FUNCTION_FIELD + 1
Const DTRG_ROTATION_FIELD = IDLE_FUNCTION_FIELD + 2

Const DTRG_DETERGENT1_FIELD = IDLE_FUNCTION_FIELD + 3
Const DTRG_DETERGENT2_FIELD = IDLE_FUNCTION_FIELD + 4
Const DTRG_DETERGENT3_FIELD = IDLE_FUNCTION_FIELD + 5
Const DTRG_DETERGENT4_FIELD = IDLE_FUNCTION_FIELD + 6
Const DTRG_DETERGENT5_FIELD = IDLE_FUNCTION_FIELD + 7
Const DTRG_DETERGENT6_FIELD = IDLE_FUNCTION_FIELD + 8
Const DTRG_DETERGENT7_FIELD = IDLE_FUNCTION_FIELD + 9
Const DTRG_DETERGENT8_FIELD = IDLE_FUNCTION_FIELD + 10
Const DTRG_DETERGENT9_FIELD = IDLE_FUNCTION_FIELD + 11

Const DTRG_ROTATIONTIME_FIELD = IDLE_FUNCTION_FIELD + 12
Const DTRG_PAUSETIME_FIELD = IDLE_FUNCTION_FIELD + 13
Const DTRG_DRUMSPEED_FIELD = IDLE_FUNCTION_FIELD + 14

Const DTRG_PARAMETERS_COUNT = IDLE_PARAMETERS_COUNT + 14

' Дискретные параметры
Const DTRG_PARAMETER_DESCR_PAUSE = "Разр. паузы"
Const DTRG_PARAMETER_DESCR_ROTATION = "Разр. вращ."

' Аналоговые параметры
Const DTRG_PARAMETER_DESCR_DETERGENT1 = "Время подачи моющ. 1"
Const DTRG_PARAMETER_DESCR_DETERGENT2 = "Время подачи моющ. 2"
Const DTRG_PARAMETER_DESCR_DETERGENT3 = "Время подачи моющ. 3"
Const DTRG_PARAMETER_DESCR_DETERGENT4 = "Время подачи моющ. 4"
Const DTRG_PARAMETER_DESCR_DETERGENT5 = "Время подачи моющ. 5"
Const DTRG_PARAMETER_DESCR_DETERGENT6 = "Время подачи моющ. 6"
Const DTRG_PARAMETER_DESCR_DETERGENT7 = "Время подачи моющ. 7"
Const DTRG_PARAMETER_DESCR_DETERGENT8 = "Время подачи моющ. 8"
Const DTRG_PARAMETER_DESCR_DETERGENT9 = "Время подачи моющ. 9"

Const DTRG_PARAMETER_DESCR_ROTATIONTIME = "Время вращ. мотора"
Const DTRG_PARAMETER_DESCR_PAUSETIME = "Время паузы вращ. мот."
Const DTRG_PARAMETER_DESCR_DRUMSPEED = "Скорость вращ. барабана"

' Название секции
Const DTRG_SECTION_NAME = "DTRG"

' Настройки по умолчанию
' Дискретные типы
Const PAUSE_DEFAULT = False
Const ROTATION_DEFAULT = True

' Аналоговые типы
Const DETERGENT_MIN = 0
Const DETERGENT_MAX = 255
Const DETERGENT_DEFAULT = 0
Const DETERGENT_DIMENSION = "сек"

Const ROTATIONTIME_MIN = 1
Const ROTATIONTIME_MAX = 250
Const ROTATIONTIME_DEFAULT = 6
Const ROTATIONTIME_DIMENSION = "сек"

Const PAUSETIME_MIN = 1
Const PAUSETIME_MAX = 250
Const PAUSETIME_DEFAULT = 12
Const PAUSETIME_DIMENSION = "сек"

Const DRUMSPEED_MIN = 30
Const DRUMSPEED_MAX = 70
Const DRUMSPEED_DEFAULT = 50
Const DRUMSPEED_DIMENSION = "об/мин"

' Параметры функции шага
' Дискретные типы
Private Pause As TYPE_BOOL_DESCRIPTION
Private Rotation As TYPE_BOOL_DESCRIPTION

' Аналоговые типы
Private Detergent_1_Time As TYPE_BYTE_DESCRIPTION
Private Detergent_2_Time As TYPE_BYTE_DESCRIPTION
Private Detergent_3_Time As TYPE_BYTE_DESCRIPTION
Private Detergent_4_Time As TYPE_BYTE_DESCRIPTION
Private Detergent_5_Time As TYPE_BYTE_DESCRIPTION
Private Detergent_6_Time As TYPE_BYTE_DESCRIPTION
Private Detergent_7_Time As TYPE_BYTE_DESCRIPTION
Private Detergent_8_Time As TYPE_BYTE_DESCRIPTION
Private Detergent_9_Time As TYPE_BYTE_DESCRIPTION
Private RotationTime As TYPE_BYTE_DESCRIPTION
Private PauseTime As TYPE_BYTE_DESCRIPTION
Private DrumSpeed As TYPE_BYTE_DESCRIPTION

Public LimitsLoaded As Boolean

' Конструктор
Private Sub Class_Initialize()
    LimitsLoaded = False
End Sub

Public Sub LoadLimits(FileName As String)
    LimitsLoaded = DoesFileExist(FileName)

    If Not LimitsLoaded Then Exit Sub

    Dim LimitsFile As New CIniFiles

    LimitsFile.Create FileName

    ' Настройки функции
    ' Дискретные типы
    Pause.DefaultValue = LimitsFile.ReadBoolean(DTRG_SECTION_NAME, "Pause.Default", PAUSE_DEFAULT)
    Rotation.DefaultValue = LimitsFile.ReadBoolean(DTRG_SECTION_NAME, "Rotation.Default", ROTATION_DEFAULT)

    ' Аналоговые типы

    With Detergent_1_Time
        .MinValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_1_Time.Min", DETERGENT_MIN)
        .MaxValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_1_Time.Max", DETERGENT_MAX)
        .DefaultValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_1_Time.Default", DETERGENT_DEFAULT)
        .Dimension = LimitsFile.ReadString(DTRG_SECTION_NAME, "Detergent_1_Time.Dimension", DETERGENT_DIMENSION)
    End With

    With Detergent_2_Time
        .MinValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_2_Time.Min", DETERGENT_MIN)
        .MaxValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_2_Time.Max", DETERGENT_MAX)
        .DefaultValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_2_Time.Default", DETERGENT_DEFAULT)
        .Dimension = LimitsFile.ReadString(DTRG_SECTION_NAME, "Detergent_2_Time.Dimension", DETERGENT_DIMENSION)
    End With

    With Detergent_3_Time
        .MinValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_3_Time.Min", DETERGENT_MIN)
        .MaxValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_3_Time.Max", DETERGENT_MAX)
        .DefaultValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_3_Time.Default", DETERGENT_DEFAULT)
        .Dimension = LimitsFile.ReadString(DTRG_SECTION_NAME, "Detergent_3_Time.Dimension", DETERGENT_DIMENSION)
    End With

    With Detergent_4_Time
        .MinValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_4_Time.Min", DETERGENT_MIN)
        .MaxValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_4_Time.Max", DETERGENT_MAX)
        .DefaultValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_4_Time.Default", DETERGENT_DEFAULT)
        .Dimension = LimitsFile.ReadString(DTRG_SECTION_NAME, "Detergent_4_Time.Dimension", DETERGENT_DIMENSION)
    End With

    With Detergent_5_Time
        .MinValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_5_Time.Min", DETERGENT_MIN)
        .MaxValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_5_Time.Max", DETERGENT_MAX)
        .DefaultValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_5_Time.Default", DETERGENT_DEFAULT)
        .Dimension = LimitsFile.ReadString(DTRG_SECTION_NAME, "Detergent_5_Time.Dimension", DETERGENT_DIMENSION)
    End With

    With Detergent_6_Time
        .MinValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_6_Time.Min", DETERGENT_MIN)
        .MaxValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_6_Time.Max", DETERGENT_MAX)
        .DefaultValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_6_Time.Default", DETERGENT_DEFAULT)
        .Dimension = LimitsFile.ReadString(DTRG_SECTION_NAME, "Detergent_6_Time.Dimension", DETERGENT_DIMENSION)
    End With

    With Detergent_7_Time
        .MinValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_7_Time.Min", DETERGENT_MIN)
        .MaxValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_7_Time.Max", DETERGENT_MAX)
        .DefaultValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_7_Time.Default", DETERGENT_DEFAULT)
        .Dimension = LimitsFile.ReadString(DTRG_SECTION_NAME, "Detergent_7_Time.Dimension", DETERGENT_DIMENSION)
    End With

    With Detergent_8_Time
        .MinValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_8_Time.Min", DETERGENT_MIN)
        .MaxValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_8_Time.Max", DETERGENT_MAX)
        .DefaultValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_8_Time.Default", DETERGENT_DEFAULT)
        .Dimension = LimitsFile.ReadString(DTRG_SECTION_NAME, "Detergent_8_Time.Dimension", DETERGENT_DIMENSION)
    End With

    With Detergent_9_Time
        .MinValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_9_Time.Min", DETERGENT_MIN)
        .MaxValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_9_Time.Max", DETERGENT_MAX)
        .DefaultValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "Detergent_9_Time.Default", DETERGENT_DEFAULT)
        .Dimension = LimitsFile.ReadString(DTRG_SECTION_NAME, "Detergent_9_Time.Dimension", DETERGENT_DIMENSION)
    End With

    With RotationTime
        .MinValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "RotationTime.Min", ROTATIONTIME_MIN)
        .MaxValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "RotationTime.Max", ROTATIONTIME_MAX)
        .DefaultValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "RotationTime.Default", ROTATIONTIME_DEFAULT)
        .Dimension = LimitsFile.ReadString(DTRG_SECTION_NAME, "RotationTime.Dimension", ROTATIONTIME_DIMENSION)
    End With

    With PauseTime
        .MinValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "PauseTime.Min", PAUSETIME_MIN)
        .MaxValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "PauseTime.Max", PAUSETIME_MAX)
        .DefaultValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "PauseTime.Default", PAUSETIME_DEFAULT)
        .Dimension = LimitsFile.ReadString(DTRG_SECTION_NAME, "PauseTime.Dimension", PAUSETIME_DIMENSION)
    End With

    With DrumSpeed
        .MinValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "DrumSpeed.Min", DRUMSPEED_MIN)
        .MaxValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "DrumSpeed.Max", DRUMSPEED_MAX)
        .DefaultValue = LimitsFile.ReadInteger(DTRG_SECTION_NAME, "DrumSpeed.Default", DRUMSPEED_DEFAULT)
        .Dimension = LimitsFile.ReadString(DTRG_SECTION_NAME, "DrumSpeed.Dimension", DRUMSPEED_DIMENSION)
    End With

    Set LimitsFile = Nothing
End Sub

' Функция-посредник
Private Sub func_SetDefaults(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordDetergent As TYPE_WPC_DETERGENT)

    Dim StepPointer As Long

    StepPointer = Manager.DataPointer + _
       Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
       HEADER_SIZE_IN_BYTES + _
       Manager.StepIndex * STEP_SIZE_IN_BYTES

    PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

    ' Дискретные типы

    Select Case Pause.DefaultValue
        Case False: RecordDetergent.Bits = RecordDetergent.Bits And &HFFEF
        Case True: RecordDetergent.Bits = RecordDetergent.Bits Or &H10
    End Select

    Select Case Rotation.DefaultValue
        Case False: RecordDetergent.Bits = RecordDetergent.Bits And &HFFDF
        Case True: RecordDetergent.Bits = RecordDetergent.Bits Or &H20
    End Select

    ' Аналоговые типы

    RecordDetergent.Detergent_1_Time = Detergent_1_Time.DefaultValue
    RecordDetergent.Detergent_2_Time = Detergent_2_Time.DefaultValue
    RecordDetergent.Detergent_3_Time = Detergent_3_Time.DefaultValue
    RecordDetergent.Detergent_4_Time = Detergent_4_Time.DefaultValue
    RecordDetergent.Detergent_5_Time = Detergent_5_Time.DefaultValue
    RecordDetergent.Detergent_6_Time = Detergent_6_Time.DefaultValue
    RecordDetergent.Detergent_7_Time = Detergent_7_Time.DefaultValue
    RecordDetergent.Detergent_8_Time = Detergent_8_Time.DefaultValue
    RecordDetergent.Detergent_9_Time = Detergent_9_Time.DefaultValue
    RecordDetergent.RotationTime = RotationTime.DefaultValue
    RecordDetergent.PauseTime = PauseTime.DefaultValue
    RecordDetergent.DrumSpeed = DrumSpeed.DefaultValue

End Sub

Public Sub SetDefaults(frm As FormMain)
    Dim RecordDetergent As TYPE_WPC_DETERGENT

    func_SetDefaults frm, 0&, RecordDetergent
End Sub

' Функция-посредник
Private Sub func_ShowPropertyTableForDTRG(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordDetergent As TYPE_WPC_DETERGENT)

    Dim I As Integer, J As Integer
    Dim ParamStr As String, s As String
    Dim StepPointer As Long

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

                Case DTRG_ROTATIONTIME_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_ROTATIONTIME

                Case DTRG_PAUSETIME_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_PAUSETIME

                Case DTRG_DRUMSPEED_FIELD:
                    ParamStr = ParamStr & DTRG_PARAMETER_DESCR_DRUMSPEED

                Case Else
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_UNKNOWN
            End Select

            If (I < DTRG_PARAMETERS_COUNT) Then ParamStr = ParamStr & "|"
        Next

        ' Эта строка автоматически изменяет количество строк
        .PropertyTable.FormatString = ParamStr

        StepPointer = Manager.DataPointer + Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

        StepPointer = Manager.DataPointer + _
           Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
           HEADER_SIZE_IN_BYTES + _
           Manager.StepIndex * STEP_SIZE_IN_BYTES

        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer

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

                    .PropertyTable.CellBackColor = &HE0E0E0

                Case IDLE_DOORUNLOCK_FIELD:

                    If (RecordTitle.LowBits And &H2) / &H2 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                    .PropertyTable.CellBackColor = &HE0E0E0

                Case IDLE_PROGNAME_FIELD:
                    s = ""

                    For J = 1 To PROG_NAME_LENGTH - 1
                        s = s & Chr$(CLng(RecordTitle.ProgName(J)))
                    Next

                    .PropertyTable.Text = s
                    .PropertyTable.CellBackColor = &HE0E0E0

                Case IDLE_STEP_FIELD:
                    .PropertyTable.CellFontBold = True
                    .PropertyTable.Text = "" & Manager.StepIndex + 1

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

                    If LimitsLoaded Then

                        If RecordDetergent.Detergent_1_Time < Detergent_1_Time.MinValue Or _
                           RecordDetergent.Detergent_1_Time > Detergent_1_Time.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_1_Time

                Case DTRG_DETERGENT2_FIELD:

                    If LimitsLoaded Then

                        If RecordDetergent.Detergent_2_Time < Detergent_2_Time.MinValue Or _
                           RecordDetergent.Detergent_2_Time > Detergent_2_Time.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_2_Time

                Case DTRG_DETERGENT3_FIELD:

                    If LimitsLoaded Then

                        If RecordDetergent.Detergent_3_Time < Detergent_3_Time.MinValue Or _
                           RecordDetergent.Detergent_3_Time > Detergent_3_Time.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_3_Time

                Case DTRG_DETERGENT4_FIELD:

                    If LimitsLoaded Then

                        If RecordDetergent.Detergent_1_Time < Detergent_1_Time.MinValue Or _
                           RecordDetergent.Detergent_1_Time > Detergent_1_Time.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_4_Time

                Case DTRG_DETERGENT5_FIELD:

                    If LimitsLoaded Then

                        If RecordDetergent.Detergent_4_Time < Detergent_4_Time.MinValue Or _
                           RecordDetergent.Detergent_4_Time > Detergent_4_Time.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_5_Time

                Case DTRG_DETERGENT6_FIELD:

                    If LimitsLoaded Then

                        If RecordDetergent.Detergent_6_Time < Detergent_6_Time.MinValue Or _
                           RecordDetergent.Detergent_6_Time > Detergent_6_Time.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_6_Time

                Case DTRG_DETERGENT7_FIELD:

                    If LimitsLoaded Then

                        If RecordDetergent.Detergent_7_Time < Detergent_7_Time.MinValue Or _
                           RecordDetergent.Detergent_7_Time > Detergent_7_Time.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_7_Time

                Case DTRG_DETERGENT8_FIELD:

                    If LimitsLoaded Then

                        If RecordDetergent.Detergent_8_Time < Detergent_8_Time.MinValue Or _
                           RecordDetergent.Detergent_8_Time > Detergent_8_Time.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_8_Time

                Case DTRG_DETERGENT9_FIELD:

                    If LimitsLoaded Then

                        If RecordDetergent.Detergent_9_Time < Detergent_9_Time.MinValue Or _
                           RecordDetergent.Detergent_9_Time > Detergent_9_Time.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordDetergent.Detergent_9_Time

                Case DTRG_ROTATIONTIME_FIELD:

                    If LimitsLoaded Then

                        If RecordDetergent.RotationTime < RotationTime.MinValue Or _
                           RecordDetergent.RotationTime > RotationTime.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordDetergent.RotationTime

                Case DTRG_PAUSETIME_FIELD:

                    If LimitsLoaded Then

                        If RecordDetergent.PauseTime < PauseTime.MinValue Or _
                           RecordDetergent.PauseTime > PauseTime.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordDetergent.PauseTime

                Case DTRG_DRUMSPEED_FIELD:

                    If LimitsLoaded Then

                        If RecordDetergent.DrumSpeed < DrumSpeed.MinValue Or _
                           RecordDetergent.DrumSpeed > DrumSpeed.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordDetergent.DrumSpeed

                Case Else
                    .PropertyTable.CellBackColor = &H8000000F

            End Select
        Next
    End With
End Sub

Public Sub ShowPropertyTableForDTRG(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordDetergent As TYPE_WPC_DETERGENT

    func_ShowPropertyTableForDTRG frm, 0&, RecordTitle, RecordDetergent
End Sub

' Функция-посредник
Private Sub func_EditPropertyForDTRG(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordDetergent As TYPE_WPC_DETERGENT)

    Dim I As Integer
    Dim StepPointer As Long
    Dim s As String

    With frm
        StepPointer = Manager.DataPointer + Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

        StepPointer = Manager.DataPointer + _
           Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
           HEADER_SIZE_IN_BYTES + _
           Manager.StepIndex * STEP_SIZE_IN_BYTES

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

        .ShapeDescription.Left = .LabelDescription.Left
        .ShapeDescription.Width = .LabelDescription.Width
        .ShapeDescription.Height = .LabelDescription.Height
        .ShapeDescription.Top = .LabelDescription.Top

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
                    s = s & Chr$(CLng(RecordTitle.ProgName(I)))
                Next
                .TextCell.Text = s
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case IDLE_STEP_FIELD:

                For I = 1 To MAX_NUMBER_OF_STEPS
                    .ComboCell.AddItem ("Шаг " & I)
                Next
                .ComboCell.ListIndex = Manager.StepIndex
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case IDLE_FUNCTION_FIELD:

                For I = 1 To NUMBER_OF_FUNCS
                    .ComboCell.AddItem (FunctionsStrings(I - 1))
                Next
                .ComboCell.ListIndex = RecordDetergent.Bits And &HF
                .ComboCell.Visible = True
                .ComboCell.SetFocus

                ' Специальные параметры шага
            Case DTRG_PAUSE_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordDetergent.Bits And &H10) / &H10
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case DTRG_ROTATION_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordDetergent.Bits And &H20) / &H20
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case DTRG_DETERGENT1_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & Detergent_1_Time.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & Detergent_1_Time.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & Detergent_1_Time.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & Detergent_1_Time.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If

                .TextCell.Text = "" & RecordDetergent.Detergent_1_Time
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case DTRG_DETERGENT2_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & Detergent_2_Time.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & Detergent_2_Time.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & Detergent_2_Time.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & Detergent_2_Time.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If

                .TextCell.Text = "" & RecordDetergent.Detergent_2_Time
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case DTRG_DETERGENT3_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & Detergent_3_Time.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & Detergent_3_Time.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & Detergent_3_Time.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & Detergent_3_Time.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If

                .TextCell.Text = "" & RecordDetergent.Detergent_3_Time
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case DTRG_DETERGENT4_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & Detergent_4_Time.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & Detergent_4_Time.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & Detergent_4_Time.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & Detergent_4_Time.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If

                .TextCell.Text = "" & RecordDetergent.Detergent_4_Time
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case DTRG_DETERGENT5_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & Detergent_5_Time.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & Detergent_5_Time.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & Detergent_5_Time.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & Detergent_5_Time.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If

                .TextCell.Text = "" & RecordDetergent.Detergent_5_Time
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case DTRG_DETERGENT6_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & Detergent_6_Time.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & Detergent_6_Time.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & Detergent_6_Time.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & Detergent_6_Time.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If

                .TextCell.Text = "" & RecordDetergent.Detergent_6_Time
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case DTRG_DETERGENT7_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & Detergent_7_Time.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & Detergent_7_Time.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & Detergent_7_Time.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & Detergent_7_Time.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If

                .TextCell.Text = "" & RecordDetergent.Detergent_7_Time
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case DTRG_DETERGENT8_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & Detergent_8_Time.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & Detergent_8_Time.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & Detergent_8_Time.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & Detergent_8_Time.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If

                .TextCell.Text = "" & RecordDetergent.Detergent_8_Time
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case DTRG_DETERGENT9_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & Detergent_9_Time.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & Detergent_9_Time.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & Detergent_9_Time.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & Detergent_9_Time.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If

                .TextCell.Text = "" & RecordDetergent.Detergent_9_Time
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case DTRG_ROTATIONTIME_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & RotationTime.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & RotationTime.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & RotationTime.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & RotationTime.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If

                .TextCell.Text = "" & RecordDetergent.RotationTime
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case DTRG_PAUSETIME_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & PauseTime.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & PauseTime.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & PauseTime.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & PauseTime.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If

                .TextCell.Text = "" & RecordDetergent.PauseTime
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case DTRG_DRUMSPEED_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & DrumSpeed.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & DrumSpeed.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & DrumSpeed.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & DrumSpeed.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If

                .TextCell.Text = "" & RecordDetergent.DrumSpeed
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case Else

        End Select

        .TextCell.SelStart = 0
        .TextCell.SelLength = Len(.TextCell.Text)
    End With
End Sub

Public Sub EditPropertyForDTRG(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordDetergent As TYPE_WPC_DETERGENT

    func_EditPropertyForDTRG frm, 0&, RecordTitle, RecordDetergent
End Sub

' Функция-посредник
Private Sub func_SetCheckBoxForDTRG(frm As FormMain, _
       ByVal ValveNumber As Integer, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordDetergent As TYPE_WPC_DETERGENT)

    Dim StepPointer As Long
    
    StepPointer = Manager.DataPointer + _
       Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
       HEADER_SIZE_IN_BYTES + _
       Manager.StepIndex * STEP_SIZE_IN_BYTES
    
    PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer

    Select Case ValveNumber
    
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
        Case 15:
        Case 16:
        Case 17:
        Case 18: RecordDetergent.Bits = RecordDetergent.Bits Xor &H20 ' "Разр. вращ."
    
    End Select
    
    SetModified True

End Sub

Public Sub SetCheckBoxForDTRG(frm As FormMain, ValveNumber As Integer)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordDetergent As TYPE_WPC_DETERGENT

    func_SetCheckBoxForDTRG frm, ValveNumber, 0&, RecordTitle, RecordDetergent
End Sub

' Функция-посредник
Private Sub func_SetComboPropertyForDTRG(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordDetergent As TYPE_WPC_DETERGENT)

    On Error GoTo ErrorHandler

    Dim I As Integer
    Dim StepPointer As Long

    With frm
        StepPointer = Manager.DataPointer + Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

        StepPointer = Manager.DataPointer + _
           Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
           HEADER_SIZE_IN_BYTES + _
           Manager.StepIndex * STEP_SIZE_IN_BYTES

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
                        RecordTitle.ProgName(I) = Asc(Mid$(.TextCell.Text, I, 1))
                    Else
                        RecordTitle.ProgName(I) = 0
                    End If
                Next
                RecordTitle.ProgName(PROG_NAME_LENGTH) = 0

            Case IDLE_STEP_FIELD:
                Manager.StepIndex = .ComboCell.ListIndex
                Exit Sub

            Case IDLE_FUNCTION_FIELD:
                ' Если функция та же, то ничего не делаем

                If (RecordDetergent.Bits And &HF) = (.ComboCell.ListIndex And &HF) Then Exit Sub

                ' При выборе новой функции для шага мы должны обнулить структуру,
                ZeroMemory RecordDetergent, STEP_SIZE_IN_BYTES

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

'<Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:21:57
'Причина: Модуль аналогичен по функционалу с ModuleWashOrRinsOrJolt>
'                            Case WPC_OPERATION_PAUS ' пауза
'</Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:21:57>

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

                'записать новое значение в поле типа функции
                RecordDetergent.Bits = RecordDetergent.Bits Or (.ComboCell.ListIndex And &HF)

                ' Специальные параметры шага
            Case DTRG_PAUSE_FIELD:

                Select Case .ComboCell.ListIndex
                    Case 0: RecordDetergent.Bits = RecordDetergent.Bits And &HFFEF
                    Case 1: RecordDetergent.Bits = RecordDetergent.Bits Or &H10
                End Select

            Case DTRG_ROTATION_FIELD:

                Select Case .ComboCell.ListIndex
                    Case 0: RecordDetergent.Bits = RecordDetergent.Bits And &HFFDF
                    Case 1: RecordDetergent.Bits = RecordDetergent.Bits Or &H20
                End Select

            Case DTRG_DETERGENT1_FIELD:
                RecordDetergent.Detergent_1_Time = Val(.TextCell.Text)

            Case DTRG_DETERGENT2_FIELD:
                RecordDetergent.Detergent_2_Time = Val(.TextCell.Text)

            Case DTRG_DETERGENT3_FIELD:
                RecordDetergent.Detergent_3_Time = Val(.TextCell.Text)

            Case DTRG_DETERGENT4_FIELD:
                RecordDetergent.Detergent_4_Time = Val(.TextCell.Text)

            Case DTRG_DETERGENT5_FIELD:
                RecordDetergent.Detergent_5_Time = Val(.TextCell.Text)

            Case DTRG_DETERGENT6_FIELD:
                RecordDetergent.Detergent_6_Time = Val(.TextCell.Text)

            Case DTRG_DETERGENT7_FIELD:
                RecordDetergent.Detergent_7_Time = Val(.TextCell.Text)

            Case DTRG_DETERGENT8_FIELD:
                RecordDetergent.Detergent_8_Time = Val(.TextCell.Text)

            Case DTRG_DETERGENT9_FIELD:
                RecordDetergent.Detergent_9_Time = Val(.TextCell.Text)

            Case DTRG_ROTATIONTIME_FIELD:
                RecordDetergent.RotationTime = Val(.TextCell.Text)

            Case DTRG_PAUSETIME_FIELD:
                RecordDetergent.PauseTime = Val(.TextCell.Text)

            Case DTRG_DRUMSPEED_FIELD:
                RecordDetergent.DrumSpeed = Val(.TextCell.Text)

            Case Else

        End Select

        SetModified True
    End With
    Exit Sub

ErrorHandler:
    Err.Clear
End Sub

Public Sub SetComboPropertyForDTRG(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordDetergent As TYPE_WPC_DETERGENT

    func_SetComboPropertyForDTRG frm, 0&, RecordTitle, RecordDetergent
End Sub

' Функция-посредник
Private Function func_ValveEnabled(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordDetergent As TYPE_WPC_DETERGENT, _
       ByVal StepIndex As Integer, _
       ByVal Num As Integer) As Boolean

    Dim StepPointer As Long

    With frm
        StepPointer = Manager.DataPointer + Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

        StepPointer = Manager.DataPointer + _
           Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
           HEADER_SIZE_IN_BYTES + _
           StepIndex * STEP_SIZE_IN_BYTES

        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer

        Select Case Num
            Case 1:
            Case 2:
            Case 3:
            Case 4: func_ValveEnabled = RecordDetergent.Detergent_1_Time > 0
            Case 5: func_ValveEnabled = RecordDetergent.Detergent_2_Time > 0
            Case 6: func_ValveEnabled = RecordDetergent.Detergent_3_Time > 0
            Case 7: func_ValveEnabled = RecordDetergent.Detergent_4_Time > 0
            Case 8: func_ValveEnabled = RecordDetergent.Detergent_5_Time > 0
            Case 9: func_ValveEnabled = RecordDetergent.Detergent_6_Time > 0
            Case 10: func_ValveEnabled = RecordDetergent.Detergent_7_Time > 0
            Case 11: func_ValveEnabled = RecordDetergent.Detergent_8_Time > 0
            Case 12: func_ValveEnabled = RecordDetergent.Detergent_9_Time > 0
            Case 13:
            Case 14:
            Case 15:
            Case 16:
            Case 17:
            Case 18: func_ValveEnabled = (RecordDetergent.Bits And &H20) > 0

            Case Else
                func_ValveEnabled = False
        End Select
    End With
End Function

Public Function ValveEnabled(frm As FormMain, StepIndex As Integer, Num As Integer) As Boolean
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordDetergent As TYPE_WPC_DETERGENT

    ValveEnabled = func_ValveEnabled(frm, 0&, RecordTitle, RecordDetergent, StepIndex, Num)
End Function
