VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CModuleSpin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const SPIN_PAUSE_FIELD = IDLE_FUNCTION_FIELD + 1
Private Const SPIN_DRAINGATE1_FIELD = IDLE_FUNCTION_FIELD + 2
Private Const SPIN_DRAINGATE2_FIELD = IDLE_FUNCTION_FIELD + 3

Private Const SPIN_DRUMSPEED_FIELD = IDLE_FUNCTION_FIELD + 4
Private Const SPIN_TIME_FIELD = IDLE_FUNCTION_FIELD + 5

Private Const SPIN_PARAMETERS_COUNT = IDLE_PARAMETERS_COUNT + 5

' Дискретные параметры
Private Const SPIN_PARAMETER_DESCR_PAUSE = "Разр. паузы"
Private Const SPIN_PARAMETER_DESCR_DRAINGATE1 = "Разр. кл. слива 1"
Private Const SPIN_PARAMETER_DESCR_DRAINGATE2 = "Разр. кл. слива 2"

' Аналоговые параметры
Private Const SPIN_PARAMETER_DESCR_DRUMSPEED = "Скорость вращ. барабана"
Private Const SPIN_PARAMETER_DESCR_TIME = "Время отжима"

' Название секции
Private Const SPIN_SECTION_NAME = "Spin"

' Настройки по умолчанию
' Дискретные типы
Private Const PAUSE_DEFAULT = False
Private Const DRAINGATE1_DEFAULT = True
Private Const DRAINGATE2_DEFAULT = False

' Аналоговые типы
Private Const DRUMSPEED_MIN = 500
Private Const DRUMSPEED_MAX = 1500
Private Const DRUMSPEED_DEFAULT = 800
Private Const DRUMSPEED_DIMENSION = "об/мин"

Private Const TIME_MIN = 1
Private Const TIME_MAX = 255
Private Const TIME_DEFAULT = 2
Private Const TIME_DIMENSION = "сек"

' Параметры функции шага
' Дискретные типы
Private Pause As TYPE_BOOL_DESCRIPTION
Private DrainGate1 As TYPE_BOOL_DESCRIPTION
Private DrainGate2 As TYPE_BOOL_DESCRIPTION

' Аналоговые типы
Private DrumSpeed As TYPE_INT_DESCRIPTION
Private Time As TYPE_BYTE_DESCRIPTION

Public LimitsLoaded As Boolean

' Конструктор
Private Sub Class_Initialize()
    LimitsLoaded = False
End Sub

Public Sub LoadLimits(FileName As String)
    LimitsLoaded = DoesFileExist(FileName)

    If Not LimitsLoaded Then Exit Sub

    Dim LimitsFile As New CIniFile

    LimitsFile.Create FileName

    ' Настройки функции
    ' Дискретные типы
    Pause.DefaultValue = LimitsFile.ReadBoolean(SPIN_SECTION_NAME, "Pause.Default", PAUSE_DEFAULT)
    DrainGate1.DefaultValue = LimitsFile.ReadBoolean(SPIN_SECTION_NAME, "DrainGate1.Default", DRAINGATE1_DEFAULT)
    DrainGate2.DefaultValue = LimitsFile.ReadBoolean(SPIN_SECTION_NAME, "DrainGate2.Default", DRAINGATE2_DEFAULT)

    ' Аналоговые типы

    With DrumSpeed
        .MinValue = LimitsFile.ReadInteger(SPIN_SECTION_NAME, "DrumSpeed.Min", DRUMSPEED_MIN)
        .MaxValue = LimitsFile.ReadInteger(SPIN_SECTION_NAME, "DrumSpeed.Max", DRUMSPEED_MAX)
        .DefaultValue = LimitsFile.ReadInteger(SPIN_SECTION_NAME, "DrumSpeed.Default", DRUMSPEED_DEFAULT)
        .Dimension = LimitsFile.ReadString(SPIN_SECTION_NAME, "DrumSpeed.Dimension", DRUMSPEED_DIMENSION)
    End With

    With Time
        .MinValue = LimitsFile.ReadInteger(SPIN_SECTION_NAME, "Time.Min", TIME_MIN)
        .MaxValue = LimitsFile.ReadInteger(SPIN_SECTION_NAME, "Time.Max", TIME_MAX)
        .DefaultValue = LimitsFile.ReadInteger(SPIN_SECTION_NAME, "Time.Default", TIME_DEFAULT)
        .Dimension = LimitsFile.ReadString(SPIN_SECTION_NAME, "Time.Dimension", TIME_DIMENSION)
    End With

    Set LimitsFile = Nothing
End Sub

' Функция-посредник
Private Sub func_SetDefaults(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordSpin As TYPE_WPC_SPIN)

    Dim StepPointer As Long

    StepPointer = Manager.DataPointer + _
       Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
       HEADER_SIZE_IN_BYTES + _
       Manager.StepIndex * STEP_SIZE_IN_BYTES

    PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

    ' Дискретные типы

    Select Case Pause.DefaultValue
        Case False: RecordSpin.Bits = RecordSpin.Bits And &HFFEF
        Case True: RecordSpin.Bits = RecordSpin.Bits Or &H10
    End Select

    Select Case DrainGate1.DefaultValue
        Case False: RecordSpin.Bits = RecordSpin.Bits And &HFFDF
        Case True: RecordSpin.Bits = RecordSpin.Bits Or &H20
    End Select

    Select Case DrainGate2.DefaultValue
        Case False: RecordSpin.Bits = RecordSpin.Bits And &HFFBF
        Case True: RecordSpin.Bits = RecordSpin.Bits Or &H40
    End Select

    ' Аналоговые типы
    RecordSpin.DrumSpeed = DrumSpeed.DefaultValue
    RecordSpin.Time = Time.DefaultValue

End Sub

Public Sub SetDefaults(frm As FormMain)
    Dim RecordSpin As TYPE_WPC_SPIN

    func_SetDefaults frm, 0&, RecordSpin
End Sub

' Функция-посредник
Private Sub func_ShowPropertyTableForSpin(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordSpin As TYPE_WPC_SPIN)

    Dim I As Integer
    Dim j As Integer
    Dim ParamStr As String
    Dim S As String
    Dim StepPointer As Long

    With frm
        ' Формируем столбик с названиями параметров
        ParamStr = ";Параметр|"

        For I = 1 To SPIN_PARAMETERS_COUNT

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
                Case SPIN_PAUSE_FIELD:
                    ParamStr = ParamStr & SPIN_PARAMETER_DESCR_PAUSE

                Case SPIN_DRAINGATE1_FIELD:
                    ParamStr = ParamStr & SPIN_PARAMETER_DESCR_DRAINGATE1

                Case SPIN_DRAINGATE2_FIELD:
                    ParamStr = ParamStr & SPIN_PARAMETER_DESCR_DRAINGATE2

                Case SPIN_DRUMSPEED_FIELD:
                    ParamStr = ParamStr & SPIN_PARAMETER_DESCR_DRUMSPEED

                Case SPIN_TIME_FIELD:
                    ParamStr = ParamStr & SPIN_PARAMETER_DESCR_TIME

                Case Else
                    ParamStr = ParamStr & IDLE_PARAMETER_DESCR_UNKNOWN
            End Select

            If (I < SPIN_PARAMETERS_COUNT) Then ParamStr = ParamStr & "|"
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

        .PropertyTable.Col = 1
        .PropertyTable.row = 0
        .PropertyTable.CellAlignment = flexAlignRightCenter
        
        For I = 1 To SPIN_PARAMETERS_COUNT
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
                    S = ""

                    For j = 1 To PROG_NAME_LENGTH - 1
                        S = S & Chr$(CLng(RecordTitle.ProgName(j)))
                    Next

                    .PropertyTable.Text = S
                    .PropertyTable.CellBackColor = &HE0E0E0

                Case IDLE_STEP_FIELD:
                    .PropertyTable.CellFontBold = True
                    .PropertyTable.Text = "" & Manager.StepIndex + 1

                Case IDLE_FUNCTION_FIELD:
                    .PropertyTable.CellFontBold = True
                    .PropertyTable.Text = FunctionsStrings(RecordSpin.Bits And &HF)

                    ' Специальные параметры
                Case SPIN_PAUSE_FIELD:

                    If (RecordSpin.Bits And &H10) / &H10 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                Case SPIN_DRAINGATE1_FIELD:

                    If (RecordSpin.Bits And &H20) / &H20 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                Case SPIN_DRAINGATE2_FIELD:

                    If (RecordSpin.Bits And &H40) / &H40 Then
                        .PropertyTable.Text = STRING_YES
                    Else
                        .PropertyTable.Text = STRING_NO
                    End If

                Case SPIN_DRUMSPEED_FIELD:

                    If LimitsLoaded Then

                        If RecordSpin.DrumSpeed < DrumSpeed.MinValue Or RecordSpin.DrumSpeed > DrumSpeed.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordSpin.DrumSpeed

                Case SPIN_TIME_FIELD:

                    If LimitsLoaded Then

                        If RecordSpin.Time < Time.MinValue Or RecordSpin.Time > Time.MaxValue Then
                            .PropertyTable.CellBackColor = &H8080FF
                        Else
                            .PropertyTable.CellBackColor = &H80000005
                        End If
                    End If
                    .PropertyTable.Text = "" & RecordSpin.Time

                Case Else
                    .PropertyTable.CellBackColor = &H8000000F

            End Select
        Next
    End With
End Sub

Public Sub ShowPropertyTableForSpin(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordSpin As TYPE_WPC_SPIN

    func_ShowPropertyTableForSpin frm, 0&, RecordTitle, RecordSpin
End Sub

' Функция-посредник
Private Sub func_EditPropertyForSpin(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordSpin As TYPE_WPC_SPIN)

    Dim I As Integer
    Dim StepPointer As Long
    Dim S As String

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
                S = ""

                For I = 1 To PROG_NAME_LENGTH - 1
                    S = S & Chr$(CLng(RecordTitle.ProgName(I)))
                Next
                .TextCell.Text = S
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
                .ComboCell.ListIndex = RecordSpin.Bits And &HF
                .ComboCell.Visible = True
                .ComboCell.SetFocus

                ' Специальные параметры шага
            Case SPIN_PAUSE_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordSpin.Bits And &H10) / &H10
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case SPIN_DRAINGATE1_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordSpin.Bits And &H20) / &H20
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case SPIN_DRAINGATE2_FIELD:
                .ComboCell.AddItem STRING_NO
                .ComboCell.AddItem STRING_YES
                .ComboCell.ListIndex = (RecordSpin.Bits And &H40) / &H40
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case SPIN_DRUMSPEED_FIELD:

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
                .TextCell.Text = "" & RecordSpin.DrumSpeed
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case SPIN_TIME_FIELD:

                If LimitsLoaded Then
                    .PropertyTable.Height = .PropertyTable.Height - .LabelDescription.Height
                    .LabelDescription.Caption = VBA.Constants.vbCrLf & _
                       DESCR_MIN_VALUE & Time.MinValue & VBA.Constants.vbCrLf & _
                       DESCR_MAX_VALUE & Time.MaxValue & VBA.Constants.vbCrLf & _
                       DESCR_DEFAULT_VALUE & Time.DefaultValue & VBA.Constants.vbCrLf & _
                       DESCR_DIMENSION & "[" & Time.Dimension & "]"
                    .LabelDescription.Visible = True
                    .ShapeDescription.Visible = True
                End If
                .TextCell.Text = "" & RecordSpin.Time
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case Else

        End Select

        .TextCell.SelStart = 0
        .TextCell.SelLength = Len(.TextCell.Text)
    End With
End Sub

Public Sub EditPropertyForSpin(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordSpin As TYPE_WPC_SPIN

    func_EditPropertyForSpin frm, 0&, RecordTitle, RecordSpin
End Sub

' Функция-посредник
Private Sub func_SetCheckBoxForSpin(frm As FormMain, _
       ByVal ValveNumber As Integer, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordSpin As TYPE_WPC_SPIN)

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
        Case 15: RecordSpin.Bits = RecordSpin.Bits Xor &H20 ' "Разр. кл. слива 1"
        Case 16: RecordSpin.Bits = RecordSpin.Bits Xor &H40 ' "Разр. кл. слива 2"
        Case 17:
        Case 18:
    
    End Select
    
    SetModified True

End Sub

Public Sub SetCheckBoxForSpin(frm As FormMain, ValveNumber As Integer)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordSpin As TYPE_WPC_SPIN

    func_SetCheckBoxForSpin frm, ValveNumber, 0&, RecordTitle, RecordSpin
End Sub

' Функция-посредник
Private Sub func_SetComboPropertyForSpin(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordSpin As TYPE_WPC_SPIN)

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

                If (RecordSpin.Bits And &HF) = (.ComboCell.ListIndex And &HF) Then Exit Sub

                ' При выборе новой функции для шага мы должны обнулить структуру,
                ZeroMemory RecordSpin, STEP_SIZE_IN_BYTES

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

'<Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:23:27
'Причина: Модуль аналогичен по функционалу с ModuleWashOrRinsOrJolt>
'                        Case WPC_OPERATION_PAUS ' пауза
'</Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:23:27>

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
                RecordSpin.Bits = RecordSpin.Bits Or (.ComboCell.ListIndex And &HF)

                ' Специальные параметры шага
            Case SPIN_PAUSE_FIELD:

                Select Case .ComboCell.ListIndex
                    Case 0: RecordSpin.Bits = RecordSpin.Bits And &HFFEF
                    Case 1: RecordSpin.Bits = RecordSpin.Bits Or &H10
                End Select

            Case SPIN_DRAINGATE1_FIELD:

                Select Case .ComboCell.ListIndex
                    Case 0: RecordSpin.Bits = RecordSpin.Bits And &HFFDF
                    Case 1: RecordSpin.Bits = RecordSpin.Bits Or &H20
                End Select

            Case SPIN_DRAINGATE2_FIELD:

                Select Case .ComboCell.ListIndex
                    Case 0: RecordSpin.Bits = RecordSpin.Bits And &HFFBF
                    Case 1: RecordSpin.Bits = RecordSpin.Bits Or &H40
                End Select

            Case SPIN_DRUMSPEED_FIELD:
                RecordSpin.DrumSpeed = Val(.TextCell.Text)

            Case SPIN_TIME_FIELD:
                RecordSpin.Time = Val(.TextCell.Text)

            Case Else

        End Select

        SetModified True
    End With

    Exit Sub

ErrorHandler:
    Err.Clear
End Sub

Public Sub SetComboPropertyForSpin(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordSpin As TYPE_WPC_SPIN

    func_SetComboPropertyForSpin frm, 0&, RecordTitle, RecordSpin
End Sub

' Функция-посредник
Private Function func_ValveEnabled(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordSpin As TYPE_WPC_SPIN, _
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
            Case 15: func_ValveEnabled = (RecordSpin.Bits And &H20) > 0
            Case 16: func_ValveEnabled = (RecordSpin.Bits And &H40) > 0
            Case 17:
            Case 18: func_ValveEnabled = RecordSpin.DrumSpeed > 0

            Case Else
                func_ValveEnabled = False
        End Select
    End With
End Function

Public Function ValveEnabled(frm As FormMain, StepIndex As Integer, Num As Integer) As Boolean
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordSpin As TYPE_WPC_SPIN

    ValveEnabled = func_ValveEnabled(frm, 0&, RecordTitle, RecordSpin, StepIndex, Num)
End Function

' Функция-посредник
Private Function func_ShowStepTableForSpin(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordSpin As TYPE_WPC_SPIN, _
       ByRef StepGrid As MSFlexGrid) As Boolean

    Dim I As Integer, j As Integer
    Dim StepPointer As Long
    Dim S As String
    
    ' ------
    ' Загружаем установки в таблицу
    With StepGrid
    
        ' Отключаем перерисовку
        .Redraw = False
    
        .rows = 12
    
        ' Очищаем таблицу установок
        .Clear
    
        Dim row As Long
    
        .FormatString = "<Параметр|Значение"
        .ColWidth(0) = 2 * .Width / 3
        .ColWidth(1) = .Width / 3
        
        .row = 0
        .Col = 0
        '.CellFontBold = True
        .CellForeColor = &HFFFFFF
        
        .Col = 1
        .CellAlignment = flexAlignRightCenter
        '.CellFontBold = True
        .CellForeColor = &HFFFFFF
        .RowHeight(.row) = Settings.StepsRowHeight
    
        row = .row
    
        ' Общие параметры [Программа]
        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .Text = "Программа"
        .CellFontBold = True
        .RowHeight(.row) = Settings.StepsRowHeight
        
        .Col = 1
        .Text = Manager.ProgramIndex + 1
        .CellFontBold = True
        
        ' -----------------------------------------------
        .Col = 0
        .row = Inc(row)
        .RowData(.row) = IDLE_ENDSOUND_FIELD
        .Text = IDLE_PARAMETER_DESCR_ENDSOUND
        .CellBackColor = &HFFFFFF
        .RowHeight(.row) = Settings.StepsRowHeight
        
        ' -----------------------------------------------
        .row = Inc(row)
        .RowData(.row) = IDLE_DOORUNLOCK_FIELD
        .Text = IDLE_PARAMETER_DESCR_DOORUNLOCK
        .CellBackColor = &HFFFFFF
        .RowHeight(.row) = Settings.StepsRowHeight
        
        ' -----------------------------------------------
        .row = Inc(row)
        .RowData(.row) = IDLE_PROGNAME_FIELD
        .Text = IDLE_PARAMETER_DESCR_PROGNAME
        .CellBackColor = &HFFFFFF
        .RowHeight(.row) = Settings.StepsRowHeight
        
        ' [Шаг]
        ' -----------------------------------------------
        .row = Inc(row)
        .RowData(.row) = IDLE_STEP_FIELD
        .Text = "Шаг"
        .CellFontBold = True
        .RowHeight(.row) = Settings.StepsRowHeight
        
        ' -----------------------------------------------
        .row = Inc(row)
        .RowData(.row) = IDLE_FUNCTION_FIELD
        .Text = IDLE_PARAMETER_DESCR_FUNCTION
        .CellBackColor = &HFFFFFF
        .RowHeight(.row) = Settings.StepsRowHeight
        
        ' Специальные параметры
        ' -----------------------------------------------
        .row = Inc(row)
        .RowData(.row) = SPIN_PAUSE_FIELD
        .Text = SPIN_PARAMETER_DESCR_PAUSE
        .CellBackColor = &HFFFFFF
        .RowHeight(.row) = Settings.StepsRowHeight
        
        ' -----------------------------------------------
        .row = Inc(row)
        .RowData(.row) = SPIN_DRAINGATE1_FIELD
        .Text = SPIN_PARAMETER_DESCR_DRAINGATE1
        .CellBackColor = &HFFFFFF
        .RowHeight(.row) = Settings.StepsRowHeight
        
        ' -----------------------------------------------
        .row = Inc(row)
        .RowData(.row) = SPIN_DRAINGATE2_FIELD
        .Text = SPIN_PARAMETER_DESCR_DRAINGATE2
        .CellBackColor = &HFFFFFF
        .RowHeight(.row) = Settings.StepsRowHeight
        
        ' -----------------------------------------------
        .row = Inc(row)
        .RowData(.row) = SPIN_DRUMSPEED_FIELD
        .Text = SPIN_PARAMETER_DESCR_DRUMSPEED
        .CellBackColor = &HFFFFFF
        .RowHeight(.row) = Settings.StepsRowHeight
        
        ' -----------------------------------------------
        .row = Inc(row)
        .RowData(.row) = SPIN_TIME_FIELD
        .Text = SPIN_PARAMETER_DESCR_TIME
        .CellBackColor = &HFFFFFF
        .RowHeight(.row) = Settings.StepsRowHeight
        
        .Height = Settings.StepsRowHeight * .rows

        StepPointer = Manager.DataPointer + Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

        StepPointer = Manager.DataPointer + _
           Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
           HEADER_SIZE_IN_BYTES + _
           (.Tag - 1) * STEP_SIZE_IN_BYTES

        PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer
        
        .Col = 1
        
        For I = 1 To .rows - 1
        
            .row = I
            .CellAlignment = flexAlignRightCenter
            .CellBackColor = &HFFFFFF

            Select Case .RowData(I)
                    
                ' Общие параметры
                Case IDLE_ENDSOUND_FIELD:

                    If (RecordTitle.LowBits And &H1) Then
                        .Text = STRING_YES
                    Else
                        .Text = STRING_NO
                    End If

                Case IDLE_DOORUNLOCK_FIELD:

                    If (RecordTitle.LowBits And &H2) / &H2 Then
                        .Text = STRING_YES
                    Else
                        .Text = STRING_NO
                    End If

                Case IDLE_PROGNAME_FIELD:
                    
                    S = ""

                    For j = 1 To PROG_NAME_LENGTH - 1
                        S = S & Chr$(CLng(RecordTitle.ProgName(j)))
                    Next

                    .Text = S

                Case IDLE_STEP_FIELD:
                    
                    .CellFontBold = True
                    .CellBackColor = &HF4E0E0
                    .Text = "" & .Tag

                Case IDLE_FUNCTION_FIELD:
                
                    .CellFontBold = True
                    .Text = FunctionsStrings(RecordSpin.Bits And &HF)

                    ' Специальные параметры
                Case SPIN_PAUSE_FIELD:

                    If (RecordSpin.Bits And &H10) / &H10 Then
                        .Text = STRING_YES
                    Else
                        .Text = STRING_NO
                    End If

                Case SPIN_DRAINGATE1_FIELD:

                    If (RecordSpin.Bits And &H20) / &H20 Then
                        .Text = STRING_YES
                    Else
                        .Text = STRING_NO
                    End If

                Case SPIN_DRAINGATE2_FIELD:

                    If (RecordSpin.Bits And &H40) / &H40 Then
                        .Text = STRING_YES
                    Else
                        .Text = STRING_NO
                    End If

                Case SPIN_DRUMSPEED_FIELD:

                    If LimitsLoaded Then

                        If RecordSpin.DrumSpeed < DrumSpeed.MinValue Or RecordSpin.DrumSpeed > DrumSpeed.MaxValue Then
                            .CellBackColor = &H8080FF
                        End If
                    
                    End If
                    
                    .Text = "" & RecordSpin.DrumSpeed

                Case SPIN_TIME_FIELD:

                    If LimitsLoaded Then

                        If RecordSpin.Time < Time.MinValue Or RecordSpin.Time > Time.MaxValue Then
                            .CellBackColor = &H8080FF
                        End If
                    
                    End If
                    
                    .Text = "" & RecordSpin.Time

                Case Else
                
                    .CellBackColor = &HF4E0E0

            End Select
            
        Next

        .Redraw = True

    End With
    
End Function
    
Public Function ShowStepTableForSpin(ByRef frm As FormMain, ByRef StepGrid As MSFlexGrid) As Boolean
    
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordSpin As TYPE_WPC_SPIN

    ShowStepTableForSpin = func_ShowStepTableForSpin(frm, 0&, RecordTitle, RecordSpin, StepGrid)
    
End Function

