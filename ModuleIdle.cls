VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CModuleIdle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Название секции
Private Const IDLE_SECTION_NAME = "Idle"

' Настройки по умолчанию
' Дискретные типы
Private Const PAUSE_DEFAULT = False

' Параметры функции шага
' Дискретные типы
Private Pause As TYPE_BOOL_DESCRIPTION

' Аналоговые типы

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
    Pause.DefaultValue = LimitsFile.ReadBoolean(IDLE_SECTION_NAME, "Pause.Default", PAUSE_DEFAULT)

    ' Аналоговые типы

    Set LimitsFile = Nothing
End Sub

' Функция-посредник
Private Sub func_SetDefaults(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordStep As TYPE_WPC_STEP)

    Dim StepPointer As Long

    StepPointer = Manager.DataPointer + _
       Manager.ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
       HEADER_SIZE_IN_BYTES + _
       Manager.StepIndex * STEP_SIZE_IN_BYTES

    PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

    ' Дискретные типы

    Select Case Pause.DefaultValue
        Case False: RecordStep.Bits = RecordStep.Bits And &HFFEF
        Case True: RecordStep.Bits = RecordStep.Bits Or &H10
    End Select

    ' Аналоговые типы

End Sub

Public Sub SetDefaults(frm As FormMain)
    Dim RecordStep As TYPE_WPC_STEP

    func_SetDefaults frm, 0&, RecordStep
End Sub

' Функция-посредник
Private Sub func_ShowPropertyTableForIdle(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordStep As TYPE_WPC_STEP)

    Dim I As Integer
    Dim j As Integer
    Dim StepPointer As Long
    Dim ParamStr As String
    Dim S As String

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
            
        Next

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

        For I = 1 To IDLE_PARAMETERS_COUNT
        
            .PropertyTable.row = I

            Select Case I
            
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
                    .PropertyTable.Text = FunctionsStrings(RecordStep.Bits And &HF)

                Case Else
                    .PropertyTable.CellBackColor = &H8000000F

            End Select
            
        Next
        
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
            Case IDLE_ENDSOUND_FIELD:

            Case IDLE_DOORUNLOCK_FIELD:

            Case IDLE_PROGNAME_FIELD:
                S = ""

                For I = 1 To PROG_NAME_LENGTH - 1
                    S = S & Chr$(CLng(RecordTitle.ProgName(I)))
                Next
                .TextCell.Text = S
                .TextCell.Visible = True
                .TextCell.SetFocus

            Case IDLE_STEP_FIELD:
                .ComboCell.Left = .PropertyTable.Left + .PropertyTable.CellLeft
                .ComboCell.Top = .PropertyTable.Top + .PropertyTable.CellTop
                .ComboCell.Width = .PropertyTable.CellWidth
                .ComboCell.Clear

                For I = 1 To MAX_NUMBER_OF_STEPS
                    .ComboCell.AddItem ("Шаг " & I)
                Next
                .ComboCell.ListIndex = Manager.StepIndex
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case IDLE_FUNCTION_FIELD:
                .ComboCell.Left = .PropertyTable.Left + .PropertyTable.CellLeft
                .ComboCell.Top = .PropertyTable.Top + .PropertyTable.CellTop
                .ComboCell.Width = .PropertyTable.CellWidth
                .ComboCell.Clear

                For I = 1 To NUMBER_OF_FUNCS
                    .ComboCell.AddItem (FunctionsStrings(I - 1))
                Next
                .ComboCell.ListIndex = RecordStep.Bits And &HF
                .ComboCell.Visible = True
                .ComboCell.SetFocus

            Case Else
            
        End Select
        
        .TextCell.SelStart = 0
        .TextCell.SelLength = Len(.TextCell.Text)
        
    End With
    
End Sub

Public Sub EditPropertyForIdle(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP

    func_EditPropertyForIdle frm, 0&, RecordTitle, RecordStep
    
End Sub

' Функция-посредник
Private Sub func_SetCheckBoxForIdle(frm As FormMain, _
       ByVal ValveNumber As Integer, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordStep As TYPE_WPC_STEP)

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
        Case 18:
    
    End Select

End Sub

Public Sub SetCheckBoxForIdle(frm As FormMain, ValveNumber As Integer)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP

    func_SetCheckBoxForIdle frm, ValveNumber, 0&, RecordTitle, RecordStep
End Sub

' Функция-посредник
Private Sub func_SetComboPropertyForIdle(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordStep As TYPE_WPC_STEP)

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
                ' В режиме пропуска доступны для изменения только два поля
            Case IDLE_ENDSOUND_FIELD:
            Case IDLE_DOORUNLOCK_FIELD:

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

'<Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:23:07
'Причина: Модуль аналогичен по функционалу с ModuleWashOrRinsOrJolt>
'                        Case WPC_OPERATION_PAUS ' пауза
'</Удалил: Мезенцев Вячеслав, 17.06.2011 г. в 17:23:07>

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

            Case Else

        End Select

        SetModified True
    End With

    Exit Sub

ErrorHandler:
    Err.Clear
End Sub

Public Sub SetComboPropertyForIdle(frm As FormMain)
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP

    func_SetComboPropertyForIdle frm, 0&, RecordTitle, RecordStep
End Sub

' Функция-посредник
Private Function func_ValveEnabled(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordStep As TYPE_WPC_STEP, _
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
            Case 15:
            Case 16:
            Case 17:
            Case 18:

            Case Else
                func_ValveEnabled = False
        End Select
    End With
End Function

Public Function ValveEnabled(frm As FormMain, StepIndex As Integer, Num As Integer) As Boolean
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP

    ValveEnabled = func_ValveEnabled(frm, 0&, RecordTitle, RecordStep, StepIndex, Num)
End Function
