VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProgramManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'**
'@author <a href="mailto:unihomelab@ya.ru">Мезенцев Вячеслав</a>
'@revision Дата ревизии: 16.06.2011 г., время: 5:50:31
'@rem <h1><b>CProgramManager</b></h1>
'Модуль для работы с бинарным образом файла проекта. Здесь используются
'недокументированные приёмы для непосредственной работы с неуправляемыми указателями.
'<pre>
'--------------------------------------------------------------------------------
' Проект   :       Конфигуратор управляющих программ
' Модуль   :       CProgramManager
' Описание :       Модуль для работы с бинарным образом файла проекта
' Автор    :       Мезенцев Вячеслав
' Изменён  :       16.06.2011 г., время: 5:50:31
'--------------------------------------------------------------------------------
'
'[29.12.2010] Слава , привет!
'В приложении файл Си со структурами шагов и программ.
'
'Дополнительно сообщаю:
'
'1) На карте памяти расположен файл "data.bin"
'2) Размер файла 32кБ
'3) Все параметры программ расположены последовательно друг за другом...
'сначала идут параметры первой УП, затем второй.. затем третьей.... и т.д..
'4) Количество программ на карте памяти - 25
'5) Количество шагов в программе - 80
'6) Размер шага - 16 байт
'7) Дополнительно каждая программа имеет заголовок 16 байт
'8) Размер одной программы = 80 * 16 + 16  = 1296 байт
'
' В программной утилите должны быть предусмотрены следующие функции:
'
'1) Прочитать данные из файла data.bin (путь до файла указывает пользователь)
'2) Сохранить данные в файл data.bin (путь до файла указывает пользователь)
'3) Скопировать параметры одной УП в другую (пользователь должен указать,
'какая УП - исходник, а какая - приемник)
'4) Очистить параметры УП (пользователь должен указать номер УП)
'5) Вставить шаг со сдвигом всех шагов вправо
'6) Удалить шаг со сдвигом всех шагов влево
'7) Назначить шагу функцию
'8) Задать параметры шага с учетом назначенной функции
'9) Визуально отобразить, какие нагрузки будут включены на шаге, исходя
'из функции шага (и в зависимости от настроенных параметров)
'10)  ...
'
' Список того, что нужно доделать:
'- [+] в меню "Файл" добавить подменю с ранее открытыми файлами;
'- [+] при создании нового проекта не спрашивается о закрытии старого и сохранении
'  изменений, если они были;
'- [+] добавить форму редактирования свойств интерфейса программы: размеров ячеек,
'  шрифта;
'- добавить возможность вывода на печать (в pdf) в компактном виде образа программ;
'- при двойном клике на ячейке в таблице шагов сделать возможным локальное
'  редактирование параметра шага;
'- [+] придумать формат сохранения структуры программ в виде JSON файла для импорта и
'  экспорта данных в читабельном виде;
'- сделать возможность отката изменений на некоторую глубину, задаваемую в
'  настройках;
'- [+] ведение лога программы;
'- ...
'
'  Минусы Visual Basic 6.0 Pro:
'  1) Не поддерживаются пользовательские типы как поля в классах, из-за чего
'  очень не удобно организовывать иерархию в коде (например, CSettings).
'  2) Работа с неуправляемыми указателями невозможна стандартыми средствами.
'  3) Стандартные компоненты устарели. Например, чтобы организовать отлов клавиш,
'  кое-где приходится использовать хуки, что ведёт к нестабильности приложения.
'  4) Неудобная навигация по коду (устарела).
'  5) Язык не располагает наглядными средствами для побитовой работы с данными.
'  6) Убогий отладчик, который позволяет лишь поверхностно отлаживать код.
'  ...
'
'  Плюсы:
'  1) AxTools CodeSMART 2001 Add-in: AutoText, Code Analyzer, Code Builders, Code Formatter,
'  Code Snippets Manager, Code Statistics, Commenting Tools, Custom Code, Designer Edit Tools,
'  Enhanced Immediate Window, Error Handling, Extended Clipboard, Extended Find,
'  Enhanced Project Explorer, Spelling Checker, Spies, Task List.
'  2) VBdocman 2.0 Add-in by Peter Macej, http://www.vbdocman.com
'  3) VBcodePrint Add-in that saves system development and maintenance time and
'  money by enhancing source code printing.
'</pre>
Option Explicit

Private mvarFileLoaded As Boolean
Private mvarImageSize As Long
Private mvarProgramIndex As Long
Private mvarStepIndex As Long
Private mvarFileName As String
Private mvarData() As Byte

Public ProgramsCount As Long

' Конструктор
Private Sub Class_Initialize()
    ' Инициализация полей начальными значениями
    mvarImageSize = 0
    mvarProgramIndex = 0
    mvarStepIndex = 0
    ProgramsCount = 0
    FileLoaded = False

    ' Инициализируем массив текстовых шаблонов для упрощённого
    ' создания структуры данных для экспорта в JSON

    JSONStepsTemplates(WPC_OPERATION_IDLE) = ""

    JSONStepsTemplates(WPC_OPERATION_FILL) = "" _
       & """Type"": 1," _
       & """Pause"": false," _
       & """ColdWaterGate"": false," _
       & """HotWaterGate"": false," _
       & """RecycledWaterGate"": false," _
       & """Rotation"": true," _
       & """Level"": 15," _
       & """RotationTime"": 6," _
       & """PauseTime"": 12," _
       & """DrumSpeed"": 50"

    JSONStepsTemplates(WPC_OPERATION_DTRG) = "" _
       & """Type"": 2," _
       & """Pause"": false," _
       & """Rotation"": true," _
       & """Detergent_1_Time"": 0," _
       & """Detergent_2_Time"": 0," _
       & """Detergent_3_Time"": 0," _
       & """Detergent_4_Time"": 0," _
       & """Detergent_5_Time"": 0," _
       & """Detergent_6_Time"": 0," _
       & """Detergent_7_Time"": 0," _
       & """Detergent_8_Time"": 0," _
       & """Detergent_9_Time"": 0," _
       & """RotationTime"": 6," _
       & """PauseTime"": 12," _
       & """DrumSpeed"": 50"

    JSONStepsTemplates(WPC_OPERATION_HEAT) = "" _
       & """Type"": 3," _
       & """Pause"": false," _
       & """Rotation"": true," _
       & """Temperature"": 40," _
       & """RotationTime"": 6," _
       & """PauseTime"": 12," _
       & """DrumSpeed"": 50"

    JSONStepsTemplates(WPC_OPERATION_WASH) = "" _
       & """Type"": 4," _
       & """Pause"": false," _
       & """Rotation"": true," _
       & """Time"": 30," _
       & """RotationTime"": 6," _
       & """PauseTime"": 12," _
       & """DrumSpeed"": 50"

    JSONStepsTemplates(WPC_OPERATION_RINS) = JSONStepsTemplates(WPC_OPERATION_WASH)
    JSONStepsTemplates(WPC_OPERATION_JOLT) = JSONStepsTemplates(WPC_OPERATION_WASH)
    JSONStepsTemplates(WPC_OPERATION_PAUS) = JSONStepsTemplates(WPC_OPERATION_WASH)

    JSONStepsTemplates(WPC_OPERATION_DRAIN) = "" _
       & """Type"": 8," _
       & """Pause"": false," _
       & """DrainGate1"": true," _
       & """DrainGate2"": false," _
       & """Rotation"": true," _
       & """Level"": 0," _
       & """RotationTime"": 6," _
       & """PauseTime"": 12," _
       & """DrumSpeed1"": 50"

    JSONStepsTemplates(WPC_OPERATION_SPIN) = "" _
       & """Type"": 9," _
       & """Pause"": false," _
       & """DrainGate1"": true," _
       & """DrainGate2"": false," _
       & """DrumSpeed"": 800," _
       & """Time"": 2"

    JSONStepsTemplates(WPC_OPERATION_COOL) = "" _
       & """Type"": 10," _
       & """Pause"": false," _
       & """Fast"": false," _
       & """Rotation"": true," _
       & """Temperature"": 40," _
       & """ColdWaterTime"": 10," _
       & """RotationTime"": 6," _
       & """PauseTime"": 12," _
       & """DrumSpeed"": 50"
       
End Sub

Public Sub CopyProgram(ByVal FromIndex As Integer, ByVal ToIndex As Integer)

    ' Копируем программу
    CopyMemory ByVal DataPointer + ToIndex * PROGRAM_SIZE_IN_BYTES, _
       ByVal DataPointer + FromIndex * PROGRAM_SIZE_IN_BYTES, _
       PROGRAM_SIZE_IN_BYTES
       
End Sub

Public Sub DeletePrograms(index As Long, Count As Long)

    If (index + Count) < MAX_NUMBER_OF_PROGRAMS Then

        ' Копируем все программы вверх, замещая текущую
        CopyMemory ByVal DataPointer + _
           index * PROGRAM_SIZE_IN_BYTES, _
           ByVal DataPointer + _
           (index + Count) * PROGRAM_SIZE_IN_BYTES, _
           (MAX_NUMBER_OF_PROGRAMS - (index + Count)) * PROGRAM_SIZE_IN_BYTES
    
    End If

    ' Обнуляем всё после блока копировани
    ZeroMemory ByVal DataPointer + _
        (MAX_NUMBER_OF_PROGRAMS - Count) * PROGRAM_SIZE_IN_BYTES, Count * PROGRAM_SIZE_IN_BYTES

    Dim CRC8Value As Byte
    Dim I As Integer
    Dim Address As Long
    Dim Size As Long
        
    For I = MAX_NUMBER_OF_PROGRAMS - Count To MAX_NUMBER_OF_PROGRAMS - 1
    
        ' Установка заголовка программы по умолчанию
        SetDefaultProgramHeader I
    
        ' Пересчитываем CRC поле записи программы
        Address = I * PROGRAM_SIZE_IN_BYTES
        
        Size = PROGRAM_SIZE_IN_BYTES - 1
        
        CRC8Value = CalculateCRC8(Address + 1, Size)
        
        SetByte Address, CRC8Value
        
    Next
        
End Sub

Public Sub DeleteSteps(Program As Long, Step As Long, Count As Long)

    If (Step + Count) < MAX_NUMBER_OF_STEPS Then

        ' Копируем все шаги вверх, замещая текущий
        CopyMemory ByVal DataPointer + _
           Program * PROGRAM_SIZE_IN_BYTES + HEADER_SIZE_IN_BYTES + _
           Step * STEP_SIZE_IN_BYTES, _
           ByVal DataPointer + _
           Program * PROGRAM_SIZE_IN_BYTES + HEADER_SIZE_IN_BYTES + _
           (Step + Count) * STEP_SIZE_IN_BYTES, _
           (MAX_NUMBER_OF_STEPS - (Step + Count)) * STEP_SIZE_IN_BYTES

    End If
    
    ' Обнуляем последний шаг
    ZeroMemory ByVal DataPointer + _
       Program * PROGRAM_SIZE_IN_BYTES + HEADER_SIZE_IN_BYTES + _
       (MAX_NUMBER_OF_STEPS - Count) * STEP_SIZE_IN_BYTES, Count * STEP_SIZE_IN_BYTES
       
    ' Пересчитываем CRC поле записи программы
    Dim CRC8Value As Byte
    Dim Address As Long
    Dim Size As Long
    
    Address = ProgramIndex * PROGRAM_SIZE_IN_BYTES
    Size = PROGRAM_SIZE_IN_BYTES - 1
    
    CRC8Value = CalculateCRC8(Address + 1, Size)
    SetByte Address, CRC8Value
    
End Sub

' Функция-посредник
Private Sub func_InsertProgram(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE)
       
    Dim StepPointer As Long
    
    ' Копируем все программы вниз, начиная с текущей
    CopyMemory ByVal DataPointer + _
       (ProgramIndex + 1) * PROGRAM_SIZE_IN_BYTES, _
       ByVal DataPointer + ProgramIndex * PROGRAM_SIZE_IN_BYTES, _
       (MAX_NUMBER_OF_PROGRAMS - ProgramIndex - 1) * PROGRAM_SIZE_IN_BYTES
           
    ' Обнуляем текущую программу
    ZeroMemory ByVal DataPointer + _
        ProgramIndex * PROGRAM_SIZE_IN_BYTES, PROGRAM_SIZE_IN_BYTES
        
    SetDefaultProgramHeader ProgramIndex
        
    ' Пересчитываем CRC поле записи программы
    Dim CRC8Value As Byte
    Dim Address As Long
    Dim Size As Long
    
    Address = ProgramIndex * PROGRAM_SIZE_IN_BYTES
    Size = PROGRAM_SIZE_IN_BYTES - 1
    
    CRC8Value = CalculateCRC8(Address + 1, Size)
    SetByte Address, CRC8Value
    
End Sub
       
Public Sub InsertProgram(frm As FormMain)

    Dim RecordTitle As TYPE_WPC_TITLE
    
    func_InsertProgram frm, 0&, RecordTitle
    
End Sub

' Функция-посредник
Private Sub func_InsertStep(frm As FormMain, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordStep As TYPE_WPC_STEP, _
       ByVal FuncNumber As Integer)
       
    Dim StepPointer As Long
    
    ' Копируем все шаги вниз, относительно текущей
    CopyMemory ByVal DataPointer + _
       ProgramIndex * PROGRAM_SIZE_IN_BYTES + HEADER_SIZE_IN_BYTES + _
       (StepIndex + 1) * STEP_SIZE_IN_BYTES, _
       ByVal DataPointer + _
       ProgramIndex * PROGRAM_SIZE_IN_BYTES + HEADER_SIZE_IN_BYTES + _
       StepIndex * STEP_SIZE_IN_BYTES, _
       (MAX_NUMBER_OF_STEPS - (StepIndex + 1)) * STEP_SIZE_IN_BYTES

    ' Обнуляем текущий шаг
    ZeroMemory ByVal DataPointer + _
       ProgramIndex * PROGRAM_SIZE_IN_BYTES + HEADER_SIZE_IN_BYTES + _
       StepIndex * STEP_SIZE_IN_BYTES, STEP_SIZE_IN_BYTES
       
    StepPointer = DataPointer + ProgramIndex * PROGRAM_SIZE_IN_BYTES
       
    PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

    StepPointer = DataPointer + _
       ProgramIndex * PROGRAM_SIZE_IN_BYTES + HEADER_SIZE_IN_BYTES + _
       StepIndex * STEP_SIZE_IN_BYTES

    PutMem4 VarPtr(begin_of_pointers) + 8, ByVal StepPointer
    
    ' записать новое значение в поле типа функции
    RecordStep.Bits = RecordStep.Bits Or (FuncNumber And &HF)
    
    With frm
    
        Select Case FuncNumber And &HF
        
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
    
    End With
    
    ' Пересчитываем CRC поле записи программы
    Dim CRC8Value As Byte
    Dim Address As Long
    Dim Size As Long
    
    Address = ProgramIndex * PROGRAM_SIZE_IN_BYTES
    Size = PROGRAM_SIZE_IN_BYTES - 1
    
    CRC8Value = CalculateCRC8(Address + 1, Size)
    SetByte Address, CRC8Value
    
End Sub

' Вставляем пустой шаг в текущую позицию со сдвигом всех шагов вправо
Public Sub InsertStep(frm As FormMain, Optional FuncNumber As Integer = 0)
    
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP
    
    func_InsertStep frm, 0&, RecordTitle, RecordStep, FuncNumber

End Sub

Public Sub CloseFile()

    If FileLoaded Then
    
        FileName = ""
        FileLoaded = False
        
    End If
    
End Sub

Public Property Let FileLoaded(ByVal vData As Boolean)
    mvarFileLoaded = vData
End Property

Public Property Get FileLoaded() As Boolean
    FileLoaded = mvarFileLoaded
End Property

Public Function GetFunctionType(ByVal ProgNum As Long, ByVal StepNum As Long) As Integer
    
    Dim StepPointer As Long
    Dim Step As TYPE_WPC_STEP

    StepPointer = DataPointer + (ProgNum - 1) * PROGRAM_SIZE_IN_BYTES + _
       HEADER_SIZE_IN_BYTES + _
       (StepNum - 1) * STEP_SIZE_IN_BYTES

    CopyMemory Step, ByVal StepPointer, STEP_SIZE_IN_BYTES

    GetFunctionType = Step.Bits And &HF
    
End Function

Public Property Get DataPointer() As Long

    DataPointer = VarPtr(mvarData(0))
    
End Property

Public Sub CreateNewFile(FileName As String)
    
    ' Закрываем файл, если открыт
    If FileLoaded Then
    
        CloseFile

        ' Пользователь отменил операцию

        If FileLoaded Then Exit Sub
        
    End If

    ' Внутреннее имя файла
    mvarFileName = FileName

    'mvarImageSize = MAX_NUMBER_OF_PROGRAMS * PROGRAM_SIZE_IN_BYTES
    mvarImageSize = IMAGE_SIZE

    ' Инициализируем счётчики
    mvarProgramIndex = 0
    
    mvarStepIndex = 0
    
    ' Признак загрузки файла
    FileLoaded = True

    ' Устанавливаем количество программ
    ProgramsCount = MAX_NUMBER_OF_PROGRAMS
    
    ' Создаём пустой образ управляющей программы
    ReDim mvarData(0 To mvarImageSize - 1)

    ' Очистка программ и установка заголовков по умолчанию
    DeletePrograms 0, ProgramsCount
    
End Sub

Public Sub ClearProgramN(ByVal Num As Long)

    Dim b As Byte
    Dim ProgPointer As Long

    ProgPointer = DataPointer + (Num - 1) * PROGRAM_SIZE_IN_BYTES
    ZeroMemory ByVal ProgPointer, PROGRAM_SIZE_IN_BYTES

    ' Пересчитываем CRC поле записи программы
    b = CalculateCRC8((Num - 1) * PROGRAM_SIZE_IN_BYTES + 1, PROGRAM_SIZE_IN_BYTES - 1)
    SetByte (Num - 1) * PROGRAM_SIZE_IN_BYTES, b
    
End Sub

Public Sub ClearAll()

    Dim Cnt As Long

    For Cnt = 1 To ProgramsCount
    
        ClearProgramN (Cnt)
        
    Next
    
End Sub

Public Property Let StepIndex(ByVal vData As Long)
    mvarStepIndex = vData
End Property

Public Property Get StepIndex() As Long
    StepIndex = mvarStepIndex
End Property

Public Property Let ProgramIndex(ByVal vData As Long)
    mvarProgramIndex = vData
End Property

Public Property Get ProgramIndex() As Long
    ProgramIndex = mvarProgramIndex
End Property

Public Property Get ImageSize() As Long
    ImageSize = mvarImageSize
End Property

Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property

Public Property Get FileName() As String
    FileName = mvarFileName
End Property

Public Function GetByte(ByVal Offset As Long) As Byte

    Dim BytePointer As Long
    Dim ByteValue As Byte

    BytePointer = DataPointer + Offset
    CopyMemory ByteValue, ByVal BytePointer, 1

    GetByte = ByteValue
    
End Function

Public Sub SetByte(ByVal Offset As Long, ByVal ByteValue As Byte)

    Dim BytePointer As Long

    BytePointer = DataPointer + Offset
    CopyMemory ByVal BytePointer, ByteValue, 1

End Sub

Public Sub LoadFromFile(FileName As String)

    Dim fileid% ' Идентификатор файла
    Dim ProgsCount%

    ' Проверяем существование файла
    ' ...

    mvarFileName = FileName

    ' Получаем свободный идентификатор
    fileid% = FreeFile

    ' Получаем доступ к файлу
    'TODO: Дописать проверку на существование файла
    Open FileName For Binary Access Read As #fileid%

    mvarImageSize = LOF(fileid%)

    ' Изменяем размер динамического массива
    ReDim mvarData(0 To LOF(fileid%) - 1)

    ' Загружаем данные в байтовый массив
    Get #fileid%, , mvarData

    Close #fileid%

    ' Признак загрузки файла
    FileLoaded = True

    ' Устанавливаем количество программ
    ProgsCount% = mvarImageSize / PROGRAM_SIZE_IN_BYTES

    ' Инициализируем счётчики
    mvarProgramIndex = 0
    mvarStepIndex = 0
    ProgramsCount = ProgsCount%
    
End Sub

'**
'@param        begin_of_pointers Required. Long.
'@param        RecordTitle Required. TYPE_WPC_TITLE object.
'@param        RecordStep Required. TYPE_WPC_STEP object.
'@param        RecordFill Required. TYPE_WPC_FILL object.
'@param        RecordDetergent Required. TYPE_WPC_DETERGENT object.
'@param        RecordHeat Required. TYPE_WPC_HEAT object.
'@param        RecordWash Required. TYPE_WPC_WASH object.
'@param        RecordDrain Required. TYPE_WPC_DRAIN object.
'@param        RecordSpin Required. TYPE_WPC_SPIN object.
'@param        RecordCool Required. TYPE_WPC_COOL object.
'@return       Boolean.
'@rem Вспомогательная функция экспорта в JSON формат.
Private Function func_ConvertToJSON(ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordStep As TYPE_WPC_STEP, _
       ByRef RecordFill As TYPE_WPC_FILL, _
       ByRef RecordDetergent As TYPE_WPC_DETERGENT, _
       ByRef RecordHeat As TYPE_WPC_HEAT, _
       ByRef RecordWash As TYPE_WPC_WASH, _
       ByRef RecordDrain As TYPE_WPC_DRAIN, _
       ByRef RecordSpin As TYPE_WPC_SPIN, _
       ByRef RecordCool As TYPE_WPC_COOL) As String

    On Local Error GoTo errHandler

    Dim b As Byte
    Dim I As Integer
    Dim J As Integer
    Dim FNum As Integer
    Dim StepPointer As Long
    Dim S As String
    Dim sInputJson As String
    Dim ProgsArray As Object
    Dim ProgElem As Object
    Dim StepElem As Object

    sInputJson = "[{""ProgID"":""""}]"

    Set ProgsArray = JSON.parse(sInputJson)

    ' Записываем уникальный идентификатор программы
    ProgsArray.Item(1).Item("ProgID") = ProgramGUID

    I = ProgramsCount

    Do While I > 0
    
        ' Формируем шаблон для программы с пустым массивом шагов
        sInputJson = "{""Title"":"""", ""EndSound"": true, ""DoorUnlock"": true, ""Steps"":[]}"

        ' С помощью шаблона создаём объект в памяти для описания программы
        Set ProgElem = JSON.parse(sInputJson)

        ' Инициализируем указатель RecordTitle адресом начала текущей программы
        StepPointer = DataPointer + (I - 1) * PROGRAM_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

        ' Вычисляем признак пустой программы
        b = CalculateCRC8((I - 1) * PROGRAM_SIZE_IN_BYTES, PROGRAM_SIZE_IN_BYTES)

        ' Дополнительная проверка для пустой программы
        If b = CRC8_FOR_DEFAULT_PROGRAM Then
        
            Dim k As Long, Sum As Long
            
            Sum = 0
            
            For k = 2 To 31
            
                Sum = Sum + GetByte((I - 1) * PROGRAM_SIZE_IN_BYTES + k)
            
            Next
            
            If Sum > 0 Then b = 0
            
        End If

        If b = CRC8_FOR_DEFAULT_PROGRAM Then
        
            ' Пропускаем пустые программы начиная с конца списка
            ' Проверка осуществляется по CRC от всей программы

            If ProgsArray.Count > 1 Then
            
                ProgsArray.Add ProgElem, before:=2
                
            End If
            
        Else
        
            S = ""

            For J = 1 To PROG_NAME_LENGTH - 1

                If RecordTitle.ProgName(J) <> 0 Then S = S & Chr$(RecordTitle.ProgName(J))
                
            Next

            ProgElem.Item("Title") = CStr(S)
            ProgElem.Item("EndSound") = CBool(RecordTitle.LowBits And &H1)
            ProgElem.Item("DoorUnlock") = CBool((RecordTitle.LowBits And &H2) / 2 ^ 1)

            ' Добавляем шаги к программе
            J = MAX_NUMBER_OF_STEPS

            Do While J > 0
            
                PutMem4 VarPtr(begin_of_pointers) + 8, _
                   ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                FNum = (RecordStep.Bits And &HF)

                ' Формируем объект для шага в зависимости от типа функции шага
                Set StepElem = JSON.parse("{" & JSONStepsTemplates(FNum) & "}")

                Select Case FNum
                
                    Case WPC_OPERATION_IDLE
                        PutMem4 VarPtr(begin_of_pointers) + 8, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                    Case WPC_OPERATION_FILL
                        PutMem4 VarPtr(begin_of_pointers) + 12, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        StepElem.Item("Pause") = CBool((RecordFill.Bits And &H10) / &H10)
                        StepElem.Item("ColdWaterGate") = CBool((RecordFill.Bits And &H20) / &H20)
                        StepElem.Item("HotWaterGate") = CBool((RecordFill.Bits And &H40) / &H40)
                        StepElem.Item("RecycledWaterGate") = CBool((RecordFill.Bits And &H80) / &H80)
                        StepElem.Item("Rotation") = CBool((RecordFill.Bits And &H100) / &H100)

                        StepElem.Item("Level") = CByte(RecordFill.Level)
                        StepElem.Item("RotationTime") = CByte(RecordFill.RotationTime)
                        StepElem.Item("PauseTime") = CByte(RecordFill.PauseTime)
                        StepElem.Item("DrumSpeed") = CByte(RecordFill.DrumSpeed)

                    Case WPC_OPERATION_DTRG
                        PutMem4 VarPtr(begin_of_pointers) + 16, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        StepElem.Item("Pause") = CBool((RecordDetergent.Bits And &H10) / &H10)
                        StepElem.Item("Rotation") = CBool((RecordDetergent.Bits And &H20) / &H20)

                        StepElem.Item("Detergent_1_Time") = CByte(RecordDetergent.Detergent_1_Time)
                        StepElem.Item("Detergent_2_Time") = CByte(RecordDetergent.Detergent_2_Time)
                        StepElem.Item("Detergent_3_Time") = CByte(RecordDetergent.Detergent_3_Time)
                        StepElem.Item("Detergent_4_Time") = CByte(RecordDetergent.Detergent_4_Time)
                        StepElem.Item("Detergent_5_Time") = CByte(RecordDetergent.Detergent_5_Time)
                        StepElem.Item("Detergent_6_Time") = CByte(RecordDetergent.Detergent_6_Time)
                        StepElem.Item("Detergent_7_Time") = CByte(RecordDetergent.Detergent_7_Time)
                        StepElem.Item("Detergent_8_Time") = CByte(RecordDetergent.Detergent_8_Time)
                        StepElem.Item("Detergent_9_Time") = CByte(RecordDetergent.Detergent_9_Time)

                        StepElem.Item("RotationTime") = CByte(RecordDetergent.RotationTime)
                        StepElem.Item("PauseTime") = CByte(RecordDetergent.PauseTime)
                        StepElem.Item("DrumSpeed") = CByte(RecordDetergent.DrumSpeed)

                    Case WPC_OPERATION_HEAT
                        PutMem4 VarPtr(begin_of_pointers) + 20, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        StepElem.Item("Pause") = CBool((RecordHeat.Bits And &H10) / &H10)
                        StepElem.Item("Rotation") = CBool((RecordHeat.Bits And &H20) / &H20)

                        StepElem.Item("Temperature") = CByte(RecordHeat.Temperature)
                        StepElem.Item("RotationTime") = CByte(RecordHeat.RotationTime)
                        StepElem.Item("PauseTime") = CByte(RecordHeat.PauseTime)
                        StepElem.Item("DrumSpeed") = CByte(RecordHeat.DrumSpeed)

                    Case WPC_OPERATION_WASH, WPC_OPERATION_RINS, WPC_OPERATION_JOLT, WPC_OPERATION_PAUS
                        PutMem4 VarPtr(begin_of_pointers) + 24, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        StepElem.Item("Type") = CByte(RecordWash.Bits And &HF)
                        StepElem.Item("Pause") = CBool((RecordWash.Bits And &H10) / &H10)
                        StepElem.Item("Rotation") = CBool((RecordWash.Bits And &H20) / &H20)

                        StepElem.Item("Time") = CByte(RecordWash.Time)
                        StepElem.Item("RotationTime") = CByte(RecordWash.RotationTime)
                        StepElem.Item("PauseTime") = CByte(RecordWash.PauseTime)
                        StepElem.Item("DrumSpeed") = CByte(RecordWash.DrumSpeed)

                    Case WPC_OPERATION_DRAIN
                        PutMem4 VarPtr(begin_of_pointers) + 28, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        StepElem.Item("Pause") = CBool((RecordDrain.Bits And &H10) / &H10)
                        StepElem.Item("DrainGate1") = CBool((RecordDrain.Bits And &H20) / &H20)
                        StepElem.Item("DrainGate2") = CBool((RecordDrain.Bits And &H40) / &H40)
                        StepElem.Item("Rotation") = CBool((RecordDrain.Bits And &H80) / &H80)

                        StepElem.Item("Level") = CByte(RecordDrain.Level)
                        StepElem.Item("RotationTime") = CByte(RecordDrain.RotationTime)
                        StepElem.Item("PauseTime") = CByte(RecordDrain.PauseTime)
                        StepElem.Item("DrumSpeed1") = CByte(RecordDrain.DrumSpeed1)

                    Case WPC_OPERATION_SPIN
                        PutMem4 VarPtr(begin_of_pointers) + 32, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        StepElem.Item("Pause") = CBool((RecordSpin.Bits And &H10) / &H10)
                        StepElem.Item("DrainGate1") = CBool((RecordSpin.Bits And &H20) / &H20)
                        StepElem.Item("DrainGate2") = CBool((RecordSpin.Bits And &H40) / &H40)

                        StepElem.Item("DrumSpeed") = CInt(RecordSpin.DrumSpeed)
                        StepElem.Item("Time") = CByte(RecordSpin.Time)

                    Case WPC_OPERATION_COOL
                        PutMem4 VarPtr(begin_of_pointers) + 36, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        StepElem.Item("Pause") = CBool((RecordCool.Bits And &H10) / &H10)
                        StepElem.Item("Fast") = CBool((RecordCool.Bits And &H20) / &H20)
                        StepElem.Item("Rotation") = CBool((RecordCool.Bits And &H40) / &H40)

                        StepElem.Item("Temperature") = CByte(RecordCool.Temperature)
                        StepElem.Item("ColdWaterTime") = CByte(RecordCool.ColdWaterTime)
                        StepElem.Item("RotationTime") = CByte(RecordCool.RotationTime)
                        StepElem.Item("PauseTime") = CByte(RecordCool.PauseTime)
                        StepElem.Item("DrumSpeed") = CByte(RecordCool.DrumSpeed)

                End Select

                ' Выводим информацию в сжатом виде, без пустых шагов

                If ProgElem.Item("Steps").Count > 0 Then

                    If StepElem Is Nothing Then
                    
                        ProgElem.Item("Steps").Add JSON.parse("{}"), before:=1
                        
                    Else
                    
                        ProgElem.Item("Steps").Add StepElem, before:=1
                        
                    End If
                    
                Else

                    If Not StepElem Is Nothing Then ProgElem.Item("Steps").Add StepElem
                    
                End If

                Set StepElem = Nothing

                J = J - 1
                
            Loop

            ' Выводим информацию

            If ProgsArray.Count = 1 Then
                ProgsArray.Add ProgElem
            Else
                ProgsArray.Add ProgElem, before:=2
                
            End If
            
        End If

        Set ProgElem = Nothing

        I = I - 1
        
    Loop

    func_ConvertToJSON = JSON.toString(ProgsArray)
    
    Set ProgsArray = Nothing
    
    Exit Function

errHandler:

    Set ProgsArray = Nothing
    
    func_ConvertToJSON = ""
    
End Function

Public Function ToJSON() As String

    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP
    Dim RecordFill As TYPE_WPC_FILL
    Dim RecordDetergent As TYPE_WPC_DETERGENT
    Dim RecordHeat As TYPE_WPC_HEAT
    Dim RecordWash As TYPE_WPC_WASH
    Dim RecordDrain As TYPE_WPC_DRAIN
    Dim RecordSpin As TYPE_WPC_SPIN
    Dim RecordCool As TYPE_WPC_COOL
    
    Dim Result As String
    
    ToJSON = func_ConvertToJSON(0&, RecordTitle, RecordStep, RecordFill, RecordDetergent, _
       RecordHeat, RecordWash, RecordDrain, RecordSpin, RecordCool)
       
End Function

Public Function toString() As String

    Dim sAns As String
    Dim iPos As String
    
    sAns = StrConv(mvarData, vbUnicode)
    
    toString = sAns

End Function

Public Sub FromString(Value As String)

    Dim I As Long
    
    ReDim mvarData(Len(Value) - 1) As Byte
    
    mvarData = StrConv(Value, vbFromUnicode)
    
End Sub

'**
'@param        FileName Required. String.
Private Function func_ExportToJSON(FileName As String) As Boolean

    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP
    Dim RecordFill As TYPE_WPC_FILL
    Dim RecordDetergent As TYPE_WPC_DETERGENT
    Dim RecordHeat As TYPE_WPC_HEAT
    Dim RecordWash As TYPE_WPC_WASH
    Dim RecordDrain As TYPE_WPC_DRAIN
    Dim RecordSpin As TYPE_WPC_SPIN
    Dim RecordCool As TYPE_WPC_COOL
    
    Dim Result As String
    
    Result = func_ConvertToJSON(0&, RecordTitle, RecordStep, RecordFill, RecordDetergent, _
       RecordHeat, RecordWash, RecordDrain, RecordSpin, RecordCool)
    
    ' Сохраняем данные в формате JSON
    SaveToJSONFile FileName, Result
    
    func_ExportToJSON = True

End Function

Public Function ExportToJSON(FileName As String) As Boolean

    ' Вызываем вспомогательную функцию, которая позволяет работать
    ' с указателями напрямую
    ExportToJSON = func_ExportToJSON(FileName)
    
End Function

'**
'@param        FileName Required. String.
'@param        begin_of_pointers Required. Long.
'@param        RecordTitle Required. TYPE_WPC_TITLE object.
'@param        RecordStep Required. TYPE_WPC_STEP object.
'@param        RecordFill Required. TYPE_WPC_FILL object.
'@param        RecordDetergent Required. TYPE_WPC_DETERGENT object.
'@param        RecordHeat Required. TYPE_WPC_HEAT object.
'@param        RecordWash Required. TYPE_WPC_WASH object.
'@param        RecordDrain Required. TYPE_WPC_DRAIN object.
'@param        RecordSpin Required. TYPE_WPC_SPIN object.
'@param        RecordCool Required. TYPE_WPC_COOL object.
'@return       Boolean.
'@rem Вспомогательная функция для импорта из JSON формата.
Private Function func_ImportFromJSON(sInputJson As String, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordStep As TYPE_WPC_STEP, _
       ByRef RecordFill As TYPE_WPC_FILL, _
       ByRef RecordDetergent As TYPE_WPC_DETERGENT, _
       ByRef RecordHeat As TYPE_WPC_HEAT, _
       ByRef RecordWash As TYPE_WPC_WASH, _
       ByRef RecordDrain As TYPE_WPC_DRAIN, _
       ByRef RecordSpin As TYPE_WPC_SPIN, _
       ByRef RecordCool As TYPE_WPC_COOL) As Boolean

    On Local Error GoTo errHandler

    Dim b As Byte
    Dim ProgsCount As Integer
    Dim I As Integer
    Dim J As Integer
    Dim StepPointer As Long
    Dim Title As String

    Dim p As Object

    ' Производим разбор данных из файла
    Set p = JSON.parse(sInputJson)

    ' Если уникальный идентификатор формата файла не совпадает
    ' с поддерживаемыми программой, то прекращаем разбор

    If ProgramGUID <> p.Item(1).Item("ProgID") Then
    
        Set p = Nothing
        FileLoaded = False
        func_ImportFromJSON = False
        
        Exit Function
        
    End If
    
    For I = 1 To ProgramsCount
    
        If LimitsLoaded Then SetDefaultProgramHeader I

        ' Пересчитываем CRC поле записи программы

        b = CalculateCRC8((I - 1) * PROGRAM_SIZE_IN_BYTES + 1, PROGRAM_SIZE_IN_BYTES - 1)
        SetByte (I - 1) * PROGRAM_SIZE_IN_BYTES, b
        
    Next

    'TODO: Добавить ловушки для возможных ошибок на всех этапах считывания
    ' Узнаём количество программ
    ProgsCount = p.Count - 1

    For I = 1 To ProgsCount
    
        StepPointer = DataPointer + (I - 1) * PROGRAM_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

        ' Общие параметры программы

        Select Case CBool(p.Item(I + 1).Item("EndSound"))
            Case False: RecordTitle.LowBits = RecordTitle.LowBits And &HFFFE
            Case True: RecordTitle.LowBits = RecordTitle.LowBits Or &H1
        End Select

        Select Case CBool(p.Item(I + 1).Item("DoorUnlock"))
            Case False: RecordTitle.LowBits = RecordTitle.LowBits And &HFFFD
            Case True: RecordTitle.LowBits = RecordTitle.LowBits Or &H2
        End Select

        Title = p.Item(I + 1).Item("Title")

        For J = 1 To PROG_NAME_LENGTH - 1

            If J <= Len(Title) Then
                RecordTitle.ProgName(J) = Asc(Mid$(Title, J, 1))
            Else
                RecordTitle.ProgName(J) = 0
            End If
        Next

        ' По соглашению, последний символ в имени программы должен быть нулевым
        ' Это сделано для совместимости с ASCIIZ строками в языке С
        RecordTitle.ProgName(PROG_NAME_LENGTH) = 0

        Dim FType As Integer
        Dim Step As Object

        ' В зависимости от типа шага заполняем структуры шагов

        For J = 1 To p.Item(I + 1).Item("Steps").Count
        
            Set Step = p.Item(I + 1).Item("Steps").Item(J)

            If Not Step Is Nothing Then
            
                FType = CByte(Step.Item("Type"))

                Select Case FType
                
                    Case WPC_OPERATION_IDLE
                        PutMem4 VarPtr(begin_of_pointers) + 8, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                    Case WPC_OPERATION_FILL
                        PutMem4 VarPtr(begin_of_pointers) + 12, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        RecordFill.Bits = FType

                        ' Специальные параметры шага

                        Select Case CBool(Step.Item("Pause"))
                            Case False: RecordFill.Bits = RecordFill.Bits And &HFFEF
                            Case True: RecordFill.Bits = RecordFill.Bits Or &H10
                        End Select

                        Select Case CBool(Step.Item("ColdWaterGate"))
                            Case False: RecordFill.Bits = RecordFill.Bits And &HFFDF
                            Case True: RecordFill.Bits = RecordFill.Bits Or &H20
                        End Select

                        Select Case CBool(Step.Item("HotWaterGate"))
                            Case False: RecordFill.Bits = RecordFill.Bits And &HFFBF
                            Case True: RecordFill.Bits = RecordFill.Bits Or &H40
                        End Select

                        Select Case CBool(Step.Item("RecycledWaterGate"))
                            Case False: RecordFill.Bits = RecordFill.Bits And &HFF7F
                            Case True: RecordFill.Bits = RecordFill.Bits Or &H80
                        End Select

                        Select Case CBool(Step.Item("Rotation"))
                            Case False: RecordFill.Bits = RecordFill.Bits And &HFEFF
                            Case True: RecordFill.Bits = RecordFill.Bits Or &H100
                        End Select

                        RecordFill.Level = CByte(Step.Item("Level"))
                        RecordFill.RotationTime = CByte(Step.Item("RotationTime"))
                        RecordFill.PauseTime = CByte(Step.Item("PauseTime"))
                        RecordFill.DrumSpeed = CByte(Step.Item("DrumSpeed"))

                    Case WPC_OPERATION_DTRG
                        PutMem4 VarPtr(begin_of_pointers) + 16, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        RecordDetergent.Bits = FType

                        ' Специальные параметры шага

                        Select Case CBool(Step.Item("Pause"))
                            Case False: RecordDetergent.Bits = RecordDetergent.Bits And &HFFEF
                            Case True: RecordDetergent.Bits = RecordDetergent.Bits Or &H10
                        End Select

                        Select Case CBool(Step.Item("Rotation"))
                            Case False: RecordDetergent.Bits = RecordDetergent.Bits And &HFFDF
                            Case True: RecordDetergent.Bits = RecordDetergent.Bits Or &H20
                        End Select

                        RecordDetergent.Detergent_1_Time = CByte(Step.Item("Detergent_1_Time"))
                        RecordDetergent.Detergent_2_Time = CByte(Step.Item("Detergent_2_Time"))
                        RecordDetergent.Detergent_3_Time = CByte(Step.Item("Detergent_3_Time"))
                        RecordDetergent.Detergent_4_Time = CByte(Step.Item("Detergent_4_Time"))
                        RecordDetergent.Detergent_5_Time = CByte(Step.Item("Detergent_5_Time"))
                        RecordDetergent.Detergent_6_Time = CByte(Step.Item("Detergent_6_Time"))
                        RecordDetergent.Detergent_7_Time = CByte(Step.Item("Detergent_7_Time"))
                        RecordDetergent.Detergent_8_Time = CByte(Step.Item("Detergent_8_Time"))
                        RecordDetergent.Detergent_9_Time = CByte(Step.Item("Detergent_9_Time"))

                        RecordDetergent.RotationTime = CByte(Step.Item("RotationTime"))
                        RecordDetergent.PauseTime = CByte(Step.Item("PauseTime"))
                        RecordDetergent.DrumSpeed = CByte(Step.Item("DrumSpeed"))

                    Case WPC_OPERATION_HEAT
                        PutMem4 VarPtr(begin_of_pointers) + 20, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        RecordHeat.Bits = FType

                        ' Специальные параметры шага

                        Select Case CBool(Step.Item("Pause"))
                            Case False: RecordHeat.Bits = RecordHeat.Bits And &HFFEF
                            Case True: RecordHeat.Bits = RecordHeat.Bits Or &H10
                        End Select

                        Select Case CBool(Step.Item("Rotation"))
                            Case False: RecordHeat.Bits = RecordHeat.Bits And &HFFDF
                            Case True: RecordHeat.Bits = RecordHeat.Bits Or &H20
                        End Select

                        RecordHeat.Temperature = CByte(Step.Item("Temperature"))
                        RecordHeat.RotationTime = CByte(Step.Item("RotationTime"))
                        RecordHeat.PauseTime = CByte(Step.Item("PauseTime"))
                        RecordHeat.DrumSpeed = CByte(Step.Item("DrumSpeed"))

                    Case WPC_OPERATION_WASH, WPC_OPERATION_RINS, WPC_OPERATION_JOLT, WPC_OPERATION_PAUS
                        PutMem4 VarPtr(begin_of_pointers) + 24, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        RecordWash.Bits = FType

                        ' Специальные параметры шага

                        Select Case CBool(Step.Item("Pause"))
                            Case False: RecordWash.Bits = RecordWash.Bits And &HFFEF
                            Case True: RecordWash.Bits = RecordWash.Bits Or &H10
                        End Select

                        Select Case CBool(Step.Item("Rotation"))
                            Case False: RecordWash.Bits = RecordWash.Bits And &HFFDF
                            Case True: RecordWash.Bits = RecordWash.Bits Or &H20
                        End Select

                        RecordWash.Time = CByte(Step.Item("Time"))
                        RecordWash.RotationTime = CByte(Step.Item("RotationTime"))
                        RecordWash.PauseTime = CByte(Step.Item("PauseTime"))
                        RecordWash.DrumSpeed = CByte(Step.Item("DrumSpeed"))

                    Case WPC_OPERATION_DRAIN
                        PutMem4 VarPtr(begin_of_pointers) + 28, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        RecordDrain.Bits = FType

                        ' Специальные параметры шага

                        Select Case CBool(Step.Item("Pause"))
                            Case False: RecordDrain.Bits = RecordDrain.Bits And &HFFEF
                            Case True: RecordDrain.Bits = RecordDrain.Bits Or &H10
                        End Select

                        Select Case CBool(Step.Item("DrainGate1"))
                            Case False: RecordDrain.Bits = RecordDrain.Bits And &HFFDF
                            Case True: RecordDrain.Bits = RecordDrain.Bits Or &H20
                        End Select

                        Select Case CBool(Step.Item("DrainGate2"))
                            Case False: RecordDrain.Bits = RecordDrain.Bits And &HFFBF
                            Case True: RecordDrain.Bits = RecordDrain.Bits Or &H40
                        End Select

                        Select Case CBool(Step.Item("Rotation"))
                            Case False: RecordDrain.Bits = RecordDrain.Bits And &HFF7F
                            Case True: RecordDrain.Bits = RecordDrain.Bits Or &H80
                        End Select

                        RecordDrain.Level = CByte(Step.Item("Level"))
                        RecordDrain.RotationTime = CByte(Step.Item("RotationTime"))
                        RecordDrain.PauseTime = CByte(Step.Item("PauseTime"))
                        RecordDrain.DrumSpeed1 = CByte(Step.Item("DrumSpeed1"))

                    Case WPC_OPERATION_SPIN
                        PutMem4 VarPtr(begin_of_pointers) + 32, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        RecordSpin.Bits = FType

                        ' Специальные параметры шага

                        Select Case CBool(Step.Item("Pause"))
                            Case False: RecordSpin.Bits = RecordSpin.Bits And &HFFEF
                            Case True: RecordSpin.Bits = RecordSpin.Bits Or &H10
                        End Select

                        Select Case CBool(Step.Item("DrainGate1"))
                            Case False: RecordSpin.Bits = RecordSpin.Bits And &HFFDF
                            Case True: RecordSpin.Bits = RecordSpin.Bits Or &H20
                        End Select

                        Select Case CBool(Step.Item("DrainGate2"))
                            Case False: RecordSpin.Bits = RecordSpin.Bits And &HFFBF
                            Case True: RecordSpin.Bits = RecordSpin.Bits Or &H40
                        End Select

                        RecordSpin.DrumSpeed = CInt(Step.Item("DrumSpeed"))
                        RecordSpin.Time = CByte(Step.Item("Time"))

                    Case WPC_OPERATION_COOL
                        PutMem4 VarPtr(begin_of_pointers) + 36, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        RecordCool.Bits = FType

                        ' Специальные параметры шага

                        Select Case CBool(Step.Item("Pause"))
                            Case False: RecordCool.Bits = RecordCool.Bits And &HFFEF
                            Case True: RecordCool.Bits = RecordCool.Bits Or &H10
                        End Select

                        Select Case CBool(Step.Item("Fast"))
                            Case False: RecordCool.Bits = RecordCool.Bits And &HFFDF
                            Case True: RecordCool.Bits = RecordCool.Bits Or &H20
                        End Select

                        Select Case CBool(Step.Item("Rotation"))
                            Case False: RecordCool.Bits = RecordCool.Bits And &HFFBF
                            Case True: RecordCool.Bits = RecordCool.Bits Or &H40
                        End Select

                        RecordCool.Temperature = CByte(Step.Item("Temperature"))
                        RecordCool.ColdWaterTime = CByte(Step.Item("ColdWaterTime"))
                        RecordCool.RotationTime = CByte(Step.Item("RotationTime"))
                        RecordCool.PauseTime = CByte(Step.Item("PauseTime"))
                        RecordCool.DrumSpeed = CByte(Step.Item("DrumSpeed"))

                End Select
            End If
        Next

        ' Пересчитываем CRC поле записи программы
        b = CalculateCRC8((I - 1) * PROGRAM_SIZE_IN_BYTES + 1, PROGRAM_SIZE_IN_BYTES - 1)
        SetByte (I - 1) * PROGRAM_SIZE_IN_BYTES, b
    Next

    FileLoaded = True
    Set p = Nothing
    func_ImportFromJSON = True

    Exit Function

errHandler:
    Set p = Nothing
    ' Признак загрузки файла
    FileLoaded = False
    func_ImportFromJSON = False
    
End Function

Public Function FromJSON(Value As String)

    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP
    Dim RecordFill As TYPE_WPC_FILL
    Dim RecordDetergent As TYPE_WPC_DETERGENT
    Dim RecordHeat As TYPE_WPC_HEAT
    Dim RecordWash As TYPE_WPC_WASH
    Dim RecordDrain As TYPE_WPC_DRAIN
    Dim RecordSpin As TYPE_WPC_SPIN
    Dim RecordCool As TYPE_WPC_COOL

    ' Вызываем вспомогательную функцию, которая позволяет работать
    ' с указателями напрямую
    FromJSON = func_ImportFromJSON(Value, 0&, RecordTitle, RecordStep, RecordFill, RecordDetergent, _
       RecordHeat, RecordWash, RecordDrain, RecordSpin, RecordCool)

End Function

Public Function ImportFromJSON(FileName As String) As Boolean
    
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP
    Dim RecordFill As TYPE_WPC_FILL
    Dim RecordDetergent As TYPE_WPC_DETERGENT
    Dim RecordHeat As TYPE_WPC_HEAT
    Dim RecordWash As TYPE_WPC_WASH
    Dim RecordDrain As TYPE_WPC_DRAIN
    Dim RecordSpin As TYPE_WPC_SPIN
    Dim RecordCool As TYPE_WPC_COOL

    Dim sInputJson As String
       
    ' Создаём новый образ проекта в памяти и инициализируем его
    CreateNewFile Left$(FileName, InStrRev(FileName, ".")) + "bin"

    ' Обнуляем образ в памяти
    ClearAll
    
    ' Считываем файл и декодируем его
    sInputJson = FromUTF8(LoadFromJSONFile(FileName))
    
    ' Вызываем вспомогательную функцию, которая позволяет работать
    ' с указателями напрямую
    ImportFromJSON = func_ImportFromJSON(sInputJson, 0&, RecordTitle, RecordStep, RecordFill, RecordDetergent, _
       RecordHeat, RecordWash, RecordDrain, RecordSpin, RecordCool)
       
    ' Если преобразование прошло неудачно, то закрываем файл
    If ImportFromJSON = False Then CloseFile
    
End Function

Private Function func_ImportProgramFromClipboard(frm As FormMain, _
       ByVal Program As Long, _
       ByRef ProgramObj As Object, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE) As Boolean
       
    On Local Error GoTo errHandler
    
    Dim b As Byte
    Dim J As Integer
    Dim ProgsCount As Integer
    Dim FType As Integer
    Dim StepPointer As Long
    Dim Title As String
    Dim Step As Object
                               
    Dim RecordStep As TYPE_WPC_STEP
    Dim RecordFill As TYPE_WPC_FILL
    Dim RecordDetergent As TYPE_WPC_DETERGENT
    Dim RecordHeat As TYPE_WPC_HEAT
    Dim RecordWash As TYPE_WPC_WASH
    Dim RecordDrain As TYPE_WPC_DRAIN
    Dim RecordSpin As TYPE_WPC_SPIN
    Dim RecordCool As TYPE_WPC_COOL
                               
    StepPointer = DataPointer + Program * PROGRAM_SIZE_IN_BYTES
    PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer
    
    InsertProgram frm
    
    ' Общие параметры программы

    Select Case CBool(ProgramObj.Item("EndSound"))
        Case False: RecordTitle.LowBits = RecordTitle.LowBits And &HFFFE
        Case True: RecordTitle.LowBits = RecordTitle.LowBits Or &H1
    End Select

    Select Case CBool(ProgramObj.Item("DoorUnlock"))
        Case False: RecordTitle.LowBits = RecordTitle.LowBits And &HFFFD
        Case True: RecordTitle.LowBits = RecordTitle.LowBits Or &H2
    End Select
    
    Title = ProgramObj.Item("Title")

    For J = 1 To PROG_NAME_LENGTH - 1

        If J <= Len(Title) Then
            RecordTitle.ProgName(J) = Asc(Mid$(Title, J, 1))
        Else
            RecordTitle.ProgName(J) = 0
        End If
        
    Next
        
    ' По соглашению, последний символ в имени программы должен быть нулевым
    ' Это сделано для совместимости с ASCIIZ строками в языке С
    RecordTitle.ProgName(PROG_NAME_LENGTH) = 0

    ' В зависимости от типа шага заполняем структуры шагов

    For J = 1 To ProgramObj.Item("Steps").Count
    
        Set Step = ProgramObj.Item("Steps").Item(J)

        ' Вызываем вспомогательную функцию, которая позволяет работать
        ' с указателями напрямую
        func_ImportStepFromClipboard frm, Program, Step, 0&, RecordTitle, RecordStep, RecordFill, RecordDetergent, _
           RecordHeat, RecordWash, RecordDrain, RecordSpin, RecordCool
        
        Set Step = Nothing
                
    Next
            
    ' Пересчитываем CRC поле записи программы
    b = CalculateCRC8(Program * PROGRAM_SIZE_IN_BYTES + 1, PROGRAM_SIZE_IN_BYTES - 1)
    SetByte Program * PROGRAM_SIZE_IN_BYTES, b
    
    func_ImportProgramFromClipboard = False
    
    Exit Function
    
errHandler:
    
    func_ImportProgramFromClipboard = True
    
End Function
       
Private Function func_ImportStepFromClipboard(frm As FormMain, _
       ByVal Program As Long, _
       ByRef Step As Object, _
       ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordStep As TYPE_WPC_STEP, _
       ByRef RecordFill As TYPE_WPC_FILL, _
       ByRef RecordDetergent As TYPE_WPC_DETERGENT, _
       ByRef RecordHeat As TYPE_WPC_HEAT, _
       ByRef RecordWash As TYPE_WPC_WASH, _
       ByRef RecordDrain As TYPE_WPC_DRAIN, _
       ByRef RecordSpin As TYPE_WPC_SPIN, _
       ByRef RecordCool As TYPE_WPC_COOL) As Boolean

    On Local Error GoTo errHandler
    
    Dim b As Byte
    Dim ProgsCount As Integer
    Dim StepPointer As Long
    Dim FType As Integer
                               
    StepPointer = DataPointer + Program * PROGRAM_SIZE_IN_BYTES
    PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer
    
    FType = CByte(Step.Item("Type"))

    InsertStep frm, FType

    Select Case FType
    
        Case WPC_OPERATION_IDLE
        
            PutMem4 VarPtr(begin_of_pointers) + 8, _
               ByVal StepPointer + HEADER_SIZE_IN_BYTES + StepIndex * STEP_SIZE_IN_BYTES

        Case WPC_OPERATION_FILL
        
            PutMem4 VarPtr(begin_of_pointers) + 12, _
               ByVal StepPointer + HEADER_SIZE_IN_BYTES + StepIndex * STEP_SIZE_IN_BYTES

            RecordFill.Bits = FType

            ' Специальные параметры шага

            Select Case CBool(Step.Item("Pause"))
                Case False: RecordFill.Bits = RecordFill.Bits And &HFFEF
                Case True: RecordFill.Bits = RecordFill.Bits Or &H10
            End Select

            Select Case CBool(Step.Item("ColdWaterGate"))
                Case False: RecordFill.Bits = RecordFill.Bits And &HFFDF
                Case True: RecordFill.Bits = RecordFill.Bits Or &H20
            End Select

            Select Case CBool(Step.Item("HotWaterGate"))
                Case False: RecordFill.Bits = RecordFill.Bits And &HFFBF
                Case True: RecordFill.Bits = RecordFill.Bits Or &H40
            End Select

            Select Case CBool(Step.Item("RecycledWaterGate"))
                Case False: RecordFill.Bits = RecordFill.Bits And &HFF7F
                Case True: RecordFill.Bits = RecordFill.Bits Or &H80
            End Select

            Select Case CBool(Step.Item("Rotation"))
                Case False: RecordFill.Bits = RecordFill.Bits And &HFEFF
                Case True: RecordFill.Bits = RecordFill.Bits Or &H100
            End Select

            RecordFill.Level = CByte(Step.Item("Level"))
            RecordFill.RotationTime = CByte(Step.Item("RotationTime"))
            RecordFill.PauseTime = CByte(Step.Item("PauseTime"))
            RecordFill.DrumSpeed = CByte(Step.Item("DrumSpeed"))

        Case WPC_OPERATION_DTRG
        
            PutMem4 VarPtr(begin_of_pointers) + 16, _
               ByVal StepPointer + HEADER_SIZE_IN_BYTES + StepIndex * STEP_SIZE_IN_BYTES

            RecordDetergent.Bits = FType

            ' Специальные параметры шага

            Select Case CBool(Step.Item("Pause"))
                Case False: RecordDetergent.Bits = RecordDetergent.Bits And &HFFEF
                Case True: RecordDetergent.Bits = RecordDetergent.Bits Or &H10
            End Select

            Select Case CBool(Step.Item("Rotation"))
                Case False: RecordDetergent.Bits = RecordDetergent.Bits And &HFFDF
                Case True: RecordDetergent.Bits = RecordDetergent.Bits Or &H20
            End Select

            RecordDetergent.Detergent_1_Time = CByte(Step.Item("Detergent_1_Time"))
            RecordDetergent.Detergent_2_Time = CByte(Step.Item("Detergent_2_Time"))
            RecordDetergent.Detergent_3_Time = CByte(Step.Item("Detergent_3_Time"))
            RecordDetergent.Detergent_4_Time = CByte(Step.Item("Detergent_4_Time"))
            RecordDetergent.Detergent_5_Time = CByte(Step.Item("Detergent_5_Time"))
            RecordDetergent.Detergent_6_Time = CByte(Step.Item("Detergent_6_Time"))
            RecordDetergent.Detergent_7_Time = CByte(Step.Item("Detergent_7_Time"))
            RecordDetergent.Detergent_8_Time = CByte(Step.Item("Detergent_8_Time"))
            RecordDetergent.Detergent_9_Time = CByte(Step.Item("Detergent_9_Time"))

            RecordDetergent.RotationTime = CByte(Step.Item("RotationTime"))
            RecordDetergent.PauseTime = CByte(Step.Item("PauseTime"))
            RecordDetergent.DrumSpeed = CByte(Step.Item("DrumSpeed"))

        Case WPC_OPERATION_HEAT
        
            PutMem4 VarPtr(begin_of_pointers) + 20, _
               ByVal StepPointer + HEADER_SIZE_IN_BYTES + StepIndex * STEP_SIZE_IN_BYTES

            RecordHeat.Bits = FType

            ' Специальные параметры шага

            Select Case CBool(Step.Item("Pause"))
                Case False: RecordHeat.Bits = RecordHeat.Bits And &HFFEF
                Case True: RecordHeat.Bits = RecordHeat.Bits Or &H10
            End Select

            Select Case CBool(Step.Item("Rotation"))
                Case False: RecordHeat.Bits = RecordHeat.Bits And &HFFDF
                Case True: RecordHeat.Bits = RecordHeat.Bits Or &H20
            End Select

            RecordHeat.Temperature = CByte(Step.Item("Temperature"))
            RecordHeat.RotationTime = CByte(Step.Item("RotationTime"))
            RecordHeat.PauseTime = CByte(Step.Item("PauseTime"))
            RecordHeat.DrumSpeed = CByte(Step.Item("DrumSpeed"))

        Case WPC_OPERATION_WASH, WPC_OPERATION_RINS, WPC_OPERATION_JOLT, WPC_OPERATION_PAUS
        
            PutMem4 VarPtr(begin_of_pointers) + 24, _
               ByVal StepPointer + HEADER_SIZE_IN_BYTES + StepIndex * STEP_SIZE_IN_BYTES

            RecordWash.Bits = FType

            ' Специальные параметры шага

            Select Case CBool(Step.Item("Pause"))
                Case False: RecordWash.Bits = RecordWash.Bits And &HFFEF
                Case True: RecordWash.Bits = RecordWash.Bits Or &H10
            End Select

            Select Case CBool(Step.Item("Rotation"))
                Case False: RecordWash.Bits = RecordWash.Bits And &HFFDF
                Case True: RecordWash.Bits = RecordWash.Bits Or &H20
            End Select

            RecordWash.Time = CByte(Step.Item("Time"))
            RecordWash.RotationTime = CByte(Step.Item("RotationTime"))
            RecordWash.PauseTime = CByte(Step.Item("PauseTime"))
            RecordWash.DrumSpeed = CByte(Step.Item("DrumSpeed"))

        Case WPC_OPERATION_DRAIN
        
            PutMem4 VarPtr(begin_of_pointers) + 28, _
               ByVal StepPointer + HEADER_SIZE_IN_BYTES + StepIndex * STEP_SIZE_IN_BYTES

            RecordDrain.Bits = FType

            ' Специальные параметры шага

            Select Case CBool(Step.Item("Pause"))
                Case False: RecordDrain.Bits = RecordDrain.Bits And &HFFEF
                Case True: RecordDrain.Bits = RecordDrain.Bits Or &H10
            End Select

            Select Case CBool(Step.Item("DrainGate1"))
                Case False: RecordDrain.Bits = RecordDrain.Bits And &HFFDF
                Case True: RecordDrain.Bits = RecordDrain.Bits Or &H20
            End Select

            Select Case CBool(Step.Item("DrainGate2"))
                Case False: RecordDrain.Bits = RecordDrain.Bits And &HFFBF
                Case True: RecordDrain.Bits = RecordDrain.Bits Or &H40
            End Select

            Select Case CBool(Step.Item("Rotation"))
                Case False: RecordDrain.Bits = RecordDrain.Bits And &HFF7F
                Case True: RecordDrain.Bits = RecordDrain.Bits Or &H80
            End Select

            RecordDrain.Level = CByte(Step.Item("Level"))
            RecordDrain.RotationTime = CByte(Step.Item("RotationTime"))
            RecordDrain.PauseTime = CByte(Step.Item("PauseTime"))
            RecordDrain.DrumSpeed1 = CByte(Step.Item("DrumSpeed1"))

        Case WPC_OPERATION_SPIN
        
            PutMem4 VarPtr(begin_of_pointers) + 32, _
               ByVal StepPointer + HEADER_SIZE_IN_BYTES + StepIndex * STEP_SIZE_IN_BYTES

            RecordSpin.Bits = FType

            ' Специальные параметры шага

            Select Case CBool(Step.Item("Pause"))
                Case False: RecordSpin.Bits = RecordSpin.Bits And &HFFEF
                Case True: RecordSpin.Bits = RecordSpin.Bits Or &H10
            End Select

            Select Case CBool(Step.Item("DrainGate1"))
                Case False: RecordSpin.Bits = RecordSpin.Bits And &HFFDF
                Case True: RecordSpin.Bits = RecordSpin.Bits Or &H20
            End Select

            Select Case CBool(Step.Item("DrainGate2"))
                Case False: RecordSpin.Bits = RecordSpin.Bits And &HFFBF
                Case True: RecordSpin.Bits = RecordSpin.Bits Or &H40
            End Select

            RecordSpin.DrumSpeed = CInt(Step.Item("DrumSpeed"))
            RecordSpin.Time = CByte(Step.Item("Time"))

        Case WPC_OPERATION_COOL
        
            PutMem4 VarPtr(begin_of_pointers) + 36, _
               ByVal StepPointer + HEADER_SIZE_IN_BYTES + StepIndex * STEP_SIZE_IN_BYTES

            RecordCool.Bits = FType

            ' Специальные параметры шага

            Select Case CBool(Step.Item("Pause"))
                Case False: RecordCool.Bits = RecordCool.Bits And &HFFEF
                Case True: RecordCool.Bits = RecordCool.Bits Or &H10
            End Select

            Select Case CBool(Step.Item("Fast"))
                Case False: RecordCool.Bits = RecordCool.Bits And &HFFDF
                Case True: RecordCool.Bits = RecordCool.Bits Or &H20
            End Select

            Select Case CBool(Step.Item("Rotation"))
                Case False: RecordCool.Bits = RecordCool.Bits And &HFFBF
                Case True: RecordCool.Bits = RecordCool.Bits Or &H40
            End Select

            RecordCool.Temperature = CByte(Step.Item("Temperature"))
            RecordCool.ColdWaterTime = CByte(Step.Item("ColdWaterTime"))
            RecordCool.RotationTime = CByte(Step.Item("RotationTime"))
            RecordCool.PauseTime = CByte(Step.Item("PauseTime"))
            RecordCool.DrumSpeed = CByte(Step.Item("DrumSpeed"))

    End Select
    
    ' Пересчитываем CRC поле записи программы
    b = CalculateCRC8(Program * PROGRAM_SIZE_IN_BYTES + 1, PROGRAM_SIZE_IN_BYTES - 1)
    SetByte Program * PROGRAM_SIZE_IN_BYTES, b
       
    func_ImportStepFromClipboard = False
    
    Exit Function
    
errHandler:
    
    func_ImportStepFromClipboard = True
    
End Function
    
Public Function ImportFromClipboard(frm As FormMain, Value As String) As Boolean
    
    On Local Error GoTo errHandler

    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP
    Dim RecordFill As TYPE_WPC_FILL
    Dim RecordDetergent As TYPE_WPC_DETERGENT
    Dim RecordHeat As TYPE_WPC_HEAT
    Dim RecordWash As TYPE_WPC_WASH
    Dim RecordDrain As TYPE_WPC_DRAIN
    Dim RecordSpin As TYPE_WPC_SPIN
    Dim RecordCool As TYPE_WPC_COOL

    Dim J As Integer
    Dim Program As Object, Step As Object
    Dim p As Object
    
    ' Производим разбор данных из файла
    Set p = JSON.parse(Value)
    
    ' Если объект пустой или не содержит элементов, то выходим
    If p Is Nothing Or p.Count = 0 Then
        
        Set p = Nothing

        ImportFromClipboard = False
        
        Exit Function
        
    End If
    
    For J = 1 To p.Count
    
        ' Если уникальный идентификатор формата элемента не совпадает
        ' с поддерживаемыми программой, переходим к следующему
        If ObjectGUID <> p.Item(J).Item(CLIPBOARD_OBJECT_GUID_STRING) Then
                    
        Else
        
            Select Case CInt(p.Item(J).Item(CLIPBOARD_OBJECT_TYPE_STRING))
            
                Case CLIPBOARD_OBJECT_TYPE_STEP:
            
                    Set Step = p.Item(J)
                    
                    ' Вызываем вспомогательную функцию, которая позволяет работать
                    ' с указателями напрямую
                    func_ImportStepFromClipboard frm, ProgramIndex, Step, 0&, RecordTitle, RecordStep, RecordFill, RecordDetergent, _
                       RecordHeat, RecordWash, RecordDrain, RecordSpin, RecordCool
                    
                    Set Step = Nothing
                    
                
                Case CLIPBOARD_OBJECT_TYPE_PROGRAM:
                
                    Set Program = p.Item(J)
            
                    ' Вызываем вспомогательную функцию, которая позволяет работать
                    ' с указателями напрямую
                    func_ImportProgramFromClipboard frm, ProgramIndex, Program, 0&, RecordTitle
                    
                    Set Program = Nothing
            
            
            End Select
            
        End If
        
    Next
                
    Set p = Nothing

    ImportFromClipboard = False
    
    Exit Function
    
errHandler:
    
    Set p = Nothing
    
    ImportFromClipboard = True
    
End Function

Public Sub SaveToFile(FileName As String)

    Dim handle As Integer ' Идентификатор файла

    ' Внутреннее имя
    mvarFileName = FileName

    ' Получаем свободный идентификатор
    handle = FreeFile

    ' Получаем доступ к файлу
    Open FileName For Binary Access Write As #handle

    ' Сохраняем образ в файле
    Put #handle, , mvarData

    ' Завершаем работу с файлом
    Close #handle
    
End Sub

Public Sub SetStepFunctionType(ByVal FunctionType As Integer)
    Dim ByteValue As Byte
    Dim Offset As Long

    Offset = ProgramIndex * PROGRAM_SIZE_IN_BYTES + _
       HEADER_SIZE_IN_BYTES + _
       StepIndex * STEP_SIZE_IN_BYTES

    ' Считываем байт
    ByteValue = GetByte(Offset)

    ' Обнулям младшие 4 бита
    ByteValue = ByteValue And &HF0

    ' Добавлям тип функции
    ByteValue = ByteValue Or (CByte(FunctionType) And &HF)

    ' Сохраняем изменения
    SetByte Offset, ByteValue
    
End Sub

' Для памяти: как сделать на бейсике циклический сдвиг влево на 1 разряд
' Dim crc As Integer
' crc = 2 * crc
' crc = ((crc And &H100) / &H100) Or (crc And &HFE)
Public Function CalculateCRC8(ByVal Offset As Long, ByVal Size As Long) As Byte
    
    Dim b As Byte
    Dim CRC As Integer
    Dim I As Long

    CRC = &HFF

    Do While Size > 0
    
        Size = Size - 1

        CopyMemory b, ByVal DataPointer + Offset, 1
        
        Offset = Offset + 1

        CRC = CRC Xor b

        For I = 0 To 7

            If CRC And &H80 Then
            
                CRC = (2 * CRC) And &HFF
                CRC = CRC Xor &H31
                
            Else
            
                CRC = (2 * CRC) And &HFF
                
            End If
            
        Next
        
    Loop

    CalculateCRC8 = CByte(CRC)
    
End Function

Private Sub func_SetDefaultProgramHeader(N As Integer, ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE)

    Dim StepPointer As Long

    StepPointer = Manager.DataPointer + N * PROGRAM_SIZE_IN_BYTES
    PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer

    Select Case EndSound.DefaultValue
        Case False: RecordTitle.LowBits = RecordTitle.LowBits And &HFFFE
        Case True: RecordTitle.LowBits = RecordTitle.LowBits Or &H1
    End Select

    Select Case DoorUnlock.DefaultValue
        Case False: RecordTitle.LowBits = RecordTitle.LowBits And &HFFFD
        Case True: RecordTitle.LowBits = RecordTitle.LowBits Or &H2
    End Select

End Sub

Public Sub SetDefaultProgramHeader(N As Integer)

    Dim RecordTitle As TYPE_WPC_TITLE

    func_SetDefaultProgramHeader N, 0&, RecordTitle

End Sub

Private Function func_StepsSelectionToString(ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordStep As TYPE_WPC_STEP, _
       ByRef RecordFill As TYPE_WPC_FILL, _
       ByRef RecordDetergent As TYPE_WPC_DETERGENT, _
       ByRef RecordHeat As TYPE_WPC_HEAT, _
       ByRef RecordWash As TYPE_WPC_WASH, _
       ByRef RecordDrain As TYPE_WPC_DRAIN, _
       ByRef RecordSpin As TYPE_WPC_SPIN, _
       ByRef RecordCool As TYPE_WPC_COOL, _
       SelStepsCount As Long) As String

    Dim FNum As Integer
    Dim I As Long, StepPointer As Long
    Dim Result As String, sInputJson As String
    Dim StepsArray As Object, StepElem As Object
    
    Result = ""
    
    sInputJson = "[]"
    
    Set StepsArray = JSON.parse(sInputJson)
    
    I = SelStepsCount
    
    ' Инициализируем указатель RecordTitle адресом начала текущей программы
    StepPointer = DataPointer + ProgramIndex * PROGRAM_SIZE_IN_BYTES
    PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer
        
    Do While I > 0
    
        PutMem4 VarPtr(begin_of_pointers) + 8, _
           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (StepIndex + I - 1) * STEP_SIZE_IN_BYTES

        FNum = (RecordStep.Bits And &HF)
        
        ' Формируем объект для шага в зависимости от типа функции шага
        Set StepElem = JSON.parse("{""" & CLIPBOARD_OBJECT_GUID_STRING & _
                                    """: " & """" & ObjectGUID & """," _
                                    & """" & CLIPBOARD_OBJECT_TYPE_STRING & _
                                    """: " & CLIPBOARD_OBJECT_TYPE_STEP & _
                                    ", " & JSONStepsTemplates(FNum) & "}")

        Select Case FNum
        
            Case WPC_OPERATION_IDLE
            
                PutMem4 VarPtr(begin_of_pointers) + 8, _
                   ByVal StepPointer + HEADER_SIZE_IN_BYTES + (Manager.StepIndex + I - 1) * STEP_SIZE_IN_BYTES

            Case WPC_OPERATION_FILL
            
                PutMem4 VarPtr(begin_of_pointers) + 12, _
                   ByVal StepPointer + HEADER_SIZE_IN_BYTES + (Manager.StepIndex + I - 1) * STEP_SIZE_IN_BYTES

                StepElem.Item("Pause") = CBool((RecordFill.Bits And &H10) / &H10)
                StepElem.Item("ColdWaterGate") = CBool((RecordFill.Bits And &H20) / &H20)
                StepElem.Item("HotWaterGate") = CBool((RecordFill.Bits And &H40) / &H40)
                StepElem.Item("RecycledWaterGate") = CBool((RecordFill.Bits And &H80) / &H80)
                StepElem.Item("Rotation") = CBool((RecordFill.Bits And &H100) / &H100)

                StepElem.Item("Level") = CByte(RecordFill.Level)
                StepElem.Item("RotationTime") = CByte(RecordFill.RotationTime)
                StepElem.Item("PauseTime") = CByte(RecordFill.PauseTime)
                StepElem.Item("DrumSpeed") = CByte(RecordFill.DrumSpeed)

            Case WPC_OPERATION_DTRG
            
                PutMem4 VarPtr(begin_of_pointers) + 16, _
                   ByVal StepPointer + HEADER_SIZE_IN_BYTES + (Manager.StepIndex + I - 1) * STEP_SIZE_IN_BYTES

                StepElem.Item("Pause") = CBool((RecordDetergent.Bits And &H10) / &H10)
                StepElem.Item("Rotation") = CBool((RecordDetergent.Bits And &H20) / &H20)

                StepElem.Item("Detergent_1_Time") = CByte(RecordDetergent.Detergent_1_Time)
                StepElem.Item("Detergent_2_Time") = CByte(RecordDetergent.Detergent_2_Time)
                StepElem.Item("Detergent_3_Time") = CByte(RecordDetergent.Detergent_3_Time)
                StepElem.Item("Detergent_4_Time") = CByte(RecordDetergent.Detergent_4_Time)
                StepElem.Item("Detergent_5_Time") = CByte(RecordDetergent.Detergent_5_Time)
                StepElem.Item("Detergent_6_Time") = CByte(RecordDetergent.Detergent_6_Time)
                StepElem.Item("Detergent_7_Time") = CByte(RecordDetergent.Detergent_7_Time)
                StepElem.Item("Detergent_8_Time") = CByte(RecordDetergent.Detergent_8_Time)
                StepElem.Item("Detergent_9_Time") = CByte(RecordDetergent.Detergent_9_Time)

                StepElem.Item("RotationTime") = CByte(RecordDetergent.RotationTime)
                StepElem.Item("PauseTime") = CByte(RecordDetergent.PauseTime)
                StepElem.Item("DrumSpeed") = CByte(RecordDetergent.DrumSpeed)

            Case WPC_OPERATION_HEAT
            
                PutMem4 VarPtr(begin_of_pointers) + 20, _
                   ByVal StepPointer + HEADER_SIZE_IN_BYTES + (Manager.StepIndex + I - 1) * STEP_SIZE_IN_BYTES

                StepElem.Item("Pause") = CBool((RecordHeat.Bits And &H10) / &H10)
                StepElem.Item("Rotation") = CBool((RecordHeat.Bits And &H20) / &H20)

                StepElem.Item("Temperature") = CByte(RecordHeat.Temperature)
                StepElem.Item("RotationTime") = CByte(RecordHeat.RotationTime)
                StepElem.Item("PauseTime") = CByte(RecordHeat.PauseTime)
                StepElem.Item("DrumSpeed") = CByte(RecordHeat.DrumSpeed)

            Case WPC_OPERATION_WASH, WPC_OPERATION_RINS, WPC_OPERATION_JOLT, WPC_OPERATION_PAUS
            
                PutMem4 VarPtr(begin_of_pointers) + 24, _
                   ByVal StepPointer + HEADER_SIZE_IN_BYTES + (Manager.StepIndex + I - 1) * STEP_SIZE_IN_BYTES

                StepElem.Item("Type") = CByte(RecordWash.Bits And &HF)
                StepElem.Item("Pause") = CBool((RecordWash.Bits And &H10) / &H10)
                StepElem.Item("Rotation") = CBool((RecordWash.Bits And &H20) / &H20)

                StepElem.Item("Time") = CByte(RecordWash.Time)
                StepElem.Item("RotationTime") = CByte(RecordWash.RotationTime)
                StepElem.Item("PauseTime") = CByte(RecordWash.PauseTime)
                StepElem.Item("DrumSpeed") = CByte(RecordWash.DrumSpeed)

            Case WPC_OPERATION_DRAIN
            
                PutMem4 VarPtr(begin_of_pointers) + 28, _
                   ByVal StepPointer + HEADER_SIZE_IN_BYTES + (Manager.StepIndex + I - 1) * STEP_SIZE_IN_BYTES

                StepElem.Item("Pause") = CBool((RecordDrain.Bits And &H10) / &H10)
                StepElem.Item("DrainGate1") = CBool((RecordDrain.Bits And &H20) / &H20)
                StepElem.Item("DrainGate2") = CBool((RecordDrain.Bits And &H40) / &H40)
                StepElem.Item("Rotation") = CBool((RecordDrain.Bits And &H80) / &H80)

                StepElem.Item("Level") = CByte(RecordDrain.Level)
                StepElem.Item("RotationTime") = CByte(RecordDrain.RotationTime)
                StepElem.Item("PauseTime") = CByte(RecordDrain.PauseTime)
                StepElem.Item("DrumSpeed1") = CByte(RecordDrain.DrumSpeed1)

            Case WPC_OPERATION_SPIN
            
                PutMem4 VarPtr(begin_of_pointers) + 32, _
                   ByVal StepPointer + HEADER_SIZE_IN_BYTES + (Manager.StepIndex + I - 1) * STEP_SIZE_IN_BYTES

                StepElem.Item("Pause") = CBool((RecordSpin.Bits And &H10) / &H10)
                StepElem.Item("DrainGate1") = CBool((RecordSpin.Bits And &H20) / &H20)
                StepElem.Item("DrainGate2") = CBool((RecordSpin.Bits And &H40) / &H40)

                StepElem.Item("DrumSpeed") = CInt(RecordSpin.DrumSpeed)
                StepElem.Item("Time") = CByte(RecordSpin.Time)

            Case WPC_OPERATION_COOL
            
                PutMem4 VarPtr(begin_of_pointers) + 36, _
                   ByVal StepPointer + HEADER_SIZE_IN_BYTES + (Manager.StepIndex + I - 1) * STEP_SIZE_IN_BYTES

                StepElem.Item("Pause") = CBool((RecordCool.Bits And &H10) / &H10)
                StepElem.Item("Fast") = CBool((RecordCool.Bits And &H20) / &H20)
                StepElem.Item("Rotation") = CBool((RecordCool.Bits And &H40) / &H40)

                StepElem.Item("Temperature") = CByte(RecordCool.Temperature)
                StepElem.Item("ColdWaterTime") = CByte(RecordCool.ColdWaterTime)
                StepElem.Item("RotationTime") = CByte(RecordCool.RotationTime)
                StepElem.Item("PauseTime") = CByte(RecordCool.PauseTime)
                StepElem.Item("DrumSpeed") = CByte(RecordCool.DrumSpeed)

        End Select

        StepsArray.Add StepElem
        
        Set StepElem = Nothing
        
        Dec I
        
    Loop
    
    ' Сохраняем данные в формате JSON
    Result = JSON.toString(StepsArray)
    
    Set StepsArray = Nothing
    
    func_StepsSelectionToString = Result
    
End Function

Public Function StepsSelectionToString(SelStepsCount As Long) As String

    Dim Result As String
    
    Result = ""
    
    ' Если ничего не выделено, то результат - пустая строка
    If SelStepsCount <= 0 Then
    
        StepsSelectionToString = ""
        
        Exit Function
    
    End If
    
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP
    Dim RecordFill As TYPE_WPC_FILL
    Dim RecordDetergent As TYPE_WPC_DETERGENT
    Dim RecordHeat As TYPE_WPC_HEAT
    Dim RecordWash As TYPE_WPC_WASH
    Dim RecordDrain As TYPE_WPC_DRAIN
    Dim RecordSpin As TYPE_WPC_SPIN
    Dim RecordCool As TYPE_WPC_COOL
    
    ' Вызываем вспомогательную функцию, которая позволяет работать
    ' с указателями напрямую
    StepsSelectionToString = func_StepsSelectionToString(0&, RecordTitle, RecordStep, RecordFill, RecordDetergent, _
       RecordHeat, RecordWash, RecordDrain, RecordSpin, RecordCool, SelStepsCount)

End Function

Private Function func_ProgsSelectionToString(ByVal begin_of_pointers As Long, _
       ByRef RecordTitle As TYPE_WPC_TITLE, _
       ByRef RecordStep As TYPE_WPC_STEP, _
       ByRef RecordFill As TYPE_WPC_FILL, _
       ByRef RecordDetergent As TYPE_WPC_DETERGENT, _
       ByRef RecordHeat As TYPE_WPC_HEAT, _
       ByRef RecordWash As TYPE_WPC_WASH, _
       ByRef RecordDrain As TYPE_WPC_DRAIN, _
       ByRef RecordSpin As TYPE_WPC_SPIN, _
       ByRef RecordCool As TYPE_WPC_COOL, _
       SelProgsCount As Long) As String

    Dim b As Byte
    Dim FNum As Integer
    Dim I As Long, J As Long, StepPointer As Long
    Dim Result As String, sInputJson As String, S As String
    Dim ProgsArray As Object, StepsArray As Object
    Dim ProgElem As Object, StepElem As Object
    
    Result = ""
    
    sInputJson = "[]"
    
    Set ProgsArray = JSON.parse(sInputJson)
    
    I = SelProgsCount
    
    ' Сохраняем данные в формате JSON
    Result = JSON.toString(ProgsArray)
    
    Do While I > 0
    
        ' Формируем объект для шага в зависимости от типа функции шага
        Set ProgElem = JSON.parse("{""" & CLIPBOARD_OBJECT_GUID_STRING & _
                                    """: " & """" & ObjectGUID & """," _
                                    & """" & CLIPBOARD_OBJECT_TYPE_STRING & _
                                    """: " & CLIPBOARD_OBJECT_TYPE_PROGRAM & _
                                    ", ""Title"":"""", ""EndSound"": true, ""DoorUnlock"": true, ""Steps"":[]}")
    
        ' Инициализируем указатель RecordTitle адресом начала текущей программы
        StepPointer = DataPointer + (ProgramIndex + I - 1) * PROGRAM_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer
    
        ' Вычисляем признак пустой программы
        b = CalculateCRC8((ProgramIndex + I - 1) * PROGRAM_SIZE_IN_BYTES, PROGRAM_SIZE_IN_BYTES)
        
        ' Дополнительная проверка для пустой программы
        If b = CRC8_FOR_DEFAULT_PROGRAM Then
        
            Dim k As Long, Sum As Long
            
            Sum = 0
            
            For k = 2 To 31
            
                Sum = Sum + GetByte((ProgramIndex + I - 1) * PROGRAM_SIZE_IN_BYTES + k)
            
            Next
            
            If Sum > 0 Then b = 0
            
        End If
        
        If b = CRC8_FOR_DEFAULT_PROGRAM Then
        
            ' Пропускаем пустые программы начиная с конца списка
            ' Проверка осуществляется по CRC от всей программы

            If ProgsArray.Count > 1 Then
            
                ProgsArray.Add ProgElem, before:=2
                
            End If
            
        Else
        
            S = ""

            For J = 1 To PROG_NAME_LENGTH - 1

                If RecordTitle.ProgName(J) <> 0 Then S = S & Chr$(RecordTitle.ProgName(J))
                
            Next

            ProgElem.Item("Title") = CStr(S)
            ProgElem.Item("EndSound") = CBool(RecordTitle.LowBits And &H1)
            ProgElem.Item("DoorUnlock") = CBool((RecordTitle.LowBits And &H2) / 2 ^ 1)
        
            ' Добавляем шаги к программе
            J = MAX_NUMBER_OF_STEPS
            
            Do While J > 0
            
                PutMem4 VarPtr(begin_of_pointers) + 8, _
                   ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                FNum = (RecordStep.Bits And &HF)
                
                ' Формируем объект для шага в зависимости от типа функции шага
                Set StepElem = JSON.parse("{" & JSONStepsTemplates(FNum) & "}")
                
                Select Case FNum
                
                    Case WPC_OPERATION_IDLE
                        
                        PutMem4 VarPtr(begin_of_pointers) + 8, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES


                    Case WPC_OPERATION_FILL
                        
                        PutMem4 VarPtr(begin_of_pointers) + 12, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        StepElem.Item("Pause") = CBool((RecordFill.Bits And &H10) / &H10)
                        StepElem.Item("ColdWaterGate") = CBool((RecordFill.Bits And &H20) / &H20)
                        StepElem.Item("HotWaterGate") = CBool((RecordFill.Bits And &H40) / &H40)
                        StepElem.Item("RecycledWaterGate") = CBool((RecordFill.Bits And &H80) / &H80)
                        StepElem.Item("Rotation") = CBool((RecordFill.Bits And &H100) / &H100)

                        StepElem.Item("Level") = CByte(RecordFill.Level)
                        StepElem.Item("RotationTime") = CByte(RecordFill.RotationTime)
                        StepElem.Item("PauseTime") = CByte(RecordFill.PauseTime)
                        StepElem.Item("DrumSpeed") = CByte(RecordFill.DrumSpeed)


                    Case WPC_OPERATION_DTRG
                        
                        PutMem4 VarPtr(begin_of_pointers) + 16, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        StepElem.Item("Pause") = CBool((RecordDetergent.Bits And &H10) / &H10)
                        StepElem.Item("Rotation") = CBool((RecordDetergent.Bits And &H20) / &H20)

                        StepElem.Item("Detergent_1_Time") = CByte(RecordDetergent.Detergent_1_Time)
                        StepElem.Item("Detergent_2_Time") = CByte(RecordDetergent.Detergent_2_Time)
                        StepElem.Item("Detergent_3_Time") = CByte(RecordDetergent.Detergent_3_Time)
                        StepElem.Item("Detergent_4_Time") = CByte(RecordDetergent.Detergent_4_Time)
                        StepElem.Item("Detergent_5_Time") = CByte(RecordDetergent.Detergent_5_Time)
                        StepElem.Item("Detergent_6_Time") = CByte(RecordDetergent.Detergent_6_Time)
                        StepElem.Item("Detergent_7_Time") = CByte(RecordDetergent.Detergent_7_Time)
                        StepElem.Item("Detergent_8_Time") = CByte(RecordDetergent.Detergent_8_Time)
                        StepElem.Item("Detergent_9_Time") = CByte(RecordDetergent.Detergent_9_Time)

                        StepElem.Item("RotationTime") = CByte(RecordDetergent.RotationTime)
                        StepElem.Item("PauseTime") = CByte(RecordDetergent.PauseTime)
                        StepElem.Item("DrumSpeed") = CByte(RecordDetergent.DrumSpeed)


                    Case WPC_OPERATION_HEAT
                        
                        PutMem4 VarPtr(begin_of_pointers) + 20, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        StepElem.Item("Pause") = CBool((RecordHeat.Bits And &H10) / &H10)
                        StepElem.Item("Rotation") = CBool((RecordHeat.Bits And &H20) / &H20)

                        StepElem.Item("Temperature") = CByte(RecordHeat.Temperature)
                        StepElem.Item("RotationTime") = CByte(RecordHeat.RotationTime)
                        StepElem.Item("PauseTime") = CByte(RecordHeat.PauseTime)
                        StepElem.Item("DrumSpeed") = CByte(RecordHeat.DrumSpeed)


                    Case WPC_OPERATION_WASH, WPC_OPERATION_RINS, WPC_OPERATION_JOLT, WPC_OPERATION_PAUS
                        
                        PutMem4 VarPtr(begin_of_pointers) + 24, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        StepElem.Item("Type") = CByte(RecordWash.Bits And &HF)
                        StepElem.Item("Pause") = CBool((RecordWash.Bits And &H10) / &H10)
                        StepElem.Item("Rotation") = CBool((RecordWash.Bits And &H20) / &H20)

                        StepElem.Item("Time") = CByte(RecordWash.Time)
                        StepElem.Item("RotationTime") = CByte(RecordWash.RotationTime)
                        StepElem.Item("PauseTime") = CByte(RecordWash.PauseTime)
                        StepElem.Item("DrumSpeed") = CByte(RecordWash.DrumSpeed)


                    Case WPC_OPERATION_DRAIN
                    
                        PutMem4 VarPtr(begin_of_pointers) + 28, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        StepElem.Item("Pause") = CBool((RecordDrain.Bits And &H10) / &H10)
                        StepElem.Item("DrainGate1") = CBool((RecordDrain.Bits And &H20) / &H20)
                        StepElem.Item("DrainGate2") = CBool((RecordDrain.Bits And &H40) / &H40)
                        StepElem.Item("Rotation") = CBool((RecordDrain.Bits And &H80) / &H80)

                        StepElem.Item("Level") = CByte(RecordDrain.Level)
                        StepElem.Item("RotationTime") = CByte(RecordDrain.RotationTime)
                        StepElem.Item("PauseTime") = CByte(RecordDrain.PauseTime)
                        StepElem.Item("DrumSpeed1") = CByte(RecordDrain.DrumSpeed1)
                        

                    Case WPC_OPERATION_SPIN
                    
                        PutMem4 VarPtr(begin_of_pointers) + 32, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        StepElem.Item("Pause") = CBool((RecordSpin.Bits And &H10) / &H10)
                        StepElem.Item("DrainGate1") = CBool((RecordSpin.Bits And &H20) / &H20)
                        StepElem.Item("DrainGate2") = CBool((RecordSpin.Bits And &H40) / &H40)

                        StepElem.Item("DrumSpeed") = CInt(RecordSpin.DrumSpeed)
                        StepElem.Item("Time") = CByte(RecordSpin.Time)
                        

                    Case WPC_OPERATION_COOL
                    
                        PutMem4 VarPtr(begin_of_pointers) + 36, _
                           ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES

                        StepElem.Item("Pause") = CBool((RecordCool.Bits And &H10) / &H10)
                        StepElem.Item("Fast") = CBool((RecordCool.Bits And &H20) / &H20)
                        StepElem.Item("Rotation") = CBool((RecordCool.Bits And &H40) / &H40)

                        StepElem.Item("Temperature") = CByte(RecordCool.Temperature)
                        StepElem.Item("ColdWaterTime") = CByte(RecordCool.ColdWaterTime)
                        StepElem.Item("RotationTime") = CByte(RecordCool.RotationTime)
                        StepElem.Item("PauseTime") = CByte(RecordCool.PauseTime)
                        StepElem.Item("DrumSpeed") = CByte(RecordCool.DrumSpeed)

                End Select

                ' Выводим информацию в сжатом виде, без пустых шагов

                If ProgElem.Item("Steps").Count > 0 Then

                    If StepElem Is Nothing Then
                    
                        ProgElem.Item("Steps").Add JSON.parse("{}")
                        
                    Else
                    
                        ProgElem.Item("Steps").Add StepElem
                        
                    End If
                    
                Else

                    If Not StepElem Is Nothing Then ProgElem.Item("Steps").Add StepElem
                    
                End If

                Set StepElem = Nothing

                J = J - 1
                
            Loop
                
            ' Выводим информацию

            If ProgsArray.Count = 0 Then
            
                ProgsArray.Add ProgElem
                
            Else
            
                ProgsArray.Add ProgElem
                
            End If
            
        End If
        
        Set ProgElem = Nothing
        
        Dec I
        
    Loop
    
    Result = JSON.toString(ProgsArray)
    
    Set ProgsArray = Nothing
    
    func_ProgsSelectionToString = Result
    
End Function
    
Public Function ProgsSelectionToString(SelProgsCount As Long) As String

    Dim Result As String
    
    Result = ""
    
    ' Если ничего не выделено, то результат - пустая строка
    If SelProgsCount <= 0 Then
    
        ProgsSelectionToString = ""
        
        Exit Function
    
    End If
    
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP
    Dim RecordFill As TYPE_WPC_FILL
    Dim RecordDetergent As TYPE_WPC_DETERGENT
    Dim RecordHeat As TYPE_WPC_HEAT
    Dim RecordWash As TYPE_WPC_WASH
    Dim RecordDrain As TYPE_WPC_DRAIN
    Dim RecordSpin As TYPE_WPC_SPIN
    Dim RecordCool As TYPE_WPC_COOL
    
    ' Вызываем вспомогательную функцию, которая позволяет работать
    ' с указателями напрямую
    ProgsSelectionToString = func_ProgsSelectionToString(0&, RecordTitle, RecordStep, RecordFill, RecordDetergent, _
       RecordHeat, RecordWash, RecordDrain, RecordSpin, RecordCool, SelProgsCount)

End Function

