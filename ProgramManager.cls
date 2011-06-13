VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TProgramManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' [29.12.2010] Слава , привет! _
В приложении файл Си со структурами шагов и программ. _

' Дополнительно сообщаю: _
 _
1) На карте памяти расположен файл "data.bin" _
2) Размер файла 32кБ _
3) Все параметры программ расположены последовательно друг за другом... _
сначала идут параметры первой УП, затем второй.. затем третьей.... и т.д.. _
4) Количество программ на карте памяти - 25 _
5) Количество шагов в программе - 80 _
6) Размер шага - 16 байт _
7) Дополнительно каждая программа имеет заголовок 16 байт _
8) Размер одной программы = 80 * 16 + 16  = 1296 байт _

' В программной утилите должны быть предусмотрены следующие функции: _
 _
1) Прочитать данные из файла data.bin (путь до файла указывает пользователь) _
2) Сохранить данные в файл data.bin (путь до файла указывает пользователь) _
3) Скопировать параметры одной УП в другую (пользователь должен указать, _
какая УП - исходник, а какая - приемник) _
4) Очистить параметры УП (пользователь должен указать номер УП) _
5) Вставить шаг со сдвигом всех шагов вправо _
6) Удалить шаг со сдвигом всех шагов влево _
7) Назначить шагу функцию _
8) Задать параметры шага с учетом назначенной функции _
9) Визуально отобразить, какие нагрузки будут включены на шаге, исходя _
из функции шага (и в зависимости от настроенных параметров) _
10)  ...

' Список того, что нужно доделать: _
- [+] в меню "Файл" добавить подменю с ранее открытыми файлами; _
- [+] при создании нового проекта не спрашивается о закрытии старого и сохранении _
  изменений, если они были; _
- добавить форму редактирования свойств интерфейса программы: размеров ячеек, _
  шрифта; _
- добавить возможность вывода на печать (в pdf) в компактном виде всего файла; _
- при двойном клике на ячейке в таблице шагов сделать возможным локальное _
  редактирование параметра шага; _
- придумать формат сохранения структуры программ в виде JSON файла для импорта и _
  экспорта данных в читабельном виде; _
- сделать возможность отката изменений на некоторую глубину, задаваемую в _
  настройках; _
-

Option Explicit

Private mvarFileLoaded As Boolean
Private mvarImageSize As Long
Private mvarProgramIndex As Long
Private mvarStepIndex As Long
Private mvarFileName As String
Private mvarData() As Byte

Public ProgramsCount As Long

Public Sub CopyProgram(ByVal FromIndex As Integer, ByVal ToIndex As Integer)
    ' Копируем программу
    CopyMemory ByVal DataPointer + ToIndex * PROGRAM_SIZE_IN_BYTES, _
        ByVal DataPointer + FromIndex * PROGRAM_SIZE_IN_BYTES, _
        PROGRAM_SIZE_IN_BYTES
End Sub

Public Sub DeleteStep()
    ' Копируем все шаги вверх, замещая текущий
    CopyMemory ByVal DataPointer + _
        ProgramIndex * PROGRAM_SIZE_IN_BYTES + HEADER_SIZE_IN_BYTES + _
        StepIndex * STEP_SIZE_IN_BYTES, _
        ByVal DataPointer + _
        ProgramIndex * PROGRAM_SIZE_IN_BYTES + HEADER_SIZE_IN_BYTES + _
        (StepIndex + 1) * STEP_SIZE_IN_BYTES, _
        (MAX_NUMBER_OF_STEPS - (StepIndex + 1)) * STEP_SIZE_IN_BYTES
    
    ' Обнуляем последний шаг
    ZeroMemory ByVal DataPointer + _
        ProgramIndex * PROGRAM_SIZE_IN_BYTES + HEADER_SIZE_IN_BYTES + _
        (MAX_NUMBER_OF_STEPS - 1) * STEP_SIZE_IN_BYTES, STEP_SIZE_IN_BYTES
End Sub

' Вставляем пустой шаг в текущую позицию со сдвигом всех шагов вправо
Public Sub InsertStep()
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
End Sub

Public Sub CloseFile()
    If FileLoaded Then
        FileName = ""
        FileLoaded = False
    End If
End Sub

Public Property Let FileLoaded(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileLoaded = 5
    mvarFileLoaded = vData
End Property


Public Property Get FileLoaded() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileLoaded
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
'Syntax: Debug.Print X.DataPointer
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
    
    ' Создаём пустой образ управляющей программы
    ReDim mvarData(0 To mvarImageSize - 1)

    ZeroMemory ByVal DataPointer, mvarImageSize
    
    ' Признак загрузки файла
    FileLoaded = True
    
    ' Инициализируем счётчики
    mvarProgramIndex = 0
    mvarStepIndex = 0
    
    ' Устанавливаем количество программ
    ProgramsCount = MAX_NUMBER_OF_PROGRAMS
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
    Dim ProgPointer As Long
    
    For Cnt = 1 To ProgramsCount
        ClearProgramN (Cnt)
    Next Cnt
End Sub

Public Property Let StepIndex(ByVal vData As Long)
'Syntax: X.StepIndex = 5
    mvarStepIndex = vData
End Property


Public Property Get StepIndex() As Long
'Syntax: Debug.Print X.StepIndex
    StepIndex = mvarStepIndex
End Property

Public Property Let ProgramIndex(ByVal vData As Long)
'Syntax: X.ProgramIndex = 5
    mvarProgramIndex = vData
End Property

Public Property Get ProgramIndex() As Long
'Syntax: Debug.Print X.ProgramIndex
    ProgramIndex = mvarProgramIndex
End Property

Public Property Get ImageSize() As Long
'Syntax: Debug.Print X.ImageSize
    ImageSize = mvarImageSize
End Property

Public Property Let FileName(ByVal vData As String)
'Syntax: X.FileName = 5
    mvarFileName = vData
End Property

Public Property Get FileName() As String
'Syntax: Debug.Print X.FileName
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

Public Sub Create()
    mvarImageSize = 0
    
    mvarProgramIndex = 0
    mvarStepIndex = 0
    
    ProgramsCount = 0
    
    FileLoaded = False
End Sub

Public Sub LoadFromFile(FileName As String)
    Dim fileid% ' Идентификатор файла
    Dim ProgsCount%, ii%
    
    ' Проверяем существование файла
    ' ...
    
    mvarFileName = FileName
    
    ' Получаем свободный идентификатор
    fileid% = FreeFile
    
    ' Получаем доступ к файлу
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

Private Function func_ExportToJSON(FileName As String, _
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
    Dim i As Integer
    Dim J As Integer
    Dim FNum As Integer
    Dim StepPointer As Long
    Dim s As String
    Dim sInputJson As String
    Dim ProgsArray As Object
    Dim ProgElem As Object
    Dim StepElem As Object
    
    JSONStepsTemplates(WPC_OPERATION_IDLE) = ""
    JSONStepsTemplates(WPC_OPERATION_FILL) = "{" _
        & """Type"": 1," _
        & """Pause"": false," _
        & """ColdWaterGate"": false," _
        & """HotWaterGate"": false," _
        & """RecycledWaterGate"": false," _
        & """Rotation"": true," _
        & """Level"": 15," _
        & """RotationTime"": 6," _
        & """PauseTime"": 12," _
        & """DrumSpeed"": 50}"
    
    JSONStepsTemplates(WPC_OPERATION_DTRG) = "{" _
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
        & """DrumSpeed"": 50}"
'    JSONStepsTemplates(WPC_OPERATION_DTRG) = "{" _
'        & """Type"": 2}" _

    JSONStepsTemplates(WPC_OPERATION_HEAT) = "{" _
        & """Type"": 3," _
        & """Pause"": false," _
        & """Rotation"": true," _
        & """Temperature"": 40," _
        & """RotationTime"": 6," _
        & """PauseTime"": 12," _
        & """DrumSpeed"": 50}"
    
    JSONStepsTemplates(WPC_OPERATION_WASH) = "{" _
        & """Type"": 4," _
        & """Pause"": false," _
        & """Rotation"": true," _
        & """Time"": 30," _
        & """RotationTime"": 6," _
        & """PauseTime"": 12," _
        & """DrumSpeed"": 50}"
    
    JSONStepsTemplates(WPC_OPERATION_RINS) = JSONStepsTemplates(WPC_OPERATION_WASH)
    JSONStepsTemplates(WPC_OPERATION_JOLT) = JSONStepsTemplates(WPC_OPERATION_WASH)
    JSONStepsTemplates(WPC_OPERATION_PAUS) = JSONStepsTemplates(WPC_OPERATION_WASH)
    
    JSONStepsTemplates(WPC_OPERATION_DRAIN) = "{" _
        & """Type"": 8," _
        & """Pause"": false," _
        & """DrainGate1"": true," _
        & """DrainGate2"": false," _
        & """Rotation"": true," _
        & """Level"": 0," _
        & """RotationTime"": 6," _
        & """PauseTime"": 12," _
        & """DrumSpeed1"": 50}"
    
    JSONStepsTemplates(WPC_OPERATION_SPIN) = "{" _
        & """Type"": 9," _
        & """Pause"": false," _
        & """DrainGate1"": true," _
        & """DrainGate2"": false," _
        & """DrumSpeed"": 800," _
        & """Time"": 2}"
    
    JSONStepsTemplates(WPC_OPERATION_COOL) = "{" _
        & """Type"": 10," _
        & """Pause"": false," _
        & """Fast"": false," _
        & """Rotation"": true," _
        & """Temperature"": 40," _
        & """ColdWaterTime"": 10," _
        & """RotationTime"": 6," _
        & """PauseTime"": 12," _
        & """DrumSpeed"": 50}"
    
    sInputJson = "[{""ProgID"":""""}]"
    
    Set ProgsArray = JSON.parse(sInputJson)
    
    ProgsArray.Item(1).Item("ProgID") = ProgramGUID
    
    i = ProgramsCount
    
    Do While i > 0
        sInputJson = "{""Title"":"""", ""EndSound"": true, ""DoorUnlock"": true, ""Steps"":[]}"
        Set ProgElem = JSON.parse(sInputJson)
        
        StepPointer = DataPointer + (i - 1) * PROGRAM_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer
        
        b = CalculateCRC8((i - 1) * PROGRAM_SIZE_IN_BYTES, PROGRAM_SIZE_IN_BYTES)
        
        If b = 29 Then
            ' Пропускаем пустые программы начиная с конца списка
            ' Проверка осуществляется по CRC от всей программы
        Else
            s = ""
            
            For J = 1 To PROG_NAME_LENGTH - 1
                If RecordTitle.ProgName(J) <> 0 Then s = s & Chr(RecordTitle.ProgName(J))
            Next J
            
            ProgElem.Item("Title") = CStr(s)
            ProgElem.Item("EndSound") = CBool(RecordTitle.LowBits And &H1)
            ProgElem.Item("DoorUnlock") = CBool((RecordTitle.LowBits And &H2) / 2 ^ 1)
                       
            ' Добавляем шаги к программе
            J = MAX_NUMBER_OF_STEPS
            
            Do While J > 0
                PutMem4 VarPtr(begin_of_pointers) + 8, _
                    ByVal StepPointer + HEADER_SIZE_IN_BYTES + (J - 1) * STEP_SIZE_IN_BYTES
                
                FNum = (RecordStep.Bits And &HF)
                Set StepElem = JSON.parse(JSONStepsTemplates(FNum))
                
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
                    ProgElem.Item("Steps").Add StepElem, before:=1
                Else
                    If Not StepElem Is Nothing Then ProgElem.Item("Steps").Add StepElem
                End If
                
                Set StepElem = Nothing
                 
                J = J - 1
            Loop
            
            ' Выводим информацию в сжатом виде, без пустых программ
            If ProgsArray.Count = 1 Then
                ProgsArray.Add ProgElem
            Else
                ProgsArray.Add ProgElem, before:=2
            End If
        End If
        
        Set ProgElem = Nothing
        
        i = i - 1
    Loop
    
    func_ExportToJSON = True
    
    SaveToJSONFile FileName, JSON.toString(ProgsArray)
    
    Set ProgsArray = Nothing
    Exit Function
    
errHandler:
    Set ProgsArray = Nothing
    func_ExportToJSON = False
End Function

Public Function ExportToJSON(FileName As String) As Boolean
    Dim RecordTitle As TYPE_WPC_TITLE
    Dim RecordStep As TYPE_WPC_STEP
    Dim RecordFill As TYPE_WPC_FILL
    Dim RecordDetergent As TYPE_WPC_DETERGENT
    Dim RecordHeat As TYPE_WPC_HEAT
    Dim RecordWash As TYPE_WPC_WASH
    Dim RecordDrain As TYPE_WPC_DRAIN
    Dim RecordSpin As TYPE_WPC_SPIN
    Dim RecordCool As TYPE_WPC_COOL
    
    ExportToJSON = func_ExportToJSON(FileName, 0&, RecordTitle, RecordStep, RecordFill, RecordDetergent, _
        RecordHeat, RecordWash, RecordDrain, RecordSpin, RecordCool)
End Function

Private Function func_ImportFromJSON(FileName As String, _
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
    Dim i As Integer
    Dim J As Integer
    Dim StepPointer As Long
    Dim sInputJson As String
    Dim Title As String
    
    Dim p As Object
    
    ' Проверяем существование файла
    ' ...
    
    sInputJson = FromUTF8(LoadFromJSONFile(FileName))
    
    Set p = JSON.parse(sInputJson)
    
    If ProgramGUID <> p.Item(1).Item("ProgID") Then
        Set p = Nothing
        FileLoaded = False
        func_ImportFromJSON = False
    End If
    
    CreateNewFile Left(FileName, InStrRev(FileName, ".")) + "bin"
    ClearAll
    
    For i = 1 To ProgramsCount
        ' Пересчитываем CRC поле записи программы
        
        b = CalculateCRC8((i - 1) * PROGRAM_SIZE_IN_BYTES + 1, PROGRAM_SIZE_IN_BYTES - 1)
        SetByte (i - 1) * PROGRAM_SIZE_IN_BYTES, b
    Next i
    
    ' Узнаём количество программ
    ProgsCount = p.Count - 1
        
    For i = 1 To ProgsCount
        StepPointer = DataPointer + (i - 1) * PROGRAM_SIZE_IN_BYTES
        PutMem4 VarPtr(begin_of_pointers) + 4, ByVal StepPointer
    
        ' Общие параметры программы
        Select Case CBool(p.Item(i + 1).Item("EndSound"))
            Case False: RecordTitle.LowBits = RecordTitle.LowBits And &HFFFE
            Case True: RecordTitle.LowBits = RecordTitle.LowBits Or &H1
        End Select
        
        Select Case CBool(p.Item(i + 1).Item("DoorUnlock"))
            Case False: RecordTitle.LowBits = RecordTitle.LowBits And &HFFFD
            Case True: RecordTitle.LowBits = RecordTitle.LowBits Or &H2
        End Select
        
        Title = p.Item(i + 1).Item("Title")
        
        For J = 1 To PROG_NAME_LENGTH - 1
            If J <= Len(Title) Then
                RecordTitle.ProgName(J) = Asc(Mid(Title, J, 1))
            Else
                RecordTitle.ProgName(J) = 0
            End If
        Next J
        
        RecordTitle.ProgName(PROG_NAME_LENGTH) = 0
        
        Dim FType As Integer
        Dim Step As Object
        
        ' В зависимости от типа шага заполняем структуры шагов
        For J = 1 To p.Item(i + 1).Item("Steps").Count
            Set Step = p.Item(i + 1).Item("Steps").Item(J)
            
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
        Next J
        
        ' Пересчитываем CRC поле записи программы
        b = CalculateCRC8((i - 1) * PROGRAM_SIZE_IN_BYTES + 1, PROGRAM_SIZE_IN_BYTES - 1)
        SetByte (i - 1) * PROGRAM_SIZE_IN_BYTES, b
    Next i
    
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
    
    ImportFromJSON = func_ImportFromJSON(FileName, 0&, RecordTitle, RecordStep, RecordFill, RecordDetergent, _
        RecordHeat, RecordWash, RecordDrain, RecordSpin, RecordCool)
End Function

Public Function LoadFromJSONFile(sFilePath As String) As String
   On Error Resume Next
   
   Dim handle As Integer
   
   If LenB(Dir$(sFilePath)) > 0 Then
      handle = FreeFile
      Open sFilePath For Binary As #handle
      LoadFromJSONFile = Space$(LOF(handle))
      Get #handle, , LoadFromJSONFile
      Close #handle
   End If
End Function

Public Sub SaveToJSONFile(ByVal FileName As String, ByVal NewValue As String)
    Dim handle As Integer
    
    handle = FreeFile
    Open FileName For Output As #handle
    Print #handle, ToUTF8(NewValue)
    Close #handle
End Sub

Public Sub SaveToFile(FileName As String)
    Dim fileid% ' Идентификатор файла
    
    ' Внутреннее имя
    mvarFileName = FileName
    
    ' Получаем свободный идентификатор
    fileid% = FreeFile
    
    ' Получаем доступ к файлу
    Open FileName For Binary Access Write As #fileid%
    
    ' Сохраняем образ в файле
    Put #fileid%, , mvarData
    
    ' Завершаем работу с файлом
    Close #fileid%
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
    Dim i As Long
    
    CRC = &HFF

    Do While Size > 0
        Size = Size - 1
        
        CopyMemory b, ByVal DataPointer + Offset, 1
        Offset = Offset + 1
        
        CRC = CRC Xor b
                
        For i = 0 To 7
            If CRC And &H80 Then
                CRC = (2 * CRC) And &HFF
                CRC = CRC Xor &H31
            Else
                CRC = (2 * CRC) And &HFF
            End If
        Next i
    Loop

    CalculateCRC8 = CByte(CRC)
End Function
