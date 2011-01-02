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
        Next I
    Loop

    CalculateCRC8 = CByte(CRC)
End Function
