Attribute VB_Name = "ModuleUserTypes"
'**
'@author Мезенцев В. Н.
'@revision Wed Jun 15 19:13:32 GMT 2011
'@rem Модуль содержит описание всех необходимых констант,
'которые используются в программе.
Option Explicit

' *****************************************
' *  КОНСТАНТЫ
' *  ~~~~~~~~~
' *****************************************

'**
'@rem Минимальное количество программ
Public Const Program_MIN = 0
'**
'@rem Максимальное количество программ
Public Const Program_MAX = 49

Public Const ProgramFlash_MIN = 0
Public Const ProgramFlash_MAX = 24

Public Const Program_Step_MIN = 0
Public Const Program_Step_MAX = 79

Public Const MAX_NUMBER_OF_PROGRAMS = ProgramFlash_MAX + 1
Public Const HEADER_SIZE_IN_BYTES = 16
Public Const STEP_SIZE_IN_BYTES = 16
Public Const MAX_NUMBER_OF_STEPS = Program_Step_MAX + 1

Public Const PROGRAM_SIZE_IN_BYTES = HEADER_SIZE_IN_BYTES + STEP_SIZE_IN_BYTES * MAX_NUMBER_OF_STEPS

Public Const IMAGE_SIZE = 32768

Public Const STRING_YES = "Да"
Public Const STRING_NO = "Нет"

Public Const PROG_NAME_LENGTH = 11

Public Const IDLE_ENDSOUND_FIELD = 1
Public Const IDLE_DOORUNLOCK_FIELD = 2
Public Const IDLE_PROGNAME_FIELD = 3
Public Const IDLE_STEP_FIELD = 4
Public Const IDLE_FUNCTION_FIELD = 5

Public Const IDLE_PARAMETERS_COUNT = 5

Public Const IDLE_PARAMETER_DESCR_UNKNOWN = "Неизвестный"
Public Const IDLE_PARAMETER_DESCR_ENDSOUND = "Звук в конце"
Public Const IDLE_PARAMETER_DESCR_DOORUNLOCK = "Разбл. люк"
Public Const IDLE_PARAMETER_DESCR_PROGNAME = "Название"
Public Const IDLE_PARAMETER_DESCR_STEP = "Шаг"
Public Const IDLE_PARAMETER_DESCR_FUNCTION = "Тип опер."

Public Type OSVERSIONINFO

    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    
End Type

' *****************************************
' *  ПОЛЬЗОВАТЕЛЬСКИЕ ТИПЫ
' *  ~~~~~~~~~~~~~~~~ ~~~~
' *****************************************

'**
'@rem Структура заголовка программы<br><pre>
'Public Type <b>TYPE_WPC_TITLE</b>
'  CRC As Byte
'  LowBits As Byte
'  HiBits As Byte
'  ProgName(1 To PROG_NAME_LENGTH) As Byte
'  Reserved(1 To 2) As Byte
'End Type</pre>

Public Type FILEINFO
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OriginalFileName As String
    ProductName As String
    ProductVersion As String
End Type

Public Enum VerisonReturnValue
    eOK = 1
    eNoVersion = 2
End Enum

Public TYPE_WPC_TITLE As TYPE_WPC_TITLE

' Структура заголовка программы
Public Type TYPE_WPC_TITLE
  CRC As Byte
  LowBits As Byte
  HiBits As Byte
  ProgName(1 To PROG_NAME_LENGTH) As Byte
  reserved(1 To 2) As Byte
End Type

' Общая структура шага
Public Type TYPE_WPC_STEP
  Bits As Integer
  
  reserved(1 To 14) As Byte
End Type

' Структура программы
Public Type TYPE_WPC_PROG
  Title As TYPE_WPC_TITLE
  Step(1 To MAX_NUMBER_OF_STEPS) As TYPE_WPC_STEP
End Type

' Структура шага НАЛИВ
Public Type TYPE_WPC_FILL
    Bits As Integer
    Level As Byte ' уровень моющего раствора
    RotationTime As Byte ' время вращения мотора
    PauseTime As Byte ' время паузы вращения мотора
    DrumSpeed As Byte ' скорость вращения барабана
    
    reserved(1 To 10) As Byte
End Type

' структура шага МОЮЩИЕ
Public Type TYPE_WPC_DETERGENT
    Bits As Integer
    Detergent_1_Time As Byte ' время подачи моющих 1
    Detergent_2_Time As Byte ' время подачи моющих 2
    Detergent_3_Time As Byte ' время подачи моющих 3
    Detergent_4_Time As Byte ' время подачи моющих 4
    Detergent_5_Time As Byte ' время подачи моющих 5
    Detergent_6_Time As Byte ' время подачи моющих 6
    Detergent_7_Time As Byte ' время подачи моющих 7
    Detergent_8_Time As Byte ' время подачи моющих 8
    Detergent_9_Time As Byte ' время подачи моющих 9
    RotationTime As Byte ' время вращения мотора
    PauseTime As Byte ' время паузы вращения мотора
    DrumSpeed As Byte ' скорость вращения барабана
    
    reserved(1 To 2) As Byte
End Type

' структура шага НАГРЕВ
Public Type TYPE_WPC_HEAT
    Bits As Integer
    Temperature As Byte ' температура моющего раствора
    RotationTime As Byte ' время вращения мотора
    PauseTime As Byte ' время паузы вращения мотора
    DrumSpeed As Byte ' скорость вращения барабана
    
    reserved(1 To 10) As Byte
End Type

' структура шага СТИРКИ (ПОЛОСКАНИЕ, РАССТРЯСКА)
Public Type TYPE_WPC_WASH
    Bits As Integer
    Time As Byte ' время стирки
    RotationTime As Byte ' время вращения мотора
    PauseTime As Byte ' время паузы вращения мотора
    DrumSpeed As Byte ' скорость вращения барабана
    
    reserved(1 To 10) As Byte
End Type

' структура шага СЛИВА
Public Type TYPE_WPC_DRAIN
    Bits As Integer
    Level As Byte ' Уровень моющего раствора после слива
    RotationTime As Byte ' время вращения мотора
    PauseTime As Byte ' время паузы вращения мотора
    DrumSpeed1 As Byte ' скорость вращения барабана при реверсе
'    DrumSpeed2 As Integer ' скорость вращения барабана при раскладке
'    Time2 As Byte ' время раскладки
    
    reserved(1 To 10) As Byte
End Type

' структура шага ОТЖИМА
Public Type TYPE_WPC_SPIN
    Bits As Integer
    DrumSpeed As Integer ' скорость вращения барабана при отжиме
    Time As Byte ' время отжима
    
    reserved(1 To 11) As Byte
End Type

' структура шага ОХЛАЖДЕНИЕ
Public Type TYPE_WPC_COOL
    Bits As Integer
    Temperature As Byte '
    ColdWaterTime As Byte ' время открытия клапана холодной воды
    RotationTime As Byte ' время вращения мотора
    PauseTime As Byte ' время паузы вращения мотора
    DrumSpeed As Byte ' скорость вращения барабана при реверсе
    
    reserved(1 To 9) As Byte
End Type

Public Type TYPE_BOOL_DESCRIPTION
    DefaultValue As Boolean
End Type

Public Type TYPE_BYTE_DESCRIPTION
    MinValue As Byte
    MaxValue As Byte
    DefaultValue As Byte
    Dimension As String
End Type

Public Type TYPE_INT_DESCRIPTION
    MinValue As Integer
    MaxValue As Integer
    DefaultValue As Integer
    Dimension As String
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

