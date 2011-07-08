Attribute VB_Name = "Constants"
'**
'@author <a href="mailto:unihomelab@ya.ru">Мезенцев Вячеслав</a>
'@revision Дата ревизии: 16.06.2011 г., время: 3:25:01
'@rem <h1><b>Constants</b></h1>
'<pre>
'--------------------------------------------------------------------------------
' Проект   :       Конфигуратор управляющих программ
' Модуль   :       Constants
' Описание :       Набор глобальных констант и переменных программы
' Автор    :       Мезенцев Вячеслав
' Изменён  :       16.06.2011 г., время: 3:25:01
'--------------------------------------------------------------------------------
'</pre>
Option Explicit

' -=[ Суффиксы в Visual basic 6 ]=-

' ( Техника указания типа данных с использованием знака типа _
(%, &, !, #, @, $) считается устаревшей )

' Название типа: [Символ в качестве суффикса] _
Integer: [%], Long: [&], Currency: [@], Single: [!], Double: [#], String: [$]

' *****************************************
' *  КОНСТАНТЫ
' *  ~~~~~~~~~
' *****************************************

'**
'@rem Идентификатор ошибки
Public Const ОШИБОК_НЕТ = vbObjectError + 1000
'**
'@rem Идентификатор ошибки
Public Const ОШИБКА_НЕИЗВЕСТНАЯ = vbObjectError + 1001

'**
'@rem Символ-разделитель пути в файловой системе
Public Const PATH_SEPARATOR As String = "\"
'**
'@rem Название программы
Public Const APP_NAME As String = "Конфигуратор УП"
'**
'@rem Имя нового файла по умолчанию
Public Const DEFAULT_FILE_NAME As String = "Без имени"
'**
'@rem
Public Const TITLE_SECTION_NAME = "Title"
'**
'@rem GUID для программы "Конфигуратор УП 1.x (*.json)"
Public Const ProgramGUID As String = "{43CE9E0A-3657-4258-B573-8B18F6AC3B42}"

'**
'@rem
Public Const CRC8_FOR_DEFAULT_PROGRAM As Byte = 6

'**
'@rem
Public Const DESCR_MIN_VALUE As String = "  Минимальное значение: "
'**
'@rem
Public Const DESCR_MAX_VALUE As String = "  Максимальное значение: "
'**
'@rem
Public Const DESCR_DEFAULT_VALUE As String = "  Значение по умолчанию: "
'**
'@rem
Public Const DESCR_DIMENSION As String = "  Размерность: "

'**
'@rem Режим отображения средней панели: таблица шагов.
Public Const STEPS_VIEW = 0
'**
'@rem Режим отображения средней панели: HEX-редактор. <pre>
'------------------------------------------------------|
'|     00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F |
'------------------------------------------------------|
'|0000 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 |
'|0001 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 |
'<pre>
Public Const CODE_VIEW = 1

'**
'@rem Режим отображения таблицы шагов: текст
Public Const TEXT_VIEW = 0
'**
'@rem Режим отображения таблицы шагов: галочки
Public Const CHECKS_VIEW = 1

'**
'@rem Настройка по умолчанию
Public Const ENDSOUND_DEFAULT As Boolean = True
'**
'@rem Настройка по умолчанию
Public Const DOORUNLOCK_DEFAULT As Boolean = True

'**
'@rem Тип операции: пропуск
Public Const WPC_OPERATION_IDLE = 0
'**
'@rem Тип операции: налив
Public Const WPC_OPERATION_FILL = 1
'**
'@rem Тип операции: моющие
Public Const WPC_OPERATION_DTRG = 2
'**
'@rem Тип операции: нагрев
Public Const WPC_OPERATION_HEAT = 3
'**
'@rem Тип операции: стирка
Public Const WPC_OPERATION_WASH = 4
'**
'@rem Тип операции: полоскание
Public Const WPC_OPERATION_RINS = 5
'**
'@rem Тип операции: расстряска
Public Const WPC_OPERATION_JOLT = 6
'**
'@rem Тип операции: пауза
Public Const WPC_OPERATION_PAUS = 7
'**
'@rem Тип операции: слив
Public Const WPC_OPERATION_DRAIN = 8
'**
'@rem Тип операции: отжим
Public Const WPC_OPERATION_SPIN = 9
'**
'@rem Тип операции: охлаждение
Public Const WPC_OPERATION_COOL = 10
'**
'@rem Тип операции: тех.полоскание
Public Const WPC_OPERATION_TRIN = 11

'**
'@rem Тип нагрузки: клапан горячей воды
Public Const LOADING_W_HOT = 0
'**
'@rem Тип нагрузки: клапан холодной воды 1
Public Const LOADING_W_COLD_1 = 1
'**
'@rem Тип нагрузки: клапан холодной воды 2
Public Const LOADING_W_COLD_2 = 2
'**
'@rem Тип нагрузки: клапан МС 1
Public Const LOADING_WD_1 = 3
'**
'@rem Тип нагрузки: клапан МС 2
Public Const LOADING_WD_2 = 4
'**
'@rem Тип нагрузки: клапан МС 3
Public Const LOADING_WD_3 = 5
'**
'@rem Тип нагрузки: клапан МС 4
Public Const LOADING_WD_4 = 6
'**
'@rem Тип нагрузки: клапан МС 5
Public Const LOADING_WD_5 = 7
'**
'@rem Тип нагрузки: клапан МС 6
Public Const LOADING_WD_6 = 8
'**
'@rem Тип нагрузки: клапан МС 7
Public Const LOADING_WD_7 = 9
'**
'@rem Тип нагрузки: клапан МС 8
Public Const LOADING_WD_8 = 10
'**
'@rem Тип нагрузки: клапан МС 9
Public Const LOADING_WD_9 = 11
'**
'@rem Тип нагрузки: Замок люка 1
Public Const LOADING_LOCK_1 = 12
'**
'@rem Тип нагрузки: Замок люка 2
Public Const LOADING_LOCK_2 = 13
'**
'@rem Тип нагрузки: Слив 1
Public Const LOADING_PUMP_1 = 14
'**
'@rem Тип нагрузки: Слив 2
Public Const LOADING_PUMP_2 = 15
'**
'@rem Тип нагрузки: Нагрев
Public Const LOADING_HEAT = 16
'**
'@rem Тип нагрузки: Движок
Public Const LOADING_DRIVE = 17

'**
'@rem
Public Const WC_COMPOSITECHECK = &H200
'**
'@rem
Public Const WC_DEFAULTCHAR = &H40
'**
'@rem
Public Const WC_DISCARDNS = &H10
'**
'@rem
Public Const WC_SEPCHARS = &H20

' Reg Key Security Options...
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const ERROR_SUCCESS = 0
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number

Public Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Public Const gREGVALSYSINFOLOC = "MSINFO"
Public Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Public Const gREGVALSYSINFO = "PATH"

'**
'@rem Количество типов функций шага программы
Public Const NUMBER_OF_FUNCS = 11

' *****************************************
' *  ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ
' *  ~~~~~~~~~~ ~~~~~~~~~~
' *****************************************

'**
'@rem Признак изменённости файла.
Public Modified As Boolean
'**
'@rem Режим проектирования.
Public DesignMode As Boolean
'**
'@rem Признак загрузки данных об уставках из файла limits.ini
Public LimitsLoaded As Boolean
'**
'@rem Хук для клавиатуры.
Public Hook As Long
'**
'@rem Текущая папка
Public CurrentDir As String
'**
'@rem Массив названий функций шагов программы. Используется для
'наглядности описания шагов в столбцах таблицы шагов.
Public FunctionsStrings(0 To NUMBER_OF_FUNCS - 1) As String
'**
'@rem Набор текстовых шаблонов для шагов программы. С помощью
'простых шаблонов упрощается функция экспорта данных в JSON формат.
'<br>
'Пример такого шаблона:<pre>
'    JSONStepsTemplates(WPC_OPERATION_FILL) = "{" _
'       & """Type"": 1," _
'       & """Pause"": false," _
'       & """ColdWaterGate"": false," _
'       & """HotWaterGate"": false," _
'       & """RecycledWaterGate"": false," _
'       & """Rotation"": true," _
'       & """Level"": 15," _
'       & """RotationTime"": 6," _
'       & """PauseTime"": 12," _
'       & """DrumSpeed"": 50}"
'</pre>
Public JSONStepsTemplates(0 To NUMBER_OF_FUNCS - 1) As String

'**
'@rem
Public EndSound As TYPE_BOOL_DESCRIPTION
'**
'@rem
Public DoorUnlock As TYPE_BOOL_DESCRIPTION

'**
'@rem
Public tMessage As Timer
'**
'@rem
Public ВекторОшибок As JVector
'**
'@rem Установки программы
Public Settings As CSettings
'**
'@rem Конфигурационный файл программы
Public IniFile As CIniFiles
'**
'@rem Конструктор бинарного образа файла проекта
Public Manager As CProgramManager
'**
'@rem Класс для работы со списком ранее открытых файлов
Public MRUFileList As cMRUFileList
