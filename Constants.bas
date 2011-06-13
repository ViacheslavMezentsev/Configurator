Attribute VB_Name = "Constants"
' Задание явного объявления переменных
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

Public Const PATH_SEPARATOR As String = "\"
Public Const APP_NAME As String = "Конфигуратор УП"
Public Const DEFAULT_FILE_NAME As String = "Без имени"
Public Const ProgramGUID As String = "{43CE9E0A-3657-4258-B573-8B18F6AC3B42}"

' Режимы отображения средней панели
Public Const STEPS_VIEW = 0
Public Const CODE_VIEW = 1

' Режимы отображения таблицы шагов
Public Const TEXT_VIEW = 0
Public Const CHECKS_VIEW = 1

' Типы операций
Public Const WPC_OPERATION_IDLE = 0 '// пропуск
Public Const WPC_OPERATION_FILL = 1 '// налив
Public Const WPC_OPERATION_DTRG = 2 '// моющие
Public Const WPC_OPERATION_HEAT = 3 '// нагрев
Public Const WPC_OPERATION_WASH = 4 '// стирка
Public Const WPC_OPERATION_RINS = 5 '// полоскание
Public Const WPC_OPERATION_JOLT = 6 '// расстряска
Public Const WPC_OPERATION_PAUS = 7 '// пауза
Public Const WPC_OPERATION_DRAIN = 8 '// слив
Public Const WPC_OPERATION_SPIN = 9 '// отжим
Public Const WPC_OPERATION_COOL = 10 '// охлаждение
Public Const WPC_OPERATION_TRIN = 11 '// тех.полоскание

' Типы нагрузок
Public Const LOADING_W_HOT = 0  ',    // клапан горячей воды
Public Const LOADING_W_COLD_1 = 1  ', // клапан холодной воды 1
Public Const LOADING_W_COLD_2 = 2 ', // клапан холодной воды 2
Public Const LOADING_WD_1 = 3 ',     // клапан МС 1
Public Const LOADING_WD_2 = 4 ',     // клапан МС 2
Public Const LOADING_WD_3 = 5 ',     // клапан МС 3
Public Const LOADING_WD_4 = 6 ',     // клапан МС 4
Public Const LOADING_WD_5 = 7 ',     // клапан МС 5
Public Const LOADING_WD_6 = 8 ',     // клапан МС 6
Public Const LOADING_WD_7 = 9 ',     // клапан МС 7
Public Const LOADING_WD_8 = 10 ',     // клапан МС 8
Public Const LOADING_WD_9 = 11 ',     // клапан МС 9
Public Const LOADING_LOCK_1 = 12 ',   // Замок люка 1
Public Const LOADING_LOCK_2 = 13 ',   // Замок люка 2
Public Const LOADING_PUMP_1 = 14 ',   // Слив 1
Public Const LOADING_PUMP_2 = 15 ',   // Слив 2
Public Const LOADING_HEAT = 16 '      // Нагрев
Public Const LOADING_DRIVE = 17 '// Движок

Public Const WC_COMPOSITECHECK = &H200
Public Const WC_DEFAULTCHAR = &H40
Public Const WC_DISCARDNS = &H10
Public Const WC_SEPCHARS = &H20

Public Const CP_ACP = 0
Public Const CP_OEMCP = 1
Public Const CP_MACCP = 2
Public Const CP_THREAD_ACP = 3
Public Const CP_SYMBOL = 42
Public Const CP_UTF7 = 65000
Public Const CP_UTF8 = 65001
 
' *****************************************
' *  ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ
' *  ~~~~~~~~~~ ~~~~~~~~~~
' *****************************************

Public Const NUMBER_OF_FUNCS = 11

' Признак изменённости файла
Public Modified As Boolean
Public DesignMode As Boolean
Public FunctionsStrings(0 To NUMBER_OF_FUNCS - 1) As String
Public JSONStepsTemplates(0 To NUMBER_OF_FUNCS - 1) As String
Public Hook As Long, tMessage As Timer

