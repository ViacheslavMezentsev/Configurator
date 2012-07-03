Attribute VB_Name = "ModuleGlobalVariables"
'**
'@author <a href="mailto:unihomelab@ya.ru">Мезенцев Вячеслав</a>
'@revision Дата ревизии: 27.06.2012 г., время: 00:34:47
'@rem <h1><b>ModuleGlobalVariables</b></h1>
'<pre>
'--------------------------------------------------------------------------------
' Проект   :       Сервер обновлений
' Модуль   :       ModuleGlobalVariables
' Описание :       Набор глобальных переменных программы
' Автор    :       Мезенцев Вячеслав
' Изменён  :       27.06.2012 г., время: 00:34:47
'--------------------------------------------------------------------------------
'</pre>
Option Explicit


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
'@rem
Public SetCancel As Boolean
'**
'@rem Состояние процесса автообновления
Public AutoUpdateState As Byte
'**
'@rem Хук для клавиатуры.
Public Hook As Long
'**
'@rem Текущая папка
Public CurrentDir As String
'**
'@rem Буфер для хранения пользовательских данных
Public ListProgramsRowData(1 To MAX_NUMBER_OF_PROGRAMS) As Integer
'**
'@rem Массив названий функций шагов программы. Используется для
'наглядности описания шагов в столбцах таблицы шагов.
Public FunctionsStrings(0 To NUMBER_OF_FUNCS - 1) As String
'**
'@rem Набор текстовых шаблонов для шагов программы. С помощью
'простых шаблонов упрощается функция экспорта данных в JSON формат.
'<br>
'Пример такого шаблона:<pre>
'    JSONStepsTemplates(WPC_OPERATION_FILL) = "{" '       & """Type"": 1," '       & """Pause"": false," '       & """ColdWaterGate"": false," '       & """HotWaterGate"": false," '       & """RecycledWaterGate"": false," '       & """Rotation"": true," '       & """Level"": 15," '       & """RotationTime"": 6," '       & """PauseTime"": 12," '       & """DrumSpeed"": 50}"
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
'@rem Конструктор бинарного образа файла проекта
Public Manager As CProgramManager
'**
'@rem Класс для работы со списком ранее открытых файлов
Public MRUFileList As cMRUFileList
'**
'@rem Флаг закрытия программы
Public isClose As Boolean
'**
'@rem Журнал проекта
Public Logger As CLogger
'**
'@rem Настройки программы
Public Settings As CSettings



