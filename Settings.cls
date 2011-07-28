VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**
'@author Мезенцев Вячеслав
'@revision 16.06.2011-3:06:51
'@rem <h1><b>CSettings</b></h1>
'Модуль настроек.
'<pre>
'--------------------------------------------------------------------------------
' Проект   :       Конфигуратор управляющих программ
' Модуль   :       CSettings
' Описание :       Модуль настроек
' Автор    :       Мезенцев Вячеслав
' Изменён  :       16.06.2011-3:06:51
'--------------------------------------------------------------------------------
'</pre>
Option Explicit

' Инициализация настроек
Public StepColWidth As Long
Public RowHeight As Long

Public StepsViewFontBold As Boolean
Public StepsViewFontItalic As Boolean
Public RewriteLogFile As Boolean
Public FilesHistoryLimitPaths As Boolean

Public AutoUpdateEnabled As Boolean
Public AutoUpdatePeriod As Long
Public AutoUpdateLink As String
Public AutoUpdateLastDate As String

Public StepsViewFontSize As Long

Public StepsViewFontName As String
Public LogFilePath As String

'**
'@param
'@rem <h2>Class_Initialize</h2> - конструктор класса
Private Sub Class_Initialize()

    SetDefaultSettings
    
End Sub

'**
'@param
'@rem <h2>SetDefaultSettings</h2> - установка настроек по умолчанию
Private Sub SetDefaultSettings()

    ' [Основные]
    StepColWidth = 250
    RowHeight = 250

    StepsViewFontBold = False
    StepsViewFontItalic = False
    StepsViewFontName = "Courier New"
    StepsViewFontSize = 10

    ' [История файлов]
    ' Ограничивать длину путей в меню
    FilesHistoryLimitPaths = True

    ' [Лог]
    RewriteLogFile = True

    ' Формируем путь к файлу лога
    Dim FilePath As String, Result As String

    FilePath = String$(255, 0)
    Result = GetModuleFileName(0, FilePath, 254)
    CurrentDir = MiscExtractPathName(FilePath, True)
    FilePath = StrConv(FilePath, vbLowerCase)
    FilePath = Replace(FilePath, ".exe", ".log")

    LogFilePath = FilePath
    
    ' [Обновление]
    AutoUpdateEnabled = False
    AutoUpdatePeriod = AUP_EVERY_DAY
    AutoUpdateLink = "http://x-uni-x.chat.ru/cop/update"
    
End Sub

'**
'@param
'@rem <h2>LoadSettings</h2>
'Загрузка настроек из конфигурационного файла
Public Sub LoadSettings()

    ' [Основные]
    StepColWidth = IniFile.ReadInteger("Settings", "StepColWidth", 300)
    RowHeight = IniFile.ReadInteger("Settings", "RowHeight", 300)

    StepsViewFontBold = IniFile.ReadBoolean("Settings", "StepsViewFontBold", False)
    StepsViewFontItalic = IniFile.ReadBoolean("Settings", "StepsViewFontItalic", False)
    StepsViewFontSize = IniFile.ReadInteger("Settings", "StepsViewFontSize", 10)
    StepsViewFontName = IniFile.ReadString("Settings", "StepsViewFontName", "Courier New")

    ' [История файлов]
    ' Ограничивать длину путей в меню
    FilesHistoryLimitPaths = IniFile.ReadBoolean("Settings", "FilesHistoryLimitPaths", True)

    ' [Лог]
    ' Перезаписывать лог файл при каждом запуске программы
    RewriteLogFile = IniFile.ReadBoolean("Settings", "RewriteLogFile", True)

    ' Формируем путь к файлу лога
    Dim FilePath As String, Result As String

    FilePath = String$(255, 0)
    Result = GetModuleFileName(0, FilePath, 254)
    CurrentDir = MiscExtractPathName(FilePath, True)
    FilePath = StrConv(FilePath, vbLowerCase)
    FilePath = Replace(FilePath, ".exe", ".log")

    LogFilePath = IniFile.ReadString("Settings", "LogFilePath", FilePath)
    
    ' [Обновление]
    AutoUpdateEnabled = IniFile.ReadBoolean("Settings", "AutoUpdateEnabled", False)
    AutoUpdatePeriod = IniFile.ReadInteger("Settings", "AutoUpdatePeriod", AUP_EVERY_DAY)
    AutoUpdateLink = IniFile.ReadString("Settings", "AutoUpdateLink", _
            "http://x-uni-x.chat.ru/cop/update")
            
End Sub

'**
'@param
'@rem <h2>SaveSettings</h2>
'Сохранение настроек в конфигурационном файле
Public Sub SaveSettings()

    ' [Основные]
    IniFile.WriteInteger "Settings", "StepColWidth", StepColWidth
    IniFile.WriteInteger "Settings", "RowHeight", RowHeight

    IniFile.WriteBoolean "Settings", "StepsViewFontBold", StepsViewFontBold
    IniFile.WriteBoolean "Settings", "StepsViewFontItalic", StepsViewFontItalic
    IniFile.WriteInteger "Settings", "StepsViewFontSize", StepsViewFontSize
    IniFile.WriteString "Settings", "StepsViewFontName", StepsViewFontName

    ' [История файлов]
    ' Ограничивать длину путей в меню
    IniFile.WriteBoolean "Settings", "FilesHistoryLimitPaths", FilesHistoryLimitPaths

    ' [Лог]
    IniFile.WriteBoolean "Settings", "RewriteLogFile", RewriteLogFile
    IniFile.WriteString "Settings", "LogFilePath", LogFilePath
    
    ' [Обновление]
    IniFile.WriteBoolean "Settings", "AutoUpdateEnabled", AutoUpdateEnabled
    IniFile.WriteInteger "Settings", "AutoUpdatePeriod", AutoUpdatePeriod
    
End Sub
