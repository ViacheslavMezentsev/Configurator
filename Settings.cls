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
Public StepsViewFontBold As Boolean
Public StepsViewFontItalic As Boolean
Public StepsVSelectorEnabled As Boolean
Public StepsHSelectorEnabled As Boolean
Public StepsSelectorWidth As Long
Public StepsColWidth As Long
Public StepsRowHeight As Long

Public RewriteLogFile As Boolean
Public FilesHistoryLimitPaths As Boolean

Public AutoUpdateEnabled As Boolean
Public AutoUpdatePeriod As Long
Public AutoUpdateLink As String
Public AutoUpdateLastDate As String

Public StepsViewFontSize As Long

Public StepsViewFontName As String
Public LogFilePath As String

Public SplittersThickness As Long

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
    StepsColWidth = 250
    StepsRowHeight = 250
    
    StepsVSelectorEnabled = True
    StepsHSelectorEnabled = False
    StepsSelectorWidth = 20
    
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
    AutoUpdateLastDate = CStr(Date)
    AutoUpdateLink = "http://82.193.156.30/cop/update"
    
    ' [Интерфейс]
    SplittersThickness = 60
    
End Sub

'**
'@param
'@rem <h2>LoadSettings</h2>
'Загрузка настроек из конфигурационного файла
Public Sub LoadSettings()

    ' [Основные]
    StepsColWidth = IniFile.ReadInteger("Settings", "StepsColWidth", 300)
    StepsRowHeight = IniFile.ReadInteger("Settings", "StepsRowHeight", 300)

    StepsVSelectorEnabled = IniFile.ReadBoolean("Settings", "StepsVSelectorEnabled", True)
    StepsHSelectorEnabled = IniFile.ReadBoolean("Settings", "StepsHSelectorEnabled", False)
    StepsSelectorWidth = IniFile.ReadInteger("Settings", "StepsSelectorWidth", 20)

    StepsViewFontBold = IniFile.ReadBoolean("Settings", "StepsViewFontBold", False)
    StepsViewFontItalic = IniFile.ReadBoolean("Settings", "StepsViewFontItalic", False)
    StepsViewFontSize = IniFile.ReadInteger("Settings", "StepsViewFontSize", 10)
    StepsViewFontName = IniFile.ReadString("Settings", "StepsViewFontName", "Courier New")

    ' [История файлов]
    ' Ограничивать длину путей в меню
    FilesHistoryLimitPaths = IniFile.ReadBoolean("Settings", "FilesHistoryLimitPaths", _
            True)

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
    AutoUpdateLastDate = IniFile.ReadString("Settings", "AutoUpdateLastDate", CStr(Date))
    AutoUpdateLink = IniFile.ReadString("Settings", "AutoUpdateLink", _
            "http://82.193.156.30/cop/update")
            
    ' [Интерфейс]
    SplittersThickness = IniFile.ReadInteger("Settings", "SplittersThickness", 60)

End Sub

'**
'@param
'@rem <h2>SaveSettings</h2>
'Сохранение настроек в конфигурационном файле
Public Sub SaveSettings()

    ' [Основные]
    IniFile.WriteInteger "Settings", "StepsColWidth", StepsColWidth
    IniFile.WriteInteger "Settings", "StepsRowHeight", StepsRowHeight

    IniFile.WriteBoolean "Settings", "StepsVSelectorEnabled", StepsVSelectorEnabled
    IniFile.WriteBoolean "Settings", "StepsHSelectorEnabled", StepsHSelectorEnabled
    IniFile.WriteInteger "Settings", "StepsSelectorWidth", StepsSelectorWidth

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
    IniFile.WriteString "Settings", "AutoUpdateLastDate", AutoUpdateLastDate
    IniFile.WriteString "Settings", "AutoUpdateLink", AutoUpdateLink
    
    ' [Интерфейс]
    IniFile.WriteInteger "Settings", "SplittersThickness", SplittersThickness
    
End Sub
