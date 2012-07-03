Attribute VB_Name = "ModuleMain"
'<CSCC>
'**
'@author <a href="mailto:unihomelab@ya.ru">Мезенцев Вячеслав</a>
'@revision Дата ревизии: 12.07.2011 г., время: 4:49:47
'@rem <h1><b>ModuleMain</b></h1>
'@rem Главный модуль с функцией Main()
'<pre>
'--------------------------------------------------------------------------------
' Проект   :       cop
' Модуль   :       ModuleMain
' Описание :       Главный модуль с функцией Main()
' Автор    :       Мезенцев Вячеслав
' Изменён  :       12.07.2011 г., время: 4:49:47
'--------------------------------------------------------------------------------
'</pre>
'</CSCC>
Option Explicit

'**
'@rem Инициализируем компоненты для работы XP Manifest
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

'**
'@rem Точка входа в программу
Private Sub Main()
       
    ' Создаём экземпляр объекта настроек (инициализация по умолчанию)
    Set Settings = New CSettings
        
    ' Загружаем сохранённые настройки
    Settings.LoadSettings
    
    ' Перезаписываем лог файл, если флаг установлен
    If Settings.RewriteLogFile Then
    
        Debug.Print Date & " " & Time & " [cop.ModuleMain.Main]: " & _
                "Файл лога удалён."
                
        DeleteFile Settings.LogFilePath
        
    End If

    ' Создаём экземпляр журнала
    Set Logger = New CLogger
    
    ' Запускаем ведение журнала
    Logger.StartLogging Settings.LogFilePath, VBRUN.LogModeConstants.vbLogToFile
        
    ' Инициализация компонентов для правильной работы интерфейса
    InitCommonControls

    ' Показываем основное окно программы
    FormMain.Show
    
End Sub
