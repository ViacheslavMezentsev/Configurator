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
       
    InitCommonControls
    
    FormMain.Show
    
End Sub
