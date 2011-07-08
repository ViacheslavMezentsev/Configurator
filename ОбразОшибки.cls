VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ОбразОшибки"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ОбразОшибки"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**
'@author <a href="mailto:unihomelab@ya.ru">Мезенцев Вячеслав</a>
'@revision Дата ревизии: 16.06.2011 г., время: 16:02:48
'@rem <h1><b>ОбразОшибки</b></h1>
'Класс-обёртка для хранении информации об ошибке.
'<pre>
'--------------------------------------------------------------------------------
' Проект   :       Конфигуратор управляющих программ
' Модуль   :       ОбразОшибки
' Описание :       Класс-обёртка для хранении информации об ошибке
' Автор    :       Мезенцев Вячеслав
' Изменён  :       16.06.2011 г., время: 16:02:48
'--------------------------------------------------------------------------------
'</pre>
Option Explicit

'**
'@rem
Public НомерОшибки As Long
'**
'@rem
Public Сообщение As String

'**
'@rem Конструктор
Private Sub Class_Initialize()
End Sub

'**
'@rem Деструктор
Private Sub Class_Terminate()
End Sub

'**
'@param        Свойства Required. ОбразОшибки object.
'@rem Конструктор копирования.
Public Sub Создать(ByVal Свойства As ОбразОшибки)
    Me.НомерОшибки = Свойства.НомерОшибки
    Me.Сообщение = Свойства.Сообщение
End Sub
