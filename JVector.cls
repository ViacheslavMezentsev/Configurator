VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"CString"
Attribute VB_Ext_KEY = "Member0" ,"CString"
'**
'@author <a href="mailto:unihomelab@ya.ru">Мезенцев Вячеслав</a>
'@revision Дата ревизии: 16.06.2011 г., время: 15:54:40
'@rem <h1><b>JVector</b></h1>
'Класс-контейнер для работы с коллекциями объектов.
'<pre>
'--------------------------------------------------------------------------------
' Проект   :       cop
' Модуль   :       JVector
' Описание :       Класс-контейнер для работы с коллекциями объектов
' Автор    :       Мезенцев Вячеслав
' Изменён  :       16.06.2011 г., время: 15:54:40
'--------------------------------------------------------------------------------
' Класс JVector
' J - от Java, так как работет с объектами:
' Set elementAt = mCol.Item(Index)
'</pre>
Option Explicit

Private mCol As Collection

Public Sub addElement(elem As Variant)

    mCol.Add elem

End Sub

Public Function Size() As Long

    Size = mCol.Count
    
End Function

Public Sub removeElementAt(Index As Integer)
Attribute removeElementAt.VB_UserMemId = -4
Attribute removeElementAt.VB_MemberFlags = "40"

    mCol.Remove (Index)
    
End Sub

Public Function elementAt(Index As Integer) As Variant

    Dim obj As Object
    
    Set elementAt = mCol.Item(Index)
    
End Function

Private Sub Class_Initialize()

    Set mCol = New Collection
    
End Sub

Public Function elements() As Collection

    Set elements = mCol
    
End Function

Public Sub removeAllElements()

    Dim i As Long
    
    For i = 1 To mCol.Count
    
        removeElementAt (1)
        
    Next
    
End Sub

Private Sub Class_Terminate()

    Set mCol = Nothing
    
End Sub

