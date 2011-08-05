VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIniFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Класс для работы с ini файлами"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**
'@author <a href="mailto:unihomelab@ya.ru">Мезенцев Вячеслав</a>
'@revision Дата ревизии: 16.06.2011 г., время: 6:02:24
'@rem <h1><b>CIniFiles</b></h1>
'Класс для работы с ini-файлами.
'<pre>
'--------------------------------------------------------------------------------
' Проект   :       Конфигуратор управляющих программ
' Модуль   :       CIniFiles
' Описание :       Класс для работы с ini-файлами
' Автор    :       Мезенцев Вячеслав
' Изменён  :       16.06.2011 г., время: 6:02:24
'--------------------------------------------------------------------------------
'</pre>
Option Explicit

Private FFileName As String

'**
'@param        FileName Required. String.
'@rem <h2>Create</h2>
'Конструктор с параметром имени файла.
Public Sub Create(FileName As String)

    FFileName = FileName
    
End Sub

'**
'@param        Section Required. String.
'@param        Ident Required. String.
'@param        Value Required. String.
'@rem <h2>WriteString</h2>
'Запись значения ключа в виде строки.
Public Sub WriteString(Section As String, Ident As String, Value As String)

    Dim Result As Boolean
    Result = WritePrivateProfileStringByKeyName(Section, Ident, Value, FFileName)
    
End Sub

'**
'@param        Section Required. String.
'@param        Ident Required. String.
'@param        Default Required. String.
'@return       String.
'@rem <h2>ReadString</h2>
'Чтние значения ключа в виде строки.
Public Function ReadString(Section As String, Ident As String, Default As String) As String
    Dim KeyValue$
    Dim characters As Long

    ' Инициализируем буфер
    KeyValue$ = String$(128, 0)

    characters = GetPrivateProfileStringByKeyName(Section, Ident, "", _
       KeyValue$, 127, FFileName)

    If (characters > 0) Then
    
        KeyValue$ = Left$(KeyValue$, characters)
        
    Else
    
        KeyValue$ = Default
        
    End If

    ReadString = KeyValue$
    
End Function

'**
'@param        Section Required. String.
'@param        Ident Required. String.
'@param        Value Required. Long.
'@rem <h2>WriteInteger</h2>
'Запись значения ключа в виде целого числа.
Public Sub WriteInteger(Section As String, Ident As String, Value As Long)

    Dim Result As Boolean
    Dim StrVal As String
    StrVal = str$(Value)
    Result = WritePrivateProfileStringByKeyName(Section, Ident, StrVal, FFileName)
    
End Sub

'**
'@param        Section Required. String.
'@param        Ident Required. String.
'@param        Default Required. Long.
'@return       Long.
'@rem <h2>ReadInteger</h2>
'Чтение значения ключа в виде целого числа.
Public Function ReadInteger(Section As String, Ident As String, Default As Long) As Long

    Dim KeyValue$
    Dim characters As Long

    ' Инициализируем буфер
    KeyValue$ = String$(128, 0)

    characters = GetPrivateProfileStringByKeyName(Section, Ident, "", KeyValue$, 127, _
            FFileName)

    If (characters > 0) Then

        KeyValue$ = Left$(KeyValue$, characters)
        ReadInteger = Round(Val(KeyValue$))
        
    Else
        ReadInteger = Default

    End If

End Function

'**
'@param        Section Required. String.
'@param        Ident Required. String.
'@param        Value Required. Boolean.
'@rem <h2>WriteBoolean</h2>
'Запись значения ключа в виде логического значения.
Public Sub WriteBoolean(Section As String, Ident As String, Value As Boolean)

    Select Case Value
        Case False: WriteInteger Section, Ident, 0
        Case True: WriteInteger Section, Ident, 1
    End Select
    
End Sub

'**
'@param        Section Required. String.
'@param        Ident Required. String.
'@param        Default Required. Boolean.
'@return       Boolean.
'@rem <h2>ReadBoolean</h2>
'Чтение значения ключа в виде логического значения.
Public Function ReadBoolean(Section As String, Ident As String, Default As Boolean) As Boolean

    Select Case Default
        Case False: ReadBoolean = ReadInteger(Section, Ident, 0) > 0
        Case True: ReadBoolean = ReadInteger(Section, Ident, 1) > 0
    End Select
    
End Function
