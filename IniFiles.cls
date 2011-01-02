VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TIniFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Класс для работы с ini файлами"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private FFileName As String

Public Sub Create(FileName As String)
    FFileName = FileName
End Sub

Public Sub WriteString(Section As String, Ident As String, Value As String)
    Dim Result As Boolean
    Result = WritePrivateProfileStringByKeyName(Section, Ident, Value, FFileName)
End Sub

Public Function ReadString(Section, Ident, Default As String) As String
    Dim KeyValue$
    Dim characters As Long
    
    ' Инициализируем буфер
    KeyValue$ = String$(128, 0)
    
    characters = GetPrivateProfileStringByKeyName(Section, Ident, "", _
        KeyValue$, 127, FFileName)
        
    If (characters > 1) Then
        KeyValue$ = Left$(KeyValue$, characters)
    End If
    
    ReadString = KeyValue$
End Function

Public Sub WriteInteger(Section As String, Ident As String, Value As Long)
    Dim Result As Boolean
    Dim StrVal As String
    StrVal = Str(Value)
    Result = WritePrivateProfileStringByKeyName(Section, Ident, StrVal, FFileName)
End Sub

Public Function ReadInteger(Section As String, Ident As String, Default As Long) As Long
    Dim KeyValue$
    Dim characters As Long
    
    ' Инициализируем буфер
    KeyValue$ = String$(128, 0)
    
    characters = GetPrivateProfileStringByKeyName(Section, Ident, "", _
        KeyValue$, 127, FFileName)
        
    If (characters > 1) Then
        KeyValue$ = Left$(KeyValue$, characters)
        ReadInteger = Round(Val(KeyValue$))
    Else
        ReadInteger = Default
    End If
End Function

