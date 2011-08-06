VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection

Public Sub addElement(elem As Variant)

    mCol.Add elem
    
End Sub

Public Function Size() As Long

    Size = mCol.Count
    
End Function

Public Sub removeElementAt(index As Long)

    mCol.Remove (index)

End Sub

Public Function elementAt(index As Long) As Variant

    Dim obj As Object
    
    elementAt = mCol.Item(index)
    
End Function

Private Sub Class_Initialize()

    Set mCol = New Collection
    
End Sub

Public Function elements() As Collection

    Set elements = mCol
    
End Function

Public Sub removeAllElements()

    Dim I As Long
    
    For I = 1 To mCol.Count
    
        removeElementAt (1)
        
    Next
    
End Sub


Private Sub Class_Terminate()

    Set mCol = Nothing
    
End Sub


