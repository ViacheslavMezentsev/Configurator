VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private FLogMode As Long
Private FLogFileName As String


Property Get FileName() As String

    FileName = FLogFileName
    
End Property


Public Sub StartLogging(FileName As String, LogMode As Long)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    
    FLogFileName = FileName

    If LogMode = VBRUN.LogModeConstants.vbLogToFile Then
        
        ' Проверяем доступ к лог-файлу
        Dim fso As Scripting.FileSystemObject
        Dim TS As TextStream
           
        Set fso = New Scripting.FileSystemObject
        Set TS = fso.OpenTextFile(FileName, ForAppending, True)
        
        If Not TS Is Nothing Then
    
            TS.Close
            
            FLogMode = LogMode
    
        End If
        
        Set TS = Nothing
        Set fso = Nothing
    
    End If

End Sub


Public Sub Info(Text As String)

    Select Case FLogMode
    
        Case VBRUN.LogModeConstants.vbLogOff:
    
            Debug.Print CStr(Date) & " " & CStr(Time) & " [INFO ] " & Text
            
            
        Case VBRUN.LogModeConstants.vbLogToFile:
        
            Dim fso As Scripting.FileSystemObject
            Dim TS As TextStream
               
            Set fso = New Scripting.FileSystemObject
            Set TS = fso.OpenTextFile(FLogFileName, ForAppending, True)
            
            TS.WriteLine CStr(Date) & " " & CStr(Time) & " [INFO ] " & Text
            
            TS.Close
            
            Set TS = Nothing
            Set fso = Nothing
    
    
    End Select
    
End Sub


Public Sub Warn(Text As String)

    Select Case FLogMode
    
        Case VBRUN.LogModeConstants.vbLogOff:
    
            Debug.Print CStr(Date) & " " & CStr(Time) & " [INFO ] " & Text
            
            
        Case VBRUN.LogModeConstants.vbLogToFile:
            
            Dim fso As Scripting.FileSystemObject
            Dim TS As TextStream
               
            Set fso = New Scripting.FileSystemObject
            Set TS = fso.OpenTextFile(FileName, ForAppending, True)
            
            TS.WriteLine CStr(Date) & " " & CStr(Time) & " [WARN ] " & Text
            
            TS.Close
            
            Set TS = Nothing
            Set fso = Nothing
    
    
    End Select
    
End Sub


Public Sub Error(Text As String)

    Select Case FLogMode
    
        Case VBRUN.LogModeConstants.vbLogOff:
    
            Debug.Print CStr(Date) & " " & CStr(Time) & " [INFO ] " & Text
            
            
        Case VBRUN.LogModeConstants.vbLogToFile:
         
            Dim fso As Scripting.FileSystemObject
            Dim TS As TextStream
               
            Set fso = New Scripting.FileSystemObject
            Set TS = fso.OpenTextFile(FileName, ForAppending, True)
            
            TS.WriteLine CStr(Date) & " " & CStr(Time) & " [ERROR] " & Text
            
            TS.Close
            
            Set TS = Nothing
            Set fso = Nothing
    
    
    End Select
    
End Sub

Private Sub Class_Initialize()
    
    FLogMode = VBRUN.LogModeConstants.vbLogOff

End Sub
