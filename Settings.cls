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
'@author �������� ��������
'@revision 16.06.2011-3:06:51
'@rem <h1><b>CSettings</b></h1>
'������ ��������.
'<pre>
'--------------------------------------------------------------------------------
' ������   :       ������������ ����������� ��������
' ������   :       CSettings
' �������� :       ������ ��������
' �����    :       �������� ��������
' �������  :       16.06.2011-3:06:51
'--------------------------------------------------------------------------------
'</pre>
Option Explicit

' ������������� ��������
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

Public AutoUpdateHost As String
Public AutoUpdatePort As Long

Public StepsViewFontSize As Long

Public StepsViewFontName As String
Public LogFilePath As String

Public SplittersThickness As Long

Public IniFilePath As String

'**
'@rem ��������� ���������
Private FIniFile As CIniFile


Property Get IniFile() As CIniFile

    Set IniFile = FIniFile

End Property


Property Set IniFile(Value As CIniFile)

    If Value Is Nothing Then Exit Property

    Set FIniFile = Value

End Property


'**
'@param
'@rem <h2>Class_Initialize</h2> - ����������� ������
Private Sub Class_Initialize()

    SetDefaultSettings
    
    ' ������ ���� ��� �������� ���������
    Set FIniFile = New CIniFile
    
    FIniFile.Create IniFilePath
    
End Sub


Private Sub Class_Terminate()
    
    Set FIniFile = Nothing
    
End Sub


'**
'@param
'@rem <h2>SetDefaultSettings</h2> - ��������� �������� �� ���������
Private Sub SetDefaultSettings()

    ' [��������]
    StepsColWidth = 250
    StepsRowHeight = 250
    
    StepsVSelectorEnabled = True
    StepsHSelectorEnabled = False
    StepsSelectorWidth = 20
    
    StepsViewFontBold = False
    StepsViewFontItalic = False
    StepsViewFontName = "Courier New"
    StepsViewFontSize = 10

    ' [������� ������]
    ' ������������ ����� ����� � ����
    FilesHistoryLimitPaths = True

    ' [���]
    RewriteLogFile = True

    ' ��������� ���� � ����� ����
    Dim FilePath As String, Result As String

    FilePath = String$(255, 0)
    Result = GetModuleFileName(0, FilePath, 254)
    CurrentDir = MiscExtractPathName(FilePath, True)
    FilePath = StrConv(FilePath, vbLowerCase)

    LogFilePath = Replace(FilePath, ".exe", ".log")
    IniFilePath = Replace(FilePath, ".exe", ".ini")
    
    Dim WE_ARE_IN_IDE As Boolean
    
    Debug.Assert MakeTrue(WE_ARE_IN_IDE)
    
    ' [����������]
    If WE_ARE_IN_IDE Then
        
        AutoUpdateHost = "192.168.0.14"
        AutoUpdateLink = "ftp://192.168.0.14/update.json"
        
    Else
        
        AutoUpdateHost = "82.193.156.30"
        AutoUpdateLink = "ftp://82.193.156.30/update.json"
    
    End If
    
    AutoUpdatePort = 1980
    AutoUpdateEnabled = False
    AutoUpdatePeriod = AUP_EVERY_DAY
    AutoUpdateLastDate = CStr(Date)
    

    ' [���������]
    SplittersThickness = 60
    
End Sub

'**
'@param
'@rem <h2>LoadSettings</h2>
'�������� �������� �� ����������������� �����
Public Sub LoadSettings()

    ' ���� ������ ����� �������� �� ������, �� ������� �� ������
    If FIniFile Is Nothing Then Exit Sub
    
    ' [��������]
    StepsColWidth = FIniFile.ReadInteger("Settings", "StepsColWidth", 300)
    StepsRowHeight = FIniFile.ReadInteger("Settings", "StepsRowHeight", 300)

    StepsVSelectorEnabled = FIniFile.ReadBoolean("Settings", "StepsVSelectorEnabled", True)
    StepsHSelectorEnabled = FIniFile.ReadBoolean("Settings", "StepsHSelectorEnabled", False)
    StepsSelectorWidth = FIniFile.ReadInteger("Settings", "StepsSelectorWidth", 20)

    StepsViewFontBold = FIniFile.ReadBoolean("Settings", "StepsViewFontBold", False)
    StepsViewFontItalic = FIniFile.ReadBoolean("Settings", "StepsViewFontItalic", False)
    StepsViewFontSize = FIniFile.ReadInteger("Settings", "StepsViewFontSize", 10)
    StepsViewFontName = FIniFile.ReadString("Settings", "StepsViewFontName", "Courier New")

    ' [������� ������]
    ' ������������ ����� ����� � ����
    FilesHistoryLimitPaths = FIniFile.ReadBoolean("Settings", "FilesHistoryLimitPaths", _
            True)

    ' [���]
    ' �������������� ��� ���� ��� ������ ������� ���������
    RewriteLogFile = FIniFile.ReadBoolean("Settings", "RewriteLogFile", True)

    ' ��������� ���� � ����� ����
    Dim FilePath As String, Result As String

    FilePath = String$(255, 0)
    Result = GetModuleFileName(0, FilePath, 254)
    CurrentDir = MiscExtractPathName(FilePath, True)
    FilePath = StrConv(FilePath, vbLowerCase)

    LogFilePath = FIniFile.ReadString("Settings", "LogFilePath", Replace(FilePath, ".exe", ".log"))
    IniFilePath = FIniFile.ReadString("Settings", "IniFilePath", Replace(FilePath, ".exe", ".ini"))
    
    ' [����������]
    Dim WE_ARE_IN_IDE As Boolean
    
    Debug.Assert MakeTrue(WE_ARE_IN_IDE)
    
    If WE_ARE_IN_IDE Then
        
        AutoUpdateHost = FIniFile.ReadString("Settings", "AutoUpdateHost", "192.168.0.14")
        AutoUpdateLink = FIniFile.ReadString("Settings", "AutoUpdateLink", "ftp://192.168.0.14/update.json")

    Else
        
        AutoUpdateHost = FIniFile.ReadString("Settings", "AutoUpdateHost", "82.193.156.30")
        AutoUpdateLink = FIniFile.ReadString("Settings", "AutoUpdateLink", "ftp://82.193.156.30/update.json")
    
    End If

    AutoUpdateEnabled = FIniFile.ReadBoolean("Settings", "AutoUpdateEnabled", False)
    AutoUpdatePeriod = FIniFile.ReadInteger("Settings", "AutoUpdatePeriod", AUP_EVERY_DAY)
    AutoUpdateLastDate = FIniFile.ReadString("Settings", "AutoUpdateLastDate", CStr(Date))
    
    AutoUpdatePort = FIniFile.ReadInteger("Settings", "AutoUpdatePort", 1980)
        
    ' [���������]
    SplittersThickness = IniFile.ReadInteger("Settings", "SplittersThickness", 60)

End Sub

'**
'@param
'@rem <h2>SaveSettings</h2>
'���������� �������� � ���������������� �����
Public Sub SaveSettings()

    ' ���� ������ ����� �������� �� ������, �� ������� �� ������
    If FIniFile Is Nothing Then Exit Sub

    ' [��������]
    FIniFile.WriteInteger "Settings", "StepsColWidth", StepsColWidth
    FIniFile.WriteInteger "Settings", "StepsRowHeight", StepsRowHeight

    FIniFile.WriteBoolean "Settings", "StepsVSelectorEnabled", StepsVSelectorEnabled
    FIniFile.WriteBoolean "Settings", "StepsHSelectorEnabled", StepsHSelectorEnabled
    FIniFile.WriteInteger "Settings", "StepsSelectorWidth", StepsSelectorWidth

    FIniFile.WriteBoolean "Settings", "StepsViewFontBold", StepsViewFontBold
    FIniFile.WriteBoolean "Settings", "StepsViewFontItalic", StepsViewFontItalic
    FIniFile.WriteInteger "Settings", "StepsViewFontSize", StepsViewFontSize
    FIniFile.WriteString "Settings", "StepsViewFontName", StepsViewFontName

    ' [������� ������]
    ' ������������ ����� ����� � ����
    FIniFile.WriteBoolean "Settings", "FilesHistoryLimitPaths", FilesHistoryLimitPaths

    ' [���]
    FIniFile.WriteBoolean "Settings", "RewriteLogFile", RewriteLogFile
    FIniFile.WriteString "Settings", "LogFilePath", LogFilePath
    FIniFile.WriteString "Settings", "IniFilePath", IniFilePath
    
    ' [����������]
    FIniFile.WriteString "Settings", "AutoUpdateHost", AutoUpdateHost
    FIniFile.WriteInteger "Settings", "AutoUpdatePort", AutoUpdatePort
    
    FIniFile.WriteBoolean "Settings", "AutoUpdateEnabled", AutoUpdateEnabled
    FIniFile.WriteInteger "Settings", "AutoUpdatePeriod", AutoUpdatePeriod
    FIniFile.WriteString "Settings", "AutoUpdateLastDate", AutoUpdateLastDate
    FIniFile.WriteString "Settings", "AutoUpdateLink", AutoUpdateLink
    
    ' [���������]
    FIniFile.WriteInteger "Settings", "SplittersThickness", SplittersThickness
    
End Sub

