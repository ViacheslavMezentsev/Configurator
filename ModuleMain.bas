Attribute VB_Name = "ModuleMain"
'<CSCC>
'**
'@author <a href="mailto:unihomelab@ya.ru">�������� ��������</a>
'@revision ���� �������: 12.07.2011 �., �����: 4:49:47
'@rem <h1><b>ModuleMain</b></h1>
'@rem ������� ������ � �������� Main()
'<pre>
'--------------------------------------------------------------------------------
' ������   :       cop
' ������   :       ModuleMain
' �������� :       ������� ������ � �������� Main()
' �����    :       �������� ��������
' ������  :       12.07.2011 �., �����: 4:49:47
'--------------------------------------------------------------------------------
'</pre>
'</CSCC>
Option Explicit

'**
'@rem �������������� ���������� ��� ������ XP Manifest
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

'**
'@rem ����� ����� � ���������
Private Sub Main()
       
    ' ������ ��������� ������� �������� (������������� �� ���������)
    Set Settings = New CSettings
        
    ' ��������� ���������� ���������
    Settings.LoadSettings
    
    ' �������������� ��� ����, ���� ���� ����������
    If Settings.RewriteLogFile Then
    
        Debug.Print Date & " " & Time & " [cop.ModuleMain.Main]: " & _
                "���� ���� �����."
                
        DeleteFile Settings.LogFilePath
        
    End If

    ' ������ ��������� �������
    Set Logger = New CLogger
    
    ' ��������� ������� �������
    Logger.StartLogging Settings.LogFilePath, VBRUN.LogModeConstants.vbLogToFile
        
    ' ������������� ����������� ��� ���������� ������ ����������
    InitCommonControls

    ' ���������� �������� ���� ���������
    FormMain.Show
    
End Sub
