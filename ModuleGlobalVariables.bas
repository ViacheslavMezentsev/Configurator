Attribute VB_Name = "ModuleGlobalVariables"
'**
'@author <a href="mailto:unihomelab@ya.ru">�������� ��������</a>
'@revision ���� �������: 27.06.2012 �., �����: 00:34:47
'@rem <h1><b>ModuleGlobalVariables</b></h1>
'<pre>
'--------------------------------------------------------------------------------
' ������   :       ������ ����������
' ������   :       ModuleGlobalVariables
' �������� :       ����� ���������� ���������� ���������
' �����    :       �������� ��������
' ������  :       27.06.2012 �., �����: 00:34:47
'--------------------------------------------------------------------------------
'</pre>
Option Explicit


' *****************************************
' *  ���������� ����������
' *  ~~~~~~~~~~ ~~~~~~~~~~
' *****************************************


'**
'@rem ������� ����������� �����.
Public Modified As Boolean
'**
'@rem ����� ��������������.
Public DesignMode As Boolean
'**
'@rem ������� �������� ������ �� �������� �� ����� limits.ini
Public LimitsLoaded As Boolean
'**
'@rem
Public SetCancel As Boolean
'**
'@rem ��������� �������� ��������������
Public AutoUpdateState As Byte
'**
'@rem ��� ��� ����������.
Public Hook As Long
'**
'@rem ������� �����
Public CurrentDir As String
'**
'@rem ����� ��� �������� ���������������� ������
Public ListProgramsRowData(1 To MAX_NUMBER_OF_PROGRAMS) As Integer
'**
'@rem ������ �������� ������� ����� ���������. ������������ ���
'����������� �������� ����� � �������� ������� �����.
Public FunctionsStrings(0 To NUMBER_OF_FUNCS - 1) As String
'**
'@rem ����� ��������� �������� ��� ����� ���������. � �������
'������� �������� ���������� ������� �������� ������ � JSON ������.
'<br>
'������ ������ �������:<pre>
'    JSONStepsTemplates(WPC_OPERATION_FILL) = "{" '       & """Type"": 1," '       & """Pause"": false," '       & """ColdWaterGate"": false," '       & """HotWaterGate"": false," '       & """RecycledWaterGate"": false," '       & """Rotation"": true," '       & """Level"": 15," '       & """RotationTime"": 6," '       & """PauseTime"": 12," '       & """DrumSpeed"": 50}"
'</pre>
Public JSONStepsTemplates(0 To NUMBER_OF_FUNCS - 1) As String

'**
'@rem
Public EndSound As TYPE_BOOL_DESCRIPTION
'**
'@rem
Public DoorUnlock As TYPE_BOOL_DESCRIPTION

'**
'@rem
Public tMessage As Timer
'**
'@rem
Public ������������ As JVector
'**
'@rem ����������� ��������� ������ ����� �������
Public Manager As CProgramManager
'**
'@rem ����� ��� ������ �� ������� ����� �������� ������
Public MRUFileList As cMRUFileList
'**
'@rem ���� �������� ���������
Public isClose As Boolean
'**
'@rem ������ �������
Public Logger As CLogger
'**
'@rem ��������� ���������
Public Settings As CSettings



