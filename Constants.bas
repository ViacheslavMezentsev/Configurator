Attribute VB_Name = "Constants"
'**
'@author <a href="mailto:unihomelab@ya.ru">�������� ��������</a>
'@revision ���� �������: 16.06.2011 �., �����: 3:25:01
'@rem <h1><b>Constants</b></h1>
'<pre>
'--------------------------------------------------------------------------------
' ������   :       ������������ ����������� ��������
' ������   :       Constants
' �������� :       ����� ���������� �������� � ���������� ���������
' �����    :       �������� ��������
' ������  :       16.06.2011 �., �����: 3:25:01
'--------------------------------------------------------------------------------
'</pre>
Option Explicit

' -=[ �������� � Visual basic 6 ]=-

' ( ������� �������� ���� ������ � �������������� ����� ���� _
(%, &, !, #, @, $) ��������� ���������� )

' �������� ����: [������ � �������� ��������] _
Integer: [%], Long: [&], Currency: [@], Single: [!], Double: [#], String: [$]

' *****************************************
' *  ���������
' *  ~~~~~~~~~
' *****************************************

'**
'@rem ������������� ������
Public Const ������_��� = vbObjectError + 1000
'**
'@rem ������������� ������
Public Const ������_����������� = vbObjectError + 1001

'**
'@rem ������-����������� ���� � �������� �������
Public Const PATH_SEPARATOR As String = "\"
'**
'@rem �������� ���������
Public Const APP_NAME As String = "������������ ��"
'**
'@rem ��� ������ ����� �� ���������
Public Const DEFAULT_FILE_NAME As String = "��� �����"
'**
'@rem
Public Const TITLE_SECTION_NAME = "Title"
'**
'@rem GUID ��� ��������� "������������ �� 1.x (*.json)"
Public Const ProgramGUID As String = "{43CE9E0A-3657-4258-B573-8B18F6AC3B42}"

'**
'@rem
Public Const CRC8_FOR_DEFAULT_PROGRAM As Byte = 6

'**
'@rem
Public Const DESCR_MIN_VALUE As String = "  ����������� ��������: "
'**
'@rem
Public Const DESCR_MAX_VALUE As String = "  ������������ ��������: "
'**
'@rem
Public Const DESCR_DEFAULT_VALUE As String = "  �������� �� ���������: "
'**
'@rem
Public Const DESCR_DIMENSION As String = "  �����������: "

'**
'@rem ����� ����������� ������� ������: ������� �����.
Public Const STEPS_VIEW = 0
'**
'@rem ����� ����������� ������� ������: HEX-��������. <pre>
'------------------------------------------------------|
'|     00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F |
'------------------------------------------------------|
'|0000 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 |
'|0001 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 |
'<pre>
Public Const CODE_VIEW = 1

'**
'@rem ����� ����������� ������� �����: �����
Public Const TEXT_VIEW = 0
'**
'@rem ����� ����������� ������� �����: �������
Public Const CHECKS_VIEW = 1

'**
'@rem ��������� �� ���������
Public Const ENDSOUND_DEFAULT As Boolean = True
'**
'@rem ��������� �� ���������
Public Const DOORUNLOCK_DEFAULT As Boolean = True

'**
'@rem ��� ��������: �������
Public Const WPC_OPERATION_IDLE = 0
'**
'@rem ��� ��������: �����
Public Const WPC_OPERATION_FILL = 1
'**
'@rem ��� ��������: ������
Public Const WPC_OPERATION_DTRG = 2
'**
'@rem ��� ��������: ������
Public Const WPC_OPERATION_HEAT = 3
'**
'@rem ��� ��������: ������
Public Const WPC_OPERATION_WASH = 4
'**
'@rem ��� ��������: ����������
Public Const WPC_OPERATION_RINS = 5
'**
'@rem ��� ��������: ����������
Public Const WPC_OPERATION_JOLT = 6
'**
'@rem ��� ��������: �����
Public Const WPC_OPERATION_PAUS = 7
'**
'@rem ��� ��������: ����
Public Const WPC_OPERATION_DRAIN = 8
'**
'@rem ��� ��������: �����
Public Const WPC_OPERATION_SPIN = 9
'**
'@rem ��� ��������: ����������
Public Const WPC_OPERATION_COOL = 10
'**
'@rem ��� ��������: ���.����������
Public Const WPC_OPERATION_TRIN = 11

'**
'@rem ��� ��������: ������ ������� ����
Public Const LOADING_W_HOT = 0
'**
'@rem ��� ��������: ������ �������� ���� 1
Public Const LOADING_W_COLD_1 = 1
'**
'@rem ��� ��������: ������ �������� ���� 2
Public Const LOADING_W_COLD_2 = 2
'**
'@rem ��� ��������: ������ �� 1
Public Const LOADING_WD_1 = 3
'**
'@rem ��� ��������: ������ �� 2
Public Const LOADING_WD_2 = 4
'**
'@rem ��� ��������: ������ �� 3
Public Const LOADING_WD_3 = 5
'**
'@rem ��� ��������: ������ �� 4
Public Const LOADING_WD_4 = 6
'**
'@rem ��� ��������: ������ �� 5
Public Const LOADING_WD_5 = 7
'**
'@rem ��� ��������: ������ �� 6
Public Const LOADING_WD_6 = 8
'**
'@rem ��� ��������: ������ �� 7
Public Const LOADING_WD_7 = 9
'**
'@rem ��� ��������: ������ �� 8
Public Const LOADING_WD_8 = 10
'**
'@rem ��� ��������: ������ �� 9
Public Const LOADING_WD_9 = 11
'**
'@rem ��� ��������: ����� ���� 1
Public Const LOADING_LOCK_1 = 12
'**
'@rem ��� ��������: ����� ���� 2
Public Const LOADING_LOCK_2 = 13
'**
'@rem ��� ��������: ���� 1
Public Const LOADING_PUMP_1 = 14
'**
'@rem ��� ��������: ���� 2
Public Const LOADING_PUMP_2 = 15
'**
'@rem ��� ��������: ������
Public Const LOADING_HEAT = 16
'**
'@rem ��� ��������: ������
Public Const LOADING_DRIVE = 17

'**
'@rem
Public Const WC_COMPOSITECHECK = &H200
'**
'@rem
Public Const WC_DEFAULTCHAR = &H40
'**
'@rem
Public Const WC_DISCARDNS = &H10
'**
'@rem
Public Const WC_SEPCHARS = &H20

' Reg Key Security Options...
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const ERROR_SUCCESS = 0
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number

Public Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Public Const gREGVALSYSINFOLOC = "MSINFO"
Public Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Public Const gREGVALSYSINFO = "PATH"

'**
'@rem ���������� ����� ������� ���� ���������
Public Const NUMBER_OF_FUNCS = 11

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
'@rem ��� ��� ����������.
Public Hook As Long
'**
'@rem ������� �����
Public CurrentDir As String
'**
'@rem ������ �������� ������� ����� ���������. ������������ ���
'����������� �������� ����� � �������� ������� �����.
Public FunctionsStrings(0 To NUMBER_OF_FUNCS - 1) As String
'**
'@rem ����� ��������� �������� ��� ����� ���������. � �������
'������� �������� ���������� ������� �������� ������ � JSON ������.
'<br>
'������ ������ �������:<pre>
'    JSONStepsTemplates(WPC_OPERATION_FILL) = "{" _
'       & """Type"": 1," _
'       & """Pause"": false," _
'       & """ColdWaterGate"": false," _
'       & """HotWaterGate"": false," _
'       & """RecycledWaterGate"": false," _
'       & """Rotation"": true," _
'       & """Level"": 15," _
'       & """RotationTime"": 6," _
'       & """PauseTime"": 12," _
'       & """DrumSpeed"": 50}"
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
'@rem ��������� ���������
Public Settings As CSettings
'**
'@rem ���������������� ���� ���������
Public IniFile As CIniFiles
'**
'@rem ����������� ��������� ������ ����� �������
Public Manager As CProgramManager
'**
'@rem ����� ��� ������ �� ������� ����� �������� ������
Public MRUFileList As cMRUFileList
