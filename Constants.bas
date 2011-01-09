Attribute VB_Name = "Constants"
' ������� ������ ���������� ����������
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

Public Const PATH_SEPARATOR As String = "\"
Public Const APP_NAME As String = "������������ ��"
Public Const DEFAULT_FILE_NAME As String = "��� �����"

' ������ ����������� ������� ������
Public Const STEPS_VIEW = 0
Public Const CODE_VIEW = 1

' ������ ����������� ������� �����
Public Const TEXT_VIEW = 0
Public Const CHECKS_VIEW = 1

' ���� ��������
Public Const WPC_OPERATION_IDLE = 0 '// �������
Public Const WPC_OPERATION_FILL = 1 '// �����
Public Const WPC_OPERATION_DTRG = 2 '// ������
Public Const WPC_OPERATION_HEAT = 3 '// ������
Public Const WPC_OPERATION_WASH = 4 '// ������
Public Const WPC_OPERATION_RINS = 5 '// ����������
Public Const WPC_OPERATION_JOLT = 6 '// ����������
Public Const WPC_OPERATION_PAUS = 7 '// �����
Public Const WPC_OPERATION_DRAIN = 8 '// ����
Public Const WPC_OPERATION_SPIN = 9 '// �����
Public Const WPC_OPERATION_COOL = 10 '// ����������
Public Const WPC_OPERATION_TRIN = 11 '// ���.����������

' ���� ��������
Public Const LOADING_W_HOT = 0  ',    // ������ ������� ����
Public Const LOADING_W_COLD_1 = 1  ', // ������ �������� ���� 1
Public Const LOADING_W_COLD_2 = 2 ', // ������ �������� ���� 2
Public Const LOADING_WD_1 = 3 ',     // ������ �� 1
Public Const LOADING_WD_2 = 4 ',     // ������ �� 2
Public Const LOADING_WD_3 = 5 ',     // ������ �� 3
Public Const LOADING_WD_4 = 6 ',     // ������ �� 4
Public Const LOADING_WD_5 = 7 ',     // ������ �� 5
Public Const LOADING_WD_6 = 8 ',     // ������ �� 6
Public Const LOADING_WD_7 = 9 ',     // ������ �� 7
Public Const LOADING_WD_8 = 10 ',     // ������ �� 8
Public Const LOADING_WD_9 = 11 ',     // ������ �� 9
Public Const LOADING_LOCK_1 = 12 ',   // ����� ���� 1
Public Const LOADING_LOCK_2 = 13 ',   // ����� ���� 2
Public Const LOADING_PUMP_1 = 14 ',   // ���� 1
Public Const LOADING_PUMP_2 = 15 ',   // ���� 2
Public Const LOADING_HEAT = 16 '      // ������
Public Const LOADING_DRIVE = 17 '// ������

' *****************************************
' *  ���������� ����������
' *  ~~~~~~~~~~ ~~~~~~~~~~
' *****************************************

' ������� ������������ �����
Public Modified As Boolean
Public DesignMode As Boolean
Public FunctionsStrings(0 To 11) As String
Public Hook As Long, tMessage As Timer


