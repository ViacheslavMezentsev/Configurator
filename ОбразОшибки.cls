VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "�����������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"�����������"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**
'@author <a href="mailto:unihomelab@ya.ru">�������� ��������</a>
'@revision ���� �������: 16.06.2011 �., �����: 16:02:48
'@rem <h1><b>�����������</b></h1>
'�����-������ ��� �������� ���������� �� ������.
'<pre>
'--------------------------------------------------------------------------------
' ������   :       ������������ ����������� ��������
' ������   :       �����������
' �������� :       �����-������ ��� �������� ���������� �� ������
' �����    :       �������� ��������
' �������  :       16.06.2011 �., �����: 16:02:48
'--------------------------------------------------------------------------------
'</pre>
Option Explicit

'**
'@rem
Public ����������� As Long
'**
'@rem
Public ��������� As String

'**
'@rem �����������
Private Sub Class_Initialize()
End Sub

'**
'@rem ����������
Private Sub Class_Terminate()
End Sub

'**
'@param        �������� Required. ����������� object.
'@rem ����������� �����������.
Public Sub �������(ByVal �������� As �����������)
    Me.����������� = ��������.�����������
    Me.��������� = ��������.���������
End Sub
