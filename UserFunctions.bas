Attribute VB_Name = "UserFunctions"
Option Explicit

' *****************************************
' *  ���������� ������� ������������
' *  ~~~~~~~~~~ ~~~~~~~ ~~~~~~~~~~~~
' *****************************************

' [VB] ��� ���������, ��� ��� �������� � IDE?
' �� �������� � �����-������ ������ ������� �� ��������� �����:

Public Function MakeTrue(ByRef bvar As Boolean) As Boolean
    bvar = True
    MakeTrue = True
End Function

' � �����, ����� ��� ����������� ���������� ���, ������� ������ _
�� ������� ��������, ������� ���������:

'    Dim WE_ARE_IN_IDE As Boolean
'    Debug.Assert MakeTrue(WE_ARE_IN_IDE)
    
'    If WE_ARE_IN_IDE Then
'        MsgBox "�� � IDE"
'        MakeSomethingBad
'    Else
'        MsgBox "�� � ���������������� �����"
'        MakeSomethingGood
'    End If

Public Function GetCellIndex(MSFlexGrid As MSFlexGrid, _
    row As Integer, col As Integer)
    
    GetCellIndex = row * MSFlexGrid.Cols + col
End Function

Public Function FromUTF8(instring As String) As String
    Dim iStrSize    As Long
    Dim s           As String
    
    iStrSize = Len(instring)
    s = String$(iStrSize, 0&)

    If iStrSize Then iStrSize = MultiByteToWideChar(CP_UTF8, 0&, instring, &HFFFF, StrPtr(s), iStrSize)
    If iStrSize >= 0 Then
        FromUTF8 = Left$(s, iStrSize - 1)
    Else
        '���, ��������, ���� ���-�� ���� �������?
    End If
End Function

Public Function ToUTF8(srcStr As String) As String
  Dim utfStr() As Byte ' ��� ���������� ������, ��������� � �������������� ���������� ����� � VBA, _
                         ���������� �� ������, � ������ ������
  Dim pAnsiStr As Long ' � API �������� �� ������, � ��������� �� ���
  Dim pUtfStr As Long
  
  Dim sLen As Long     ' ������ ����������
  
  pAnsiStr = StrPtr(srcStr) '  ��������� �� ������
  ' ���������� ����� ����, ������ ��� ���������� ������ � UTF-����
  sLen = WideCharToMultiByte(CP_UTF8, _
                             0, _
                             pAnsiStr, _
                             -1, _
                             0, _
                             0, _
                             0, _
                             0)
  If sLen > 0 Then
    ReDim utfStr(sLen) ' �������� ����� ��� ��������������
    pUtfStr = VarPtr(utfStr(0)) ' ��������� �� ���� �����
    
    ' ��������� ��������������
    sLen = WideCharToMultiByte(CP_UTF8, _
                             0, _
                             pAnsiStr, _
                             Len(srcStr), _
                             pUtfStr, _
                             sLen, _
                             0, _
                             0)
  End If
  
  If sLen > 0 Then
    ' �����-�� ��������� 0 ��������� � ��������� - ������ ���
    ReDim Preserve utfStr(sLen - 1)
    ToUTF8 = StrConv(utfStr, vbUnicode)
    ReDim utfStr(0)
  Else
    ToUTF8 = ""
  End If
End Function

Public Function getGUID() As String
    Dim Buffer(0 To 15) As Byte
    Dim s As String
    Dim ret As Long

    s = String$(128, 0)

    ' �������� ��������� ���
    ret = CoCreateGuid(Buffer(0))
    
    ' ����������� ��� � �����,
    ' ��������� ������������������� ������� StrPtr
    ret = StringFromGUID2(Buffer(0), StrPtr(s), 128)

    getGUID = Left$(s, ret - 1) ' �������� "�����"
End Function

Public Function DoesFileExist(ByVal strPath As String) As Boolean
    DoesFileExist = PathFileExists(strPath)
End Function

Public Function MiscExtractPathName(strPath As String, ByVal bFlag) As String
    'The string is treated as if it contains                   '
    'a path and file name.                                     '
    ''''''''''''''''''''''''''''''�'''''''''''''''''''''''''''''
    ' If bFlag = TRUE:                                         '
    '                   Function extracts the path from        '
    '                   the input string and returns it.       '
    ' If bFlag = FALSE:                                        '
    '                   Function extracts the File name from   '
    '                   the input string and returns it.       '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim lPos As Long
    Dim lOldPos As Long
    'Shorten the path one level'
    lPos = 1
    lOldPos = 1
    Do
        lPos = InStr(lPos, strPath, PATH_SEPARATOR)
        If lPos > 0 Then
            lOldPos = lPos
            lPos = lPos + 1
        Else
            If lOldPos = 1 And Not bFlag Then
                lOldPos = 0
            End If
            Exit Do
        End If
    Loop
    If bFlag Then
        MiscExtractPathName = Left$(strPath, lOldPos - 1)
    Else
        MiscExtractPathName = Mid$(strPath, lOldPos + 1)
    End If
End Function

' ����� �������� ����� ���������� � ����������� �� �������:
Public Function GetLoadingsFromFuncN(ByVal FuncN As Long) As Long
    Dim Num As Long
    Select Case FuncN
        ' �� ���� "������" ����� ���� �������� ��������: ������� ��1, ��2, ��, �����
        Case WPC_OPERATION_FILL ' �����
            Num = 2 ^ LOADING_W_COLD_1 _
            Or 2 ^ LOADING_W_COLD_2 _
            Or 2 ^ LOADING_W_HOT _
            Or 2 ^ LOADING_DRIVE
            
        ' �� ���� "������" ����� ���� �������� ��������: ������� ��1...��9, �����
        Case WPC_OPERATION_DTRG
            Num = 2 ^ LOADING_WD_1 _
            Or 2 ^ LOADING_WD_2 _
            Or 2 ^ LOADING_WD_3 _
            Or 2 ^ LOADING_WD_4 _
            Or 2 ^ LOADING_WD_5 _
            Or 2 ^ LOADING_WD_6 _
            Or 2 ^ LOADING_WD_7 _
            Or 2 ^ LOADING_WD_8 _
            Or 2 ^ LOADING_WD_9 _
            Or 2 ^ LOADING_DRIVE
        
        ' �� ���� "�������" ����� ���� �������� ��������: ���, �����
        Case WPC_OPERATION_HEAT ' ������
            Num = 2 ^ LOADING_HEAT _
            Or 2 ^ LOADING_DRIVE
        
        ' �� ���� "������", "����������", "����������", "�����" ����� ���� _
        �������� ��������: �����
        Case WPC_OPERATION_WASH ' ������
            Num = 2 ^ LOADING_DRIVE
        
        Case WPC_OPERATION_RINS ' ����������
            Num = 2 ^ LOADING_DRIVE
        
        Case WPC_OPERATION_JOLT ' ����������
            Num = 2 ^ LOADING_DRIVE
        
        Case WPC_OPERATION_PAUS ' �����
            Num = 2 ^ LOADING_DRIVE
        
        ' �� ���� "�����" ����� ���� �������� ��������: ������� ����1, ����2, �����
        Case WPC_OPERATION_DRAIN ' ����
            Num = 2 ^ LOADING_PUMP_1 _
            Or 2 ^ LOADING_PUMP_2 _
            Or 2 ^ LOADING_DRIVE
        
        ' �� ���� "������" ����� ���� �������� ��������: ������� ����1, ����2, �����
        Case WPC_OPERATION_SPIN ' �����
            Num = 2 ^ LOADING_PUMP_1 _
            Or 2 ^ LOADING_PUMP_2 _
            Or 2 ^ LOADING_DRIVE
        
        ' �� ���� "����������" ����� ���� �������� ��������: ������� ��1, �����
        Case WPC_OPERATION_COOL ' ����������
            Num = 2 ^ LOADING_W_COLD_1 _
            Or 2 ^ LOADING_DRIVE
        
        Case Else
    End Select

    GetLoadingsFromFuncN = Num
End Function

'
' Note that we swap the low order bytes in a long so that
' we don't have to worry about overflow problems
'
Public Function SwapInteger(ByVal i As Long) As Long
    SwapInteger = ((i \ &H100) And &HFF) Or ((i And &HFF) * &H100&)
End Function

'
' Swap a long value from Motorola to Intel format or vice versa
'
Public Function SwapLong(ByVal l As Long) As Long
    Dim addbit%
    Dim newlow&, newhigh&

    newlow& = l \ &H10000
    newlow& = SwapInteger(newlow& And &HFFFF&)

    newhigh& = SwapInteger(l And &HFFFF&)
    If newhigh& And &H8000& Then
        ' This would overflow
        newhigh& = newhigh And &H7FFF
        addbit% = True
    End If
    newhigh& = (newhigh& * &H10000) Or newlow&
    If addbit% Then newhigh = newhigh Or &H80000000
    SwapLong = newhigh&

End Function

Public Sub HookKeyboard(t1 As Timer)
    Set tMessage = t1
    
    Hook = SetWindowsHookEx(WH_KEYBOARD, _
        AddressOf KeyboardProc, App.hInstance, App.ThreadID)
End Sub

Public Sub UnHookKeyboard()
    UnhookWindowsHookEx Hook
End Sub

Public Function KeyboardProc(ByVal ncode As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
Dim keyCode As Long, CtrlDown As Boolean
    If ncode >= 0 Then
        keyCode = wParam And &HFF&
        
        If (keyCode = VBRUN.KeyCodeConstants.vbKeyLeft Or _
            keyCode = VBRUN.KeyCodeConstants.vbKeyRight Or _
            keyCode = VBRUN.KeyCodeConstants.vbKeyUp Or _
            keyCode = VBRUN.KeyCodeConstants.vbKeyDown) _
            And (lParam And &H80000000) = 0 Then
            
            ' at this point we have detected a PgUp or PgDown "key down"
            ' event, so now we need to check the state of the Ctrl key
            
            'CtrlDown = (GetAsyncKeyState(vbKeyControl) And &H8000&) <> 0
            
            'If CtrlDown Then
                ' fire our Timer to indicate a Ctrl/PgDown or Ctrl/PgUp
                tMessage.Tag = CInt(keyCode)
                tMessage.Interval = 10 ' ) one shot
                tMessage.Interval = 0  ' ) timer
                
                ' now "eat" the key
                'KeyboardProc = -1
        'Else
            ' allow the key to be processed as normal
            KeyboardProc = CallNextHookEx _
                (Hook, ncode, wParam, ByVal lParam)
        End If
    Else
        KeyboardProc = CallNextHookEx _
            (Hook, ncode, wParam, ByVal lParam)
    End If
End Function

