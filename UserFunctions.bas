Attribute VB_Name = "UserFunctions"
'**
'@author <a href="mailto:unihomelab@ya.ru">�������� ��������</a>
'@revision ���� �������: 16.06.2011 �., �����: 5:08:47
'@rem <h1><b>UserFunctions</b></h1>
'<pre>
'--------------------------------------------------------------------------------
' ������   :       ������������ ����������� ��������
' ������   :       UserFunctions
' �������� :       ���������� ���������������� �������
' �����    :       �������� ��������
' ������  :       16.06.2011 �., �����: 5:08:47
'--------------------------------------------------------------------------------
'</pre>
Option Explicit

' *****************************************
' *  ���������� ������� ������������
' *  ~~~~~~~~~~ ~~~~~~~ ~~~~~~~~~~~~
' *****************************************

Public Sub Dec(ByRef Variable As Long, Optional Amount As Long = 1)
    Variable = Variable - Amount
End Sub

Public Sub Inc(ByRef Variable As Long, Optional Amount As Long = 1)
    Variable = Variable + Amount
End Sub

Public Sub SetModified(Value As Boolean)
    Modified = Value
End Sub

'CSEH: ErrResumeNext
Public Sub TxtLog(sText As String, Optional bNoDateTime As Boolean = False)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

    ' This routine is provided to be used in conjunction with the ErrReportAndTrace error handling scheme
    ' as well as for any other tasks that require logging.

    Dim iFF%, sTrailer$
    Static bNewSession As Boolean

    sTrailer = ""

    If Not bNoDateTime Then sTrailer = Date & " - " & Time & " --- "
    
    iFF = FreeFile
    Open App.Path & "\Log.txt" For Append As #iFF
        
    If Not bNewSession Then

        bNewSession = True
        Print #iFF, sTrailer & "New session....................................................."

    End If

    Print #iFF, sTrailer & sText

    Close #iFF

End Sub

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
       row As Integer, col As Integer) As Integer
    
    GetCellIndex = row * MSFlexGrid.Cols + col
End Function

Function FromUTF8(ByVal cnvUni As String) As String
    Dim cnvUni2 As String
    
    If cnvUni = vbNullString Then Exit Function
    cnvUni2 = WToA(cnvUni, CP_ACP)
    FromUTF8 = AToW(cnvUni2, CP_UTF8)
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
    Dim buffer(0 To 15) As Byte
    Dim s As String
    Dim ret As Long

    s = String$(128, 0)

    ' �������� ��������� ���
    ret = CoCreateGuid(buffer(0))
    
    ' ����������� ��� � �����,
    ' ��������� ������������������� ������� StrPtr
    ret = StringFromGUID2(buffer(0), StrPtr(s), 128)

    getGUID = Left$(s, ret - 1) ' �������� "�����"
End Function

Public Function DoesFileExist(ByVal strPath As String) As Boolean
    DoesFileExist = PathFileExists(strPath)
End Function

'**
'@param strPath
'@param bFlag
'@rem <h2>MiscExtractPathName</h2>
'The string is treated as if it contains a path and file name.
'<pre>
' If bFlag = TRUE:
'                   Function extracts the path from
'                   the input string and returns it.
' If bFlag = FALSE:
'                   Function extracts the File name from
'                   the input string and returns it.
'</pre>
Public Function MiscExtractPathName(strPath As String, ByVal bFlag As Boolean) As String

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
    
    Dim KeyCode As Long

    If ncode >= 0 Then
        KeyCode = wParam And &HFF&
        
        If (KeyCode = VBRUN.KeyCodeConstants.vbKeyLeft Or _
           KeyCode = VBRUN.KeyCodeConstants.vbKeyRight Or _
           KeyCode = VBRUN.KeyCodeConstants.vbKeyUp Or _
           KeyCode = VBRUN.KeyCodeConstants.vbKeyDown) _
           And (lParam And &H80000000) = 0 Then
            
            ' at this point we have detected a PgUp or PgDown "key down"
            ' event, so now we need to check the state of the Ctrl key
            
            'CtrlDown = (GetAsyncKeyState(vbKeyControl) And &H8000&) <> 0
            
            'If CtrlDown Then
            ' fire our Timer to indicate a Ctrl/PgDown or Ctrl/PgUp
            tMessage.Tag = CInt(KeyCode)
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

'**
'@param        ErrNum Required. Long.
'@param        ErrDescription Required. String.
'@return       String.
'@rem <h2>GetErrorMessageById</h2>
'������� ���������� �������� �������������� �������� �� ������ ������.
Public Function GetErrorMessageById(ErrNum As Long, ErrDescription As String) As String
    
    Dim ����������� As String
    
    ����������� = "(" & CStr(ErrNum) & ") "

    Select Case ErrNum
        
        '����������� ������������� ������
        Case ������_�����������:
            ����������� = ����������� & "����������� ������"
        
        '��������� ������
        Case Else
            ����������� = ����������� & ErrDescription
            
    End Select

    GetErrorMessageById = �����������
End Function

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
       KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid$(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left$(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left$(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------

    Select Case KeyValType                                  ' Search Data Types...
        Case REG_SZ                                             ' String Registry Key Data Type
            KeyVal = tmpVal                                     ' Copy String Value
        Case REG_DWORD                                          ' Double Word Registry Key Data Type

            For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                KeyVal = KeyVal + Hex$(Asc(Mid$(tmpVal, i, 1)))   ' Build Value Char. By Char.
            Next
            KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:              ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Public Sub StartSysInfo()
    '<EhHeader>
    On Error GoTo StartSysInfo_Err
    '</EhHeader>
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...

    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version

        If (Dir$(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
            ' Error - File Can Not Be Found...
        Else
            GoTo StartSysInfo_Err
        End If
        ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo StartSysInfo_Err
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    '<EhFooter>
    Exit Sub

StartSysInfo_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.FormAbout.StartSysInfo]: " _
       & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Sub

'**************************************
' Name: Get Version Number for EXE, DLL or OCX files
' Description:This function will retrieve the version number, product name, original program name (like if you right click on the EXE file and select properties, then select Version tab, it shows you all that information) etc
' By: Serge
'
' Returns:FileInfo structure
'
' Assumes:Label (named Label1 and make it wide enough, also increase the height of the label to have size of the form), Common Dilaog Box (CommonDialog1) and a Command Button (Command1)
'
'This code is copyrighted and has' limited warranties.Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=4976&lngWId=1'for details.'**************************************

Public Function GetFileVersionInformation(ByRef pstrFieName As String, _
                                          ByRef tFileInfo As FILEINFO) As VerisonReturnValue
    '<EhHeader>
    On Error GoTo GetFileVersionInformation_Err
    '</EhHeader>
    
    Dim lBufferLen          As Long, lDummy As Long
    Dim sBuffer()           As Byte
    Dim lVerPointer         As Long
    Dim lRet                As Long
    Dim Lang_Charset_String As String
    Dim HexNumber           As Long
    Dim i                   As Integer
    Dim strTemp             As String
    
    'Clear the Buffer tFileInfo
    tFileInfo.CompanyName = ""
    tFileInfo.FileDescription = ""
    tFileInfo.FileVersion = ""
    tFileInfo.InternalName = ""
    tFileInfo.LegalCopyright = ""
    tFileInfo.OriginalFileName = ""
    tFileInfo.ProductName = ""
    tFileInfo.ProductVersion = ""
    
    lBufferLen = GetFileVersionInfoSize(pstrFieName, lDummy)

    If lBufferLen < 1 Then
        GetFileVersionInformation = eNoVersion
        Exit Function
    End If
    
    ReDim sBuffer(lBufferLen)
    
    lRet = GetFileVersionInfo(pstrFieName, 0&, lBufferLen, sBuffer(0))

    If lRet = 0 Then
        GetFileVersionInformation = eNoVersion
        Exit Function
    End If
    
    lRet = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lVerPointer, lBufferLen)

    If lRet = 0 Then
        GetFileVersionInformation = eNoVersion
        Exit Function
    End If
    
    Dim bytebuffer(255) As Byte
    
    MoveMemory bytebuffer(0), lVerPointer, lBufferLen
    HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
    Lang_Charset_String = Hex(HexNumber)

    'Pull it all apart:
    '04------= SUBLANG_ENGLISH_USA
    '--09----= LANG_ENGLISH
    ' ----04E4 = 1252 = Codepage for Windows:Multilingual
    Do While Len(Lang_Charset_String) < 8
        Lang_Charset_String = "0" & Lang_Charset_String
    Loop
    
    Dim strVersionInfo(7) As String
    
    strVersionInfo(0) = "CompanyName"
    strVersionInfo(1) = "FileDescription"
    strVersionInfo(2) = "FileVersion"
    strVersionInfo(3) = "InternalName"
    strVersionInfo(4) = "LegalCopyright"
    strVersionInfo(5) = "OriginalFileName"
    strVersionInfo(6) = "ProductName"
    strVersionInfo(7) = "ProductVersion"
    
    Dim buffer As String

'<Modified by: Project Administrator at 7.10.2011-21:01:35 on machine: ALPHA>
    For i = 2 To 2
'</Modified by: Project Administrator at 7.10.2011-21:01:35 on machine: ALPHA>
        buffer = String(255, 0)
        strTemp = "\StringFileInfo\" & Lang_Charset_String & "\" & strVersionInfo(i)
        lRet = VerQueryValue(sBuffer(0), strTemp, lVerPointer, lBufferLen)

        If lRet = 0 Then
            GetFileVersionInformation = eNoVersion
            Exit Function
        End If
        
        lstrcpy buffer, lVerPointer
        buffer = Mid$(buffer, 1, InStr(buffer, vbNullChar) - 1)

        Select Case i

            Case 0
                tFileInfo.CompanyName = buffer

            Case 1
                tFileInfo.FileDescription = buffer

            Case 2
                tFileInfo.FileVersion = buffer

            Case 3
                tFileInfo.InternalName = buffer

            Case 4
                tFileInfo.LegalCopyright = buffer

            Case 5
                tFileInfo.OriginalFileName = buffer

            Case 6
                tFileInfo.ProductName = buffer

            Case 7
                tFileInfo.ProductVersion = buffer
        End Select
    Next i

    GetFileVersionInformation = eOK
    
    '<EhFooter>
    Exit Function

GetFileVersionInformation_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & " [INFO] [cop.UserFunctions.GetFileVersionInformation]: " _
        & GetErrorMessageById(Err.Number, Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next
    '</EhFooter>
End Function
