Attribute VB_Name = "UserFunctions"
'**
'@author <a href="mailto:unihomelab@ya.ru">Мезенцев Вячеслав</a>
'@revision Дата ревизии: 16.06.2011 г., время: 5:08:47
'@rem <h1><b>UserFunctions</b></h1>
'<pre>
'--------------------------------------------------------------------------------
' Проект   :       Конфигуратор управляющих программ
' Модуль   :       UserFunctions
' Описание :       Глобальные пользовательские функции
' Автор    :       Мезенцев Вячеслав
' Изменён  :       16.06.2011 г., время: 5:08:47
'--------------------------------------------------------------------------------
'</pre>
Option Explicit

' *****************************************
' *  ГЛОБАЛЬНЫЕ ФУНКЦИИ ПОЛЬЗОВАТЕЛЯ
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
        Print #iFF, sTrailer & _
                "New session....................................................."

    End If

    Print #iFF, sTrailer & sText

    Close #iFF

End Sub

' [VB] Как проверить, что код работает в IDE?
' Вы заводите в каком-нибудь модуле функцию со следующим кодом:

Public Function MakeTrue(ByRef bvar As Boolean) As Boolean

    bvar = True
    MakeTrue = True

End Function

' А затем, когда вам понадобится разместить код, который должен _
  по разному работать, делаете следующее:

'    Dim WE_ARE_IN_IDE As Boolean
'    Debug.Assert MakeTrue(WE_ARE_IN_IDE)
    
'    If WE_ARE_IN_IDE Then
'        MsgBox "Мы в IDE"
'        MakeSomethingBad
'    Else
'        MsgBox "Мы в скомпилированном файле"
'        MakeSomethingGood
'    End If

Public Function GetCellIndex(MSFlexGrid As MSFlexGrid, _
                             row As Integer, _
                             col As Integer) As Integer
    
    GetCellIndex = row * MSFlexGrid.Cols + col

End Function

Function FromUTF8(ByVal cnvUni As String) As String
    '<EhHeader>
    On Error GoTo FromUTF8_Err
    '</EhHeader>

    Dim cnvUni2 As String
    
    If cnvUni = vbNullString Then Exit Function

    cnvUni2 = WToA(cnvUni, CP_ACP)
    FromUTF8 = AToW(cnvUni2, CP_UTF8)

    '<EhFooter>
    Exit Function

FromUTF8_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.UserFunctions.FromUTF8]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Function

Public Function ToUTF8(srcStr As String) As String

    '<EhHeader>
    On Error GoTo ToUTF8_Err
    '</EhHeader>

    Dim utfStr() As Byte ' для исключения ошибок, связанных с дополнительной обработкой строк в VBA, используем не строку, а массив байтов

    Dim pAnsiStr As Long ' в API передаем не строки, а указатели на них
    Dim pUtfStr As Long
  
    Dim sLen As Long     ' просто переменная
  
    pAnsiStr = StrPtr(srcStr) '  указатель на строку
    ' определяем число байт, нужное для размещения строки в UTF-коде
    sLen = WideCharToMultiByte(CP_UTF8, 0, pAnsiStr, -1, 0, 0, 0, 0)

    If sLen > 0 Then

        ReDim utfStr(sLen) ' выделяем буфер для преобразования
        pUtfStr = VarPtr(utfStr(0)) ' указатель на этот буфер
    
        ' выполняем преобразование
        sLen = WideCharToMultiByte(CP_UTF8, 0, pAnsiStr, Len(srcStr), pUtfStr, sLen, 0, 0)

    End If
  
    If sLen > 0 Then

        ' почто-то финальный 0 переходит в результат - уберем его
        ReDim Preserve utfStr(sLen - 1)
        ToUTF8 = StrConv(utfStr, vbUnicode)
        ReDim utfStr(0)
    Else
        ToUTF8 = ""

    End If

    '<EhFooter>
    Exit Function

ToUTF8_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.UserFunctions.ToUTF8]: " & GetErrorMessageById(Err.Number, _
            Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Function

Public Function getGUID() As String
    '<EhHeader>
    On Error GoTo getGUID_Err
    '</EhHeader>

    Dim buffer(0 To 15) As Byte
    Dim s As String
    Dim Ret As Long

    s = String$(128, 0)

    ' получает численный код
    Ret = CoCreateGuid(buffer(0))
    
    ' преобразуем его в текст,
    ' используя недокументированную функцию StrPtr
    Ret = StringFromGUID2(buffer(0), StrPtr(s), 128)

    getGUID = Left$(s, Ret - 1) ' отсекаем "хвост"

    '<EhFooter>
    Exit Function

getGUID_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.UserFunctions.getGUID]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Function

Public Function VerifyFile(FileName$) As Boolean

  ' Проверка - существует ли указанный файл
  On Error Resume Next
  
  ' Файл открывается как выходной, последовательный
  Open FileName$ For Output As #1
  
  If Err Then ' Ошибка при открытии - нет файла
  
    VerifyFile = False
    
  Else
  
    VerifyFile = True: Close #1
    
  End If
  
End Function

Public Function DoesFileExist(ByVal strPath As String) As Boolean
    '<EhHeader>
    On Error GoTo DoesFileExist_Err
    '</EhHeader>

    DoesFileExist = PathFileExists(strPath)

    '<EhFooter>
    Exit Function

DoesFileExist_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.UserFunctions.DoesFileExist]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
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
Public Function MiscExtractPathName(strPath As String, ByVal bFlag As Boolean, Optional PathSeparator = PATH_SEPARATOR) As String
    '<EhHeader>
    On Error GoTo MiscExtractPathName_Err
    '</EhHeader>

    Dim lPos As Long
    Dim lOldPos As Long

    'Shorten the path one level'
    lPos = 1
    lOldPos = 1

    Do
        lPos = InStr(lPos, strPath, PathSeparator)

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

    '<EhFooter>
    Exit Function

MiscExtractPathName_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.UserFunctions.MiscExtractPathName]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Function

' Какие нагрузки могут включаться в зависимости от функции:
Public Function GetLoadingsFromFuncN(ByVal FuncN As Long) As Long
    '<EhHeader>
    On Error GoTo GetLoadingsFromFuncN_Err
    '</EhHeader>

    Dim Num As Long

    Select Case FuncN

            ' На шаге "налива" могут быть включены нагрузки: клапана ХВ1, ХВ2, ГВ, МОТОР
        Case WPC_OPERATION_FILL ' Налив

            Num = 2 ^ LOADING_W_COLD_1 Or 2 ^ LOADING_W_COLD_2 Or 2 ^ LOADING_W_HOT Or _
                    2 ^ LOADING_DRIVE
            
            ' На шаге "моющие" могут быть включены нагрузки: клапана МС1...МС9, МОТОР
        Case WPC_OPERATION_DTRG

            Num = 2 ^ LOADING_WD_1 Or 2 ^ LOADING_WD_2 Or 2 ^ LOADING_WD_3 Or 2 ^ _
                    LOADING_WD_4 Or 2 ^ LOADING_WD_5 Or 2 ^ LOADING_WD_6 Or 2 ^ _
                    LOADING_WD_7 Or 2 ^ LOADING_WD_8 Or 2 ^ LOADING_WD_9 Or 2 ^ _
                    LOADING_DRIVE
        
            ' На шаге "нагрева" могут быть включены нагрузки: ТЭН, МОТОР
        Case WPC_OPERATION_HEAT ' нагрев

            Num = 2 ^ LOADING_HEAT Or 2 ^ LOADING_DRIVE
        
            ' На шаге "стирки", "полоскания", "расстряски", "паузы" могут быть _
              включены нагрузки: МОТОР

        Case WPC_OPERATION_WASH ' стирка

            Num = 2 ^ LOADING_DRIVE
        
        Case WPC_OPERATION_RINS ' полоскание

            Num = 2 ^ LOADING_DRIVE
        
        Case WPC_OPERATION_JOLT ' расстряска

            Num = 2 ^ LOADING_DRIVE
        
        Case WPC_OPERATION_PAUS ' пауза

            Num = 2 ^ LOADING_DRIVE
        
            ' На шаге "слива" могут быть включены нагрузки: клапана СЛИВ1, СЛИВ2, МОТОР
        Case WPC_OPERATION_DRAIN ' слив

            Num = 2 ^ LOADING_PUMP_1 Or 2 ^ LOADING_PUMP_2 Or 2 ^ LOADING_DRIVE
        
            ' На шаге "отжима" могут быть включены нагрузки: клапана СЛИВ1, СЛИВ2, МОТОР
        Case WPC_OPERATION_SPIN ' отжим

            Num = 2 ^ LOADING_PUMP_1 Or 2 ^ LOADING_PUMP_2 Or 2 ^ LOADING_DRIVE
        
            ' На шаге "охлаждения" могут быть включены нагрузки: клапана ХВ1, МОТОР
        Case WPC_OPERATION_COOL ' охлаждение

            Num = 2 ^ LOADING_W_COLD_1 Or 2 ^ LOADING_DRIVE
        
        Case Else

    End Select

    GetLoadingsFromFuncN = Num

    '<EhFooter>
    Exit Function

GetLoadingsFromFuncN_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.UserFunctions.GetLoadingsFromFuncN]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Function

'
' Note that we swap the low order bytes in a long so that
' we don't have to worry about overflow problems
'
Public Function SwapInteger(ByVal I As Long) As Long

    SwapInteger = ((I \ &H100) And &HFF) Or ((I And &HFF) * &H100&)

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
    '<EhHeader>
    On Error GoTo HookKeyboard_Err
    '</EhHeader>

    Set tMessage = t1
    
    Hook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyboardProc, App.hInstance, _
            App.ThreadID)

    '<EhFooter>
    Exit Sub

HookKeyboard_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.UserFunctions.HookKeyboard]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Public Sub UnHookKeyboard()

    UnhookWindowsHookEx Hook

End Sub

Public Function KeyboardProc(ByVal ncode As Long, _
                             ByVal wParam As Long, _
                             ByVal lParam As Long) As Long
    
    Dim KeyCode As Long

    If ncode >= 0 Then

        KeyCode = wParam And &HFF&
        
        If (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = _
                vbKeyUp Or KeyCode = vbKeyDown) And (lParam And &H80000000) = 0 Then
            
            ' at this point we have detected a PgUp or PgDown "key down"
            ' event, so now we need to check the state of the Ctrl key
            
            'ShiftDown = (GetAsyncKeyState(vbKeyShift) And &H8000&) <> 0
            
            tMessage.Tag = CInt(KeyCode)
            tMessage.Interval = 10 ' ) one shot
            tMessage.Interval = 0  ' ) timer
                
            ' now "eat" the key
            'KeyboardProc = -1
            'Else
            ' allow the key to be processed as normal
            KeyboardProc = CallNextHookEx(Hook, ncode, wParam, ByVal lParam)

        End If

    Else
    
        KeyboardProc = CallNextHookEx(Hook, ncode, wParam, ByVal lParam)

    End If

End Function

'**
'@param        ErrNum Required. Long.
'@param        ErrDescription Required. String.
'@return       String.
'@rem <h2>GetErrorMessageById</h2>
'Функция возвращает описание исключительной ситуации по номеру ошибки.
Public Function GetErrorMessageById(ErrNum As Long, ErrDescription As String) As String
    
    Dim ТекстОшибки As String
    
    ТекстОшибки = "(" & CStr(ErrNum) & ") "

    Select Case ErrNum
        
            'Определённые пользователем ошибки
        Case ОШИБКА_НЕИЗВЕСТНАЯ:

            ТекстОшибки = ТекстОшибки & "Неизвестная ошибка"
        
            'Системные ошибки
        Case Else

            ТекстОшибки = ТекстОшибки & ErrDescription
            
    End Select

    GetErrorMessageById = ТекстОшибки

End Function

Public Function GetKeyValue(KeyRoot As Long, _
                            KeyName As String, _
                            SubKeyRef As String, _
                            ByRef KeyVal As String) As Boolean

    Dim I As Long                                           ' Loop Counter
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
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
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

            For I = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit

                KeyVal = KeyVal + Hex$(Asc(Mid$(tmpVal, I, 1)))   ' Build Value Char. By Char.
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
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, _
            SysInfoPath) Then

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
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.FormAbout.StartSysInfo]: " & GetErrorMessageById(Err.Number, _
            Err.Description), VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
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
    Dim I                   As Integer
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
    HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + bytebuffer(0) * &H10000 + _
            bytebuffer(1) * &H1000000
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
    For I = 2 To 2
    '</Modified by: Project Administrator at 7.10.2011-21:01:35 on machine: ALPHA>
    
        buffer = String(255, 0)
        strTemp = "\StringFileInfo\" & Lang_Charset_String & "\" & strVersionInfo(I)
        lRet = VerQueryValue(sBuffer(0), strTemp, lVerPointer, lBufferLen)

        If lRet = 0 Then

            GetFileVersionInformation = eNoVersion
            Exit Function

        End If
        
        lstrcpy buffer, lVerPointer
        buffer = Mid$(buffer, 1, InStr(buffer, vbNullChar) - 1)

        Select Case I

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

    Next I

    GetFileVersionInformation = eOK
    
    '<EhFooter>
    Exit Function

GetFileVersionInformation_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.UserFunctions.GetFileVersionInformation]: " & _
            GetErrorMessageById(Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation
    Resume Next

    '</EhFooter>
End Function

Public Function LoadFromJSONFile(sFilePath As String) As String
    '<EhHeader>
    On Error GoTo LoadFromJSONFile_Err
    '</EhHeader>

    Dim handle As Integer ' Идентификатор файла

    If LenB(Dir$(sFilePath)) > 0 Then
    
        ' Получаем свободный идентификатор
        handle = FreeFile

        ' Получаем доступ к файлу
        Open sFilePath For Binary As #handle

        LoadFromJSONFile = Space$(LOF(handle))

        Get #handle, , LoadFromJSONFile

        ' Завершаем работу с файлом
        Close #handle
        
    End If
    
    '<EhFooter>
    Exit Function

LoadFromJSONFile_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.UserFunctions.LoadFromJSONFile]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Function

Public Sub SaveToJSONFile(ByVal FileName As String, ByVal NewValue As String)
    '<EhHeader>
    On Error GoTo SaveToJSONFile_Err
    '</EhHeader>
    
    Dim handle As Integer ' Идентификатор файла

    ' Получаем свободный идентификатор
    handle = FreeFile

    ' Получаем доступ к файлу
    Open FileName For Output As #handle

    ' Выводим текст в файл
    Print #handle, ToUTF8(NewValue)

    ' Завершаем работу с файлом
    Close #handle
    
    '<EhFooter>
    Exit Sub

SaveToJSONFile_Err:
    App.LogEvent "" & VBA.Constants.vbCrLf & Date & " " & Time & _
            " [INFO] [cop.UserFunctions.SaveToJSONFile]: " & GetErrorMessageById( _
            Err.Number, Err.Description), _
            VBRUN.LogEventTypeConstants.vbLogEventTypeInformation

    Resume Next

    '</EhFooter>
End Sub

Public Function GetOSVersion() As String

    ' Windows Version Data:
    '
    ' To determine the operating system that is running on a given system,
    ' the following data is needed:
    '
    ' Win 95 Win 98 WinME WinNT 4 Win2000 Win XP
    ' PlatformID 1 1 1 2 2 2
    ' Major Ver 4 4 4 4 5 5
    ' Minor Ver 0 10 90 0 0 1
    '
    '
    Dim OSInfo As OSVERSIONINFO
    Dim PId As String
    Dim Ret As Long

    ' set the structure size
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    
    ' get the Windows version
    Ret = GetVersionEx(OSInfo)

    ' check for errors
    If Ret = 0 Then MsgBox "Error Getting Version Information": Exit Function

    Select Case OSInfo.dwPlatformId

        Case 0

            PId = "Pre-Windows 95"

            ' 95/98/ME
        Case 1

            Select Case OSInfo.dwMinorVersion

                Case 0

                    ' Windows 95
                    PId = "Windows 95"

                Case 10

                    ' Windows 98
                    PId = "Windows 98"

                Case 90

                    ' Windows ME
                    PId = "Windows ME"

            End Select

            ' NT/2000/XP
        Case 2

            Select Case OSInfo.dwMajorVersion

                Case 4

                    ' NT version
                    PId = "Windows NT"

                Case 5

                    ' 2000/XP version
                    Select Case OSInfo.dwMinorVersion

                        Case 0

                            ' Windows 2000
                            PId = "Windows 2000"

                        Case 1

                            ' Windows XP
                            PId = "Windows XP"

                    End Select

            End Select

    End Select

    GetOSVersion = PId & " version" & str$(OSInfo.dwMajorVersion) & "." & LTrim(str( _
            OSInfo.dwMinorVersion)) & " build " & str(OSInfo.dwBuildNumber)
 
End Function
