Attribute VB_Name = "UserFunctions"
Option Explicit

' *****************************************
' *  ГЛОБАЛЬНЫЕ ФУНКЦИИ ПОЛЬЗОВАТЕЛЯ
' *  ~~~~~~~~~~ ~~~~~~~ ~~~~~~~~~~~~
' *****************************************

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
        'тут, наверное, тоже что-то надо сделать?
    End If
End Function

Public Function ToUTF8(srcStr As String) As String
  Dim utfStr() As Byte ' для исключения ошибок, связанных с дополнительной обработкой строк в VBA, _
                         используем не строку, а массив байтов
  Dim pAnsiStr As Long ' в API передаем не строки, а указатели на них
  Dim pUtfStr As Long
  
  Dim sLen As Long     ' просто переменная
  
  pAnsiStr = StrPtr(srcStr) '  указатель на строку
  ' определяем число байт, нужное для размещения строки в UTF-коде
  sLen = WideCharToMultiByte(CP_UTF8, _
                             0, _
                             pAnsiStr, _
                             -1, _
                             0, _
                             0, _
                             0, _
                             0)
  If sLen > 0 Then
    ReDim utfStr(sLen) ' выделяем буфер для преобразования
    pUtfStr = VarPtr(utfStr(0)) ' указатель на этот буфер
    
    ' выполняем преобразование
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
    ' почто-то финальный 0 переходит в результат - уберем его
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

    ' получает численный код
    ret = CoCreateGuid(Buffer(0))
    
    ' преобразуем его в текст,
    ' используя недокументированную функцию StrPtr
    ret = StringFromGUID2(Buffer(0), StrPtr(s), 128)

    getGUID = Left$(s, ret - 1) ' отсекаем "хвост"
End Function

Public Function DoesFileExist(ByVal strPath As String) As Boolean
    DoesFileExist = PathFileExists(strPath)
End Function

Public Function MiscExtractPathName(strPath As String, ByVal bFlag) As String
    'The string is treated as if it contains                   '
    'a path and file name.                                     '
    ''''''''''''''''''''''''''''''­'''''''''''''''''''''''''''''
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

' Какие нагрузки могут включаться в зависимости от функции:
Public Function GetLoadingsFromFuncN(ByVal FuncN As Long) As Long
    Dim Num As Long
    Select Case FuncN
        ' На шаге "налива" могут быть включены нагрузки: клапана ХВ1, ХВ2, ГВ, МОТОР
        Case WPC_OPERATION_FILL ' Налив
            Num = 2 ^ LOADING_W_COLD_1 _
            Or 2 ^ LOADING_W_COLD_2 _
            Or 2 ^ LOADING_W_HOT _
            Or 2 ^ LOADING_DRIVE
            
        ' На шаге "моющие" могут быть включены нагрузки: клапана МС1...МС9, МОТОР
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
        
        ' На шаге "нагрева" могут быть включены нагрузки: ТЭН, МОТОР
        Case WPC_OPERATION_HEAT ' нагрев
            Num = 2 ^ LOADING_HEAT _
            Or 2 ^ LOADING_DRIVE
        
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
            Num = 2 ^ LOADING_PUMP_1 _
            Or 2 ^ LOADING_PUMP_2 _
            Or 2 ^ LOADING_DRIVE
        
        ' На шаге "отжима" могут быть включены нагрузки: клапана СЛИВ1, СЛИВ2, МОТОР
        Case WPC_OPERATION_SPIN ' отжим
            Num = 2 ^ LOADING_PUMP_1 _
            Or 2 ^ LOADING_PUMP_2 _
            Or 2 ^ LOADING_DRIVE
        
        ' На шаге "охлаждения" могут быть включены нагрузки: клапана ХВ1, МОТОР
        Case WPC_OPERATION_COOL ' охлаждение
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

