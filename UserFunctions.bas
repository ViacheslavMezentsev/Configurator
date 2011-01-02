Attribute VB_Name = "UserFunctions"
Option Explicit

' *****************************************
' *  ГЛОБАЛЬНЫЕ ФУНКЦИИ ПОЛЬЗОВАТЕЛЯ
' *  ~~~~~~~~~~ ~~~~~~~ ~~~~~~~~~~~~
' *****************************************

Public Function GetCellIndex(MSFlexGrid As MSFlexGrid, _
    row As Integer, col As Integer)
    
    GetCellIndex = row * MSFlexGrid.Cols + col
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
        
        Case WPC_OPERATION_TRIN ' тех.полоскание
            Num = 2 ^ LOADING_DRIVE
        
        Case Else
    End Select

    GetLoadingsFromFuncN = Num
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

