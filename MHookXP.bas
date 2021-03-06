Attribute VB_Name = "ModuleHookXP"
' *************************************************************************
'  Copyright �2009 Karl E. Peterson
'  All Rights Reserved, http://vb.mvps.org/
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' The *Subclass APIs in comctl32 were not exported by name until XP, and
' even in XP GetWindowSubclass remains exported only by ordinal.  All four
' functions first appeared in v4.71 of comctl32.dll, which shipped with
' Windows 98 and/or IE 4.01 - more details here:
' http://www.geoffchappell.com/studies/windows/shell/comctl32/history/ords472.htm
Private Declare Function SetWindowSubclass _
                Lib "comctl32" _
                Alias "#410" (ByVal hWnd As Long, _
                              ByVal pfnSubclass As Long, _
                              ByVal uIdSubclass As Long, _
                              ByVal dwRefData As Long) As Long

Private Declare Function GetWindowSubclass _
                Lib "comctl32" _
                Alias "#411" (ByVal hWnd As Long, _
                              ByVal pfnSubclass As Long, _
                              ByVal uIdSubclass As Long, _
                              pdwRefData As Long) As Long

Private Declare Function RemoveWindowSubclass _
                Lib "comctl32" _
                Alias "#412" (ByVal hWnd As Long, _
                              ByVal pfnSubclass As Long, _
                              ByVal uIdSubclass As Long) As Long

Private Declare Function DefSubclassProc _
                Lib "comctl32" _
                Alias "#413" (ByVal hWnd As Long, _
                              ByVal uMsg As Long, _
                              ByVal wParam As Long, _
                              ByVal lParam As Long) As Long

' RemoveWindowsHook must be called prior to destruction.
Private Const WM_NCDESTROY As Long = &H82

Public Function HookSet(ByVal hWnd As Long, _
                        ByVal Thing As IHookXP, _
                        Optional dwRefData As Long) As Boolean

    ' http://msdn.microsoft.com/en-us/library/bb762102(VS.85).aspx
    HookSet = CBool(SetWindowSubclass(hWnd, AddressOf SubclassProc, ObjPtr(Thing), _
            dwRefData))

End Function

Public Function HookGetData(ByVal hWnd As Long, ByVal Thing As IHookXP) As Long

    Dim dwRefData As Long

    ' http://msdn.microsoft.com/en-us/library/bb776430(VS.85).aspx
    If GetWindowSubclass(hWnd, AddressOf SubclassProc, ObjPtr(Thing), dwRefData) Then

        HookGetData = dwRefData

    End If

End Function

Public Function HookClear(ByVal hWnd As Long, ByVal Thing As IHookXP) As Boolean

    ' http://msdn.microsoft.com/en-us/library/bb762094(VS.85).aspx
    HookClear = CBool(RemoveWindowSubclass(hWnd, AddressOf SubclassProc, ObjPtr(Thing)))

End Function

Public Function HookDefault(ByVal hWnd As Long, _
                            ByVal uiMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long

    ' http://msdn.microsoft.com/en-us/library/bb776403(VS.85).aspx
    HookDefault = DefSubclassProc(hWnd, uiMsg, wParam, lParam)

End Function

Public Function SubclassProc(ByVal hWnd As Long, _
                             ByVal uiMsg As Long, _
                             ByVal wParam As Long, _
                             ByVal lParam As Long, _
                             ByVal uIdSubclass As IHookXP, _
                             ByVal dwRefData As Long) As Long

    ' http://msdn.microsoft.com/en-us/library/bb776774(VS.85).aspx
    SubclassProc = uIdSubclass.Message(hWnd, uiMsg, wParam, lParam, dwRefData)

    ' This should *never* be necessary, but just in case client fails to...
    If uiMsg = WM_NCDESTROY Then

        Call HookClear(hWnd, uIdSubclass)

    End If

End Function

