VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHookMouseWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************************
'  Copyright ©1997-2009 Karl E. Peterson
'  All Rights Reserved, http://vb.mvps.org/
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' Win32 API Declarations
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       Source As Any, _
                                       ByVal length As Long)

Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SystemParametersInfo _
                Lib "user32" _
                Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                               ByVal uParam As Long, _
                                               lpvParam As Any, _
                                               ByVal fuWinIni As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function WindowFromPoint _
                Lib "user32" (ByVal xPoint As Long, _
                              ByVal yPoint As Long) As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

' Mousewheel constants and data structures
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_XBUTTONDOWN As Long = &H20B         '(_WIN32_WINNT >= 0x0500)
Private Const WM_XBUTTONUP As Long = &H20C           '(_WIN32_WINNT >= 0x0500)
Private Const WM_XBUTTONDBLCLK As Long = &H20D       '(_WIN32_WINNT >= 0x0500)
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_MOUSEHWHEEL As Long = &H20E         '(_WIN32_WINNT >= 0x0600)
Private Const SM_MOUSEWHEELPRESENT As Long = 75
Private Const SPI_GETWHEELSCROLLLINES As Long = 104
Private Const SPI_SETWHEELSCROLLLINES As Long = 105
Private Const SPI_GETWHEELSCROLLCHARS As Long = 108  '(_WIN32_WINNT >= 0x0600)
Private Const SPI_SETWHEELSCROLLCHARS As Long = 109  '(_WIN32_WINNT >= 0x0600)

Private Const WHEEL_PAGESCROLL As Long = -1        ' (UINT_MAX) /* Scroll one page */
Private Const WHEEL_DELTA As Long = 120            ' /* Value for rolling one detent */

Private Type POINTAPI

    X As Long
    Y As Long

End Type

' Key State Masks for Mouse Messages
Private Const MK_LBUTTON As Long = &H1
Private Const MK_RBUTTON As Long = &H2
Private Const MK_SHIFT As Long = &H4
Private Const MK_CONTROL As Long = &H8
Private Const MK_MBUTTON As Long = &H10

' Key State Masks for GetKeyState function
Private Const VK_LBUTTON As Long = &H1
Private Const VK_RBUTTON As Long = &H2
Private Const VK_MBUTTON As Long = &H4             '  NOT contiguous with L RBUTTON
Private Const VK_SHIFT As Long = &H10
Private Const VK_CONTROL As Long = &H11
Private Const VK_MENU As Long = &H12

' Subclassing interface
Implements IHookXP

' Events
Public Event MouseWheel(ByVal hWnd As Long, _
                        ByVal Delta As Long, _
                        ByVal Shift As Long, _
                        ByVal Button As Long, _
                        ByVal X As Long, _
                        ByVal Y As Long, _
                        ByRef Cancel As Boolean)

Public Event MouseWheelH(ByVal hWnd As Long, _
                         ByVal Delta As Long, _
                         ByVal Shift As Long, _
                         ByVal Button As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByRef Cancel As Boolean)

' Member variables
Private m_hWnd As Long

Private m_Enabled As Boolean

' Default values.
Private Const defEnabled As Boolean = True

' *********************************************
'  Initialize/Terminate
' *********************************************
Private Sub Class_Initialize()

    ' Set defaults
    m_Enabled = defEnabled

End Sub

Private Sub Class_Terminate()

    ' Tear down
    Call Unhook

End Sub

' *********************************************
'  Public Properties
' *********************************************
Public Property Let Enabled(ByVal NewValue As Boolean)

    m_Enabled = NewValue 'persist

End Property

Public Property Get Enabled() As Boolean

    Enabled = m_Enabled  'report

End Property

Public Property Let hWnd(ByVal NewValue As Long)

    ' Clear any existing hook, first.
    Call Unhook

    ' Set hook into new window, if possible.
    If IsWindow(NewValue) Then

        If HookSet(NewValue, Me) Then

            m_hWnd = NewValue

        End If

    End If

End Property

Public Property Get hWnd() As Long

    hWnd = m_hWnd  'report

End Property

' *********************************************
'  Public Properties (read-only)
' *********************************************
Public Property Get Present() As Boolean

    Present = CBool(GetSystemMetrics(SM_MOUSEWHEELPRESENT))

End Property

Public Property Get ScrollChars() As Long

    ' No OS support for hscroll until Vista.
    If SystemParametersInfo(SPI_GETWHEELSCROLLCHARS, 0&, ScrollChars, 0&) = False Then

        ' Err.LastDllError = 1439 (Invalid system-wide (SPI_*) parameter)
        ScrollChars = 1

    End If

End Property

Public Property Get ScrollLines() As Long

    Call SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0&, ScrollLines, 0&)

End Property

' *********************************************
'  Private Methods
' *********************************************
Private Function ReadButtonStates() As Long

    Dim Mask As Long

    Const HighBit As Long = &H80000000

    ' A series of GetKeyState calls builds mask of "shift" keys.
    If GetKeyState(VK_LBUTTON) And HighBit Then

        'Mask = Mask Or vbLeftButton
        Mask = vbLeftButton

    End If

    If GetKeyState(VK_MBUTTON) And HighBit Then

        Mask = Mask Or vbMiddleButton

    End If

    If GetKeyState(VK_RBUTTON) And HighBit Then

        Mask = Mask Or vbRightButton

    End If

    ' Return accumulated mask value
    ReadButtonStates = Mask

End Function

Private Function ReadKeyStates() As Long

    Dim Mask As Long

    Const HighBit As Long = &H80000000

    ' A series of GetKeyState calls builds mask of "shift" keys.
    If GetKeyState(VK_SHIFT) And HighBit Then

        'Mask = Mask Or vbShiftMask
        Mask = vbShiftMask

    End If

    If GetKeyState(VK_CONTROL) And HighBit Then

        Mask = Mask Or vbCtrlMask

    End If

    If GetKeyState(VK_MENU) And HighBit Then

        Mask = Mask Or vbAltMask

    End If

    ' Return accumulated mask value
    ReadKeyStates = Mask

End Function

Private Sub Unhook()

    ' Clear existing hook.
    If m_hWnd Then

        Call HookClear(m_hWnd, Me)
        m_hWnd = 0

    End If

End Sub

Public Function WordHi(ByVal DWord As Long) As Integer

    Call CopyMemory(WordHi, ByVal VarPtr(DWord) + 2, 2)

End Function

Public Function WordLo(ByVal DWord As Long) As Integer

    Call CopyMemory(WordLo, DWord, 2)

End Function

' *********************************************
'  Implemented Subclassing Interface
' *********************************************
Private Function IHookXP_Message(ByVal hWnd As Long, _
                                 ByVal uiMsg As Long, _
                                 ByVal wParam As Long, _
                                 ByVal lParam As Long, _
                                 ByVal dwRefData As Long) As Long

    'Best Practices for Supporting Microsoft Mouse and Keyboard Devices
    'http://msdn.microsoft.com/en-us/library/ms997498.aspx
    Dim EatIt As Boolean
    Dim Delta As Long
    Dim pt As POINTAPI
    Dim hWndOver As Long
    Dim Button As Long
    Dim Shift As Long
    Dim Cancel As Boolean
   
    ' Special processing for messages we care about.
    Select Case uiMsg

        Case WM_MOUSEWHEEL, WM_MOUSEHWHEEL

            If m_Enabled Then

                ' Gather all available information about event.
                Button = ReadButtonStates()
                Shift = ReadKeyStates()
                Delta = WordHi(wParam)
                pt.X = WordLo(lParam)
                pt.Y = WordHi(lParam)
                hWndOver = WindowFromPoint(pt.X, pt.Y)
            
                ' Alert client that wheel event occurred.
                If uiMsg = WM_MOUSEWHEEL Then

                    RaiseEvent MouseWheel(hWndOver, Delta, Shift, Button, pt.X, pt.Y, _
                            Cancel)
                Else
                    RaiseEvent MouseWheelH(hWndOver, Delta, Shift, Button, pt.X, pt.Y, _
                            Cancel)

                End If
            
                ' Fire default handler, just in case, but tell Windows
                ' that we handled it regardless.  VB Forms don't react
                ' at all to these messages, but the baseclass for some
                ' controls (eg, textbox) will use it, so it depends on
                ' what the client is subclassing how this will play.
                If Cancel = False Then

                    Call HookDefault(hWnd, uiMsg, wParam, lParam)

                End If

                IHookXP_Message = 1  'True
                EatIt = True

            End If
         
        Case WM_NCDESTROY

            Call Unhook  ' !!!

    End Select
   
    ' Pass back to default message handler.
    If EatIt = False Then

        IHookXP_Message = HookDefault(hWnd, uiMsg, wParam, lParam)

    End If

End Function

