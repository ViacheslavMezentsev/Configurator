VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHookMouseEvents"
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

Private Declare Function TrackMouseEvent _
                Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

' Mouse constants and data structures
Private Const WM_NCDESTROY As Long = &H82
Private Const WM_XBUTTONDOWN As Long = &H20B         '(_WIN32_WINNT >= 0x0500)
Private Const WM_XBUTTONUP As Long = &H20C           '(_WIN32_WINNT >= 0x0500)
Private Const WM_XBUTTONDBLCLK As Long = &H20D       '(_WIN32_WINNT >= 0x0500)

Private Const WM_MOUSEACTIVATE As Long = &H21
Private Const WM_MOUSEFIRST As Long = &H200
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_RBUTTONDBLCLK As Long = &H206
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_MBUTTONDBLCLK As Long = &H209
Private Const WM_MOUSELAST As Long = &H209

Private Const WM_CAPTURECHANGED As Long = &H215

Private Const WM_MOUSEHOVER As Long = &H2A1   'Based on TrackMouseEvent
Private Const WM_MOUSELEAVE As Long = &H2A3   'Based on TrackMouseEvent

Private Type TRACKMOUSEEVENT_STRUCT

    cbSize As Long
    dwFlags As Long
    hWndTrack As Long
    dwHoverTime As Long

End Type

Private Const TME_HOVER As Long = &H1
Private Const TME_LEAVE As Long = &H2
Private Const TME_NONCLIENT As Long = &H10
Private Const TME_QUERY As Long = &H40000000
Private Const TME_CANCEL As Long = &H80000000
Private Const HOVER_DEFAULT As Long = &HFFFFFFFF

' WM_MOUSEACTIVATE Return Codes
Private Const MA_ACTIVATE = 1
Private Const MA_ACTIVATEANDEAT = 2
Private Const MA_NOACTIVATE = 3
Private Const MA_NOACTIVATEANDEAT = 4

' Subclassing interface
Implements IHookXP

' Events
Public Event LostCapture(ByVal hWnd As Long, ByVal hWndCapture As Long)

Public Event MouseEnter(ByVal hWnd As Long)

Public Event MouseHover(ByVal hWnd As Long)

Public Event MouseLeave(ByVal hWnd As Long)

Public Event XButtonDblClick(ByVal hWnd As Long, _
                             ByVal Button As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long)

Public Event XButtonDown(ByVal hWnd As Long, _
                         ByVal Button As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long)

Public Event XButtonUp(ByVal hWnd As Long, _
                       ByVal Button As Long, _
                       ByVal X As Long, _
                       ByVal Y As Long)

' Member variables
Private m_hWnd As Long

Private m_Enabled As Boolean

Private m_HoverTime As Long

Private m_Objects As Collection

' Default values.
Private Const defEnabled As Boolean = True
Private Const defHoverTime As Long = HOVER_DEFAULT

' *********************************************
'  Initialize/Terminate
' *********************************************
Private Sub Class_Initialize()

    ' Set defaults
    m_Enabled = defEnabled
    m_HoverTime = defHoverTime
    Set m_Objects = New Collection

End Sub

Private Sub Class_Terminate()

    ' Tear down
    Call UnhookAll
    Set m_Objects = Nothing

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

Public Property Let HoverTime(ByVal NewValue As Boolean)

    m_HoverTime = NewValue 'persist

End Property

Public Property Get HoverTime() As Boolean

    HoverTime = m_HoverTime  'report

End Property

' *********************************************
'  Public Properties (read-only)
' *********************************************

' *********************************************
'  Public Methods
' *********************************************
Public Function Add(obj As Object) As Boolean

    On Error Resume Next
    m_Objects.Add obj, FmtHex(obj.hWnd, 8)

    If Err.Number = 0 Then 'success

        If HookSet(obj.hWnd, Me) Then

            Add = True

        End If

    Else
        Debug.Print Err.Number, Err.Description

    End If

End Function

Public Function Count() As Long

    Count = m_Objects.Count

End Function

Public Function Item(ByVal hWnd As Long) As Object

    On Error Resume Next
    Set Item = m_Objects.Item(FmtHex(hWnd, 8))

End Function

Public Function Remove(obj As Object) As Boolean

    On Error Resume Next
    Remove = UnhookOne(obj.hWnd)

End Function
   
' *********************************************
'  Private Methods
' *********************************************
Public Function FmtHex(ByVal InVal As Long, ByVal OutLen As Integer) As String

    ' Left pad with zeros to OutLen.
    FmtHex = "&h" & Right$(String$(OutLen, "0") & Hex$(InVal), OutLen)

End Function

Private Sub UnhookAll()

    Dim obj As Object

    ' Clear existing hook.
    For Each obj In m_Objects

        Call HookClear(obj.hWnd, Me)

    Next obj

End Sub

Private Function UnhookOne(ByVal hWnd As Long) As Boolean

    On Error Resume Next
    m_Objects.Remove FmtHex(hWnd, 8)

    If Err.Number = 0 Then 'success

        If HookClear(hWnd, Me) Then

            UnhookOne = True

        End If

    Else
        Debug.Print Err.Number, Err.Description

    End If

End Function

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
    Static hWndLast As Long

    Dim EatIt As Boolean
   
    ' Special processing for messages we care about.
    Select Case uiMsg

        Case WM_MOUSEMOVE

            ' The Old New Thing : Why is there no WM_MOUSEENTER message?
            ' http://blogs.msdn.com/oldnewthing/archive/2003/10/13/55279.aspx
            If hWnd <> hWndLast Then

                hWndLast = hWnd  'Set flag

                If m_Enabled Then

                    RaiseEvent MouseEnter(hWnd)

                End If

                ' Start tracking for Leave event.
                StartTrackingMouse hWnd, TME_LEAVE Or TME_HOVER

            End If
      
        Case WM_MOUSEHOVER

            If m_Enabled Then

                RaiseEvent MouseHover(hWnd)

            End If
      
        Case WM_MOUSELEAVE

            hWndLast = 0  'Clear flag

            If m_Enabled Then

                RaiseEvent MouseLeave(hWnd)

            End If
         
        Case WM_CAPTURECHANGED

            If m_Enabled Then

                RaiseEvent LostCapture(hWnd, lParam)

            End If
         
        Case WM_XBUTTONDOWN

            If m_Enabled Then

                RaiseEvent XButtonDown(hWnd, WordHi(wParam), WordLo(lParam), WordHi( _
                        lParam))

            End If
         
        Case WM_XBUTTONUP

            If m_Enabled Then

                RaiseEvent XButtonUp(hWnd, WordHi(wParam), WordLo(lParam), WordHi(lParam))

            End If
         
        Case WM_XBUTTONDBLCLK

            If m_Enabled Then

                RaiseEvent XButtonDblClick(hWnd, WordHi(wParam), WordLo(lParam), WordHi( _
                        lParam))

            End If
         
        Case WM_NCDESTROY

            Call UnhookOne(hWnd)  ' !!!
         
    End Select
   
    ' Pass back to default message handler.
    If EatIt = False Then

        IHookXP_Message = HookDefault(hWnd, uiMsg, wParam, lParam)

    End If

End Function

Private Function StartTrackingMouse(ByVal hWnd As Long, ByVal Flags As Long) As Long

    Dim tme As TRACKMOUSEEVENT_STRUCT

    ' Wrap all the mess into a tidy little procedure.
    With tme

        .cbSize = Len(tme)
        .dwFlags = Flags

        If .dwFlags And TME_HOVER Then

            .dwHoverTime = m_HoverTime

        End If

        .hWndTrack = hWnd

    End With

    ' Return results.
    StartTrackingMouse = TrackMouseEvent(tme)

End Function

Private Sub DumpMsg(ByVal hWnd As Long, _
                    ByVal uiMsg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long, _
                    ByVal dwRefData As Long)

    ' Call DumpMsg(hWnd, uiMsg, wParam, lParam, dwRefData)
    Debug.Print "hWnd: 0x"; Hex$(hWnd), "Msg: 0x"; Hex$(uiMsg), "wParam: 0x"; Hex$( _
            wParam), "lParam: 0x"; Hex$(lParam), "RefData: "; dwRefData

End Sub

