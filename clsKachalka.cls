VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKachalka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'должен задать hWnd окна-родителя для запроса пароля (по умолчанию десктоп),
' или установить hWnd в нуль и заполнить логин-пароль
Event Authenticate(hWnd As Long, Username As String, Password As String)

'возвращает статус завершения скачки
Event Complete(ByVal Status As kach_tlb.BINDSTATUS, ByVal StatusText As String)

'самое главное событие: пришли данные
' дожен установить Cancel, чтобы прекратить скачку
Event Progress(ByVal Progress As Long, _
               ByVal ProgressMax As Long, _
               ByVal Status As Long, _
               ByVal StatusText As String, _
               Cancel As Boolean)

'событие о приходе данных для push mode, передаёт сам фрагмент
' DataFormat указывает на FORMATETC
Event DataAvailable(ByVal EventType As kach_tlb.BSCF, _
                    ByVal Data As String, _
                    ByVal DataFormat As Long)

'должен задать параметры скачки, по умолчанию BINDF_RESYNCHRONIZE Or BINDF_GETFROMCACHE_IF_NET_FAIL
Event GetOptions(Options As kach_tlb.BINDF)

'метод запроса: GET, POST, или PUT
Public Method As kach_tlb.BINDVERB

'данне запроса для методов POST и PUT
Public RequestData As String

Implements IBindStatusCallback
Implements IAuthenticate

Private Offset As Long

Private ReturnCode(11) As Byte

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Sub IAuthenticate_Authenticate(hWnd As Long, _
                                       pszUsername As Long, _
                                       pszPassword As Long)

    Dim Username As String, Password As String

    hWnd = GetDesktopWindow
    RaiseEvent Authenticate(hWnd, Username, Password)

    If hWnd = 0 Then

        pszUsername = StrPtr(Username)
        PutMem4 ByVal VarPtr(Username), 0
        pszPassword = StrPtr(Password)
        PutMem4 ByVal VarPtr(Password), 0

    End If

End Sub

Private Sub IBindStatusCallback_GetBindInfo(grfBINDF As kach_tlb.BINDF, _
                                            pbindinfo As kach_tlb.BINDINFO)

    grfBINDF = BINDF_RESYNCHRONIZE Or BINDF_GETFROMCACHE_IF_NET_FAIL
    RaiseEvent GetOptions(grfBINDF)
    pbindinfo.dwBindVerb = Me.Method

    If StrPtr(Me.RequestData) Then

        pbindinfo.stgmedData.TYMED = TYMED_ISTREAM
        
        Set pbindinfo.stgmedData.pstm = CreateStreamOnHGlobal(0, 1)
        
        pbindinfo.cbstgmedData = Len(Me.RequestData)
        pbindinfo.stgmedData.pstm.Write Me.RequestData, pbindinfo.cbstgmedData
        pbindinfo.stgmedData.pstm.Seek 0, STREAM_SEEK_SET

    End If

End Sub

Private Function IBindStatusCallback_GetPriority() As Long

    Stop 'афаик не нужен

End Function

Private Sub IBindStatusCallback_OnDataAvailable(ByVal grfBSCF As kach_tlb.BSCF, _
                                                ByVal dwSize As Long, _
                                                pformatetc As kach_tlb.FORMATETC, _
                                                pstgmed As kach_tlb.STGMEDIUM)

    dwSize = dwSize - Offset

    Dim buf As String: buf = String(dwSize, 0)

    If pstgmed.TYMED <> TYMED_ISTREAM Then Error 13

    If dwSize <> pstgmed.pstm.Read(buf, dwSize) Then Error 7

    RaiseEvent DataAvailable(grfBSCF, buf, VarPtr(pformatetc))
    
    Offset = Offset + dwSize

End Sub

Private Sub IBindStatusCallback_OnLowResource(ByVal reserved As Long)

    'не нужен
End Sub

Private Sub IBindStatusCallback_OnObjectAvailable(riid As kach_tlb.UUID, _
                                                  ByVal pUnk As stdole.IUnknown)

    'не нужен
End Sub

Private Sub IBindStatusCallback_OnProgress(ByVal ulProgress As Long, _
                                           ByVal ulProgressMax As Long, _
                                           ByVal ulStatusCode As kach_tlb.BINDSTATUS, _
                                           ByVal szStatusText As Long)

    Dim Cancel As Boolean

    RaiseEvent Progress(ulProgress, ulProgressMax, ulStatusCode, SysAllocString( _
            szStatusText), Cancel)

    If Cancel Then 'whoo-hoo

        Dim RetAddr As Long, pRetAddr As Long

        pRetAddr = VarPtr(ulProgress) - 8
        PutMem4 ReturnCode(1), E_ABORT
        GetMem4 ByVal pRetAddr, RetAddr
        PutMem4 ByVal pRetAddr, VarPtr(ReturnCode(0))
        PutMem4 ReturnCode(6), RetAddr

    End If

End Sub

Private Sub IBindStatusCallback_OnStartBinding(ByVal dwReserved As Long, _
                                               ByVal pib As kach_tlb.IBinding)

    Offset = 0

End Sub

Private Sub IBindStatusCallback_OnStopBinding(ByVal hrStatus As Long, _
                                              ByVal szStatusText As Long)

    RaiseEvent Complete(hrStatus, SysAllocString(szStatusText))

End Sub

Public Sub DownloadToFile(ByVal URL As String, ByVal FileName As String)

    Dim Result As Long

    Result = URLDownloadToFile(Nothing, URL, FileName, 0, Me)

    If Result Then Err.Raise Result

End Sub

Public Function DownloadToString(ByVal URL As String) As String

    Dim Result As Long, pStream As IStream

    Result = URLOpenBlockingStream(Nothing, URL, pStream, 0, Me)

    If Result Then Err.Raise Result

    Dim s As STATSTG

    pStream.Stat s, STATFLAG_NONAME

    If s.cbSizeHi Then Error 7

    DownloadToString = String(s.cbSize, 0)

    If s.cbSize <> pStream.Read(DownloadToString, s.cbSize) Then Error 7

End Function

Public Sub DownloadPush(ByVal URL As String)

    Dim Result As Long, pStream As IStream

    Result = URLOpenStream(Nothing, URL, 0, Me)

    If Result Then Err.Raise Result

End Sub

Private Sub Class_Initialize()

    ReturnCode(0) = &HB8
    ReturnCode(5) = &HBA
    ReturnCode(10) = &HFF
    ReturnCode(11) = &HE2

End Sub

