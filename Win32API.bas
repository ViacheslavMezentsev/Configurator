Attribute VB_Name = "Win32API"
' «‡‰‡ÌËÂ ˇ‚ÌÓ„Ó Ó·˙ˇ‚ÎÂÌËˇ ÔÂÂÏÂÌÌ˚ı
Option Explicit

' -=[ —ÛÙÙËÍÒ˚ ‚ Visual basic 6 ]=-

' ( “ÂıÌËÍ‡ ÛÍ‡Á‡ÌËˇ ÚËÔ‡ ‰‡ÌÌ˚ı Ò ËÒÔÓÎ¸ÁÓ‚‡ÌËÂÏ ÁÌ‡Í‡ ÚËÔ‡ _
(%, &, !, #, @, $) Ò˜ËÚ‡ÂÚÒˇ ÛÒÚ‡Â‚¯ÂÈ )

' Õ‡Á‚‡ÌËÂ ÚËÔ‡: [—ËÏ‚ÓÎ ‚ Í‡˜ÂÒÚ‚Â ÒÛÙÙËÍÒ‡] _
Integer: [%], Long: [&], Currency: [@], Single: [!], Double: [#], String: [$]

' *****************************************
' *  MSVBVM60 ‘”Õ ÷»»
' *  ~~~~~~~~ ~~~~~~~
' *****************************************

Public Declare Function GetMem4 Lib "msvbvm60" (ByVal pSrc As Long, ByVal pDst As Long) As Long
Public Declare Function PutMem4 Lib "msvbvm60" (ByVal pDst As Long, ByVal NewValue As Long) As Long
Public Declare Function GetMem2 Lib "msvbvm60" (ByVal pSrc As Long, ByVal pDst As Long) As Long
Public Declare Function PutMem2 Lib "msvbvm60" (ByVal pDst As Long, ByVal NewValue As Long) As Long

' *****************************************
' *  WIN32API ‘”Õ ÷»»
' *  ~~~~~~~~ ~~~~~~~
' *****************************************

Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Public Const DELETE = &H10000
Public Const READ_CONTROL = &H20000
Public Const WRITE_DAC = &H40000
Public Const WRITE_OWNER = &H80000
Public Const SYNCHRONIZE = &H100000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2

Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_ALL = &H1F0000

Public Const SECTION_QUERY = &H1
Public Const SECTION_MAP_WRITE = &H2
Public Const SECTION_MAP_READ = &H4
Public Const SECTION_MAP_EXECUTE = &H8
Public Const SECTION_EXTEND_SIZE = &H10
Public Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE

Public Const FILE_MAP_COPY = SECTION_QUERY
Public Const FILE_MAP_WRITE = SECTION_MAP_WRITE
Public Const FILE_MAP_READ = SECTION_MAP_READ
Public Const FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const CREATE_ALWAYS = 2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Public Const PAGE_READWRITE = 4&

Public Const WH_KEYBOARD As Long = 2

' The WritePrivateProfileSection function replaces the keys and values
' under the specified section in an initialization file.
Public Declare Function WritePrivateProfileSection Lib "kernel32" _
    Alias "WritePrivateProfileSectionA" ( _
    ByVal lpAppName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String _
    ) As Boolean
    
' The WritePrivateProfileString function copies a string
' into the specified section of the specified initialization file.
Public Declare Function WritePrivateProfileStringByKeyName% Lib "kernel32" _
    Alias "WritePrivateProfileStringA" ( _
    ByVal lpAppName$, _
    ByVal lpKeyName$, _
    ByVal lpString$, _
    ByVal lpFileName$ _
)

Public Declare Function WritePrivateProfileStringToDeleteKey% Lib "kernel32" _
    Alias "WritePrivateProfileStringA" ( _
    ByVal lpAppName$, _
    ByVal lpKeyName$, _
    ByVal lpString&, _
    ByVal lpFileName$ _
)

Public Declare Function WritePrivateProfileStringToDeleteSection% Lib "kernel32" _
    Alias "WritePrivateProfileStringA" ( _
    ByVal lpAppName$, _
    ByVal lpKeyName&, _
    ByVal lpString&, _
    ByVal lpFileName$ _
)

' The GetPrivateProfileString function retrieves a string
' from the specified section in an initialization file.
Public Declare Function GetPrivateProfileStringByKeyName& Lib "kernel32" _
    Alias "GetPrivateProfileStringA" ( _
    ByVal lpAppName$, _
    ByVal lpKeyName$, _
    ByVal lpDefault$, _
    ByVal lpReturnedString$, _
    ByVal nSize&, _
    ByVal lpFileName$ _
)

Public Declare Function GetPrivateProfileStringKeys& Lib "kernel32" _
    Alias "GetPrivateProfileStringA" ( _
    ByVal lpAppName$, _
    ByVal lpKeyName&, _
    ByVal lpDefault$, _
    ByVal lpReturnedString$, _
    ByVal nSize&, _
    ByVal lpFileName$ _
)

Public Declare Function GetModuleFileName Lib "kernel32" _
    Alias "GetModuleFileNameA" ( _
    ByVal hModule As Long, _
    ByVal lpFileName As String, _
    ByVal nSize As Long _
) As Long

Public Declare Function PathFileExists Lib "shlwapi.dll" _
    Alias "PathFileExistsA" ( _
    ByVal pszPath As String _
) As Long

Public Declare Function CoCreateGuid Lib "ole32.dll" (Buffer As Byte) As Long
    
Public Declare Function StringFromGUID2 Lib "ole32.dll" ( _
    Buffer As Byte, _
    ByVal lpsz As Long, _
    ByVal cbMax As Long _
) As Long
 
' ******************************************************
' *  ‘”Õ ÷»» ƒÀﬂ –¿¡Œ“€ — ‘¿…À¿Ã» Œ“Œ¡–¿∆¿≈Ã€Ã» ¬ œ¿Ãﬂ“‹
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~ ~~~~~~~~~~~~~ ~ ~~~~~~
' ******************************************************

Public Declare Function CreateFileMapping Lib "kernel32" _
    Alias "CreateFileMappingA" ( _
    ByVal hFile As Long, _
    lpFileMappigAttributes As SECURITY_ATTRIBUTES, _
    ByVal flProtect As Long, _
    ByVal dwMaximumSizeHigh As Long, _
    ByVal dwMaximumSizeLow As Long, _
    ByVal lpname As String _
) As Long

Public Declare Function OpenFileMapping Lib "kernel32" _
    Alias "OpenFileMappingA" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal lpname As String _
) As Long

Public Declare Function MapViewOfFile Lib "kernel32" ( _
    ByVal hFileMappingObject As Long, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwFileOffsetHigh As Long, _
    ByVal dwFileOffsetLow As Long, _
    ByVal dwNumberOfBytesToMap As Long _
) As Long

Public Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long) As Long
Public Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long

' *********************************
' *  ‘”Õ ÷»» ƒÀﬂ –¿¡Œ“€ — ‘¿…À¿Ã»
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~
' *********************************

Public Declare Function CreateFile Lib "kernel32" _
    Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    lpSecurityAttributes As SECURITY_ATTRIBUTES, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long _
) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long _
) As Long

Public Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Long _
) As Long

Public Declare Function GetLastError Lib "kernel32" () As Long

' *********************************
' *  ‘”Õ ÷»» ƒÀﬂ –¿¡Œ“€ — œ¿Ãﬂ“‹ﬁ
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~
' *********************************

Public Declare Sub CopyMemory Lib "kernel32.dll" _
    Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal length As Long _
)

Public Declare Sub ZeroMemory Lib "kernel32.dll" _
    Alias "RtlZeroMemory" ( _
    ByRef Destination As Any, _
    ByVal length As Long _
)

' ***********************************
' *  ‘”Õ ÷»» ƒÀﬂ –¿¡Œ“€ — ÀŒ¬”ÿ ¿Ã»
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~~~
' ***********************************

Public Declare Function SetWindowsHookEx Lib "user32" _
    Alias "SetWindowsHookExA" ( _
    ByVal idHook As Long, _
    ByVal lpfn As Long, ByVal hmod As Long, _
    ByVal dwThreadId As Long _
) As Long

Public Declare Function CallNextHookEx Lib "user32" ( _
    ByVal hHook As Long, _
    ByVal ncode As Long, _
    ByVal wParam As Long, _
    lParam As Any _
) As Long

Public Declare Function UnhookWindowsHookEx Lib "user32" ( _
    ByVal hHook As Long _
) As Long

Public Declare Function GetAsyncKeyState Lib "user32" ( _
    ByVal vKey As Long _
) As Integer

' ***********************************
' *  ‘”Õ ÷»» ƒÀﬂ –¿¡Œ“€ — UNICODE
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~
' ***********************************
Public Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, _
    ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
    
Public Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, _
    ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

