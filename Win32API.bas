Attribute VB_Name = "Win32API"
'**
'@author <a href="mailto:unihomelab@ya.ru">ÃÂÁÂÌˆÂ‚ ¬ˇ˜ÂÒÎ‡‚</a>
'@revision ƒ‡Ú‡ Â‚ËÁËË: 16.06.2011 „., ‚ÂÏˇ: 5:28:47
'@rem <h1><b>Win32API</b></h1>
'<pre>
'--------------------------------------------------------------------------------
' œÓÂÍÚ   :        ÓÌÙË„Û‡ÚÓ ÛÔ‡‚Îˇ˛˘Ëı ÔÓ„‡ÏÏ
' ÃÓ‰ÛÎ¸   :       Win32API
' ŒÔËÒ‡ÌËÂ :       ¡Ë·ÎËÓÚÂÍ‡ ÒËÒÚÂÏ˚ı ÙÛÌÍˆËÈ
' ¿‚ÚÓ    :       ÃÂÁÂÌˆÂ‚ ¬ˇ˜ÂÒÎ‡‚
' »ÁÏÂÌ∏Ì  :       16.06.2011 „., ‚ÂÏˇ: 5:28:47
'--------------------------------------------------------------------------------
'</pre>
Option Explicit

' -=[ —ÛÙÙËÍÒ˚ ‚ Visual basic 6 ]=-

' ( “ÂıÌËÍ‡ ÛÍ‡Á‡ÌËˇ ÚËÔ‡ ‰‡ÌÌ˚ı Ò ËÒÔÓÎ¸ÁÓ‚‡ÌËÂÏ ÁÌ‡Í‡ ÚËÔ‡ (%, &, !, #, @, $) Ò˜ËÚ‡ÂÚÒˇ ÛÒÚ‡Â‚¯ÂÈ )

' Õ‡Á‚‡ÌËÂ ÚËÔ‡: [—ËÏ‚ÓÎ ‚ Í‡˜ÂÒÚ‚Â ÒÛÙÙËÍÒ‡] Integer: [%], Long: [&], Currency: [@], Single: [!], Double: [#], String: [$]

' *****************************************
' *  MSVBVM60 ‘”Õ ÷»»
' *  ~~~~~~~~ ~~~~~~~
' *****************************************

Public Declare Function GetMem4 _
               Lib "msvbvm60" (ByVal pSrc As Long, _
                               ByVal pDst As Long) As Long

Public Declare Function PutMem4 _
               Lib "msvbvm60" (ByVal pDst As Long, _
                               ByVal NewValue As Long) As Long

Public Declare Function GetMem2 _
               Lib "msvbvm60" (ByVal pSrc As Long, _
                               ByVal pDst As Long) As Long

Public Declare Function PutMem2 _
               Lib "msvbvm60" (ByVal pDst As Long, _
                               ByVal NewValue As Long) As Long

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
Public Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or _
        SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or _
        SECTION_EXTEND_SIZE

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

' ShowWindow() Commands
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_MAX = 10

'
' GetSystemMetrics() codes
'

Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SM_CXVSCROLL = 2
Public Const SM_CYHSCROLL = 3
Public Const SM_CYCAPTION = 4
Public Const SM_CXBORDER = 5
Public Const SM_CYBORDER = 6
Public Const SM_CXDLGFRAME = 7
Public Const SM_CYDLGFRAME = 8
Public Const SM_CYVTHUMB = 9
Public Const SM_CXHTHUMB = 10
Public Const SM_CXICON = 11
Public Const SM_CYICON = 12
Public Const SM_CXCURSOR = 13
Public Const SM_CYCURSOR = 14
Public Const SM_CYMENU = 15
Public Const SM_CXFULLSCREEN = 16
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYKANJIWINDOW = 18
Public Const SM_MOUSEPRESENT = 19
Public Const SM_CYVSCROLL = 20
Public Const SM_CXHSCROLL = 21
Public Const SM_DEBUG = 22
Public Const SM_SWAPBUTTON = 23
Public Const SM_RESERVED1 = 24
Public Const SM_RESERVED2 = 25
Public Const SM_RESERVED3 = 26
Public Const SM_RESERVED4 = 27
Public Const SM_CXMIN = 28
Public Const SM_CYMIN = 29
Public Const SM_CXSIZE = 30
Public Const SM_CYSIZE = 31
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33
Public Const SM_CXMINTRACK = 34
Public Const SM_CYMINTRACK = 35
Public Const SM_CXDOUBLECLK = 36
Public Const SM_CYDOUBLECLK = 37
Public Const SM_CXICONSPACING = 38
Public Const SM_CYICONSPACING = 39
Public Const SM_MENUDROPALIGNMENT = 40
Public Const SM_PENWINDOWS = 41
Public Const SM_DBCSENABLED = 42
Public Const SM_CMOUSEBUTTONS = 43

'/* Ternary raster operations */
'/* dest = source                   */
Public Const SRCCOPY = &HCC0020

'Public Const SRCPAINT            (DWORD)0x00EE0086 /* dest = source OR dest           */
'Public Const SRCAND              (DWORD)0x008800C6 /* dest = source AND dest          */
'Public Const SRCINVERT           (DWORD)0x00660046 /* dest = source XOR dest          */
'Public Const SRCERASE            (DWORD)0x00440328 /* dest = source AND (NOT dest )   */
'Public Const NOTSRCCOPY          (DWORD)0x00330008 /* dest = (NOT source)             */
'Public Const NOTSRCERASE         (DWORD)0x001100A6 /* dest = (NOT src) AND (NOT dest) */
'Public Const MERGECOPY           (DWORD)0x00C000CA /* dest = (source AND pattern)     */
'Public Const MERGEPAINT          (DWORD)0x00BB0226 /* dest = (NOT source) OR dest     */
'Public Const PATCOPY             (DWORD)0x00F00021 /* dest = pattern                  */
'Public Const PATPAINT            (DWORD)0x00FB0A09 /* dest = DPSnoo                   */
'Public Const PATINVERT           (DWORD)0x005A0049 /* dest = pattern XOR dest         */
'Public Const DSTINVERT           (DWORD)0x00550009 /* dest = (NOT dest)               */
'Public Const BLACKNESS           (DWORD)0x00000042 /* dest = BLACK                    */
'Public Const WHITENESS           (DWORD)0x00FF0062 /* dest = WHITE                    */

' Local system uses a modem to connect to the Internet
Public Const INTERNET_CONNECTION_MODEM = 1

' Local system uses a local area network to connect to the Internet
Public Const INTERNET_CONNECTION_LAN = 2

' Local system uses a proxy server to connect to the Internet
Public Const INTERNET_CONNECTION_PROXY = 4

' Local system's modem is busy with a non-Internet connection
Public Const INTERNET_CONNECTION_MODEM_BUSY = 8

' The WritePrivateProfileSection function replaces the keys and values
' under the specified section in an initialization file.
Public Declare Function WritePrivateProfileSection _
               Lib "Kernel32" _
               Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, _
                                                    ByVal lpString As String, _
                                                    ByVal lpFileName As String) As Boolean
    
' The WritePrivateProfileString function copies a string
' into the specified section of the specified initialization file.
Public Declare Function WritePrivateProfileStringByKeyName% _
               Lib "Kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpAppName$, _
                                                   ByVal lpKeyName$, _
                                                   ByVal lpString$, _
                                                   ByVal lpFileName$)

Public Declare Function WritePrivateProfileStringToDeleteKey% _
               Lib "Kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpAppName$, _
                                                   ByVal lpKeyName$, _
                                                   ByVal lpString&, _
                                                   ByVal lpFileName$)

Public Declare Function WritePrivateProfileStringToDeleteSection% _
               Lib "Kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpAppName$, _
                                                   ByVal lpKeyName&, _
                                                   ByVal lpString&, _
                                                   ByVal lpFileName$)

' The GetPrivateProfileString function retrieves a string
' from the specified section in an initialization file.
Public Declare Function GetPrivateProfileStringByKeyName& _
               Lib "Kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpAppName$, _
                                                 ByVal lpKeyName$, _
                                                 ByVal lpDefault$, _
                                                 ByVal lpReturnedString$, _
                                                 ByVal nSize&, _
                                                 ByVal lpFileName$)

Public Declare Function GetPrivateProfileStringKeys& _
               Lib "Kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpAppName$, _
                                                 ByVal lpKeyName&, _
                                                 ByVal lpDefault$, _
                                                 ByVal lpReturnedString$, _
                                                 ByVal nSize&, _
                                                 ByVal lpFileName$)

Public Declare Function GetModuleFileName _
               Lib "Kernel32" _
               Alias "GetModuleFileNameA" (ByVal hModule As Long, _
                                           ByVal lpFileName As String, _
                                           ByVal nSize As Long) As Long

Public Declare Function PathFileExists _
               Lib "shlwapi.dll" _
               Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Public Declare Function CoCreateGuid Lib "ole32.dll" (buffer As Byte) As Long
    
Public Declare Function StringFromGUID2 _
               Lib "ole32.dll" (buffer As Byte, _
                                ByVal lpsz As Long, _
                                ByVal cbMax As Long) As Long

' ******************************************************
' *     ‘”Õ ÷»» ƒÀﬂ –¿¡Œ“€ — –≈≈—“–ŒÃ
' *     ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~~
' ******************************************************

Public Declare Function RegOpenKeyEx _
               Lib "advapi32" _
               Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                      ByVal lpSubKey As String, _
                                      ByVal ulOptions As Long, _
                                      ByVal samDesired As Long, _
                                      ByRef phkResult As Long) As Long

Public Declare Function RegQueryValueEx _
               Lib "advapi32" _
               Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                         ByVal lpValueName As String, _
                                         ByVal lpReserved As Long, _
                                         ByRef lpType As Long, _
                                         ByVal lpData As String, _
                                         ByRef lpcbData As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

' ******************************************************
' *  ‘”Õ ÷»» ƒÀﬂ –¿¡Œ“€ — ‘¿…À¿Ã» Œ“Œ¡–¿∆¿≈Ã€Ã» ¬ œ¿Ãﬂ“‹
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~ ~~~~~~~~~~~~~ ~ ~~~~~~
' ******************************************************

Public Declare Function CreateFileMapping _
               Lib "Kernel32" _
               Alias "CreateFileMappingA" (ByVal hFile As Long, _
                                           lpFileMappigAttributes As SECURITY_ATTRIBUTES, _
                                           ByVal flProtect As Long, _
                                           ByVal dwMaximumSizeHigh As Long, _
                                           ByVal dwMaximumSizeLow As Long, _
                                           ByVal lpname As String) As Long

Public Declare Function OpenFileMapping _
               Lib "Kernel32" _
               Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, _
                                         ByVal bInheritHandle As Long, _
                                         ByVal lpname As String) As Long

Public Declare Function MapViewOfFile _
               Lib "Kernel32" (ByVal hFileMappingObject As Long, _
                               ByVal dwDesiredAccess As Long, _
                               ByVal dwFileOffsetHigh As Long, _
                               ByVal dwFileOffsetLow As Long, _
                               ByVal dwNumberOfBytesToMap As Long) As Long

Public Declare Function UnmapViewOfFile _
               Lib "Kernel32" (ByVal lpBaseAddress As Long) As Long

Public Declare Function FlushFileBuffers Lib "Kernel32" (ByVal hFile As Long) As Long

' *********************************
' *  ‘”Õ ÷»» ƒÀﬂ –¿¡Œ“€ — ‘¿…À¿Ã»
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~
' *********************************

Public Declare Function CreateFile _
               Lib "Kernel32" _
               Alias "CreateFileA" (ByVal lpFileName As String, _
                                    ByVal dwDesiredAccess As Long, _
                                    ByVal dwShareMode As Long, _
                                    lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                                    ByVal dwCreationDisposition As Long, _
                                    ByVal dwFlagsAndAttributes As Long, _
                                    ByVal hTemplateFile As Long) As Long

Public Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long

Public Declare Function WriteFile _
               Lib "Kernel32" (ByVal hFile As Long, _
                               lpBuffer As Any, _
                               ByVal nNumberOfBytesToWrite As Long, _
                               lpNumberOfBytesWritten As Long, _
                               ByVal lpOverlapped As Long) As Long

Public Declare Function ReadFile _
               Lib "Kernel32" (ByVal hFile As Long, _
                               lpBuffer As Any, _
                               ByVal nNumberOfBytesToRead As Long, _
                               lpNumberOfBytesRead As Long, _
                               ByVal lpOverlapped As Long) As Long

Public Declare Function DeleteFile _
               Lib "Kernel32" _
               Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Declare Function GetLastError Lib "Kernel32" () As Long

Public Declare Function ShellExecute _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long

Public Declare Function GetTempPath _
               Lib "kernel32.dll" _
               Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                                     ByVal lpBuffer As String) As Long

Declare Function GetTempFileName _
        Lib "kernel32.dll" _
        Alias "GetTempFileNameA" (ByVal lpszPath As String, _
                                  ByVal lpPrefixString As String, _
                                  ByVal wUnique As Long, _
                                  ByVal lpTempFileName As String) As Long
     
' *********************************
' *  ‘”Õ ÷»» ƒÀﬂ –¿¡Œ“€ — œ¿Ãﬂ“‹ﬁ
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~
' *********************************

Public Declare Sub CopyMemory _
               Lib "kernel32.dll" _
               Alias "RtlMoveMemory" (ByRef Destination As Any, _
                                      ByRef Source As Any, _
                                      ByVal length As Long)

Public Declare Sub ZeroMemory _
               Lib "kernel32.dll" _
               Alias "RtlZeroMemory" (ByRef Destination As Any, _
                                      ByVal length As Long)

' ***********************************
' *  ‘”Õ ÷»» ƒÀﬂ –¿¡Œ“€ — ÀŒ¬”ÿ ¿Ã»
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~~~
' ***********************************

Public Declare Function SetWindowsHookEx _
               Lib "user32" _
               Alias "SetWindowsHookExA" (ByVal idHook As Long, _
                                          ByVal lpfn As Long, _
                                          ByVal hmod As Long, _
                                          ByVal dwThreadId As Long) As Long

Public Declare Function CallNextHookEx _
               Lib "user32" (ByVal hHook As Long, _
                             ByVal ncode As Long, _
                             ByVal wParam As Long, _
                             lParam As Any) As Long

Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' ***********************************
' *  ‘”Õ ÷»» ƒÀﬂ –¿¡Œ“€ — UNICODE
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~
' ***********************************
Public Declare Function MultiByteToWideChar _
               Lib "kernel32.dll" (ByVal CodePage As Long, _
                                   ByVal dwFlags As Long, _
                                   ByVal lpMultiByteStr As String, _
                                   ByVal cchMultiByte As Long, _
                                   ByVal lpWideCharStr As Long, _
                                   ByVal cchWideChar As Long) As Long
    
Public Declare Function WideCharToMultiByte _
               Lib "kernel32.dll" (ByVal CodePage As Long, _
                                   ByVal dwFlags As Long, _
                                   ByVal lpWideCharStr As Long, _
                                   ByVal cchWideChar As Long, _
                                   ByVal lpMultiByteStr As Long, _
                                   ByVal cchMultiByte As Long, _
                                   ByVal lpDefaultChar As Long, _
                                   ByVal lpUsedDefaultChar As Long) As Long

'**************************************
'Windows API/Global Declarations for :Get Version Number for EXE, DLL or OCX files
'**************************************
Public Declare Function GetFileVersionInfo _
               Lib "Version.dll" _
               Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, _
                                            ByVal dwhandle As Long, _
                                            ByVal dwlen As Long, _
                                            lpData As Any) As Long

Public Declare Function GetFileVersionInfoSize _
               Lib "Version.dll" _
               Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
                                                lpdwHandle As Long) As Long

Public Declare Function VerQueryValue _
               Lib "Version.dll" _
               Alias "VerQueryValueA" (pBlock As Any, _
                                       ByVal lpSubBlock As String, _
                                       lplpBuffer As Any, _
                                       puLen As Long) As Long

Public Declare Sub MoveMemory _
               Lib "Kernel32" _
               Alias "RtlMoveMemory" (dest As Any, _
                                      ByVal Source As Long, _
                                      ByVal length As Long)

Public Declare Function lstrcpy _
               Lib "Kernel32" _
               Alias "lstrcpyA" (ByVal lpString1 As String, _
                                 ByVal lpString2 As Long) As Long

' ***********************************
' *  ‘”Õ ÷»» ƒÀﬂ –¿¡Œ“€ — »Õ“≈–Õ≈“
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~~
' ***********************************
Public Declare Function InternetGetConnectedState _
               Lib "wininet.dll" (ByRef lpdwFlags As Long, _
                                  ByVal dwReserved As Long) As Long

Public Declare Function GetVersionEx _
               Lib "Kernel32" _
               Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Function GetCurrentDirectory _
               Lib "kernel32.dll" _
               Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, _
                                             ByVal lpBuffer As String) As Long
     
Public Declare Function GetUserName _
               Lib "advapi32.dll" _
               Alias "GetUserNameA" (ByVal lpBuffer As String, _
                                     ByRef nSize As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long

Public Declare Sub keybd_event _
               Lib "user32" (ByVal bVk As Byte, _
                             ByVal bScan As Byte, _
                             ByVal dwFlags As Long, _
                             ByVal dwExtraInfo As Long)

Public Declare Function VkKeyScan _
               Lib "user32" _
               Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
 
Public Declare Function MapVirtualKey _
               Lib "user32" _
               Alias "MapVirtualKeyA" (ByVal wCode As Long, _
                                       ByVal wMapType As Long) As Long

Public Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
     
Public Declare Function BitBlt _
               Lib "gdi32.dll" (ByVal hDestDC As Long, _
                                ByVal x As Long, _
                                ByVal y As Long, _
                                ByVal nWidth As Long, _
                                ByVal nHeight As Long, _
                                ByVal hSrcDC As Long, _
                                ByVal xSrc As Long, _
                                ByVal ySrc As Long, _
                                ByVal dwRop As Long) As Long
     
Public Declare Function PlgBlt _
               Lib "gdi32.dll" (ByVal hdcDest As Long, _
                                ByRef lpPoint As POINTAPI, _
                                ByVal hdcSrc As Long, _
                                ByVal nXSrc As Long, _
                                ByVal nYSrc As Long, _
                                ByVal nWidth As Long, _
                                ByVal nHeight As Long, _
                                ByVal hbmMask As Long, _
                                ByVal xMask As Long, _
                                ByVal yMask As Long) As Long

Public Declare Function ReleaseDC _
               Lib "user32.dll" (ByVal hWnd As Long, _
                                 ByVal hdc As Long) As Long

Public Declare Function Polyline _
               Lib "gdi32" (ByVal hdc As Long, _
                            lpPoint As POINTAPI, _
                            ByVal nCount As Long) As Long

Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
    ByVal x As Long, ByVal y As Long) As Long

Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
    ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long

Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
