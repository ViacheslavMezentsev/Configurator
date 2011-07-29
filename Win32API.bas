Attribute VB_Name = "Win32API"
'**
'@author <a href="mailto:unihomelab@ya.ru">Мезенцев Вячеслав</a>
'@revision Дата ревизии: 16.06.2011 г., время: 5:28:47
'@rem <h1><b>Win32API</b></h1>
'<pre>
'--------------------------------------------------------------------------------
' Проект   :       Конфигуратор управляющих программ
' Модуль   :       Win32API
' Описание :       Библиотека системых функций
' Автор    :       Мезенцев Вячеслав
' Изменён  :       16.06.2011 г., время: 5:28:47
'--------------------------------------------------------------------------------
'</pre>
Option Explicit

' -=[ Суффиксы в Visual basic 6 ]=-

' ( Техника указания типа данных с использованием знака типа (%, &, !, #, @, $) считается устаревшей )

' Название типа: [Символ в качестве суффикса] Integer: [%], Long: [&], Currency: [@], Single: [!], Double: [#], String: [$]

' *****************************************
' *  MSVBVM60 ФУНКЦИИ
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
' *  WIN32API ФУНКЦИИ
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
               Lib "kernel32" _
               Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, _
                                                    ByVal lpString As String, _
                                                    ByVal lpFileName As String) As Boolean
    
' The WritePrivateProfileString function copies a string
' into the specified section of the specified initialization file.
Public Declare Function WritePrivateProfileStringByKeyName% _
               Lib "kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpAppName$, _
                                                   ByVal lpKeyName$, _
                                                   ByVal lpString$, _
                                                   ByVal lpFileName$)

Public Declare Function WritePrivateProfileStringToDeleteKey% _
               Lib "kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpAppName$, _
                                                   ByVal lpKeyName$, _
                                                   ByVal lpString&, _
                                                   ByVal lpFileName$)

Public Declare Function WritePrivateProfileStringToDeleteSection% _
               Lib "kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpAppName$, _
                                                   ByVal lpKeyName&, _
                                                   ByVal lpString&, _
                                                   ByVal lpFileName$)

' The GetPrivateProfileString function retrieves a string
' from the specified section in an initialization file.
Public Declare Function GetPrivateProfileStringByKeyName& _
               Lib "kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpAppName$, _
                                                 ByVal lpKeyName$, _
                                                 ByVal lpDefault$, _
                                                 ByVal lpReturnedString$, _
                                                 ByVal nSize&, _
                                                 ByVal lpFileName$)

Public Declare Function GetPrivateProfileStringKeys& _
               Lib "kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpAppName$, _
                                                 ByVal lpKeyName&, _
                                                 ByVal lpDefault$, _
                                                 ByVal lpReturnedString$, _
                                                 ByVal nSize&, _
                                                 ByVal lpFileName$)

Public Declare Function GetModuleFileName _
               Lib "kernel32" _
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
' *     ФУНКЦИИ ДЛЯ РАБОТЫ С РЕЕСТРОМ
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
' *  ФУНКЦИИ ДЛЯ РАБОТЫ С ФАЙЛАМИ ОТОБРАЖАЕМЫМИ В ПАМЯТЬ
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~ ~~~~~~~~~~~~~ ~ ~~~~~~
' ******************************************************

Public Declare Function CreateFileMapping _
               Lib "kernel32" _
               Alias "CreateFileMappingA" (ByVal hFile As Long, _
                                           lpFileMappigAttributes As SECURITY_ATTRIBUTES, _
                                           ByVal flProtect As Long, _
                                           ByVal dwMaximumSizeHigh As Long, _
                                           ByVal dwMaximumSizeLow As Long, _
                                           ByVal lpname As String) As Long

Public Declare Function OpenFileMapping _
               Lib "kernel32" _
               Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, _
                                         ByVal bInheritHandle As Long, _
                                         ByVal lpname As String) As Long

Public Declare Function MapViewOfFile _
               Lib "kernel32" (ByVal hFileMappingObject As Long, _
                               ByVal dwDesiredAccess As Long, _
                               ByVal dwFileOffsetHigh As Long, _
                               ByVal dwFileOffsetLow As Long, _
                               ByVal dwNumberOfBytesToMap As Long) As Long

Public Declare Function UnmapViewOfFile _
               Lib "kernel32" (ByVal lpBaseAddress As Long) As Long

Public Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long

' *********************************
' *  ФУНКЦИИ ДЛЯ РАБОТЫ С ФАЙЛАМИ
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~
' *********************************

Public Declare Function CreateFile _
               Lib "kernel32" _
               Alias "CreateFileA" (ByVal lpFileName As String, _
                                    ByVal dwDesiredAccess As Long, _
                                    ByVal dwShareMode As Long, _
                                    lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                                    ByVal dwCreationDisposition As Long, _
                                    ByVal dwFlagsAndAttributes As Long, _
                                    ByVal hTemplateFile As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function WriteFile _
               Lib "kernel32" (ByVal hFile As Long, _
                               lpBuffer As Any, _
                               ByVal nNumberOfBytesToWrite As Long, _
                               lpNumberOfBytesWritten As Long, _
                               ByVal lpOverlapped As Long) As Long

Public Declare Function ReadFile _
               Lib "kernel32" (ByVal hFile As Long, _
                               lpBuffer As Any, _
                               ByVal nNumberOfBytesToRead As Long, _
                               lpNumberOfBytesRead As Long, _
                               ByVal lpOverlapped As Long) As Long

Public Declare Function DeleteFile _
               Lib "kernel32" _
               Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Declare Function GetLastError Lib "kernel32" () As Long

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
' *  ФУНКЦИИ ДЛЯ РАБОТЫ С ПАМЯТЬЮ
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~
' *********************************

Public Declare Sub CopyMemory _
               Lib "kernel32.dll" _
               Alias "RtlMoveMemory" (ByRef Destination As Any, _
                                      ByRef Source As Any, _
                                      ByVal Length As Long)

Public Declare Sub ZeroMemory _
               Lib "kernel32.dll" _
               Alias "RtlZeroMemory" (ByRef Destination As Any, _
                                      ByVal Length As Long)

' ***********************************
' *  ФУНКЦИИ ДЛЯ РАБОТЫ С ЛОВУШКАМИ
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
' *  ФУНКЦИИ ДЛЯ РАБОТЫ С UNICODE
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
               Lib "kernel32" _
               Alias "RtlMoveMemory" (dest As Any, _
                                      ByVal Source As Long, _
                                      ByVal Length As Long)

Public Declare Function lstrcpy _
               Lib "kernel32" _
               Alias "lstrcpyA" (ByVal lpString1 As String, _
                                 ByVal lpString2 As Long) As Long

' ***********************************
' *  ФУНКЦИИ ДЛЯ РАБОТЫ С ИНТЕРНЕТ
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~~
' ***********************************
Public Declare Function InternetGetConnectedState _
               Lib "wininet.dll" (ByRef lpdwFlags As Long, _
                                  ByVal dwReserved As Long) As Long

Public Declare Function GetVersionEx _
               Lib "kernel32" _
               Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Function GetCurrentDirectory Lib "kernel32.dll" Alias "GetCurrentDirectoryA" ( _
     ByVal nBufferLength As Long, _
     ByVal lpBuffer As String) As Long
     
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" ( _
     ByVal lpBuffer As String, _
     ByRef nSize As Long) As Long
