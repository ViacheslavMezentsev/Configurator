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

Public Declare Function GetMem4 Lib "msvbvm60" (ByVal pSrc As Long, ByVal pDst As Long) As Long
Public Declare Function PutMem4 Lib "msvbvm60" (ByVal pDst As Long, ByVal NewValue As Long) As Long
Public Declare Function GetMem2 Lib "msvbvm60" (ByVal pSrc As Long, ByVal pDst As Long) As Long
Public Declare Function PutMem2 Lib "msvbvm60" (ByVal pDst As Long, ByVal NewValue As Long) As Long

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

' The WritePrivateProfileSection function replaces the keys and values
' under the specified section in an initialization file.
Public Declare Function WritePrivateProfileSection Lib "Kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Boolean
    
' The WritePrivateProfileString function copies a string
' into the specified section of the specified initialization file.
Public Declare Function WritePrivateProfileStringByKeyName% Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpString$, ByVal lpFileName$)

Public Declare Function WritePrivateProfileStringToDeleteKey% Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpString&, ByVal lpFileName$)

Public Declare Function WritePrivateProfileStringToDeleteSection% Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName$, ByVal lpKeyName&, ByVal lpString&, ByVal lpFileName$)

' The GetPrivateProfileString function retrieves a string
' from the specified section in an initialization file.
Public Declare Function GetPrivateProfileStringByKeyName& Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize&, ByVal lpFileName$)

Public Declare Function GetPrivateProfileStringKeys& Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName$, ByVal lpKeyName&, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize&, ByVal lpFileName$)

Public Declare Function GetModuleFileName Lib "Kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Public Declare Function CoCreateGuid Lib "ole32.dll" (buffer As Byte) As Long
    
Public Declare Function StringFromGUID2 Lib "ole32.dll" (buffer As Byte, ByVal lpsz As Long, ByVal cbMax As Long) As Long

' ******************************************************
' *     ФУНКЦИИ ДЛЯ РАБОТЫ С РЕЕСТРОМ
' *     ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~~
' ******************************************************

Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

' ******************************************************
' *  ФУНКЦИИ ДЛЯ РАБОТЫ С ФАЙЛАМИ ОТОБРАЖАЕМЫМИ В ПАМЯТЬ
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~ ~~~~~~~~~~~~~ ~ ~~~~~~
' ******************************************************

Public Declare Function CreateFileMapping Lib "Kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As SECURITY_ATTRIBUTES, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpname As String) As Long

Public Declare Function OpenFileMapping Lib "Kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpname As String) As Long

Public Declare Function MapViewOfFile Lib "Kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long

Public Declare Function UnmapViewOfFile Lib "Kernel32" (ByVal lpBaseAddress As Long) As Long
Public Declare Function FlushFileBuffers Lib "Kernel32" (ByVal hFile As Long) As Long

' *********************************
' *  ФУНКЦИИ ДЛЯ РАБОТЫ С ФАЙЛАМИ
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~
' *********************************

Public Declare Function CreateFile Lib "Kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Public Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long

Public Declare Function WriteFile Lib "Kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

Public Declare Function ReadFile Lib "Kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long

Public Declare Function DeleteFile Lib "Kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Declare Function GetLastError Lib "Kernel32" () As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


' *********************************
' *  ФУНКЦИИ ДЛЯ РАБОТЫ С ПАМЯТЬЮ
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~
' *********************************

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)

Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal length As Long)

' ***********************************
' *  ФУНКЦИИ ДЛЯ РАБОТЫ С ЛОВУШКАМИ
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~~~
' ***********************************

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long

Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' ***********************************
' *  ФУНКЦИИ ДЛЯ РАБОТЫ С UNICODE
' *  ~~~~~~~~ ~~~~~~~~~ ~ ~~~~~~~
' ***********************************
Public Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
    
Public Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

'**************************************
'Windows API/Global Declarations for :Get Version Number for EXE, DLL or OCX files
'**************************************
Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Public Declare Sub MoveMemory Lib "Kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)
Public Declare Function lstrcpy Lib "Kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long


