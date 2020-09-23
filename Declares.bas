Attribute VB_Name = "Declares"
Declare Function ShellExecute Lib "shell32.dll" Alias _
          "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
          As String, ByVal lpFile As String, ByVal lpParameters _
          As String, ByVal lpDirectory As String, ByVal nShowCmd _
          As Long) As Long

Declare Function ShowWindow% Lib "User" (ByVal hwnd%, ByVal nCmdShow%)
    Global Const SW_HIDE = 0
    Global Const SW_SHOWNORMAL = 1
    Global Const SW_SHOWMINIMIZED = 2
    Global Const SW_SHOWMAXIMIZED = 3
    Global Const SW_SHOWNOACTIVE = 4
    Global Const SW_SHOW = 5
    Global Const SW_MINIMIZE = 6
    Global Const SW_SHOWMINNOACTIVE = 7
    Global Const SW_SHOWNA = 8
    Global Const SW_RESTORE = 9

Declare Function GetShortPathName Lib "kernel32" Alias _
    "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal _
    lpszShortPath As String, ByVal cchBuffer As Long) As Long


Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

