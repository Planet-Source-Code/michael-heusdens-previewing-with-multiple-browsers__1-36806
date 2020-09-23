Attribute VB_Name = "mdShell"
Option Explicit
















Private Const INFINITE = &HFFFFFFFF
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const WAIT_TIMEOUT = &H102&
'


Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
    End Type
'


Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
'


Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'


Private Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
'


Private Declare Function CreateProcessByNum Lib "kernel32" Alias "CreateProcessA" _
    (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes _
    As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags _
    As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As _
    STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
'

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'











Public Function LaunchAppSynchronous(strExecutablePathAndName As String) As Boolean
                                                    
Dim lngResponse As Long
Dim typStartUpInfo As STARTUPINFO
Dim typProcessInfo As PROCESS_INFORMATION
LaunchAppSynchronous = False


With typStartUpInfo
    .cb = Len(typStartUpInfo)
    .lpReserved = vbNullString
    .lpDesktop = vbNullString
    .lpTitle = vbNullString
    .dwFlags = 0
End With
'Launch the application by creating a new process
lngResponse = CreateProcessByNum(vbNullString, strExecutablePathAndName, 0, 0, True, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, typStartUpInfo, typProcessInfo)


If lngResponse Then
    'Wait for the application to terminate before moving on
    Call WaitForTermination(typProcessInfo)
    LaunchAppSynchronous = True
Else
    LaunchAppSynchronous = False
End If
End Function


Private Sub WaitForTermination(typProcessInfo As PROCESS_INFORMATION)
    'This wait routine allows other application events
    'to be processed while waiting for the process to
    'complete.
    Dim lngResponse As Long
    'Let the process initialize
    Call WaitForInputIdle(typProcessInfo.hProcess, INFINITE)
    'We don't need the thread handle so get rid of it
    Call CloseHandle(typProcessInfo.hThread)
    'Wait for the application to end


    Do
        lngResponse = WaitForSingleObject(typProcessInfo.hProcess, 0)


        If lngResponse <> WAIT_TIMEOUT Then
            'No timeout, app is terminated
            Exit Do
        End If


        DoEvents
    Loop While True
    'Kill the last handle of the process
    Call CloseHandle(typProcessInfo.hProcess)
End Sub

