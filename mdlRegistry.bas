Attribute VB_Name = "mdlRegistry"
Option Explicit


Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4

Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003

Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259

Public Const ERROR_SUCCESS = 0&

Global Const KEY_ALL_ACCESS = &H3F

Global Const REG_OPTION_NON_VOLATILE = 0

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String)
Private Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String)


Public Function DeleteKey(lPredefinedKey As Long, sKeyName As String)
' Description:
'   This Function will Delete a key
'
' Syntax:
'   DeleteKey Location, KeyName
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is name of the key you wish to delete, it may include subkeys (example "Key1\SubKey1")


    Dim lRetVal As Long         'result of the SetValueEx function
    Dim hKey As Long         'handle of open key
    
    'open the specified key
    
    'lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = RegDeleteKey(lPredefinedKey, sKeyName)
    'RegCloseKey (hKey)
End Function

Public Function DeleteValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
' Description:
'   This Function will delete a value
'
' Syntax:
'   DeleteValue Location, KeyName, ValueName
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is the name of the key that the value you wish to delete is in
'   , it may include subkeys (example "Key1\SubKey1")
'
'   ValueName is the name of value you wish to delete

       Dim lRetVal As Long         'result of the SetValueEx function
       Dim hKey As Long         'handle of open key

       'open the specified key

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = RegDeleteValue(hKey, sValueName)
       RegCloseKey (hKey)
End Function

Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String

    Select Case lType
        Case REG_SZ
            sValue = vValue
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
    End Select

End Function





Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError



    ' Determine the size and type of data to be read

    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        ' For strings
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If

        ' For DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:

    QueryValueEx = lrc
    Exit Function

QueryValueExError:

    Resume QueryValueExExit

End Function
Public Function CreateNewKey(lPredefinedKey As Long, sNewKeyName As String)
' Description:
'   This Function will create a new key
'
' Syntax:
'   QueryValue Location, KeyName
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is name of the key you wish to create, it may include subkeys (example "Key1\SubKey1")

    
    
    Dim hNewKey As Long         'handle to the new key
    Dim lRetVal As Long         'result of the RegCreateKeyEx function
    
    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
End Function

Public Function SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
' Description:
'   This Function will set the data field of a value
'
' Syntax:
'   QueryValue Location, KeyName, ValueName, ValueSetting, ValueType
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is the key that the value is under (example: "Key1\SubKey1")
'
'   ValueName is the name of the value you want create, or set the value of (example: "ValueTest")
'
'   ValueSetting is what you want the value to equal
'
'   ValueType must equal either REG_SZ (a string) Or REG_DWORD (an integer)

       Dim lRetVal As Long         'result of the SetValueEx function
       Dim hKey As Long         'handle of open key

       'open the specified key

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
       RegCloseKey (hKey)

End Function

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
' Description:
'   This Function will return the data field of a value
'
' Syntax:
'   Variable = QueryValue(Location, KeyName, ValueName)
'
'   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
'   , HKEY_USERS
'
'   KeyName is the key that the value is under (example: "Software\Microsoft\Windows\CurrentVersion\Explorer")
'
'   ValueName is the name of the value you want to access (example: "link")

       Dim lRetVal As Long         'result of the API functions
       Dim hKey As Long         'handle of opened key
       Dim vValue As Variant      'setting of queried value


       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = QueryValueEx(hKey, sValueName, vValue)
       QueryValue = vValue
       RegCloseKey (hKey)
End Function

Public Function KeyExists() As Boolean
Dim ver As String, vers As String, NetscLangType As String
Dim Browsemnu As String, NetscKey As String

'Create registry key to store locations
CreateNewKey HKEY_LOCAL_MACHINE, "SOFTWARE\Shadows Inc\Glass Sailor\Browsers"

'''''''''''''''''''''''''''''''''''''
'Check to see if Netscape is loaded on this machine.
'The latest Netscape has a different registry entry,
'so we must run a check for both.
'(Netscape Gold, Netscape Communicator, Netscape 6 (en),
'and Netscape 7.0 beta 1 tested)
'''''''''''''''''''''''''''''''''''''

'Checking for Netscape Communicator
KeyExists = bCheckKeyExists(HKEY_LOCAL_MACHINE, "SOFTWARE\Netscape\Netscape Navigator\")

If KeyExists = True Then
    'Gets Netscape Communicator version
    ver = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Netscape\Netscape Navigator\", "CurrentVersion")
    ''''''''''''''''''''''''''''''''''''''
    'Cut off the language type
    '''''''''''''''''''''''''''''''''''''
    ver = CutLangType(ver)
    
    Browsemnu = "Netscape " & ver
    AddToMenu Browsemnu
    'Get Netscape Communicator location and add it to the registry
    SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Shadows Inc\Glass Sailor\Browsers\", Browsemnu, QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Netscape.exe\", ""), REG_SZ
    
Else
    'Netscape Communicator wasn't found, do nothing
    
End If

'Checking for Netscape 6
KeyExists = bCheckKeyExists(HKEY_LOCAL_MACHINE, "SOFTWARE\Netscape\Netscape 6\")

If KeyExists = True Then
    'Gets Netscape 6 version
    ver = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Netscape\Netscape 6\", "CurrentVersion")
    
    '''''''''''''''''''''''''''''''''''''
    'Get language type (needed. For some
    'reason, placing the ver as is in a
    'string won't allow us to get into
    'the subkeys in which the exe path
    'is stored). This will later be
    'compared to ISO-639 language codes.
    '''''''''''''''''''''''''''''''''''''
    vers = ver   'prevents us from rewriting the version string
    NetscLangType = GetLangType(vers)
    
    '''''''''''''''''''''''''''''''''''''
    'Cut off the language type
    '''''''''''''''''''''''''''''''''''''
    ver = CutLangType(ver)
    
    Browsemnu = "Netscape " & ver
    AddToMenu Browsemnu
    
    'Alright. Now lets compare to the ISO-639 v1 (two-letter) language codes.
    'I've provided support only for english
    If StrComp(LCase(NetscLangType), "(en)") = 1 Then
        'Get location and add it to the registry
        NetscKey = "SOFTWARE\Netscape\Netscape 6\" & ver & " (en)\Main\"
        SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Shadows Inc\Glass Sailor\Browsers\", Browsemnu, QueryValue(HKEY_LOCAL_MACHINE, NetscKey, "PathToExe"), REG_SZ
    End If
    
Else
    'Netscape 6 wasn't found, do nothing
    
End If

'Checking for Netscape 7 (This may work for later versions of
'Netscape as well...)
KeyExists = bCheckKeyExists(HKEY_LOCAL_MACHINE, "SOFTWARE\Netscape\Netscape\")

If KeyExists = True Then
    'Gets version
    ver = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Netscape\Netscape\", "CurrentVersion")
    
    '''''''''''''''''''''''''''''''''''''
    'Get language type (needed. For some
    'reason, placing the ver as is in a
    'string won't allow us to get into
    'the subkeys in which the exe path
    'is stored). This will later be
    'compared to ISO-639 language codes.
    '''''''''''''''''''''''''''''''''''''
    vers = ver   'prevents us from rewriting the version string
    NetscLangType = GetLangType(vers)
    
    '''''''''''''''''''''''''''''''''''''
    'Cut off the language type
    '''''''''''''''''''''''''''''''''''''
    ver = CutLangType(ver)
    
    Browsemnu = "Netscape " & ver
    AddToMenu Browsemnu
    
    'Alright. Now lets compare to the ISO-639 v1 (two-letter) language codes.
    'Support for chinese, english, french, german, and japanese
    If StrComp(LCase(NetscLangType), "(en)") = 1 Then
        'Get location and add it to the registry
        NetscKey = "SOFTWARE\Netscape\Netscape\" & ver & " (en)\Main\"
        SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Shadows Inc\Glass Sailor\Browsers\", Browsemnu, QueryValue(HKEY_LOCAL_MACHINE, NetscKey, "PathToExe"), REG_SZ

    ElseIf StrComp(LCase(NetscLangType), "(zh)") = 1 Then
        'Get location and add it to the registry
        NetscKey = "SOFTWARE\Netscape\Netscape\" & ver & " (zh)\Main\"
        SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Shadows Inc\Glass Sailor\Browsers\", Browsemnu, QueryValue(HKEY_LOCAL_MACHINE, NetscKey, "PathToExe"), REG_SZ

    ElseIf StrComp(LCase(NetscLangType), "(fr)") = 1 Then
        'Get location and add it to the registry
        NetscKey = "SOFTWARE\Netscape\Netscape\" & ver & " (fr)\Main\"
        SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Shadows Inc\Glass Sailor\Browsers\", Browsemnu, QueryValue(HKEY_LOCAL_MACHINE, NetscKey, "PathToExe"), REG_SZ
    
    ElseIf StrComp(LCase(NetscLangType), "(de)") = 1 Then
        'Get location and add it to the registry
        NetscKey = "SOFTWARE\Netscape\Netscape\" & ver & " (de)\Main\"
        SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Shadows Inc\Glass Sailor\Browsers\", Browsemnu, QueryValue(HKEY_LOCAL_MACHINE, NetscKey, "PathToExe"), REG_SZ
    
    ElseIf StrComp(LCase(NetscLangType), "(ja)") = 1 Then
        'Get location and add it to the registry
        NetscKey = "SOFTWARE\Netscape\Netscape\" & ver & " (ja)\Main\"
        SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Shadows Inc\Glass Sailor\Browsers\", Browsemnu, QueryValue(HKEY_LOCAL_MACHINE, NetscKey, "PathToExe"), REG_SZ
    
    End If
        
Else
    'Not found, do nothing
    
End If


'''''''''''''''''''''''''''''''''''''
'Checking to see if Amaya is loaded on this machine.
'(http://www.w3.org/Amaya/)
'''''''''''''''''''''''''''''''''''''
KeyExists = bCheckKeyExists(HKEY_LOCAL_MACHINE, "SOFTWARE\W3C - INRIA\Amaya\")

If KeyExists = True Then
    Browsemnu = "Amaya"
    AddToMenu Browsemnu
    'Get Amaya location and add it to the registry
    SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Shadows Inc\Glass Sailor\Browsers\", Browsemnu, QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\amaya.exe\", ""), REG_SZ

    
Else
    'Amaya wasn't found, do nothing
    
End If

'''''''''''''''''''''''''''''''''''''
'Checking to see if Opera is loaded on this machine.
'''''''''''''''''''''''''''''''''''''
KeyExists = bCheckKeyExists(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Opera.exe\")

If KeyExists = True Then
    'Gets Opera version
    ver = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Opera 6.0 (Win32)\", "DisplayName")
    '''''''''''''''''''''''''''''''''''''
    'Cut off the platform type
    '''''''''''''''''''''''''''''''''''''
    ver = CutLangType(ver)
    
    Browsemnu = ver
    AddToMenu Browsemnu
    'Get Opera location and add it to the registry
    SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Shadows Inc\Glass Sailor\Browsers\", Browsemnu, QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Opera.exe\", ""), REG_SZ

    
Else
    'Opera wasn't found, do nothing
    
End If

'''''''''''''''''''''''''''''''''''''
'No real need to see if IE is installed but we'll
'do it anyway.
'''''''''''''''''''''''''''''''''''''
KeyExists = bCheckKeyExists(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer\")

If KeyExists = True Then
    'Gets IE version
    ver = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer\Version Vector\", "IE")
    '''''''''''''''''''''''''''''''''''''
    'We cut off the zeros so version is displayed #.#
    'instead of #.####
    '''''''''''''''''''''''''''''''''''''
    ver = CutDecimal(ver, 1)

    Browsemnu = "Internet Explorer " & ver
    AddToMenu Browsemnu
    'Get IE location and add it to the registry
    SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Shadows Inc\Glass Sailor\Browsers\", Browsemnu, QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\IEXPLORE.exe\", ""), REG_SZ

    
Else
    'IE wasn't found, do nothing
    
End If


End Function

Public Function bCheckKeyExists(ByVal hKey As Long, ByVal strKey As String) As Boolean

Dim phkResult As Long

If RegOpenKeyEx(hKey, strKey, 0, 1, phkResult) = ERROR_SUCCESS Then
    bCheckKeyExists = True
    RegCloseKey phkResult
Else
    bCheckKeyExists = False
End If

End Function

Public Function CutDecimal(Number As String, ByPlace As Byte) As String
Dim Dec As Byte

Dec = InStr(1, Number, ".", vbBinaryCompare) ' find the Decimal

If Dec = 0 Then
    CutDecimal = Number   'if there is no decimal Then dont do anything
    Exit Function
End If
CutDecimal = Mid(Number, 1, Dec + ByPlace) 'How many places you want after the decimal point

End Function

Public Function CutLangType(Number As String)
Dim intCount As Integer
Dim boolSwitch As Boolean

If Number = "" Then
    Exit Function
End If

Number = CStr(Number)
intCount = 1

Do Until boolSwitch = True
    If Mid(Number, intCount, 1) <> "" Then
        If Mid(Number, intCount, 1) = "(" Then  'Check for the delimiter
            boolSwitch = True  'Delimiter found, stop loop
        Else
            intCount = intCount + 1  'Didn't find delimiter, goto next
        End If
    Else
        CutLangType = Number
        Exit Function
    End If
Loop

Number = Left(Number, intCount - 2) 'Cut delimiter off so we only show version number
CutLangType = Number

End Function

Public Function GetLangType(LangParse As String)
Dim intCount As Integer
Dim boolSwitch As Boolean

If LangParse = "" Then
    Exit Function
End If

LangParse = CStr(LangParse)
intCount = 1

Do Until boolSwitch = True
    If Mid(LangParse, intCount, 1) <> "" Then
        If Mid(LangParse, intCount, 1) = "(" Then  'Check for one delimiter
            LangParse = Right(LangParse, intCount - 2)
            intCount = intCount + 1
        ElseIf Mid(LangParse, intCount, 1) = ")" Then   'Check for the other delimiter
            boolSwitch = True  'Delimiter found, stop loop
        Else
            intCount = intCount + 1  'Didn't find delimiter, goto next
        End If
    Else
        GetLangType = LangParse
        Exit Function
    End If
Loop

GetLangType = LangParse

End Function

Public Function AddToMenu(cSubMenuName As String)
'GETS NEXT INDEX NUMBER FOR MENU ITEM
Dim iNextIndex As Integer
'MENU NAME INDEX PLUS ONE FOR ITEM
iNextIndex = Form1.mnuPBrowse(Form1.mnuPBrowse.UBound).Index + 1
'LOADS THE MENU NAME WITH NEW INDEX
Call Load(Form1.mnuPBrowse(iNextIndex))
'SETS THE NEW ITEM CAPTION TO MENU
Form1.mnuPBrowse(iNextIndex).Caption = cSubMenuName
End Function

'All this really does is hide and disable the menu item until program exits
Public Function RemoveMenu(cSubMenuName As String)
Dim iRIndex As Integer
Dim nElementCount As Integer

nElementCount = Form1.mnuPBrowse(Form1.mnuPBrowse.UBound).Index

For iRIndex = 1 To nElementCount
    If Form1.mnuPBrowse(iRIndex).Caption = cSubMenuName Then
        Form1.mnuPBrowse(iRIndex).Enabled = False
        Form1.mnuPBrowse(iRIndex).Visible = False
        Exit For
    End If
Next iRIndex

End Function


'Sub Main()
    'Examples of each function:
    'CreateNewKey HKEY_CURRENT_USER, "TestKey\SubKey1\SubKey2"
    'SetKeyValue HKEY_CURRENT_USER, "TestKey\SubKey1", "Test", "Testing, Testing", REG_SZ
    'MsgBox QueryValue(HKEY_CURRENT_USER, "TestKey\SubKey1", "Test")
    'DeleteKey HKEY_CURRENT_USER, "TestKey\SubKey1\SubKey2"
    'DeleteValue HKEY_CURRENT_USER, "TestKey\SubKey1", "Test"
    'KeyExists = bCheckKeyExists(HKEY_LOCAL_MACHINE, "TestKey\SubKey1\")
'End Sub
