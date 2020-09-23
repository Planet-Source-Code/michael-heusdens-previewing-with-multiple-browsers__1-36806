Attribute VB_Name = "INIMenuCalls"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Created by: Michael Heusdens
'    To work in conjunction with Zebastion's Class Mod
'
' Description:
'    Simplifies Zebastion's code so we won't have to input
' certain codes repetitively.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public ptConfigTest As INIFile '** Create config object!


Public Function LoadINI(ByVal getLoadValue As Integer)
Dim sValue As String
Dim sTemp() As String
Dim nCount As Integer
Dim nElementCount As Integer

Dim sConfigFile As String
    
    Set ptConfigTest = New INIFile
    
    '** Set config file
    sConfigFile = App.Path & "\" & "browsers.ini"
    
    '** Set Test Config Format and location!!
    ptConfigTest.SetFile sConfigFile
    ptConfigTest.Delimiter = "|"
    ptConfigTest.StartTag = "["
    ptConfigTest.EndTag = "]"
    ptConfigTest.KeyTag = "="
    ptConfigTest.TempFolder = App.Path & "\"

    '** Look for the keys
    sTemp = Split(ptConfigTest.FindKeys("Browsers"), "|")

    nElementCount = UBound(sTemp)
    
    Select Case getLoadValue
    Case "0"  'load to menu
        For nCount = 0 To nElementCount
            sValue = ptConfigTest.GetValue("Browsers", sTemp(nCount))
            AddToMenu sValue
        Next nCount
    Case "1"  'load to listbox
        Form3.List1.Clear
        For nCount = 0 To nElementCount
            sValue = ptConfigTest.GetValue("Browsers", sTemp(nCount))
            Form3.List1.AddItem sValue
        Next nCount
    End Select

End Function

Public Function SetValueINI(ivalue As String)
Dim sTemp() As String
Dim nElementCount As Integer

'Find all keys
sTemp = Split(ptConfigTest.FindKeys("Browsers"), "|")

'Get count and add 1 for next available
nElementCount = UBound(sTemp) + 1

ptConfigTest.Createkey "Browsers", "manual" & (nElementCount)
ptConfigTest.SetValue "Browsers", "manual" & (nElementCount), ivalue

End Function

Public Function RemoveFromINI(ivalue As String)
Dim sTemp() As String
Dim nCount As Integer
Dim nElementCount As Integer

    '** Look for the keys
    sTemp = Split(ptConfigTest.FindKeys("Browsers"), "|")

    nElementCount = UBound(sTemp)
    
    For nCount = 0 To nElementCount
        If ptConfigTest.GetValue("Browsers", sTemp(nCount)) = ivalue Then
                    ptConfigTest.DeleteKey "Browsers", sTemp(nCount)
                    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Shadows Inc\Glass Sailor\Browsers\", ivalue
                    RemoveMenu ivalue
                    LoadINI 1
                    GoTo keyRemoved
        End If
    Next nCount

keyRemoved:

End Function


Function FileExist(Fname As String) As Boolean
    On Local Error Resume Next
   FileExist = (Dir(Fname) <> "")
End Function

