VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INIFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************
'** (C)2001 Zebastion aka Zeb
'** Pronounced Sebastion - my parents were
'** pissed when they named me
'****************************************************************

Option Explicit

'** Public Properties (this is bad.. make accesser functions later)
Public StartTag As String
Public EndTag As String
Public Delimiter As String
Public KeyTag As String
Public TempFolder As String

Private psFile As String

Private Function CopyBack(DestinationFile As String, DataFile As String) As Boolean

Dim nFileIn As Integer
Dim nFileOut As Integer

Dim sCurrentInRecord As String
Dim sCurrentOutRecord As String

    Kill DestinationFile

    '** Get Free Files
    nFileIn = FreeFile
    
    '** Open for Input
    Open DataFile For Input As #nFileIn
    
    '** Open for Output
    nFileOut = FreeFile
    Open DestinationFile For Output As #nFileOut
        
        Do While Not EOF(nFileIn)
            
            '** Copy back to origional file!
            '** Read Line of Input File
            Input #nFileIn, sCurrentInRecord
            
            '** No wonder its blank!!
            sCurrentOutRecord = sCurrentInRecord
            
            '** Write to Output File
            Print #nFileOut, sCurrentOutRecord
            
        Loop
    
    '** Close Files
    Close #nFileIn
    Close #nFileOut
    
End Function
Public Sub Createkey(Section As String, sNewKey As String)

Dim nFileIn As Integer
Dim nFileOut As Integer
Dim sTempFile As String

Dim sCurrentInRecord As String
Dim sCurrentOutRecord As String
Dim sNewRecord As String

    '** Get Free Files
    nFileIn = FreeFile
    
    '** Open for Input
    Open psFile For Input As #nFileIn
    
    sTempFile = Path(TempFolder) & "~" & "data.tmp"
    
    '** Open for Output
    nFileOut = FreeFile
    Open sTempFile For Output As #nFileOut
    
        '** Create Temp file that will be a replicar of the config file
        '** with the exception of the NEW value!!
        
        '** Find desired section
        '** If the current in record is the wanted section
        '** then loop through keys till the wanted key is found
        Do While Not EOF(nFileIn)
            
            '** Get current InRecord
            Input #nFileIn, sCurrentInRecord
            
            If IsWantedSection(sCurrentInRecord, Section) Then
            
                '** Write section header
                sCurrentOutRecord = sCurrentInRecord
                Print #nFileOut, sCurrentOutRecord
                
                '** Get next record
                'Input #nFileIn, sCurrentInRecord
                
                'Do While Not eof(

                    'If IsWantedKey(sCurrentInRecord, Key, Section) Then
                        '** Create string with out the existing value!
                        'sNewRecord = Right$(sCurrentInRecord, InStr(1, sCurrentInRecord, KeyTag) + 1)
                        'sCurrentOutRecord = Left$(sCurrentInRecord, InStr(1, sCurrentInRecord, KeyTag) - 1) & KeyTag & Value
                    sCurrentOutRecord = sNewKey & KeyTag & "NewValue" & vbCrLf
                    Print #nFileOut, sCurrentOutRecord
                    'Exit Do
                    'Else
                        'sCurrentOutRecord = sCurrentInRecord
                        'Print #nFileOut, sCurrentOutRecord
                    'End If
                    
                    '** Get next line
                    'If EOF(nFileIn) Then Exit Do
                    
                    'Input #nFileIn, sCurrentInRecord
                    
                'Loop
            Else
                '** Copy the existing data back to the file!
                sCurrentOutRecord = sCurrentInRecord
                '** Write current out record
                Print #nFileOut, sCurrentOutRecord
            End If

        Loop
    
    Close #nFileIn
    Close #nFileOut
    
    '** Copy back to the existing file
    CopyBack psFile, sTempFile
    
    '** Hide the evidence (delete the tmp file)
    Kill (sTempFile)
    
End Sub

Public Sub CreateSection(sSectionName As String)

Dim nFileOut As Integer
Dim sCurrentOutRecord As String

    '** Open for Output
    nFileOut = FreeFile
    Open psFile For Append As #nFileOut
        sCurrentOutRecord = vbCrLf & StartTag & sSectionName & EndTag & vbCrLf
        sCurrentOutRecord = sCurrentOutRecord & "NewKey" & KeyTag & "NewValue"  '** Create a default key for the new section
        Print #nFileOut, sCurrentOutRecord
    Close #nFileOut
    
End Sub
Public Function DeleteKey(Section As String, Key As String) As Boolean

Dim nFileIn As Integer
Dim nFileOut As Integer
Dim sTempFile As String

Dim sCurrentInRecord As String
Dim sCurrentOutRecord As String
Dim sNewRecord As String

    '** Get Free Files
    nFileIn = FreeFile
    
    '** Open for Input
    Open psFile For Input As #nFileIn
    
    sTempFile = Path(TempFolder) & "~" & "data.tmp"
    
    '** Open for Output
    nFileOut = FreeFile
    Open sTempFile For Output As #nFileOut
    
        '** Create Temp file that will be a replicar of the config file
        '** with the exception of the NEW value!!
        
        '** Find desired section
        '** If the current in record is the wanted section
        '** then loop through keys till the wanted key is found
        Do While Not EOF(nFileIn)
            
            '** Get current InRecord
            Input #nFileIn, sCurrentInRecord
            
            If IsWantedSection(sCurrentInRecord, Section) Then
            
                '** Write section header
                sCurrentOutRecord = sCurrentInRecord
                Print #nFileOut, sCurrentOutRecord
                
                '** Get next record
                Input #nFileIn, sCurrentInRecord
                
                Do While Not IsSection(sCurrentInRecord)

                    If IsWantedKey(sCurrentInRecord, Key, Section) Then
                        '** dont do anything, in effect this will remove the
                        '** line when we read it all back int o the origional file!
                        Exit Do
                    Else
                        sCurrentOutRecord = sCurrentInRecord
                        Print #nFileOut, sCurrentOutRecord
                    End If
                    
                    '** Get next line
                    If EOF(nFileIn) Then Exit Do
                    Input #nFileIn, sCurrentInRecord
                    
                Loop
            Else
                '** Copy the existing data back to the file!
                sCurrentOutRecord = sCurrentInRecord
                '** Write current out record
                Print #nFileOut, sCurrentOutRecord
            End If

        Loop
    
    Close #nFileIn
    Close #nFileOut
    
    '** Copy back to the existing file
    CopyBack psFile, sTempFile
    
    '** Hide the evidence (delete the tmp file)
    Kill (sTempFile)

End Function
Public Function DeleteSection(Section As String) As Boolean

Dim nFileIn As Integer
Dim nFileOut As Integer
Dim sTempFile As String

Dim sCurrentInRecord As String
Dim sCurrentOutRecord As String
Dim sNewRecord As String

    '** Get Free Files
    nFileIn = FreeFile
    
    '** Open for Input
    Open psFile For Input As #nFileIn
    
    sTempFile = Path(TempFolder) & "~" & "data.tmp"
    
    '** Open for Output
    nFileOut = FreeFile
    Open sTempFile For Output As #nFileOut
    
        '** Create Temp file that will be a replicar of the config file
        '** with the exception of the NEW value!!
        
        '** Find desired section
        '** If the current in record is the wanted section
        '** then loop through keys till the wanted key is found
        Do While Not EOF(nFileIn)
            
            '** Get current InRecord
            Input #nFileIn, sCurrentInRecord
            
            If IsWantedSection(sCurrentInRecord, Section) Then
            
                '** Write section header
                'sCurrentOutRecord = sCurrentInRecord
                'Print #nFileOut, sCurrentOutRecord
                
                '** Get next record
                If EOF(nFileIn) Then Exit Do
                Input #nFileIn, sCurrentInRecord
                
                Do While Not IsSection(sCurrentInRecord)
                    '** Do nothing, we don't want any of these keys!
                    If EOF(nFileIn) Then Exit Do
                    Input #nFileIn, sCurrentInRecord
                    
                    If IsSection(sCurrentInRecord) Then
                        sCurrentOutRecord = sCurrentInRecord
                        Print #nFileOut, sCurrentOutRecord
                    End If
                Loop
            Else
                '** Copy the existing data back to the file!
                sCurrentOutRecord = sCurrentInRecord
                '** Write current out record
                Print #nFileOut, sCurrentOutRecord
            End If

        Loop
    
    Close #nFileIn
    Close #nFileOut
    
    '** Copy back to the existing file
    CopyBack psFile, sTempFile
    
    '** Hide the evidence (delete the tmp file)
    Kill (sTempFile)
    
End Function

Public Function RenameKey(Section As String, Key As String, NewKeyName As String) As Boolean

Dim nFileIn As Integer
Dim nFileOut As Integer
Dim sTempFile As String

Dim sCurrentInRecord As String
Dim sCurrentOutRecord As String
Dim sNewRecord As String

    '** Get Free Files
    nFileIn = FreeFile
    
    '** Open for Input
    Open psFile For Input As #nFileIn
    
    sTempFile = Path(TempFolder) & "~" & "data.tmp"
    
    '** Open for Output
    nFileOut = FreeFile
    Open sTempFile For Output As #nFileOut
    
        '** Create Temp file that will be a replicar of the config file
        '** with the exception of the NEW value!!
        
        '** Find desired section
        '** If the current in record is the wanted section
        '** then loop through keys till the wanted key is found
        Do While Not EOF(nFileIn)
            
            '** Get current InRecord
            Input #nFileIn, sCurrentInRecord
            
            If IsWantedSection(sCurrentInRecord, Section) Then
            
                '** Write section header
                sCurrentOutRecord = sCurrentInRecord
                Print #nFileOut, sCurrentOutRecord
                
                '** Get next record
                Input #nFileIn, sCurrentInRecord
                
                Do While Not IsSection(sCurrentInRecord)

                    If IsWantedKey(sCurrentInRecord, Key, Section) Then
                        '** Create string with out the existing value!
                        'sNewRecord = Right$(sCurrentInRecord, InStr(1, sCurrentInRecord, KeyTag) + 1)
                        'sCurrentOutRecord = Left$(sCurrentInRecord, InStr(1, sCurrentInRecord, KeyTag) - 1) & KeyTag & Value
                        sCurrentOutRecord = NewKeyName & Mid$(sCurrentInRecord, InStr(1, sCurrentInRecord, KeyTag), Len(sCurrentInRecord) - InStr(1, sCurrentInRecord, KeyTag) + 1)
                        Print #nFileOut, sCurrentOutRecord
                        Exit Do
                    Else
                        sCurrentOutRecord = sCurrentInRecord
                        Print #nFileOut, sCurrentOutRecord
                    End If
                    
                    '** Get next line
                    If EOF(nFileIn) Then Exit Do
                    
                    Input #nFileIn, sCurrentInRecord
                    
                Loop
            Else
                '** Copy the existing data back to the file!
                sCurrentOutRecord = sCurrentInRecord
                '** Write current out record
                Print #nFileOut, sCurrentOutRecord
            End If

        Loop
    
    Close #nFileIn
    Close #nFileOut
    
    '** Copy back to the existing file
    CopyBack psFile, sTempFile
    
    '** Hide the evidence (delete the tmp file)
    Kill (sTempFile)
    
End Function

Public Function RenameSection(Section As String, NewSectionName As String) As Boolean

Dim nFileIn As Integer
Dim nFileOut As Integer
Dim sTempFile As String

Dim sCurrentInRecord As String
Dim sCurrentOutRecord As String
Dim sNewRecord As String

    '** Get Free Files
    nFileIn = FreeFile
    
    '** Open for Input
    Open psFile For Input As #nFileIn
    
    sTempFile = Path(TempFolder) & "~" & "data.tmp"
    
    '** Open for Output
    nFileOut = FreeFile
    Open sTempFile For Output As #nFileOut
    
        '** Create Temp file that will be a replicar of the config file
        '** with the exception of the NEW value!!
        
        '** Find desired section
        '** If the current in record is the wanted section
        '** then loop through keys till the wanted key is found
        Do While Not EOF(nFileIn)
            
            '** Get current InRecord
            Input #nFileIn, sCurrentInRecord
            
            If IsWantedSection(sCurrentInRecord, Section) Then
            
                '** Build new section
                sCurrentOutRecord = StartTag & NewSectionName & EndTag
                
                Print #nFileOut, sCurrentOutRecord
                
            Else
                '** Copy the existing data back to the file!
                sCurrentOutRecord = sCurrentInRecord
                '** Write current out record
                Print #nFileOut, sCurrentOutRecord
            End If

        Loop
    
    Close #nFileIn
    Close #nFileOut
    
    '** Copy back to the existing file
    CopyBack psFile, sTempFile
    
    '** Hide the evidence (delete the tmp file)
    Kill (sTempFile)
    
End Function

Public Function SetFile(File As String) As String
    
    '** Set the file name and location
    If File <> "" Then
        psFile = File
    Else
        psFile = "nofile"
    End If
    
End Function

Public Function GetData()

'** File Variables
Dim nFreeFile As Integer

'** For Data Return
Dim sReturnString As String
Dim sDelimiter As String
Dim sCurrentRecord As String

    '** Set Delimiter
    sDelimiter = Delimiter
    
    '** Open the file one record at a time!!
    nFreeFile = FreeFile
    
    Open psFile For Input As nFreeFile
        
        Do While Not EOF(nFreeFile) '** Loop through the text file
            
            '** Get next line from the text file
            Input #nFreeFile, sCurrentRecord
            
            If InStr(1, sCurrentRecord, StartTag) <> 0 And InStr(1, sCurrentRecord, EndTag) <> 0 Then
                If sReturnString = "" Then
                    sReturnString = Mid$(sCurrentRecord, InStr(1, sCurrentRecord, StartTag) + Len(StartTag), (InStr(1, sCurrentRecord, EndTag) - Len(StartTag)) - 1)
                Else
                    sReturnString = sReturnString & sDelimiter & Mid$(sCurrentRecord, InStr(1, sCurrentRecord, StartTag) + Len(StartTag), (InStr(1, sCurrentRecord, EndTag) - Len(StartTag)) - 1)
                End If
            End If
        Loop
    Close nFreeFile
    
    GetData = sReturnString

End Function

Public Function GetSections()

'** File Variables
Dim nFreeFile As Integer

'** For Data Return
Dim sReturnString As String
Dim sDelimiter As String
Dim sCurrentRecord As String

    '** Set Delimiter
    sDelimiter = Delimiter
    
    '** Open the file one record at a time!!
    nFreeFile = FreeFile
    
    Open psFile For Input As nFreeFile
        
        Do While Not EOF(nFreeFile) '** Loop through the text file
            
            '** Get next line from the text file
            Input #nFreeFile, sCurrentRecord
            
            If InStr(1, sCurrentRecord, StartTag) <> 0 And InStr(1, sCurrentRecord, EndTag) <> 0 Then
                If sReturnString = "" Then
                    sReturnString = Mid$(sCurrentRecord, InStr(1, sCurrentRecord, StartTag) + Len(StartTag), (InStr(1, sCurrentRecord, EndTag) - Len(StartTag)) - 1)
                Else
                    sReturnString = sReturnString & sDelimiter & Mid$(sCurrentRecord, InStr(1, sCurrentRecord, StartTag) + Len(StartTag), (InStr(1, sCurrentRecord, EndTag) - Len(StartTag)) - 1)
                End If
            End If
        Loop
    Close nFreeFile
    
    GetSections = sReturnString

End Function

Public Function GetKeyData(Tag As String, Optional Delimiter As String)

'** File Variables
Dim nFreeFile As Integer

'** For Data Return
Dim sReturnString As String
Dim sDelimiter As String
Dim sCurrentRecord As String

    '** Set Delimiter
    sDelimiter = Delimiter
    
    '** Open the file one record at a time!!
    nFreeFile = FreeFile
    
    Open psFile For Input As FreeFile
        
        Do While Not EOF(nFreeFile) '** Loop through the text file
            
            '** Get next line from the text file
            Input #nFreeFile, sCurrentRecord
            
            If InStr(1, sCurrentRecord, Tag) <> 0 Then
                If sReturnString = "" Then
                    sReturnString = Left$(sCurrentRecord, InStr(1, sCurrentRecord, Tag) - 1)
                Else
                    sReturnString = sReturnString & sDelimiter & Left$(sCurrentRecord, InStr(1, sCurrentRecord, Tag) - 1)
                End If
            End If
    Loop
    
    GetKeyData = sReturnString

End Function

Public Function FindKeys(Section As String)

'** File Variables
Dim nFreeFile As Integer

'** For Data Return
Dim sReturnString As String
Dim sDelimiter As String
Dim sCurrentRecord As String

    '** Check for Delimiter
    sDelimiter = Delimiter
    
    '** Open the file one record at a time
    nFreeFile = FreeFile
    
    Open psFile For Input As FreeFile
        
        Do While Not EOF(nFreeFile) '** Loop through the text file
            
            '** Get next line from the text file
            Input #nFreeFile, sCurrentRecord
            
            '** Remove un-necessary spaceing before and after
            sCurrentRecord = Trim$(sCurrentRecord)
            
            If IsWantedSection(sCurrentRecord, Section) Then
                Do '** Needs to run at least once
                
                    '** Get next record
                    Input #nFreeFile, sCurrentRecord
                    
                    '** This is a really bad way to do this.. can someone please fix this?
                    If IsSection(sCurrentRecord) Then Exit Do
                    
                    If sCurrentRecord <> "" And InStr(1, sCurrentRecord, KeyTag) > 0 Then
                        If sReturnString = "" Then
                            sReturnString = Left$(sCurrentRecord, InStr(1, sCurrentRecord, KeyTag) - 1)
                        Else
                            sReturnString = sReturnString & sDelimiter & Left$(sCurrentRecord, InStr(1, sCurrentRecord, KeyTag) - 1)
                        End If
                    End If
                    
                Loop Until EOF(nFreeFile)
            
            End If
        Loop
    
    Close nFreeFile
    
    FindKeys = sReturnString

End Function
Public Function IsWantedSection(CurrentSection As String, WantedSection As String) As Boolean

Dim sSectionList() As String
Dim sSections As String
Dim nCount As Integer
Dim sSectionNameWithTags As String
   
    '** Check we are first dealing with a valid section
    If IsSection(CurrentSection) Then
        
        '** Addit to make a section as the user will not be pasing the tags
        sSectionNameWithTags = StartTag & WantedSection & EndTag
        
        If sSectionNameWithTags = CurrentSection Then
            IsWantedSection = True
        Else
            IsWantedSection = False
        End If
    End If

End Function
Public Function IsSection(Data As String) As Boolean

Dim sSectionList() As String
Dim sSections As String
Dim nCount As Integer
Dim sSectionNameWithTags As String

    '** Generate section list
    sSections = GetData
    sSectionList = Split(sSections, "|")
    
    For nCount = 0 To UBound(sSectionList)
        '** Call GetSections and loop through the list to see
        '** if it is actually a section.
        
        sSectionNameWithTags = StartTag & sSectionList(nCount) & EndTag
        
        If sSectionNameWithTags = Data Then
            IsSection = True
            Exit For
        Else
            IsSection = False
        End If
        
    Next nCount

End Function

Public Function GetValue(Section As String, Key As String) As String

'** File Variables
Dim nFreeFile As Integer

'** For Data Return
Dim sReturnString As String
Dim sDelimiter As String
Dim sCurrentRecord As String

    '** Check for Delimiter
    sDelimiter = Delimiter
    
    '** Open the file one record at a time!!
    nFreeFile = FreeFile
    
    Open psFile For Input As FreeFile
        
        Do While Not EOF(nFreeFile) '** Loop through the text file
            
            '** Get next line from the text file
            Input #nFreeFile, sCurrentRecord
            
            '** Remove un-necessary spaceing before and after
            sCurrentRecord = Trim$(sCurrentRecord)
            
            If IsWantedSection(sCurrentRecord, Section) Then
                Do '** Needs to run at least once
                
                    '** Get next record
                    Input #nFreeFile, sCurrentRecord
                    
                    '** This is a really bad way to do this.. can someone please fix this?
                    If IsSection(sCurrentRecord) Then Exit Do
                    
                    If IsWantedKey(sCurrentRecord, Key, Section) Then
                        sReturnString = Right$(sCurrentRecord, Len(sCurrentRecord) - (Len(KeyTag) + Len(Key)))
                        Exit Do
                    End If
                    
                Loop Until EOF(nFreeFile)
            
            End If
        Loop
    
    Close nFreeFile
    
    GetValue = sReturnString

End Function

Public Function IsWantedKey(Data As String, Key As String, Section As String) As Boolean

Dim sCurrentKey As String

    '** Actually check that this line of data is a key
    If IsValidKey(Section, Data) Then
        
        sCurrentKey = Left$(Data, InStr(1, Data, KeyTag) - 1)
        
        '** Ok its a key, but is it the one we want?
        If sCurrentKey = Key Then
            If sCurrentKey = Key Then
                IsWantedKey = True
            Else
                IsWantedKey = False
            End If
        End If
    Else
        IsWantedKey = False
    End If
    
End Function
Public Function IsValidKey(Section As String, Data As String) As Boolean

Dim sListOfValidKeys() As String
Dim nCount As Integer
Dim bReturnValue As String
Dim sCurrentKey As String

    If InStr(1, Data, KeyTag) > 0 Then
        sCurrentKey = Left$(Data, InStr(1, Data, KeyTag) - 1)
        
        '** Get list of keys for this section
        
        sListOfValidKeys = Split(FindKeys(Section), Delimiter)
        
        '** Loop through them to see if the past key is part of it
        For nCount = 0 To UBound(sListOfValidKeys)
            If sListOfValidKeys(nCount) = sCurrentKey Then
                bReturnValue = True
                Exit For
            Else
                bReturnValue = False
            End If
        Next nCount
    Else
        bReturnValue = False
    End If
    
    IsValidKey = bReturnValue
    
End Function
Public Function SetValue(Section As String, Key As String, Value As String) As Boolean

Dim nFileIn As Integer
Dim nFileOut As Integer
Dim sTempFile As String

Dim sCurrentInRecord As String
Dim sCurrentOutRecord As String
Dim sNewRecord As String

    '** Get Free Files
    nFileIn = FreeFile
    
    '** Open for Input
    Open psFile For Input As #nFileIn
    
    sTempFile = Path(TempFolder) & "~" & "data.tmp"
    
    '** Open for Output
    nFileOut = FreeFile
    Open sTempFile For Output As #nFileOut
    
        '** Create Temp file that will be a replicar of the config file
        '** with the exception of the NEW value!!
        
        '** Find desired section
        '** If the current in record is the wanted section
        '** then loop through keys till the wanted key is found
        Do While Not EOF(nFileIn)
            
            '** Get current InRecord
            Input #nFileIn, sCurrentInRecord
            
            If IsWantedSection(sCurrentInRecord, Section) Then
            
                '** Write section header
                sCurrentOutRecord = sCurrentInRecord
                Print #nFileOut, sCurrentOutRecord
                
                '** Get next record
                Input #nFileIn, sCurrentInRecord
                
                Do While Not IsSection(sCurrentInRecord)

                    If IsWantedKey(sCurrentInRecord, Key, Section) Then
                        '** Create string with out the existing value!
                        'sNewRecord = Right$(sCurrentInRecord, InStr(1, sCurrentInRecord, KeyTag) + 1)
                        sCurrentOutRecord = Left$(sCurrentInRecord, InStr(1, sCurrentInRecord, KeyTag) - 1) & KeyTag & Value
                        Print #nFileOut, sCurrentOutRecord
                        Exit Do
                    Else
                        sCurrentOutRecord = sCurrentInRecord
                        Print #nFileOut, sCurrentOutRecord
                    End If
                    
                    '** Get next line
                    Input #nFileIn, sCurrentInRecord
                    
                Loop
            Else
                '** Copy the existing data back to the file!
                sCurrentOutRecord = sCurrentInRecord
                '** Write current out record
                Print #nFileOut, sCurrentOutRecord
            End If

        Loop
    
    Close #nFileIn
    Close #nFileOut
    
    '** Copy back to the existing file
    CopyBack psFile, sTempFile
    
    '** Hide the evidence (delete the tmp file)
    Kill (sTempFile)
    
End Function

Public Function Path(PathAndFile As String)
    
    If Right(PathAndFile, 1) = "\" Then
        Path = PathAndFile
    Else
        Path = PathAndFile & "\"
    End If
    
End Function

Private Sub Class_Initialize()
    
    '** Init Fields
    psFile = ""
    
End Sub

