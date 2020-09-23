VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8705
      _Version        =   393217
      ScrollBars      =   3
      RightMargin     =   1e9
      TextRTF         =   $"Form1.frx":0000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu DummyA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPV 
         Caption         =   "&Preview in Browser"
         Begin VB.Menu mnuPBrowse 
            Caption         =   "Default"
            Index           =   1
         End
         Begin VB.Menu mnuPBrowse 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu DummyZ 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAddB 
            Caption         =   "Edit Browser List"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BrowserReg.vbp
' Created by: Michael Heusdens
'
' Copyright 2001-2002 by Shadows, Inc.
' http://www.shadowsinc.com
'
' INIFile Class Module by: Zebastion aka Zeb
'
' Modifications:
'   I've updated this to check for the latest
' Netscape version and may work for later
' releases. This is, of course, assuming the
' registry traits in Netscape 7 stay consistant
' in future upgrades in Netscape.
'   I also changed the way this gets Netscape
' 6's executable path. Netscape 7's installation
' changes the AppPaths for 6 to 7. In this code,
' we obtain Netscape 6 & 7's execute path in the
' same manner.
'   Other modifications include adding a
' richtextbox. This gives it a feel of an editor.
'
' Description:
'    Checks the registry to see if certain popular
' browsers are installed on the machine. If so, it
' creates registry entries for future reference and
' lists them in the menu.
'
' Looks for the following Browser types:
'    Netscape (Communicator 4 & up), Amaya, Opera,
' and IE
'
' Other features:
'    - Manually add/delete browsers that do not
' have any registry settings.
'
'
' Tested platforms:
'    Windows XP Professional
'    Windows ME
'    Windows 98 SE
'    Windows 95
'
' *Note: Should work in Windows NT 4+
'
' Tested Browser types:
'    Netscape (Communicator 4, 6 (en), and 7.0 beta 1),
' Amaya, Opera, and IE
'
' Fixes:
'    - For Netscape, changed the version check to end
' when "(" is encountered. This gives us "6.2.1" instead
' of "6.2." Encountered that problem with Netscape 6.2's
' version call.
'
'
' Comments:
'   This is simulates how Macromedia retrieves the browsers
' from the registry and places them in its own registry setting.
' The difference here is that instead of using array entries for
' browser names and another set of array entries for program
' location, BrowserReg first locates popular browsers on the
' machine, then places the program locations in its registry
' with the program name as the value name. Much cleaner, if
' you ask me.
'   If you choose to use or modify this, please include me in
' the credits. I'm not sure who wrote the original registry
' module before I modified it, but credit should be given to
' them as well.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Dim Success%

Private Sub Form_Load()

'Look for available browsers on the machine
KeyExists

'Look for manually inputted browsers
If Not FileExist(App.Path & "\" & "browsers.ini") Then
    'this creates the ini file with a dummy key.
    Success% = WritePrivateProfileSection("Browsers", "Default", App.Path & "\" & "browsers.ini")
Else
    LoadINI 0
End If

NoINI:

End Sub

Private Sub Command1_Click()
'From here, we get the short name (dos name) of the location of file
'This fixes the problem Opera has with long name types.
Dim sFile As String, sShortFile As String * 67
Dim lRet

CommonDialog1.Filter = "HTML Files (*.htm, *.html)|*.htm;*.html"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
    sFile = CommonDialog1.FileName
    lRet = GetShortPathName(sFile, sShortFile, Len(sShortFile))
    sFile = Left(sShortFile, lRet)
    Text1.Text = sFile
End If

End Sub

'Set to open document
Private Sub mnuOpen_Click()
'From here, we get the short name (dos name) of the location of file
'This fixes the problem Opera has with long name types.
Dim sFile As String, sShortFile As String * 67
Dim lRet
Dim FileStr As String, TempStr As String

On Error GoTo OpenCanceled

CommonDialog1.FileName = ""
CommonDialog1.Filter = "HTML Files (*.htm, *.html)|*.htm;*.html"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
        Open CommonDialog1.FileName For Input As #1
        FileStr = ""
        Do Until EOF(1)
        Line Input #1, TempStr
        FileStr = FileStr & TempStr & Chr$(13) & Chr$(10)
        Loop
        RichTextBox1.TextRTF = ""
        RichTextBox1.TextRTF = FileStr
        Close #1
        sFile = CommonDialog1.FileName
        lRet = GetShortPathName(sFile, sShortFile, Len(sShortFile))
        sFile = Left(sShortFile, lRet)
        'RichTextBox1.Text = sFile
        Form1.Caption = sFile

End If

OpenCanceled:

End Sub

Private Sub mnuAddB_Click()
Load Form3
Form3.Show
End Sub

Private Sub mnuPBrowse_Click(Index As Integer)
Dim hwnd
Dim opbrowser As String

'Check caption of menu to see if it is the default browser
If mnuPBrowse(Index).Caption = "Default" Then
    'Run default browser
    opbrowser = ShellExecute(hwnd, "open", Form1.Caption, "", "C:\", SW_SHOWNORMAL)
Else
    'It's not... run browser defined by caption
    opbrowser = ShellExecute(hwnd, "open", QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Shadows Inc\Glass Sailor\Browsers\", mnuPBrowse(Index).Caption), Form1.Caption, "C:\", SW_SHOWNORMAL)
End If

End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    RichTextBox1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

End Sub
