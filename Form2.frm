VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add a Browser"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3825
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdBrowse_Click()

CommonDialog1.Filter = "Executable Files (*.exe)|*.exe"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
    txtLoc.Text = CommonDialog1.FileName
End If

End Sub

Private Sub cmdCancel_Click()
'clear text
txtLoc.Text = ""
txtName.Text = ""
Unload Me
End Sub

Private Sub cmdOK_Click()

If txtName.Text = "" Then
    MsgBox "You need to enter the program name!", vbOKOnly
    Exit Sub
End If
If txtLoc.Text = "" Then
    MsgBox "You need to enter the program location!", vbOKOnly
    Exit Sub
End If

If txtName.Enabled = False Then
    'Only editing program's location
    SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Shadows Inc\Glass Sailor\Browsers\", txtName.Text, txtLoc.Text, REG_SZ
Else
    'Place name in ini file
    SetValueINI txtName.Text
    'add to menu
    AddToMenu txtName.Text
    'Set name and location in registry
    SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Shadows Inc\Glass Sailor\Browsers\", txtName.Text, txtLoc.Text, REG_SZ
    
    'Refresh the listbox
    LoadINI 1
End If

txtName.Text = ""
txtLoc.Text = ""
Unload Me
End Sub

