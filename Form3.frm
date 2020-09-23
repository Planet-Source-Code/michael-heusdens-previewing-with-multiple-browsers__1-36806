VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editable Browser List"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4080
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdEd 
      Caption         =   "Edit..."
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdRm 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Browsers:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()
List1.Clear
Unload Me
End Sub

Private Sub Form_Load()

If Not FileExist(App.Path & "\browsers.ini") Then
    GoTo NoINI
Else
    LoadINI 1
End If

NoINI:

End Sub

Private Sub cmdAdd_Click()
Load Form2
Form2.Show vbModal
End Sub

Private Sub cmdEd_Click()
Dim Retlst As String
'get location
Retlst = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Shadows Inc\Glass Sailor\Browsers\", List1.Text)
Load Form2
Form2.txtName.Enabled = False
Form2.txtName = List1.Text
Form2.txtLoc = Retlst
Form2.Show vbModal
End Sub

Private Sub cmdRm_Click()
RemoveFromINI List1.Text
End Sub
