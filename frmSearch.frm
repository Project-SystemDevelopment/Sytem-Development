VERSION 5.00
Begin VB.Form frmSearchLG 
   Caption         =   "Retrieve LG"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   9150
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   7200
      Width           =   11655
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10080
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   11655
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtDirectory 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5520
         TabIndex        =   9
         Top             =   600
         Width           =   3615
      End
      Begin VB.ComboBox cboFilecode 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cboDircode 
         Height          =   315
         ItemData        =   "frmSearch.frx":0000
         Left            =   5520
         List            =   "frmSearch.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label lblDirName 
         Caption         =   "Current Path"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   14
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "File Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "File Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Template Folder"
         Height          =   255
         Left            =   4200
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.FileListBox filFile 
      Height          =   5355
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   11655
   End
   Begin VB.DirListBox DirDirectory 
      Height          =   540
      Left            =   480
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Retrieve LG"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1365
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Double Click to Open Selected File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Template Folder"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmSearchLG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim path As String
Dim Extension As String

Private Sub cboDircode_Click()
    Screen.MousePointer = vbHourglass
    filFile.Enabled = False
    If cboDircode.ListIndex = "0" Then
        DirDirectory.path = crdPath + "lg"
    ElseIf cboDircode.ListIndex = "1" Then
        DirDirectory.path = crdPath + "repudiation"
    End If
    filFile.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cbofilecode_Click()
    Select Case cboFilecode.ListIndex
    Case 0
    Extension = ".*"
    filFile.Pattern = "*.*"
    Case 1
    Extension = ".txt"
    filFile.Pattern = "*.txt"
    
    Case 2
    Extension = ".doc"
    filFile.Pattern = "*.doc"
    
    End Select
    txtFileName = ""
End Sub


Private Sub cboFilecode_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



Private Sub CmdExit_Click()
Dim RecordExit
    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        'bExit = True
        Unload Me
    End If
End Sub


Private Sub DirDirectory_Change()
path = DirDirectory.path
txtDirectory.Text = DirDirectory.path
filFile.path = DirDirectory.path
End Sub

'Private Sub Drive1_Change()
'On Error GoTo driveError

'DirDirectory.path = Drive1.Drive
'Exit Sub
'driveError:
'MsgBox " Drive error or not ready!", vbOKOnly + vbExclamation, DialogBoxTitle
'Drive1.Drive = DirDirectory.path
'Exit Sub

'End Sub

Private Sub filFile_Click()
    
    txtFileName.Text = filFile.filename

End Sub

Private Sub filFile_DblClick()
On Error Resume Next
Dim str As String
Screen.MousePointer = vbHourglass
str = DirDirectory.path + "\" + filFile.filename
filFile.Enabled = False
If objword = "" Or objword = Null Or objword = 0 Then
    Set objword = New word.Application
    objword.Visible = True
    Set obdoc = objword.Documents.Add(str)
Else
    objword.Visible = True
    Set obdoc = objword.Documents.Add(str)
End If
filFile.Enabled = True
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
On Error GoTo patherror
DirDirectory.path = crdPath & "lg" 'open default directory
txtFileName = ""
Extension = ""

cboFilecode.AddItem "All files (*.*)"
cboFilecode.AddItem "Text files (*.txt)"
cboFilecode.AddItem "Doc files (*.doc)"
cboFilecode.ListIndex = 0
'DirDirectory.Visible = False
'Drive1.Visible = False

cboDircode.AddItem "Letter Of Guarantee"
'cboDircode.AddItem "ASCII"
cboDircode.AddItem "Repudiation Letter"
cboDircode.ListIndex = 0

Exit Sub

patherror: 'happens when path not found
  If ERR.Number = "76" Then
     MsgBox "Default directory " & crdPath & "LG\ not found", vbExclamation + vbOKOnly, DialogBoxTitle
  Else
     MsgBox ERR.Number & "-" & ERR.Description, vbOKOnly + vbExclamation, DialogBoxTitle
  End If
End Sub

Private Sub word(Optional str As String)
    Set objword = New word.Application
    objword.Visible = True
    Set objdoc = objword.Documents.Add(str)
End Sub



Private Sub txtFileName_Change()
If txtFileName = "" Then
filFile.Pattern = "*" + Extension

End If
End Sub
Private Sub txtFileName_KeyPress(KeyAscii As Integer)
Dim newtext As String

If KeyAscii = 46 Then
    KeyAscii = 0
    ElseIf KeyAscii = 8 And txtFileName = "" Then
    KeyAscii = 0
    Else
    KeyAscii = KeyAscii
    newtext = txtFileName.Text + Chr(KeyAscii)
End If


filFile.Pattern = newtext + "*" + Extension



End Sub
