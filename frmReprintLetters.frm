VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReprintLetters 
   Caption         =   "Reprint Letters"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   11655
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   8640
         TabIndex        =   11
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdReprint 
         Caption         =   "&Reprint"
         Height          =   375
         Left            =   9600
         TabIndex        =   6
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   10560
         TabIndex        =   5
         Top             =   180
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstReprintLetters 
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "LG No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   410
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12045
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Reprint Letters"
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
         Top             =   80
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   11655
      Begin VB.TextBox txtLetterNo 
         Height          =   285
         Left            =   4440
         TabIndex        =   9
         Top             =   180
         Width           =   1575
      End
      Begin VB.ComboBox cboLetterType 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   180
         Width           =   1935
      End
      Begin VB.Label lblLetter 
         Caption         =   "LG No"
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Types of Letter"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   180
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmReprintLetters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboLetterType_Click()
'
'    If Mid(cboLetterType, 1, 1) = "L" Then
'        lstReprintLetters.ColumnHeaders(1).Text = "LG No"
'        lstReprintLetters.ColumnHeaders(2).Text = "LG Amt"
'        lstReprintLetters.ColumnHeaders(3).Text = "Date of Generation"
'        lstReprintLetters.ColumnHeaders(4).Text = "Issued By"
'    ElseIf Mid(cboLetterType, 1, 1) = "R" Then
'        lstReprintLetters.ColumnHeaders(4).Text = ""
'        lstReprintLetters.ColumnHeaders(1).Text = "Case No"
'        'lstReprintLetters.ColumnHeaders(2).Text = " Amt"
'        lstReprintLetters.ColumnHeaders(2).Text = "Date of Generation"
'        lstReprintLetters.ColumnHeaders(3).Text = "Issued By"
        
'    ElseIf Mid(cboLetterType, 1, 1) = "E" Then
'        lstReprintLetters.ColumnHeaders(1).Text = "Excess No"
'        lstReprintLetters.ColumnHeaders(2).Text = "Excess Amt"
'        lstReprintLetters.ColumnHeaders(3).Text = "Date of Generation"
'        lstReprintLetters.ColumnHeaders(4).Text = "Issued By"

'    End If

End Sub

Private Sub CmdSearch_Click()
'Dim LG As New ADODB.Recordset
'Dim REP As New ADODB.Recordset
'Dim EXC As New ADODB.Recordset
'Dim strsqlLG As String
'Dim strsqlREP As String
'Dim strsqlEXC As String
'Dim MasterReprintList As ListItem

    'If Mid(cboLetterType, 1, 1) = "L" Then
    '    strsqlLG = "select * from CAS_LG where LGNo = '" & Trim(txtLetterNo) & "'"
    '    LG.Open strsqlLG, medixmasterconn, adOpenKeyset, adLockOptimistic
    '
    '    With LG
    '        Set MasterReprintList = lstReprintLetters.listitems.Add(, , !LGNo)
    '            MasterReprintList.SubItems(1) = !LGAmt
    '            MasterReprintList.SubItems(2) = !LGDate
    '            MasterReprintList.SubItems(3) = !LGIssuedNy
    '    End With
    
   ' ElseIf Mid(cboLetterType, 1, 1) = "R" Then
   '     strsqlREP = "select * from CASRepudiation where CaseNumber = '" & Trim(txtLetterNo) & "'"
   '     REP.Open strsqlREP, medixmasterconn, adOpenKeyset, adLockOptimistic
   '
   '     With LG
   '         Set MasterReprintList = lstReprintLetters.listitems.Add(, , !LGNo)
   '             MasterReprintList.SubItems(1) = !LGAmt
   '             MasterReprintList.SubItems(2) = !LGDate
   '             MasterReprintList.SubItems(3) = !LGIssuedNy
End Sub

Private Sub Form_Load()
'    cboLetterType.AddItem "L" & "-" & "Letter of Gurantee"
'    cboLetterType.AddItem "E" & "-" & "Excess Letter"
'    cboLetterType.AddItem "R" & "-" & "Repudiation Letter"
End Sub

