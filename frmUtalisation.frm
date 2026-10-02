VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUtalisation 
   Caption         =   "Utalisation"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   9180
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   630
      Left            =   0
      TabIndex        =   2
      Top             =   5160
      Width           =   9135
      Begin VB.CommandButton CmdSearch 
         Cancel          =   -1  'True
         Caption         =   "&Search"
         Height          =   375
         Left            =   7080
         TabIndex        =   10
         Top             =   240
         Width           =   850
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   8160
         TabIndex        =   7
         Top             =   240
         Width           =   850
      End
   End
   Begin MSComctlLib.ListView lstUtilization 
      Height          =   4095
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   7223
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hospitals"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bill Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Approved Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "DOA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DOD"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   9135
      Begin VB.ComboBox cboCRI 
         Height          =   315
         Left            =   720
         TabIndex        =   9
         Top             =   180
         Width           =   2535
      End
      Begin VB.ComboBox cboSEL 
         Height          =   315
         Left            =   4320
         TabIndex        =   5
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Criteria"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Selection"
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Utilization"
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
         Left            =   225
         TabIndex        =   1
         Top             =   45
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmUtalisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboListing()
    cboCRI.AddItem "0" & " - " & "ICD"
    cboCRI.AddItem "1" & " - " & "Doctor"
    cboCRI.AddItem "2" & " - " & "Hospital"
End Sub

Private Sub cboCRI_Click()
    
    Dim ICD As New ADODB.Recordset
    Dim HOS As New ADODB.Recordset

    HOS.Open "select HosCode, HosName from HOS", medixmasterconn, adOpenKeyset, adLockOptimistic
    ICD.Open "select ICDCOde, ICDDescription from ICD ", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If Mid(cboCRI, 1, 1) = 0 Then
       Label4.Caption = "ICD"
       Do Until ICD.EOF
          cboSEL.AddItem ICD!ICDCode '& " - " & Trim(ICD!ICDDescription)
          ICD.MoveNext
       Loop
    'lstUtilization.ColumnHeaders(1).Text = "Hospital"
    'lstUtilization.ColumnHeaders(2).Text = "BillAmount"
    'lstUtilization.ColumnHeaders(3).Text = "Hospital"
    
    ElseIf Mid(cboCRI, 1, 1) = 1 Then
       Label4.Caption = "Doctor"
    
    ElseIf Mid(cboCRI, 1, 1) = 2 Then
       Label4.Caption = "Hospital"
       cboSEL.Clear
       Do Until HOS.EOF
          cboSEL.AddItem HOS!HOSName '& " - " & Trim(HOS!HOSName)
          HOS.MoveNext
       Loop
    End If
       
    ICD.Close
    HOS.Close
    
    Set ICD = Nothing
    Set HOS = Nothing
    
End Sub

Private Sub CmdExit_Click()
Dim RecordExit
    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        bExit = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call ComboListing
End Sub
