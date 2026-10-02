VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClaimReminder 
   Caption         =   "frmClaimReminder"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   40
      Top             =   7320
      Width           =   11775
      Begin VB.CommandButton cmdExit 
         Appearance      =   0  'Flat
         Caption         =   "&Exit"
         Height          =   285
         Left            =   10920
         TabIndex        =   27
         Top             =   240
         Width           =   645
      End
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         Caption         =   "&Save"
         Height          =   285
         Left            =   10320
         TabIndex        =   26
         Top             =   240
         Width           =   645
      End
      Begin VB.CommandButton cmdEdit 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   285
         Left            =   9720
         TabIndex        =   25
         Top             =   240
         Width           =   645
      End
      Begin VB.CommandButton cmdAdd 
         Appearance      =   0  'Flat
         Caption         =   "&Add"
         Height          =   285
         Left            =   9120
         TabIndex        =   24
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   37
      Top             =   240
      Width           =   11775
      Begin VB.ComboBox cboVersion 
         Height          =   315
         Left            =   3000
         TabIndex        =   1
         Top             =   195
         Width           =   795
      End
      Begin VB.TextBox txtCASNumber 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   0
         Top             =   195
         Width           =   1575
      End
      Begin VB.CommandButton cmdShow 
         Appearance      =   0  'Flat
         Caption         =   "&Go"
         Height          =   285
         Left            =   4080
         TabIndex        =   2
         Top             =   195
         Width           =   645
      End
      Begin MSMask.MaskEdBox txtREMDate 
         Height          =   285
         Left            =   7920
         TabIndex        =   3
         Top             =   195
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label19 
         Caption         =   "Case Number"
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
         Left            =   240
         TabIndex        =   39
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Reminder Date"
         Height          =   255
         Left            =   6600
         TabIndex        =   38
         Top             =   225
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Member Details"
      TabPicture(0)   =   "frmClaimReminder.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lstMBM"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Reminder Details"
      TabPicture(1)   =   "frmClaimReminder.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtAdd3"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtAdd2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtAdd1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtAdmissionDate"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtDischargeDate"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lstReminder"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Frame3"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtPrincipleName"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtPatientName"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtPolicyNo"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtMBMNo"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      Begin VB.TextBox txtMBMNo 
         Height          =   285
         Left            =   2800
         TabIndex        =   5
         Top             =   420
         Width           =   2000
      End
      Begin VB.TextBox txtPolicyNo 
         Height          =   285
         Left            =   8040
         TabIndex        =   8
         Top             =   420
         Width           =   2000
      End
      Begin VB.TextBox txtPatientName 
         Height          =   285
         Left            =   8040
         TabIndex        =   9
         Top             =   900
         Width           =   2000
      End
      Begin VB.TextBox txtPrincipleName 
         Height          =   285
         Left            =   2800
         TabIndex        =   6
         Top             =   900
         Width           =   2000
      End
      Begin VB.Frame Frame3 
         Caption         =   "Actions"
         Height          =   1935
         Left            =   120
         TabIndex        =   30
         Top             =   1740
         Width           =   11500
         Begin VB.CheckBox chkRMF 
            Caption         =   "Member Reimbursement Form"
            Height          =   275
            Left            =   450
            TabIndex        =   11
            Top             =   360
            Width           =   3000
         End
         Begin VB.CheckBox chkOrinInv 
            Caption         =   "Original Invoice"
            Height          =   275
            Left            =   4575
            TabIndex        =   12
            Top             =   360
            Width           =   3000
         End
         Begin VB.CheckBox chkMedicalForm 
            Caption         =   "Complete Medical Form"
            Height          =   275
            Left            =   8100
            TabIndex        =   13
            Top             =   360
            Width           =   3000
         End
         Begin VB.CheckBox chkProposalForm 
            Caption         =   "Proposal Form"
            Height          =   275
            Left            =   450
            TabIndex        =   14
            Top             =   720
            Width           =   3000
         End
         Begin VB.CheckBox chkRefLetter 
            Caption         =   "Referral Letter"
            Height          =   275
            Left            =   4575
            TabIndex        =   15
            Top             =   720
            Width           =   3000
         End
         Begin VB.CheckBox chkSubroForm 
            Caption         =   "Form of Subrogation Receipt"
            Height          =   275
            Left            =   8100
            TabIndex        =   16
            Top             =   720
            Width           =   3000
         End
         Begin VB.TextBox txtOther1 
            Height          =   285
            Left            =   1400
            MaxLength       =   50
            TabIndex        =   21
            Top             =   1440
            Width           =   3500
         End
         Begin VB.TextBox txtOther2 
            Height          =   285
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   23
            Top             =   1440
            Width           =   3500
         End
         Begin VB.CheckBox chkRec 
            Caption         =   "Offical Receipt"
            Height          =   255
            Left            =   450
            TabIndex        =   17
            Top             =   1080
            Width           =   3000
         End
         Begin VB.CheckBox chkStudent 
            Caption         =   "Student ID"
            Height          =   255
            Left            =   4575
            TabIndex        =   18
            Top             =   1080
            Width           =   3000
         End
         Begin VB.CheckBox chkPolice 
            Caption         =   "Police Report"
            Height          =   255
            Left            =   8100
            TabIndex        =   19
            Top             =   1080
            Width           =   3000
         End
         Begin VB.CheckBox chkOther1 
            Caption         =   "Other 1"
            Height          =   275
            Left            =   450
            TabIndex        =   20
            Top             =   1440
            Width           =   3000
         End
         Begin VB.CheckBox chkOther2 
            Caption         =   "Other 2"
            Height          =   275
            Left            =   6480
            TabIndex        =   22
            Top             =   1440
            Width           =   3000
         End
      End
      Begin MSComctlLib.ListView lstReminder 
         Height          =   2535
         Left            =   120
         TabIndex        =   28
         Top             =   3720
         Width           =   11500
         _ExtentX        =   20294
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "REMNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CASNumber"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "RMF"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Orig Inv."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Medical Form"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Proposal Form"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Ref Letter"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Orig Receipt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Student ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Police Report"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Subrogation From"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Others 1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Others 2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "CovID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "REMDate"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSMask.MaskEdBox txtDischargeDate 
         Height          =   285
         Left            =   8040
         TabIndex        =   10
         Top             =   1380
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtAdmissionDate 
         Height          =   285
         Left            =   2800
         TabIndex        =   7
         Top             =   1380
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ListView lstMBM 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   29
         Top             =   540
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   9763
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Member No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cover ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Patient Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Policy No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date of Admission"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Date of Discharge"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtAdd1 
         Height          =   285
         Left            =   2800
         TabIndex        =   42
         Top             =   420
         Visible         =   0   'False
         Width           =   2000
      End
      Begin VB.TextBox txtAdd2 
         Height          =   285
         Left            =   2800
         TabIndex        =   43
         Top             =   900
         Visible         =   0   'False
         Width           =   2000
      End
      Begin VB.TextBox txtAdd3 
         Height          =   285
         Left            =   2800
         TabIndex        =   44
         Top             =   1380
         Visible         =   0   'False
         Width           =   2000
      End
      Begin VB.Label Label1 
         Caption         =   "Membership No."
         Height          =   255
         Left            =   1300
         TabIndex        =   36
         Top             =   420
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Policy No."
         Height          =   255
         Left            =   6480
         TabIndex        =   35
         Top             =   420
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Principle Name"
         Height          =   255
         Left            =   1305
         TabIndex        =   34
         Top             =   900
         Width           =   1395
      End
      Begin VB.Label Label4 
         Caption         =   "Patient Name"
         Height          =   255
         Left            =   6480
         TabIndex        =   33
         Top             =   900
         Width           =   1395
      End
      Begin VB.Label Label5 
         Caption         =   "Date of Admission"
         Height          =   255
         Left            =   1320
         TabIndex        =   32
         Top             =   1380
         Width           =   1395
      End
      Begin VB.Label Label6 
         Caption         =   "Date of Discharge"
         Height          =   255
         Left            =   6480
         TabIndex        =   31
         Top             =   1380
         Width           =   1395
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Claim Reminder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmClaimReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean
Dim NewFlag As Boolean
Dim FinalFlag As Boolean
Dim ShowFlag As Boolean
Dim REMCount As String
Dim TempREM As String
Dim Letter As String
Dim objword As Word.Application
Dim objdoc As Word.Document

Private Sub cboVersion_DropDown()
Dim CLM As New ADODB.Recordset
Dim strsqlCLM As String

    If txtCASNumber <> "" Then
        cboVersion.Clear
        medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        
        strsqlCLM = "EXEC CLM_CLMReminder '" & Trim(txtCASNumber) & "', 2"
        CLM.Open strsqlCLM, medixmasterconnWS, adOpenForwardOnly, adLockOptimistic
        
        If Not CLM.EOF Then
            With CLM
            Do Until .EOF
                cboVersion.AddItem !claimversion
                .MoveNext
            Loop
            End With
            cboVersion = 1
        Else
            cboVersion = 1
        End If
        
        CLM.Close
        Set CLM = Nothing
        medixmasterconnWS.Close
        Set medixmasterconnWS = Nothing
    End If
    
End Sub

Private Sub cboVersion_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboVersion.Text, cboVersion.SelStart) + Chr(KeyAscii) + Mid(cboVersion.Text, cboVersion.SelStart + cboVersion.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboVersion.ListCount - 1
 
        If NewText = LCase(Left(cboVersion.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboVersion.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboVersion.Text = ValidValue
                cboVersion.SelStart = 0
                cboVersion.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CmdAdd_Click()
    If txtCASNumber = "" Then
        MsgBox "Please enter Case Number", vbOKOnly + vbCritical, DialogBoxTitle
        txtCASNumber.SetFocus
        Exit Sub
    End If
    EditFlag = False
    Frame3.Enabled = True
    txtREMDate.Enabled = True
    chkRMF.SetFocus
    cmdSave.Enabled = True
    cmdEdit.Enabled = False
End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
    Frame3.Enabled = True
    txtREMDate.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    cmdEdit.Enabled = False
    medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    Call SaveREM
    medixmasterconnWS.Close
    Set medixmasterconnWS = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
End Sub

Private Sub cmdShow_Click()
Dim CLM As New ADODB.Recordset
Dim strsqlCLM As String

    cmdSave.Enabled = False
    cmdEdit.Enabled = False
    cmdAdd.Enabled = True
    Frame3.Enabled = False
    
    If txtCASNumber = "" Then
        MsgBox "Please key in the Case No.!", vbCritical
        txtCASNumber.SetFocus
    Else
        medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        
        strsqlCLM = "EXEC CLM_CLMReminder '" & Trim(txtCASNumber) & "', 2"
        CLM.Open strsqlCLM, medixmasterconnWS, adOpenForwardOnly, adLockOptimistic
        
        If CLM.EOF Then
            MsgBox "Case Number Not Found!", vbOKOnly + vbCritical, DialogBoxTitle
            Call CleanForm
            lstMBM.listitems.Clear
            lstReminder.listitems.Clear
            cboVersion = 1
            txtCASNumber.SetFocus
            CLM.Close
            Set CLM = Nothing
            medixmasterconnWS.Close
            Set medixmasterconnWS = Nothing
            Exit Sub
        End If

        Call CleanForm
        Call showMBM
        Call ShowMBMDetails
        Call showReminder
        CLM.Close
        Set CLM = Nothing
        medixmasterconnWS.Close
        Set medixmasterconnWS = Nothing
    End If
End Sub

Private Sub Form_Load()
    txtMBMNo.Enabled = False
    txtPolicyNo.Enabled = False
    txtPrincipleName.Enabled = False
    txtPatientName.Enabled = False
    txtAdmissionDate.Enabled = False
    txtDischargeDate.Enabled = False
    txtREMDate.Enabled = False
    Frame3.Enabled = False
    lstReminder.Enabled = False
    cmdEdit.Enabled = False
    cmdSave.Enabled = False
    cboVersion = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
            Case vbKeyReturn
                cmdShow_Click
        End Select
    End If
End Sub

Private Sub lstReminder_Click()
Dim MBM As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim strAppend As String
Dim strAppendCase As String

If lstReminder.listitems.count = 0 Then
    lstReminder.Enabled = False
Else
    Frame3.Enabled = False
    txtREMDate.Enabled = False
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
    
    chkRMF.value = IIf(lstReminder.SelectedItem.SubItems(2) = " ", 0, 1)
    chkOrinInv.value = IIf(lstReminder.SelectedItem.SubItems(3) = " ", 0, 1)
    chkMedicalForm.value = IIf(lstReminder.SelectedItem.SubItems(4) = " ", 0, 1)
    chkProposalForm.value = IIf(lstReminder.SelectedItem.SubItems(5) = " ", 0, 1)
    chkRefLetter.value = IIf(lstReminder.SelectedItem.SubItems(6) = " ", 0, 1)
    chkRec.value = IIf(lstReminder.SelectedItem.SubItems(7) = " ", 0, 1)
    chkStudent.value = IIf(lstReminder.SelectedItem.SubItems(8) = " ", 0, 1)
    chkPolice.value = IIf(lstReminder.SelectedItem.SubItems(9) = " ", 0, 1)
    chkSubroForm.value = IIf(lstReminder.SelectedItem.SubItems(10) = " ", 0, 1)
    chkOther1.value = IIf(lstReminder.SelectedItem.SubItems(11) = "", 0, 1)
    chkOther2.value = IIf(lstReminder.SelectedItem.SubItems(12) = "", 0, 1)
    
    If chkOther1.value = 1 Then
        txtOther1 = lstReminder.SelectedItem.SubItems(11)
    Else
        txtOther1 = ""
    End If
    If chkOther2.value = 1 Then
        txtOther2 = lstReminder.SelectedItem.SubItems(12)
    Else
        txtOther2 = ""
    End If
    
    txtREMDate = Format(lstReminder.SelectedItem.SubItems(14), "dd/mm/yyyy")

End If
End Sub

Private Sub lstReminder_DblClick()
    medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
    EditFlag = False
    TempREM = lstReminder.SelectedItem.Text
    If lstReminder.SelectedItem.Text = 3 Then
        Letter = "Final Reminder Letter.dot"
        FinalFlag = True
    Else
        Letter = "Reminder Letter.dot"
        FinalFlag = False
    End If
    Call funcword
    
    medixmasterconnWS.Close
    Set medixmasterconnWS = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        
    ElseIf SSTab1.Tab = 1 Then
        
    End If
End Sub

Private Sub showMBM()
Dim MBM As New ADODB.Recordset
Dim strsql As String
Dim MBMListing As ListItem

    strsql = "Exec CLM_CLMReminder '" & Trim(txtCASNumber) & "',1"
    MBM.Open strsql, medixmasterconnWS, adOpenForwardOnly, adLockOptimistic
    
    lstMBM.listitems.Clear
    If Not MBM.EOF Then
        With MBM
            Set MBMListing = lstMBM.listitems.Add(, , !MBMNumber)
            
            If Len(Trim(!MBMCOVID)) <> 0 Then
                MBMListing.SubItems(1) = "00"
            End If
            If Len(Trim(!MBMName)) <> 0 Then
                MBMListing.SubItems(2) = ChkStr(!MBMName)
            End If
            If Len(Trim(!MBMPolicyNo)) <> 0 Then
                MBMListing.SubItems(3) = !MBMPolicyNo
            End If
            If Len(Trim(!CASDateAdmitted)) Then
                MBMListing.SubItems(4) = !CASDateAdmitted
            End If
            If Len(Trim(!CasDischargeDate)) Then
                MBMListing.SubItems(5) = !CasDischargeDate
            End If
            If Trim(!INSCode) = "F" Or Trim(!INSCode) = "H" Then GoTo MBMCov
        End With
    Else
        MsgBox "Record not found!", vbOKOnly + vbCritical, DialogBoxTitle
        txtCASNumber.SetFocus
    End If
    MBM.Close
    Set MBM = Nothing
    Exit Sub
    
MBMCov:
    Do Until MBM.EOF
        With MBM
            Set MBMListing = lstMBM.listitems.Add(, , !MBMNumber)
            
            If Len(Trim(!MBMCCoverID)) <> 0 Then
                MBMListing.SubItems(1) = !MBMCCoverID
            End If
            If Len(Trim(!MBMCName)) <> 0 Then
                MBMListing.SubItems(2) = ChkStr(!MBMCName)
            End If
            If Len(Trim(!MBMPolicyNo)) <> 0 Then
                MBMListing.SubItems(3) = !MBMPolicyNo
            End If
            If Len(Trim(!CASDateAdmitted)) Then
                MBMListing.SubItems(4) = !CASDateAdmitted
            End If
            If Len(Trim(!CasDischargeDate)) Then
                MBMListing.SubItems(5) = !CasDischargeDate
            End If
            .MoveNext
        End With
    Loop

End Sub

Private Sub ShowMBMDetails()
Dim MBM As New ADODB.Recordset
Dim strsql As String
    
    strsql = "Exec CLM_CLMReminder '" & Trim(txtCASNumber) & "',1"
    MBM.Open strsql, medixmasterconnWS, adOpenForwardOnly, adLockOptimistic
    
    If Not MBM.EOF Then
    Do Until MBM.EOF
        If StrComp(MBM!MBMPatientCovID, MBM!MBMCOVID) = 0 Then
            With MBM
                txtMBMNo = !MBMNumber
                txtPolicyNo = !MBMPolicyNo
                txtPatientName = !MBMName
                txtPrincipleName = !MBMName
                txtAdmissionDate = Format(!CASDateAdmitted, "dd/mm/yyyy")
                If !CasDischargeDate <> "" Then
                    txtDischargeDate = Format(!CasDischargeDate, "dd/mm/yyyy")
                Else
                    txtDischargeDate = "__/__/____"
                End If
                If !MBMAddress1 <> "" Then
                    txtAdd1 = !MBMAddress1
                End If
                If !MBMAddress2 <> "" Then
                    txtAdd2 = !MBMAddress2
                End If
                If !MBMAddress3 <> "" Then
                    txtAdd3 = !MBMAddress3
                End If
            End With
        ElseIf StrComp(MBM!MBMPatientCovID, MBM!MBMCCoverID) = 0 Then
            With MBM
                txtMBMNo = !MBMNumber
                txtPolicyNo = !MBMPolicyNo
                txtPatientName = !MBMCName
                txtPrincipleName = !MBMName
                txtAdmissionDate = Format(!CASDateAdmitted, "dd/mm/yyyy")
                If !CasDischargeDate <> "" Then
                    txtDischargeDate = Format(!CasDischargeDate, "dd/mm/yyyy")
                Else
                    txtDischargeDate = "__/__/____"
                End If
                If !MBMAddress1 <> "" Then
                    txtAdd1 = !MBMAddress1
                End If
                If !MBMAddress2 <> "" Then
                    txtAdd2 = !MBMAddress2
                End If
                If !MBMAddress3 <> "" Then
                    txtAdd3 = !MBMAddress3
                End If
            End With
        End If
        MBM.MoveNext
    Loop
    End If
    MBM.Close
    Set MBM = Nothing

End Sub

Private Sub showReminder()
Dim REMD As New ADODB.Recordset
Dim strsqlREMD As String
Dim strsql As String
Dim REMListing As ListItem

    strsql = "" & Trim(txtCASNumber) & ", " & Trim(cboVersion) & ""
    strsqlREMD = "Exec CLM_CLMReminder '" & strsql & "', 3"
    REMD.Open strsqlREMD, medixmasterconnWS, adOpenForwardOnly, adLockOptimistic
    
    txtREMDate = Format(Date, "dd/mm/yyyy")
    lstReminder.listitems.Clear
    
    If Not REMD.EOF Then
        lstReminder.Enabled = True
        Do Until REMD.EOF
            NewFlag = False
            With REMD
            Set REMListing = lstReminder.listitems.Add(, , !REMNumber)
                
            If Len(!CasNumber) Then
                REMListing.SubItems(1) = !CasNumber
            End If
            If Len(!REMRMF) Then
                REMListing.SubItems(2) = !REMRMF
            End If
            If Len(!REMOrinInv) Then
                REMListing.SubItems(3) = !REMOrinInv
            End If
            If Len(!REMRpt) Then
                REMListing.SubItems(4) = !REMRpt
            End If
            If Len(!REMProposalForm) Then
                REMListing.SubItems(5) = !REMProposalForm
            End If
            If Len(!REMRefLetter) Then
                REMListing.SubItems(6) = !REMRefLetter
            End If
            If Len(!REMReceipt) Then
                REMListing.SubItems(7) = !REMReceipt
            End If
            If Len(!REMStudent) Then
                REMListing.SubItems(8) = !REMStudent
            End If
            If Len(!REMPolice) Then
                REMListing.SubItems(9) = !REMPolice
            End If
            If Len(!REMSubroForm) Then
                REMListing.SubItems(10) = !REMSubroForm
            End If
            If Len(!REMOthers1) Then
                REMListing.SubItems(11) = !REMOthers1
            End If
            If Len(!REMOthers2) Then
                REMListing.SubItems(12) = !REMOthers2
            End If
            If Len(!REMCOVID) Then
                REMListing.SubItems(13) = !REMCOVID
            End If
            If Len(!REMDate) Then
                REMListing.SubItems(14) = !REMDate
            End If
            REMCount = !REMNumber
            End With
            REMD.MoveNext
        Loop
    Else
        REMCount = 0
        NewFlag = True
    End If

End Sub

Private Sub SaveREM()
Dim REMD As New ADODB.Recordset
Dim RecordSaved As String
Dim RMF, INV, MED, PROPOSAL As String
Dim REF, SUBRO, OTHER1, OTHER2 As String
Dim REC, STU, POLICE As String
Dim TempDate As String
Dim TempCovID As String
Dim TempREMDate As String
Dim Tempstr As String
Dim strsql As String
Dim strsqlREM As String

    If txtCASNumber = "" Then
        MsgBox "Please Enter the Case Number", vbOKOnly + vbCritical, DialogBoxTitle
        txtCASNumber.SetFocus
        Exit Sub
    End If
    If cboVersion = "" Then
        MsgBox "Please select Claim Version Number", vbOKOnly + vbCritical, DialogBoxTitle
        SSTab1.Tab = 0
        Exit Sub
    End If
    If txtMBMNo = "" Then
        MsgBox "Please select a member", vbOKOnly + vbCritical, DialogBoxTitle
        SSTab1.Tab = 0
        Exit Sub
    End If
    If chkOther1.value = 1 Then
        If txtOther1 = "" Then
            MsgBox "Please enter your requested form", vbOKOnly + vbCritical, DialogBoxTitle
            txtOther1.SetFocus
            Exit Sub
        End If
    End If
    If chkOther2.value = 1 Then
        If txtOther2 = "" Then
            MsgBox "Please enter your requested form", vbOKOnly + vbCritical, DialogBoxTitle
            txtOther2.SetFocus
            Exit Sub
        End If
    End If
    If txtREMDate = "__/__/____" Then
        MsgBox "Please enter the Reminder Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtREMDate.SetFocus
        Exit Sub
    End If
    If IsDate(txtREMDate) = False Then
        MsgBox "Invalid Date Format!", vbOKOnly + vbCritical, DialogBoxTitle
        txtREMDate.SetFocus
        Exit Sub
    End If
        
    If chkRMF.value = 1 Then
        RMF = "Y"
    End If
    If chkOrinInv.value = 1 Then
        INV = "Y"
    End If
    If chkMedicalForm.value = 1 Then
        MED = "Y"
    End If
    If chkProposalForm.value = 1 Then
        PROPOSAL = "Y"
    End If
    If chkRefLetter.value = 1 Then
        REF = "Y"
    End If
    If chkRec.value = 1 Then
        REC = "Y"
    End If
    If chkStudent.value = 1 Then
        STU = "Y"
    End If
    If chkPolice.value = 1 Then
        POLICE = "Y"
    End If
    If chkSubroForm.value = 1 Then
        SUBRO = "Y"
    End If
    If chkOther1.value = 1 Then
        OTHER1 = Replace(ChkStr(txtOther1), "'", "''")
    End If
    If chkOther2.value = 1 Then
        OTHER2 = Replace(ChkStr(txtOther2), "'", "''")
    End If
        
    TempDate = Format(Date, "yyyy/mm/dd")
    TempREMDate = Format(txtREMDate, "yyyy/mm/dd")
    TempCovID = lstMBM.SelectedItem.SubItems(1)
    Tempstr = "" & Trim(txtCASNumber) & ", " & Trim(cboVersion) & ""
    
    strsqlREM = "Exec CLM_CLMReminder '" & Tempstr & "',3"
    REMD.Open strsqlREM, medixmasterconnWS, adOpenForwardOnly, adLockOptimistic
    If Not REMD.EOF Then
        If EditFlag = False And NewFlag = False Then
            TempREM = REMCount + 1
        ElseIf EditFlag = True And NewFlag = False Then
            TempREM = lstReminder.SelectedItem.Text
        End If
    Else
        TempREM = 0
    End If
    REMD.Close
    Set REMD = Nothing
        
    If TempREM > 3 Then
        MsgBox "There is no further reminder should be sent out.", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    ElseIf TempREM = 3 Then
        Letter = "Final Reminder Letter.dot"
        FinalFlag = True
    ElseIf TempREM >= 0 And TempREM < 3 Then
        Letter = "Reminder Letter.dot"
        FinalFlag = False
    End If
        
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            medixmasterconnWS.beginTrans
            
            If EditFlag = False Then
                strsql = "EXEC CLM_SAVEReminder '" & TempREM & "', '" & ChkStr(Trim(txtCASNumber)) & "', " & _
                        "'" & RMF & "', '" & INV & "', '" & MED & "', '" & PROPOSAL & "', " & _
                        "'" & REF & "', '" & SUBRO & "', '" & OTHER1 & "', '" & OTHER2 & "', " & _
                        "'" & TempDate & "', '" & Logon & "', '" & TempCovID & "', '" & TempREMDate & "', " & _
                        "'" & REC & "', '" & STU & "', '" & POLICE & "', '" & Trim(cboVersion) & "'"
            Else
                strsql = "EXEC CLM_UPDATEReminder '" & TempREM & "', '" & ChkStr(Trim(txtCASNumber)) & "', " & _
                        "'" & RMF & "', '" & INV & "', '" & MED & "', '" & PROPOSAL & "', " & _
                        "'" & REF & "', '" & SUBRO & "', '" & OTHER1 & "', '" & OTHER2 & "', " & _
                        "'" & TempDate & "', '" & Logon & "', '" & TempREMDate & "', " & _
                        "'" & REC & "', '" & STU & "', '" & POLICE & "', '" & Trim(cboVersion) & "'"
            End If
            
            medixmasterconnWS.Execute strsql
            medixmasterconnWS.CommitTrans
            
            If EditFlag = False Then
                MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            Else
                MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
            End If
            Call showReminder
            Call funcword
            EditFlag = False
            cmdSave.Enabled = False
        Else
            cmdSave.Enabled = True
        End If
        
        cmdAdd.Enabled = True
        cmdEdit.Enabled = False
End Sub

Private Sub funcword()
On Error GoTo openerror

If EditFlag = False Then
    Set objword = Nothing
    If objword Is Nothing Then
        Set objword = New Word.Application
        objword.Visible = True
        Set objdoc = objword.Documents.Add(LocCardPath & "Claim\" & Letter & "\")
    ElseIf Not objdoc Is Nothing Then
        If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
            vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
        ElseIf vbCancel Then
            Exit Sub
        End If
        Set objdoc = Nothing
        Set objdoc = objword.Documents.Add(LocCardPath & "Claim\" & Letter & "\")
    End If
    
    If bookmark(Letter) = True Then
        Me.show
    End If
End If
Exit Sub

openerror:
    Me.show
    If Err.Number = "5356" Then
        MsgBox Err.Description & "Please close the previous document first using Word or save it to other name", vbExclamation + vbOKOnly, DialogBoxTitle
    ElseIf Err.Number = "462" Then
        Set objword = New Word.Application
        objword.Visible = True
    ElseIf Err.Number = "4198" Then
        Err.Clear
    Else
        MsgBox Err.Description & Err.Number, vbExclamation + vbOKOnly, DialogBoxTitle
    End If
End Sub

Private Function bookmark(Optional sql As String) As Integer
Dim MBM As New ADODB.Recordset
Dim CLM As New ADODB.Recordset
Dim USR As New ADODB.Recordset
Dim strsqlMBM As String
Dim strsqlCLM As String
Dim strsqlUSR As String
Dim strsql As String

    bookmark = False
    
    If txtCASNumber <> "" Then
        objdoc.Bookmarks("CaseNo").Range.Text = StrConv(txtCASNumber, vbUpperCase)
    End If
                       
    If txtPrincipleName <> "" Then
        objdoc.Bookmarks("MBMName").Range.Text = StrConv(txtPrincipleName, vbUpperCase)
        objdoc.Bookmarks("InsuredName").Range.Text = StrConv(txtPrincipleName, vbUpperCase)
    End If
    
    If txtAdd1 <> "" Then
        objdoc.Bookmarks("Address1").Range.Text = StrConv(txtAdd1, vbUpperCase)
    End If
                
    If txtAdd2 <> "" Then
        objdoc.Bookmarks("Address2").Range.Text = StrConv(txtAdd2, vbUpperCase)
    End If
    
    If txtAdd3 <> "" Then
        objdoc.Bookmarks("Address3").Range.Text = StrConv(txtAdd3, vbUpperCase)
    End If
            
    If txtMBMNo <> "" Then
        objdoc.Bookmarks("MBMNo").Range.Text = txtMBMNo
    End If
            
    If txtPolicyNo <> "" Then
        objdoc.Bookmarks("PolicyNo").Range.Text = txtPolicyNo
    End If
            
    If txtPatientName <> "" Then
        objdoc.Bookmarks("PatientName").Range.Text = StrConv(txtPatientName, vbUpperCase)
    End If
            
    If txtAdmissionDate <> "" Then
        objdoc.Bookmarks("AdmissionDate").Range.Text = txtAdmissionDate
    End If
        
    If txtDischargeDate <> "" Then
        objdoc.Bookmarks("DischargeDate").Range.Text = txtDischargeDate
    Else
        objdoc.Bookmarks("DischargeDate").Range.Text = "-"
    End If
        
    If chkRMF.value = 1 Then
        objdoc.Bookmarks("RMF").Range.Text = "X"
    End If
        
    If chkOrinInv.value = 1 Then
        objdoc.Bookmarks("INV").Range.Text = "X"
    End If
        
    If chkMedicalForm.value = 1 Then
        objdoc.Bookmarks("MED").Range.Text = "X"
    End If
        
    If chkProposalForm.value = 1 Then
        objdoc.Bookmarks("PROPOSAL").Range.Text = "X"
    End If
        
    If chkRefLetter.value = 1 Then
        objdoc.Bookmarks("REF").Range.Text = "X"
    End If
        
    If chkRec.value = 1 Then
        objdoc.Bookmarks("REC").Range.Text = "X"
    End If
        
    If chkStudent.value = 1 Then
        objdoc.Bookmarks("STU").Range.Text = "X"
    End If
        
    If chkPolice.value = 1 Then
        objdoc.Bookmarks("POLICE").Range.Text = "X"
    End If
        
    If chkSubroForm.value = 1 Then
        objdoc.Bookmarks("SUBRO").Range.Text = "X"
    End If
        
    If chkOther1.value = 1 Then
        objdoc.Bookmarks("OTHER1").Range.Text = "X"
        objdoc.Bookmarks("TextOther1").Range.Text = StrConv(txtOther1, vbUpperCase)
    End If
        
    If chkOther2.value = 1 Then
        objdoc.Bookmarks("OTHER2").Range.Text = "X"
        objdoc.Bookmarks("TextOther2").Range.Text = StrConv(txtOther2, vbUpperCase)
    End If
                                        
    If FinalFlag = False Then
        
        strsql = "" & Trim(txtCASNumber) & ", " & Trim(cboVersion) & ""
        strsqlCLM = "Exec CLM_CLMReminder '" & strsql & "', 3"
        CLM.Open strsqlCLM, medixmasterconnWS, adOpenForwardOnly, adLockOptimistic
    
        If Not CLM.EOF Then
        Do Until CLM.EOF
            With CLM
            If StrComp(TempREM, !REMNumber) = 0 Then
                If !REMNumber = 0 Then
                    objdoc.Bookmarks("REMDate").Range.Text = Format(!REMDate, "dd mmmm yyyy")
                    objdoc.Bookmarks("REMNo").Range.Text = ""
                ElseIf !REMNumber = 1 Then
                    objdoc.Bookmarks("REMDate").Range.Text = Format(!REMDate, "dd mmmm yyyy")
                    objdoc.Bookmarks("REMNo").Range.Text = "FIRST REMINDER"
                ElseIf !REMNumber = 2 Then
                    objdoc.Bookmarks("REMDate").Range.Text = Format(!REMDate, "dd mmmm yyyy")
                    objdoc.Bookmarks("REMNo").Range.Text = "SECOND REMINDER"
                ElseIf !REMNumber = 3 Then
                    objdoc.Bookmarks("REMDate").Range.Text = Format(!REMDate, "dd mmmm yyyy")
                    objdoc.Bookmarks("REMNo").Range.Text = "FINAL REMINDER"
                End If
            End If
            .MoveNext
            End With
        Loop
        End If
        CLM.Close
        Set CLM = Nothing
        
    Else
        
        strsql = "" & Trim(txtCASNumber) & ", " & Trim(cboVersion) & ""
        strsqlCLM = "Exec CLM_CLMReminder '" & strsql & "', 3"
        CLM.Open strsqlCLM, medixmasterconnWS, adOpenForwardOnly, adLockOptimistic
        
        If Not CLM.EOF Then
        Do Until CLM.EOF
            With CLM
                If !REMNumber = 0 Then
                    objdoc.Bookmarks("LetterDate1").Range.Text = Format(!REMDate, "dd-mm-yyyy")
                ElseIf !REMNumber = 1 Then
                    objdoc.Bookmarks("LetterDate2").Range.Text = Format(!REMDate, "dd-mm-yyyy")
                ElseIf !REMNumber = 2 Then
                    objdoc.Bookmarks("LetterDate3").Range.Text = Format(!REMDate, "dd-mm-yyyy")
                ElseIf !REMNumber = 3 Then
                    objdoc.Bookmarks("REMDate").Range.Text = Format(!REMDate, "dd mmmm yyyy")
                    objdoc.Bookmarks("REMNo").Range.Text = "FINAL REMINDER"
                End If
                .MoveNext
            End With
        Loop
        End If
        CLM.Close
        Set CLM = Nothing
    End If
        
        strsqlMBM = "Exec CLM_CLMReminder '" & Trim(txtCASNumber) & "', 1"
        MBM.Open strsqlMBM, medixmasterconnWS, adOpenForwardOnly, adLockOptimistic
        
        If Not MBM.EOF Then
            With MBM
                
                If Not !PAYCompanyName = "" Then
                    objdoc.Bookmarks("Payor").Range.Text = !PAYCompanyName
                End If
                If Not !PAYMAINContact = "" Then
                    objdoc.Bookmarks("ContactPerson").Range.Text = !PAYMAINContact
                End If
                If Not !PAYFAXno = "" Then
                    objdoc.Bookmarks("FaxNo").Range.Text = !PAYFAXno
                End If
            End With
            MBM.Close
            Set MBM = Nothing
        End If
        
        strsqlUSR = "Exec CLM_CLMReminder 'A', 4"
        USR.Open strsqlUSR, medixmasterconnWS, adOpenForwardOnly, adLockOptimistic
                
        Do Until USR.EOF
            With USR
            If StrComp(Logon, !USRCode) = 0 Then
                objdoc.Bookmarks("LogonName").Range.Text = StrConv(!USRName, vbUpperCase)
            End If
            .MoveNext
            End With
        Loop
        USR.Close
        Set USR = Nothing
    
    bookmark = True
    FinalFlag = False
    
End Function

Private Sub CleanForm()
    txtMBMNo = ""
    txtPolicyNo = ""
    txtPrincipleName = ""
    txtPatientName = ""
    txtAdmissionDate = "__/__/____"
    txtDischargeDate = "__/__/____"
    txtREMDate = "__/__/____"
    txtOther1 = ""
    txtOther2 = ""
    txtAdd1 = ""
    txtAdd2 = ""
    txtAdd3 = ""
    chkRMF.value = 0
    chkOrinInv.value = 0
    chkMedicalForm.value = 0
    chkProposalForm.value = 0
    chkRefLetter.value = 0
    chkRec.value = 0
    chkStudent.value = 0
    chkPolice.value = 0
    chkSubroForm.value = 0
    chkOther1.value = 0
    chkOther2.value = 0
End Sub

Private Sub txtCASNumber_LostFocus()
    txtCASNumber = ChkStr(txtCASNumber)
End Sub

Private Sub txtREMDate_LostFocus()
    If txtREMDate <> "__/__/____" Then
        If IsDate(txtREMDate) = False Then
            MsgBox "Invalid Date Format!", vbOKOnly + vbCritical, DialogBoxTitle
            txtREMDate.SetFocus
            Exit Sub
        End If
    End If
End Sub
