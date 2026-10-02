VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMBMInvoicePosting 
   Caption         =   "Member Invoice Posting"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   23
      Top             =   240
      Width           =   11775
      Begin VB.ComboBox CboPayor 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   4095
      End
      Begin VB.ComboBox cboPostStatus 
         Height          =   315
         Left            =   7680
         TabIndex        =   5
         Top             =   240
         Width           =   3735
      End
      Begin MSMask.MaskEdBox txtInvDate1 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInvDate2 
         Height          =   285
         Left            =   4440
         TabIndex        =   2
         Top             =   520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPostDate1 
         Height          =   285
         Left            =   7680
         TabIndex        =   6
         Top             =   520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPostDate2 
         Height          =   285
         Left            =   9960
         TabIndex        =   7
         Top             =   520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtInvNo1 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox txtInvNo2 
         Height          =   285
         Left            =   4440
         TabIndex        =   4
         Top             =   780
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Payor"
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   280
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Invoice Date from"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   550
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Invoice No. from"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   820
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Post Status"
         Height          =   255
         Left            =   6360
         TabIndex        =   28
         Top             =   280
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Post Date from"
         Height          =   255
         Left            =   6360
         TabIndex        =   27
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "To"
         Height          =   255
         Left            =   3720
         TabIndex        =   26
         Top             =   580
         Width           =   350
      End
      Begin VB.Label Label8 
         Caption         =   "To"
         Height          =   255
         Left            =   3720
         TabIndex        =   25
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   9480
         TabIndex        =   24
         Top             =   580
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   7320
      Width           =   11775
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   330
         Left            =   7080
         TabIndex        =   9
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   330
         Left            =   8640
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9240
         TabIndex        =   12
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10920
         TabIndex        =   14
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10200
         TabIndex        =   13
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   7800
         TabIndex        =   10
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdPost 
         Caption         =   "&Post"
         Height          =   330
         Left            =   8520
         TabIndex        =   11
         Top             =   200
         Width           =   735
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Debit Amt"
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblDbAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3000
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   960
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblCrAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   5040
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Debit Amt"
         Height          =   255
         Left            =   4200
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstHospInvList 
      Height          =   5775
      Left            =   60
      TabIndex        =   8
      Top             =   1560
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10186
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   23
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Journal Type"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ac Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Credit"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Hosp"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Payor"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Month"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "HealthCode"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "MBMCode"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Ac Period"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "WS No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "BordxNo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ClaimNo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "InvDate"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Amount"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "PatientName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "PlanCode"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "InsuredType"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "PostStatus"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "PostDate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "PostBy"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "CasNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "VerNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "InvIndex"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Member Invoice Posting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmMBMInvoicePosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intExit As Boolean
Dim SearchCri As String

Private Sub ComboListing()
Dim PayorRec As New ADODB.Recordset
Dim PayorSql As String

    PayorSql = "Select PayCode, HLTCode, PAYCompanyName from Pay "
    PayorRec.Open PayorSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until PayorRec.EOF
        CboPayor.AddItem Trim(PayorRec!PAYCode) & "-" & Trim(PayorRec!PAYCompanyName)
        PayorRec.MoveNext
    Loop
    
    cboPostStatus.AddItem "0 - No"
    cboPostStatus.AddItem "1 - Yes"
    cboPostStatus.ListIndex = 0
End Sub

Private Sub CboPayor_Change()
    If CboPayor.Text = "" Then
        If InStr(CboPayor.Text, "null") Then
           CboPayor.Enabled = False
        End If
    End If
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayor.Text, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor.Text, CboPayor.SelStart + CboPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayor.ListCount - 1
 
        If NewText = LCase(Left(CboPayor.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = CboPayor.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            CboPayor.Text = ValidValue
            CboPayor.SelStart = 0
            CboPayor.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub cboPostStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPostStatus.Text, cboPostStatus.SelStart) + Chr(KeyAscii) + Mid(cboPostStatus.Text, cboPostStatus.SelStart + cboPostStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPostStatus.ListCount - 1
 
        If NewText = LCase(Left(cboPostStatus.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = cboPostStatus.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
        If ValidCount = 1 Then
            cboPostStatus.Text = ValidValue
            cboPostStatus.SelStart = 0
            cboPostStatus.SelLength = Len(ValidValue)
        End If
    End If
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    cboPostStatus.ListIndex = 0
    CboPayor.Text = ""
    lblCurrent = ""
    LblDbAmt = ""
    LblCrAmt = ""
    lstHospInvList.listitems.Clear
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPost_Click()
On ERR GoTo ErrorSave
Dim StartTrans As Boolean
Dim AA As Integer
Dim INvoiceREc As New ADODB.Recordset
Dim StrINvSql As String
Dim TmpInvIndex As String
Dim UpdAsk

    If lstHospInvList.listitems.count <= 0 Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    Else
        UpdAsk = MsgBox("Do you want to Post the record(s) ?", vbYesNo + vbExclamation)
        If UpdAsk = vbYes Then
            Screen.MousePointer = vbHourglass
            medixmasterconnWS.beginTrans
            StartTrans = True
            
            For AA = 1 To lstHospInvList.listitems.count
                TmpInvIndex = lstHospInvList.listitems(AA).SubItems(21)
                StrINvSql = "Select CLMINVPostStatus, CLMINVPostDate, CLMINVPostUser " & _
                            " From ClaimInvoice WHere INvIndex='" & Trim(TmpInvIndex) & "'"
                INvoiceREc.Open StrINvSql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                If INvoiceREc.recordCount > 0 Then
                    INvoiceREc!CLMINVPostStatus = "1"
                    INvoiceREc!CLMInvPostDate = Date
                    INvoiceREc!CLMINVPostUser = Logon
                    INvoiceREc.Update
                End If
                INvoiceREc.Close
                Set INvoiceREc = Nothing
            Next AA
            
            medixmasterconnWS.CommitTrans
            StartTrans = False
            Screen.MousePointer = vbDefault
            MsgBox "Record(s) Posted Sucessfully... ! ", vbOKOnly
            cboPostStatus.ListIndex = 1
            Call HospInvListing(SearchCri & " And CLMINVPostStatus ='1'")
        End If
    End If
    Exit Sub
ErrorSave:
        Screen.MousePointer = vbDefault
        If StartTrans = True Then
            medixmasterconnWS.RollbackTrans
        End If
        MsgBox ERR.Description
End Sub

Private Sub CmdPrintFile_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelAnalysis1(lstHospInvList)
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSearch_Click()
Dim strsql As String
Dim PostSql As String
    strsql = ""
    PostSql = ""
    If Trim(CboPayor) <> "" Then
        strsql = strsql & " And substring(CASNumber,1,2)='" & Mid(CboPayor, 1, 2) & "'"
    End If
    If IsDate(txtInvDate1) Then
        strsql = strsql & " And InvDate >='" & txtInvDate1 & "'"
    End If
    If IsDate(txtInvDate2) Then
        strsql = strsql & " And InvDate <='" & txtInvDate2 & "'"
    End If
    If Trim(txtInvNo1) <> "" Then
        strsql = strsql & " And InvNo >='" & Trim(txtInvNo1) & "'"
    End If
    If Trim(txtInvNo2) <> "" Then
        strsql = strsql & " And InvNo <='" & Trim(txtInvNo2) & "'"
    End If
    If Mid(cboPostStatus, 1, 1) = "0" Then
        PostSql = " And (CLMINVPostStatus IS NULL OR CLMINVPostStatus = '0')"
    End If
    If Mid(cboPostStatus, 1, 1) = "1" Then
        PostSql = " And CLMINVPostStatus ='1'"
    End If
    If IsDate(txtPostDate1) Then
        strsql = strsql & " And CLMINVPostDate >='" & txtPostDate1 & "'"
    End If
    If IsDate(txtPostDate2) Then
        strsql = strsql & " And CLMINVPostDate <='" & txtPostDate2 & "'"
    End If
    SearchCri = strsql
    Screen.MousePointer = vbHourglass
    Call HospInvListing(strsql & PostSql)
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub CmdStop_Click()
    intExit = True
    CmdExit.Enabled = True
End Sub

Private Sub Form_Load()
    intExit = False
    Call ComboListing
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '--- Enable PrintScreen feature
    '--- IMPORTANT - Don't forget to set .KeyPreview = True for this form at design time!!
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
End Sub

Private Sub txtPostDate1_LostFocus()
    If IsDate(txtPostDate1) Then
    Else
        If txtPostDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPostDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtPostDate2_LostFocus()
    If IsDate(txtPostDate2) Then
    Else
        If txtPostDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPostDate2.SetFocus
        End If
    End If
End Sub

Private Sub HospInvListing(StrAppend As String)
Dim lstRec As ListItem
Dim HospInv As New ADODB.Recordset
Dim HospSql As String
Dim MBMXRec As New ADODB.Recordset
Dim MBMXSql As String
Dim SuppRec As New ADODB.Recordset
Dim SuppSql As String
Dim tmpPayor, tmpHLTCode, tmpINSCode, tmpPatient, tmpPlan As String
Dim JournalType, AcCode, CreditType As String
Dim AA, CurrRec As Integer
Dim TotDbAmt, TotCrAmt As Double
'Dim ChkStop As Boolean
Dim HospMBMCode As String

    tmpPayor = ""
    tmpHLTCode = ""
    tmpPlan = ""
    tmpINSCode = ""
    tmpPatient = ""
'    ChkStop = False
    lstHospInvList.listitems.Clear
    
    HospSql = "SELECT MBMNumber, MBMPatientCovID, CLMPayableByMedix2M, " & _
              "InvDate, CasNumber, INVBordxRefNo, CLMINVPostStatus, " & _
              "CLMInvPostDate, CLMINVPostUser, claimversion, INVIndex " & _
              " From View_MBMInv_Posting WHERE 0=0 "
    
    If Trim(StrAppend) <> "" Then
        HospSql = HospSql & StrAppend
    End If
    HospInv.Open HospSql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If HospInv.recordCount > 0 Then
        Do Until HospInv.EOF
        CurrRec = CurrRec + 1
            With HospInv
                MBMXSql = "Select HLTCode, PAYCode ,INSCode,PLNCode, MBMName FROM MBMCrossReference WHERE MBMNumber='" & Trim(!MBMNumber) & "'"
                MBMXRec.Open MBMXSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                If MBMXRec.recordCount > 0 Then
                    tmpPayor = MBMXRec!PAYCode
                    tmpHLTCode = MBMXRec!HLTCode
                    tmpPlan = MBMXRec!PLNCode
                    tmpINSCode = MBMXRec!INSCode
                    tmpPatient = MBMXRec!MBMName
                End If
                MBMXRec.Close
                Set MBMXRec = Nothing
                If Trim(!MBMPatientCovID) <> "00" Then
                    SuppSql = "SELECT MBMCName, MBMCPLNCode FROM MBMCoveredPersons WHERE MBMCNumber='" & Trim(!MBMNumber) & _
                            "' And MBMCCoverID='" & Trim(!MBMPatientCovID) & "'"
                    SuppRec.Open SuppSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                    If SuppRec.recordCount > 0 Then
                        tmpPatient = SuppRec!MBMCName
                        If Trim(SuppRec!MBMCPLNCode) <> "" Then
                            tmpPlan = Trim(SuppRec!MBMCPLNCode)
                        End If
                    End If
                    SuppRec.Close
                    Set SuppRec = Nothing
                End If
                JournalType = "MJC"
                AcCode = "600" & Trim(tmpPayor)
                CreditType = "D"
                HospMBMCode = "0R"
                TotDbAmt = TotDbAmt + IIf(IsNumeric(!CLMPayableByMedix2M), !CLMPayableByMedix2M, 0)

                For AA = 1 To 2
                    If AA = 2 Then
                        TotCrAmt = TotCrAmt + IIf(IsNumeric(!CLMPayableByMedix2M), !CLMPayableByMedix2M, 0)
                        AcCode = "820" & Trim(tmpPayor)
                        CreditType = "C"
                    End If
                    Set lstRec = lstHospInvList.listitems.Add(, , JournalType)
                    lstRec.SubItems(1) = Trim(AcCode)
                    lstRec.SubItems(2) = Trim(CreditType)
                    lstRec.SubItems(3) = Trim(HospMBMCode)
                    lstRec.SubItems(4) = Trim(tmpPayor)
                    lstRec.SubItems(5) = IIf(IsDate(!InvDate), Format(!InvDate, "MM"), "")
                    lstRec.SubItems(6) = Trim(tmpHLTCode)
                    lstRec.SubItems(7) = Right(!MBMNumber, 4)
                    lstRec.SubItems(8) = IIf(IsDate(!InvDate), Format(!InvDate, "MM/YYYY"), "")
                    lstRec.SubItems(9) = Trim(!CasNumber) & "-" & Trim(!claimversion)
                    lstRec.SubItems(10) = IIf(Len(Trim(!INVBordxRefNo)), Trim(!INVBordxRefNo), "")
                    lstRec.SubItems(11) = Trim(!CasNumber) & "-" & Trim(!claimversion)
                    lstRec.SubItems(12) = IIf(IsDate(!InvDate), Format(!InvDate, "MM/YYYY"), "")
                    lstRec.SubItems(13) = Format(IIf(IsNumeric(!CLMPayableByMedix2M), !CLMPayableByMedix2M, 0), "#,##0.00")
                    lstRec.SubItems(14) = Trim(tmpPatient)
                    lstRec.SubItems(15) = Trim(tmpPlan)
                    lstRec.SubItems(16) = Trim(tmpINSCode)
                    lstRec.SubItems(17) = IIf(Len(Trim(!CLMINVPostStatus)), Trim(!CLMINVPostStatus), "")
                    If Trim(!CLMInvPostDate) <> "" Then
                        lstRec.SubItems(18) = !CLMInvPostDate
                    End If
                    lstRec.SubItems(19) = IIf(Len(Trim(!CLMINVPostUser)), Trim(!CLMINVPostUser), "")
                    lstRec.SubItems(20) = Trim(!CasNumber)
                    lstRec.SubItems(21) = Trim(!claimversion)
                    lstRec.SubItems(22) = Trim(!INVIndex)
                Next AA
                LblDbAmt = TotDbAmt
                LblCrAmt = TotCrAmt
                lblCurrent = CurrRec
            End With
            HospInv.MoveNext
            DoEvents
            If intExit = True Then
'               ChkStop = True
               Exit Do
            End If
        Loop
    End If
    intExit = 0
    
    HospInv.Close
    Set HospInv = Nothing
End Sub


