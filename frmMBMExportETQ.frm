VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMBMExportETQ 
   Caption         =   "Membership Export (Etiqa/Maybank)"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   9975
      Begin VB.TextBox txtMBMNumberTO 
         Height          =   285
         Left            =   6960
         TabIndex        =   11
         Top             =   200
         Width           =   1335
      End
      Begin VB.TextBox txtMBMNumberFR 
         Height          =   285
         Left            =   5280
         TabIndex        =   10
         Top             =   200
         Width           =   1335
      End
      Begin VB.ComboBox cboPayor 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   200
         Width           =   3015
      End
      Begin MSMask.MaskEdBox txtBordxDateFR 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDateTO 
         Height          =   285
         Left            =   2880
         TabIndex        =   9
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtBatchNo1 
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
         Left            =   5280
         MaxLength       =   3
         TabIndex        =   12
         Top             =   450
         Width           =   1335
      End
      Begin VB.TextBox txtBatchNo2 
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
         Left            =   6960
         MaxLength       =   3
         TabIndex        =   13
         Top             =   450
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   20
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   19
         Top             =   560
         Width           =   315
      End
      Begin VB.Label Label26 
         Caption         =   "Bordx Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   560
         Width           =   915
      End
      Begin VB.Label Label24 
         Caption         =   "MBM Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   4320
         TabIndex        =   17
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label8 
         Caption         =   "To"
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
         Left            =   6720
         TabIndex        =   16
         Top             =   495
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Batch No."
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
         Left            =   4320
         TabIndex        =   15
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label Label24 
         Caption         =   "Payor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   250
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   9975
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7680
         TabIndex        =   5
         Top             =   210
         Width           =   720
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6240
         TabIndex        =   4
         Top             =   210
         Width           =   720
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9120
         TabIndex        =   3
         Top             =   210
         Width           =   720
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8400
         TabIndex        =   2
         Top             =   210
         Width           =   720
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         TabIndex        =   1
         Top             =   210
         Width           =   720
      End
   End
   Begin MSComctlLib.ListView lstMemberListing 
      Height          =   5055
      Left            =   120
      TabIndex        =   21
      Top             =   1125
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   24
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "MBMNumber"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "MBMName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "MBMIC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Policy No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Branch"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Eff Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Exp Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Swipe Card"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Department "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Batch No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Membership Export (Etiqa/Maybank)"
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
      TabIndex        =   22
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmMBMExportETQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpedit As Boolean
Public intExit As Integer
Private m_blnDirection(1 To 22) As Boolean
Dim RecordArray() As String


Private Sub cmdAdd_Click()
    tmpedit = False
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPayor.Text, cboPayor.SelStart) + Chr(KeyAscii) + Mid(cboPayor.Text, cboPayor.SelStart + cboPayor.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPayor.ListCount - 1
 
        If NewText = LCase(Left(cboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPayor.Text = ValidValue
                cboPayor.SelStart = 0
                cboPayor.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lstMemberListing.ListItems.Clear
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    
    Screen.MousePointer = vbHourglass
    
        If lstMemberListing.ListItems.Count <> 0 Then
            Call ExportListViewtoExcel(lstMemberListing)
        Else
            MsgBox "Report cannot be printed without any record.", vbInformation
        End If
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSearch_Click()
    cmdSearch.Enabled = False
    Screen.MousePointer = vbHourglass
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    cmdPrint.Enabled = False
    cmdClear.Enabled = False
    CmdExit.Enabled = False
    
        Call MBMInquiryList
    
    cmdPrint.Enabled = True
    cmdClear.Enabled = True
    CmdExit.Enabled = True

 '   medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    Screen.MousePointer = Default
    cmdSearch.Enabled = True

End Sub

Private Sub MBMInquiryList()
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sqlstr As String
Dim strCriteria As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim strsqlCOV As String
Dim RstCov As New ADODB.Recordset
Dim Check As Boolean
Dim rstCard As New ADODB.Recordset
Dim strsqlCard As String
Dim tmpSwipeCard As String

    strCriteria = ""
    If txtMBMNumberFR <> "" Then
        strCriteria = strCriteria & " AND MBMNumber >=  '" & Trim(txtMBMNumberFR) & "'"
    End If
    If txtMBMNumberTO <> "" Then
        strCriteria = strCriteria & " AND MBMNumber <=  '" & (txtMBMNumberTO) & "'"
    End If
    If txtBordxDateFR <> "__/__/____" Then
        strCriteria = strCriteria & " AND MBMBordxDate >= '" & Format(txtBordxDateFR, "mm/dd/yyyy") & "'"
    End If
    If txtBordxDateTO <> "__/__/____" Then
        strCriteria = strCriteria & " AND MBMBordxDate <= '" & Format(txtBordxDateTO, "mm/dd/yyyy") & "'"
    End If
    If txtBatchNo1 <> "" Then
        strCriteria = strCriteria & " AND MBMBatchNo >=  '" & (txtBatchNo1) & "'"
    End If
    If txtBatchNo2 <> "" Then
        strCriteria = strCriteria & " AND MBMBatchNo <=  '" & (txtBatchNo2) & "'"
    End If
    
    If cboPayor <> "" Then
        strCriteria = strCriteria & " AND PAYCode =  '" & Mid(cboPayor, 1, 2) & "'"
    End If
    lstMemberListing.ListItems.Clear
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Search_MBMExport_AXA"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strCriteria)
    cmd.Parameters.Append Prm1
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    Do
        Do Until rst2.EOF
            strsqlCOV = ""
            strsqlCOV = "SELECT * from MBMSCSecurity  Where MBMNumber LIKE '%" & Mid(rst2!MBMNumber, 1, 11) & "%' "
            RstCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            Do While Not RstCov.EOF
                tmpSwipeCard = RstCov!MBMSecurityCode & RstCov!MBMCardType & RstCov!MBMClientType & "  HCSB.COM.MY                   />" & RstCov!MBMNumber & "-00"

                RstCov.MoveNext
            Loop
                RstCov.Close
                Set RstCov = Nothing
            With rst2
                If Len(!MBMAXAPrintRemark) Then
                    RecordArray() = Split(!MBMAXAPrintRemark, "_")
                Else
                    RecordArray() = Split("_", "_")
                End If
                Set Record = lstMemberListing.ListItems.Add(, , !MBMNumber & "-" & "00")
                Record.SubItems(1) = !MBMName
                Record.SubItems(2) = !MBMIcBcPp
                Record.SubItems(3) = !MBMPolicyNo
                Record.SubItems(4) = IIf(Len(!MBMBranch), Trim(!MBMBranch), "")
                Record.SubItems(5) = IIf(Len(!MBMPayorEffDate), Format(!MBMPayorEffDate, "ddmmyyyy"), "")
                Record.SubItems(6) = IIf(Len(!MBMPayorEXPDate), Format(!MBMPayorEXPDate, "ddmmyyyy"), "")
                Record.SubItems(7) = IIf(Len(tmpSwipeCard), Trim(tmpSwipeCard), "")
                Record.SubItems(8) = IIf(Len(!MBMDepartment), Trim(!MBMDepartment), "")
                Record.SubItems(9) = IIf(Len(!MBMBatchNo), !MBMBatchNo, "")
                Record.SubItems(10) = ""
                Record.SubItems(11) = ""
                Record.SubItems(12) = ""
                Record.SubItems(13) = ""
                Record.SubItems(14) = ""
                Record.SubItems(15) = ""
                Record.SubItems(16) = ""
                Record.SubItems(17) = ""
                Record.SubItems(18) = ""
                Record.SubItems(19) = ""
                Record.SubItems(20) = ""
                Record.SubItems(21) = ""
                Record.SubItems(22) = ""
                

                'If Trim(!INSCode) = "F" Or Trim(!INSCode) = "H" Then
                '    strsqlCOV = ""
                    strsqlCOV = "SELECT MBMCoveredPersons.MBMCNumber, " & _
                                "MBMCoveredPersons.MBMCCoverID, MBMCoveredPersons.MBMCName," & _
                                "MBMCoveredPersonsTwo.MBMSendingMode,  MBMCoveredPersonsTwo.MBMCardDate " & _
                                "FROM MBMCoveredPersons INNER JOIN  MBMCoveredPersonsTwo ON " & _
                                "MBMCoveredPersons.MBMCNumber = MBMCoveredPersonsTwo.MBMCNumber AND " & _
                                "MBMCoveredPersons.MBMCCoverID = MBMCoveredPersonsTwo.MBMCCoverID  Where MBMCoveredPersons.MBMCNumber = '" & Trim(!MBMNumber) & "'"
                    RstCov.Open strsqlCOV, medixmasterconn, adOpenKeyset, adLockOptimistic
                
                    Do While Not RstCov.EOF
                        Set Record = lstMemberListing.ListItems.Add(, , "")
                        Record.SubItems(1) = IIf(Len(!MBMNumber), Trim(!MBMNumber), "")
                        Record.SubItems(2) = IIf(Len(RstCov!MBMCCoverID), Trim(RstCov!MBMCCoverID), "")
                        Record.SubItems(3) = IIf(Len(RstCov!MBMCName), Trim(RstCov!MBMCName), "")
                        Record.SubItems(4) = IIf(Len(!MBMBordxDate), !MBMBordxDate, "")
                        Record.SubItems(5) = IIf(Len(!MBMBatchNo), Trim(!MBMBatchNo), "")
                        Record.SubItems(6) = IIf(Len(RstCov!MBMSendingMode), Trim(RstCov!MBMSendingMode), "")
                        Record.SubItems(7) = IIf(Len(RstCov!MBMCardDate), Format(RstCov!MBMCardDate, "dd/mm/yyyy"), "")
                        Record.SubItems(8) = IIf(Len(!MBMPolicyNo), !MBMPolicyNo, "")
                    RstCov.MoveNext
                    Loop
                    RstCov.Close
                    Set RstCov = Nothing
                'End If
                                                            
                                                            
            End With
            rst2.MoveNext
            DoEvents
            If intExit = 1 Then
               Check = False
               Exit Do
            End If
        Loop
    Loop Until Check = False
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    If lstMemberListing.ListItems.Count = 0 Then
        MsgBox "Record Not Found!", vbCritical + vbOKOnly, DialogBoxTitle
    End If
End Sub

Private Sub CmdStop_Click()
    intExit = 1
End Sub

Private Sub Form_Load()
    Call ComboListing
End Sub

Private Sub lstMemberListing_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        Select Case ColumnHeader.Index
        Case 1
            SortListView lstMemberListing, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView lstMemberListing, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView lstMemberListing, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView lstMemberListing, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView lstMemberListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
        Case 6
            SortListView lstMemberListing, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView lstMemberListing, ColumnHeader.Index, ldtString, m_blnDirection(7)
        Case 8
            SortListView lstMemberListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
        Case 9
            SortListView lstMemberListing, ColumnHeader.Index, ldtString, m_blnDirection(9)
            
        End Select

End Sub

Private Sub lstMemberListing_KeyDown(KeyCode As Integer, Shift As Integer)
    'If lstMemberListing.ListItems.Count > 0 Then
    '    cbosendingmode = lstMemberListing.SelectedItem.SubItems(5)
    '    txtsentdate = Format(lstMemberListing.SelectedItem.SubItems(6), "dd/mm/yyyy")
    'End If
End Sub

Private Sub lstMemberListing_KeyUp(KeyCode As Integer, Shift As Integer)
    'If lstMemberListing.ListItems.Count > 0 Then
    '    cbosendingmode = lstMemberListing.SelectedItem.SubItems(5)
    '    txtsentdate = Format(lstMemberListing.SelectedItem.SubItems(6), "dd/mm/yyyy")
    'End If
End Sub

Private Sub smdEdit_Click()
    tmpedit = True
End Sub

Private Sub txtBordxDateFr_LostFocus()
    If IsDate(txtBordxDateFR) Then
    Else
        If txtBordxDateFR <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDateFR.SetFocus
        End If
    End If
End Sub

Private Sub txtBordxDateTo_LostFocus()
    If IsDate(txtBordxDateTO) Then
    Else
        If txtBordxDateTO <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDateTO.SetFocus
        End If
    End If
End Sub


Private Sub SetCheckAllItems(bState As Boolean)
Dim LV As LVITEM
Dim lvCount As Long
Dim lvIndex As Long
Dim lvState As Long
Dim r As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = lstMemberListing.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstMemberListing.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
End Sub

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command
   
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute

    Do While Not PAY.EOF
        With PAY
            cboPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            PAY.MoveNext
        End With
    Loop
    
    PAY.Close
    Set PAY = Nothing
    

End Sub






