VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMCOInvCRNote 
   Caption         =   "MCO Invoice Credit Note"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   8355
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   20
      Top             =   5880
      Width           =   8295
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5640
         TabIndex        =   24
         Top             =   160
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   6360
         TabIndex        =   22
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   7560
         TabIndex        =   21
         Top             =   160
         Width           =   615
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   6960
         TabIndex        =   23
         Top             =   160
         Width           =   615
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "MCO Invoice Fee Credit Note"
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
         Top             =   60
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   8295
      Begin VB.TextBox txtBatchNoFr 
         Height          =   285
         Left            =   5640
         MaxLength       =   25
         TabIndex        =   28
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtBatchNoTo 
         Height          =   285
         Left            =   7080
         MaxLength       =   25
         TabIndex        =   27
         Top             =   120
         Width           =   1095
      End
      Begin VB.ComboBox cboMCOpayor 
         Height          =   315
         Left            =   960
         TabIndex        =   26
         Top             =   160
         Width           =   3615
      End
      Begin VB.TextBox txtMBMNumberTo 
         Height          =   285
         Left            =   2880
         MaxLength       =   25
         TabIndex        =   8
         Top             =   450
         Width           =   1695
      End
      Begin MSMask.MaskEdBox txtbordxDateFR 
         Height          =   315
         Left            =   5640
         TabIndex        =   5
         Top             =   350
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtbordxDateTo 
         Height          =   315
         Left            =   7080
         TabIndex        =   3
         Top             =   350
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtINVDateTo 
         Height          =   315
         Left            =   7080
         TabIndex        =   15
         Top             =   630
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtINVDateFR 
         Height          =   315
         Left            =   5640
         TabIndex        =   17
         Top             =   630
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtINVNoTo 
         Height          =   285
         Left            =   2880
         MaxLength       =   25
         TabIndex        =   12
         Top             =   705
         Width           =   1695
      End
      Begin VB.TextBox txtMBMNumberFr 
         Height          =   285
         Left            =   960
         MaxLength       =   25
         TabIndex        =   7
         Top             =   450
         Width           =   1575
      End
      Begin VB.TextBox txtINVNoFr 
         Height          =   285
         Left            =   960
         MaxLength       =   25
         TabIndex        =   11
         Top             =   705
         Width           =   1575
      End
      Begin VB.Label lbl16 
         Caption         =   "To"
         Height          =   255
         Left            =   6840
         TabIndex        =   4
         Top             =   400
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Batch No.  "
         Height          =   255
         Left            =   4680
         TabIndex        =   30
         Top             =   160
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   6840
         TabIndex        =   29
         Top             =   165
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Inv Date   "
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   700
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "To"
         Height          =   255
         Left            =   6840
         TabIndex        =   16
         Top             =   700
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Inv No.  "
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   705
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   490
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "MBM No.  "
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   465
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Bordx Date   "
         Height          =   255
         Left            =   4680
         TabIndex        =   6
         Top             =   400
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lstMCOCrNote 
      Height          =   4290
      Left            =   0
      TabIndex        =   19
      Top             =   1560
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7567
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Membership No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cover ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IC No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Invoice No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Invoice Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Invoice By"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmMCOInvCRNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub cmdSearch_Click()
Dim strsql As String
Dim strAppend As String
Dim strAppendSupp As String
Dim sql As String
    
    strAppend = ""
    strAppendSupp = ""
                
    If Len(cboMCOpayor) Then
        strAppend = strAppend + " And PAYCode = '" & Mid(cboMCOpayor, 1, 2) & "'"
    End If
    
    If Len(Trim(txtMBMNumberFr)) Then
        strAppend = strAppend + " AND MBMNumber >= '" & Trim(txtMBMNumberFr) & "'"
    End If
    
    If Len(Trim(txtMBMNumberTo)) Then
        strAppend = strAppend + " And MBMNumber <= '" & Trim(txtMBMNumberTo) & "'"
    End If
        
    If Len(txtINVNoFr) Then
        strAppend = strAppend + " and MBMInvoiceNo >= '" & Trim(txtINVNoFr) & "'"
    End If
    
    If Len(txtINVNoTo) Then
        strAppend = strAppend + " and MBMInvoiceNo <= '" & Trim(txtINVNoTo) & "'"
    End If
    
    
    If Len(Trim(txtINVDateFR)) > 0 And IsDate(txtINVDateFR) = True Then
        strAppend = strAppend + " And InvDate >= '" & Format(txtINVDateFR, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtINVDateTo)) > 0 And IsDate(txtINVDateTo) = True Then
        strAppend = strAppend + " And InvDate <= '" & Format(txtINVDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtbordxDateFR)) > 0 And IsDate(txtbordxDateFR) = True Then
        strAppend = strAppend + " And MBMBordxDate >= '" & Format(txtbordxDateFR, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtbordxDateTo)) > 0 And IsDate(txtbordxDateTo) = True Then
        strAppend = strAppend + " And MBMBordxDate <= '" & Format(txtbordxDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtBatchNoFr)) > 0 And IsDate(txtBatchNoFr) = True Then
        strAppend = strAppend + " And MBMBatchNo <= '" & Trim(txtBatchNoFr) & "'"
    End If
    
    If Len(Trim(txtBatchNoFr)) > 0 And IsDate(txtBatchNoTo) = True Then
        strAppend = strAppend + " And MBMBatchNo >= '" & Trim(txtBatchNoTo) & "'"
    End If
    
    'Supplementary Search
    If Len(cboMCOpayor) Then
        strAppendSupp = strAppendSupp + " And substring(MBMCNumber,1,2) = '" & Mid(cboMCOpayor, 1, 2) & "'"
    End If

    If Len(Trim(txtMBMNumberFr)) Then
        strAppendSupp = strAppendSupp + " AND MBMCNumber >= '" & Trim(txtMBMNumberFr) & "'"
    End If
    
    If Len(Trim(txtMBMNumberTo)) Then
        strAppendSupp = strAppendSupp + " And MBMCNumber <= '" & Trim(txtMBMNumberTo) & "'"
    End If
        
    If Len(txtINVNoFr) Then
        strAppendSupp = strAppendSupp + " and MBMInvoiceNo >= '" & Trim(txtINVNoFr) & "'"
    End If
    
    If Len(txtINVNoTo) Then
        strAppendSupp = strAppendSupp + " and MBMInvoiceNo <= '" & Trim(txtINVNoTo) & "'"
    End If
    
    
    If Len(Trim(txtINVDateFR)) > 0 And IsDate(txtINVDateFR) = True Then
        strAppendSupp = strAppendSupp + " And InvDate >= '" & Format(txtINVDateFR, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtINVDateTo)) > 0 And IsDate(txtINVDateTo) = True Then
        strAppendSupp = strAppendSupp + " And InvDate <= '" & Format(txtINVDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtbordxDateFR)) > 0 And IsDate(txtbordxDateFR) = True Then
        strAppendSupp = strAppendSupp + " And MBMCBordxDate >= '" & Format(txtbordxDateFR, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtbordxDateTo)) > 0 And IsDate(txtbordxDateTo) = True Then
        strAppendSupp = strAppendSupp + " And MBMCBordxDate <= '" & Format(txtbordxDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtBatchNoFr)) > 0 And IsDate(txtBatchNoFr) = True Then
        strAppendSupp = strAppendSupp + " And MBMCBatchNo <= '" & Trim(txtBatchNoFr) & "'"
    End If
    
    If Len(Trim(txtBatchNoFr)) > 0 And IsDate(txtBatchNoTo) = True Then
        strAppendSupp = strAppendSupp + " And MBMCBatchNo >= '" & Trim(txtBatchNoTo) & "'"
    End If
    
    'searchCriteria = strAppend
    Call MCOInvListing(strAppend, strAppendSupp)

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

Private Sub Form_Load()

Call MCOComboListing

End Sub

Private Sub lstMCOCrNote_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstMCOCrNote.SortKey = ColumnHeader.Index - 1
    lstMCOCrNote.Sorted = True
    If lstMCOCrNote.SortOrder = lvwAscending Then
       lstMCOCrNote.SortOrder = lvwDescending
    Else
       lstMCOCrNote.SortOrder = lvwAscending
    End If
End Sub

Private Sub txtbordxDateFR_LostFocus()
    If IsDate(txtbordxDateFR) Then
    Else
        If txtbordxDateFR <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtbordxDateFR.SetFocus
        End If
    End If
End Sub

Private Sub txtbordxDateTo_Change()
    If IsDate(txtbordxDateTo) Then
    Else
        If txtbordxDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtbordxDateTo.SetFocus
        End If
    End If
End Sub

Private Sub txtINVDateFR_LostFocus()
    If IsDate(txtINVDateFR) Then
    Else
        If txtINVDateFR <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtINVDateFR.SetFocus
        End If
    End If
End Sub

Private Sub txtINVDateTo_LostFocus()
    If IsDate(txtINVDateTo) Then
    Else
        If txtINVDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtINVDateTo.SetFocus
        End If
    End If
End Sub

Private Sub MCOInvListing(strAppend As String, strAppendSupp As String)
Dim sqlstrMBMSupp  As String
Dim MBMSupp As New ADODB.Recordset
Dim sqlstrMBM  As String
Dim MBM As New ADODB.Recordset
Dim Record As ListItem
'sqlstrMBM = ""
    lstMCOCrNote.listitems.Clear
    
    sqlstrMBM = " Select Paycode, MBMNumber, MBMName, MBMIcBcPp,MBMInvoiceBy,MBMInvoiceDate,MBMInvoiceNo  " & _
             "  FROM VIEW_MBM_MBMOthers Where 0 = 0 "
        
    If Len(Trim(strAppend)) Then
        sqlstrMBM = sqlstrMBM + strAppend
    End If

    MBM.Open sqlstrMBM, medixmasterconn, adOpenKeyset, adLockOptimistic

    If MBM.recordCount Then
        Do Until MBM.EOF
            Set Record = lstMCOCrNote.listitems.Add(, , MBM!MBMNumber)
            
            'If Len(MBM!MBMName) Then
            Record.SubItems(1) = "00"
            'End If
            
            If Len(MBM!MBMName) Then
                Record.SubItems(2) = MBM!MBMName
            End If

            If Len(MBM!MBMIcBcPp) Then
                Record.SubItems(3) = MBM!MBMIcBcPp
            End If
            
            If Len(MBM!MBMInvoiceNo) Then
                Record.SubItems(4) = MBM!MBMInvoiceNo
            End If
           
            If Len(MBM!MBMInvoiceDate) Then
                Record.SubItems(5) = MBM!MBMInvoiceDate
            End If
            If Len(MBM!MBMInvoiceBy) Then
                Record.SubItems(6) = Format(MBM!MBMInvoiceBy, "#,##0.00")
            End If
            'If Len(MBM!HOSPaidAmt) Then
            '    Record.SubItems(6) = Format(MBM!HOSPaidAmt, "#,##0.00")
            'End If
        MBM.MoveNext
        Loop
    End If
        
    sqlstrMBMSupp = " Select MBMCNumber, MBMCName, MBMCIcBcPpNo,  MBMCCoverID,MBMInvoiceBy," & _
             " MBMInvoiceDate,MBMInvoiceNo,MBMCCoverID FROM MBMCoveredPersons Where 0 = 0 "
        
    If Len(Trim(strAppendSupp)) Then
        sqlstrMBMSupp = sqlstrMBMSupp + strAppendSupp
    End If

    MBMSupp.Open sqlstrMBMSupp, medixmasterconn, adOpenKeyset, adLockOptimistic

    If MBMSupp.recordCount Then
        Do Until MBMSupp.EOF
            Set Record = lstMCOCrNote.listitems.Add(, , MBMSupp!MBMCNumber)
            
            If Len(MBMSupp!MBMCCoverID) Then
                Record.SubItems(1) = MBMSupp!MBMCCoverID
            End If
    
            If Len(MBMSupp!MBMCName) Then
                Record.SubItems(2) = MBMSupp!MBMCName
            End If

            If Len(MBMSupp!MBMCIcBcPpNo) Then
                Record.SubItems(3) = MBMSupp!MBMCIcBcPpNo
            End If
            
            If Len(MBMSupp!MBMInvoiceNo) Then
                Record.SubItems(4) = MBMSupp!MBMInvoiceNo
            End If
           
            If Len(MBMSupp!MBMInvoiceDate) Then
                Record.SubItems(5) = MBMSupp!MBMInvoiceDate
            End If
            If Len(MBMSupp!MBMInvoiceBy) Then
                Record.SubItems(6) = Format(MBMSupp!MBMInvoiceBy, "#,##0.00")
            End If
            'If Len(MBM!HOSPaidAmt) Then
            '    Record.SubItems(6) = Format(MBM!HOSPaidAmt, "#,##0.00")
            'End If
        
        MBMSupp.MoveNext
        Loop
    End If

    If MBM.recordCount = 0 And MBMSupp.recordCount = 0 Then
        MsgBox "No record found", vbCritical
    End If
    
    MBM.Close
    Set MBM = Nothing
    MBMSupp.Close
    Set MBM = Nothing
End Sub

Private Sub MCOComboListing()
Dim MCO As New ADODB.Recordset
        
    MCO.Open "PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
        
    cboMCOpayor.Clear
        
    Do Until MCO.EOF
        cboMCOpayor.AddItem MCO!PAYCode & " - " & MCO!PAYCompanyName
        MCO.MoveNext
    Loop
        
    MCO.Close
    
End Sub

Private Sub UpdateRecord(MBMNo As String, MBMID As String, CRNote As String)
Dim Update
Dim strsql As String
Dim MBM As New ADODB.Recordset
    
    If MBMID = "00" Then
        strsql = " Select  MBMCRNote, MBMCRNoteDate, MBMCRNoteBy from MBMOthers " & _
                    " Where MBMNumber = '" & Trim(MBMNo) & "'"
        MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    Else
        strsql = " Select  MBMCRNote, MBMCRNoteDate, MBMCRNoteBy  from MBMCoveredPersons " & _
                    " Where MBMCNumber = '" & Trim(MBMNo) & "' AND MBMCCoverID = '" & Trim(MBMID) & "'"
        MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    End If
    
    If MBM.recordCount Then
        With MBM
            !MBMCRNote = Trim(ChkStr(CRNote))
            !MBMCRNoteDate = Date
            !MBMCRNoteBy = User
        End With
        MBM.Update
    End If
    MBM.Close
    Set MBM = Nothing
End Sub

