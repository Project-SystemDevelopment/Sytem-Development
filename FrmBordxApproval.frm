VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBordxApproval 
   Caption         =   "Bordx Approval"
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
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   120
      TabIndex        =   22
      Top             =   240
      Width           =   11655
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   5040
         TabIndex        =   3
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox TxtBordx2 
         Height          =   285
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox TxtBordx1 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtBordxAmt2 
         Height          =   285
         Left            =   9240
         MaxLength       =   8
         TabIndex        =   8
         Top             =   540
         Width           =   1695
      End
      Begin VB.TextBox TxtBordxAmt1 
         Height          =   285
         Left            =   7320
         MaxLength       =   8
         TabIndex        =   7
         Top             =   540
         Width           =   1455
      End
      Begin VB.ComboBox CboStatus 
         Height          =   315
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   3375
      End
      Begin MSMask.MaskEdBox txtBordxDate1 
         Height          =   285
         Left            =   7320
         TabIndex        =   5
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDate2 
         Height          =   285
         Left            =   9240
         TabIndex        =   6
         Top             =   210
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   30
         Top             =   650
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   650
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   8880
         TabIndex        =   27
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   6240
         TabIndex        =   26
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   8880
         TabIndex        =   25
         Top             =   630
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx Amt"
         Height          =   255
         Left            =   6240
         TabIndex        =   24
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "App. Status"
         Height          =   255
         Left            =   6240
         TabIndex        =   23
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   6960
      Width           =   11655
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   4200
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdCheck 
         Caption         =   "&Check All"
         Height          =   375
         Left            =   6120
         TabIndex        =   10
         Top             =   220
         Width           =   855
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   9240
         TabIndex        =   14
         Top             =   220
         Width           =   1095
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   7680
         TabIndex        =   12
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10320
         TabIndex        =   15
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10920
         TabIndex        =   16
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   6960
         TabIndex        =   11
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton CmdToPayor 
         Caption         =   "&Update"
         Height          =   375
         Left            =   8400
         TabIndex        =   13
         Top             =   220
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   3360
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstBordxApproval 
      Height          =   5415
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Payor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Bordx Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Bordx Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Bordx Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Approved Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Approved By"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Approved Date"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Bordx Approval"
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
      TabIndex        =   31
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmBordxApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim BordxNoArray() As String
Dim searchCriteria As String
Dim ItemChecked As Boolean
Dim listloop1 As Integer
Dim listloop2 As Integer
Dim Check As Boolean
Dim Check2 As Boolean
Dim Current2 As String
Dim TotalAmt2 As String
Private m_blnDirection(1 To 9) As Boolean

Private Sub CmdCheck_Click()
    SetCheckAllItems True
End Sub

Private Sub cboPayor_Click()
    If CboPayor.Text = "" Then
        If InStr(CboPayor.Text, "null") Then
            CboPayor.Enabled = False
        End If
        TxtBordx1.SetFocus
    End If
    TxtBordx1.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'catch both "Enter" keys on keyboard
    Call EnterToTab(KeyAscii)
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
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    Current2 = 0
    TotalAmt2 = 0
    Check = False
    Check2 = False
    Label14.Visible = False
    TxtBordx2.Visible = False
    ReDim BordxNoArray(0)
    Call ComboListing
    ItemChecked = False
End Sub


'**********************Start Command Button *******************************

Private Sub ChkScan_Click()
    If ChkScan.Value = 1 Then
        lstBordxApproval.ListItems.Clear
        Label14.Visible = False
        TxtBordx2.Visible = False
        TxtBordx1.SetFocus
    Else
        lstBordxApproval.ListItems.Clear
        Label14.Visible = True
        TxtBordx2.Visible = True
        TxtBordx1.SetFocus
    End If
End Sub

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
        'Call ExportToExcelBordxReturnPayor(lstBordxApproval)
    If lstBordxApproval.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstBordxApproval)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdClear_Click()
    CboPayor.Clear
    CboStatus.Clear
    Call ClearForm(Me)
    Current2 = 0
    TotalAmt2 = 0
    Check = False
    Check2 = False
    TxtBordx1.SetFocus
    lstBordxApproval.ListItems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    ReDim BordxNoArray(0)
    Call ComboListing
End Sub

Private Sub CmdSearch_Click()
    ItemChecked = False
    LblTotalAmt = ""
    lblTotal = ""
    lblCurrent = ""
    ReDim BordxNoArray(0)
    If ChkScan.Value = 1 Then
    
        If CboPayor = "" Then
            MsgBox "Please select a Payor", vbOKOnly + vbCritical, DialogBoxTitle
            CboPayor.SetFocus
        ElseIf TxtBordx1 = "" Then
            MsgBox "Please Enter Bordx No", vbOKOnly + vbCritical, DialogBoxTitle
            TxtBordx1.SetFocus
            CmdClear.Enabled = True
            CmdToPayor.Enabled = True
            cmdPrint.Enabled = True
            CmdExit.Enabled = True
        Else
        
            Screen.MousePointer = vbHourglass
            
            listrecord = lstBordxApproval.ListItems.Count
            If listrecord <> 0 Then
                
                For listloop1 = 1 To listrecord
                    Call BordxNoCheck
                Next listloop1
            End If
                If Check = True Then
                    MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                ElseIf Check = False Then
                    CmdClear.Enabled = False
                    CmdToPayor.Enabled = False
                    cmdPrint.Enabled = False
                    CmdExit.Enabled = False
                    Call SearchRecord
                End If
                Check = False
                CmdClear.Enabled = True
                CmdToPayor.Enabled = True
                cmdPrint.Enabled = True
                CmdExit.Enabled = True
            Screen.MousePointer = Default
            TxtBordx1 = ""
            TxtBordx1.SetFocus
        End If
        TxtBordx1.SetFocus
    Else
        
            Screen.MousePointer = vbHourglass
                CmdSearch.Enabled = False
                CmdToPayor.Enabled = False
                'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                    Call SearchRecord
                'medixmasterconnWS.Close
                'Set medixmasterconnWS = Nothing
                CmdSearch.Enabled = True
                CmdToPayor.Enabled = True
            Screen.MousePointer = Default
            
        CmdClear.Enabled = True
        CmdToPayor.Enabled = True
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    End If
End Sub

Private Function SearchRecord()
Dim StrSql As String
Dim StrAppend As String
Dim sql As String
Dim CAS As New ADODB.Recordset

    StrAppend = ""
    If Len(Trim(CboPayor)) Then
        StrAppend = StrAppend + " And BordxPAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If ChkScan.Value = 1 Then
        If Len(Trim(TxtBordx1)) Then
            StrAppend = StrAppend + " AND BordxRefNo = '" & Trim(TxtBordx1) & "'"
        End If
    Else
        If Len(Trim(TxtBordx1)) Then
            StrAppend = StrAppend + " AND BordxRefNo >= '" & Trim(TxtBordx1) & "'"
        End If
        
        If Len(Trim(TxtBordx2)) Then
            StrAppend = StrAppend + " AND BordxRefNo <= '" & Trim(TxtBordx2) & "'"
        End If
    End If
    
    If IsDate(txtBordxDate1) = True Then
        StrAppend = StrAppend + " And BordxDate >= '" & Format(txtBordxDate1, "mm/dd/yyyy") & "'"
    End If
      
    If IsDate(txtBordxDate2) = True Then
        StrAppend = StrAppend + " And BordxDate <= '" & Format(txtBordxDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtBordxAmt1)) Then
        StrAppend = StrAppend + " AND BordxClaimAmt >= '" & Trim(TxtBordxAmt1) & "'"
    End If
    
    If Len(Trim(TxtBordxAmt2)) Then
        StrAppend = StrAppend + " AND BordxClaimAmt <= '" & Trim(TxtBordxAmt2) & "'"
    End If
    
    If Mid(CboStatus, 1, 1) = "1" Then
        StrAppend = StrAppend + " AND BordxAppStatus = '1'"
    End If
    
    If Mid(CboStatus, 1, 1) = "0" Then
        StrAppend = StrAppend + " AND BordxAppStatus = '0'"
    End If
    
    searchCriteria = StrAppend
    
    If ChkScan.Value = 0 Then
        Call BordxApprovalListing(StrAppend)
    Else
        Call BordxApprovalListing2(StrAppend)
    End If
   
End Function

Private Sub CmdStop_Click()
    intExit = 1
    TxtBordx1.SetFocus
    CmdClear.Enabled = True
    CmdToPayor.Enabled = True
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub CmdExit_Click()
'Dim RecordExit
'    TxtBordx1.SetFocus
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub CmdToPayor_Click()
Dim ii As Integer
Dim Cnt As Integer
Dim recordFound As Boolean
Dim recordPassed As Boolean
Dim Update

    CmdSearch.Enabled = False
    CmdToPayor.Enabled = False
    
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
    If ItemChecked = True Then
        Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
        
        If Update = vbYes Then
            Cnt = 0
            recordFound = False
            If lstBordxApproval.ListItems.Count <> 0 Then
                For ii = 0 To UBound(BordxNoArray)
                    
                    If Trim(BordxNoArray(ii)) <> "" Then
                        UpdateApprovalRecords BordxNoArray(ii)
                        recordFound = True
                    End If
                Next ii
                If recordFound = True Then
                    MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                
                Erase BordxNoArray
                Arraycnt = 0
                ReDim BordxNoArray(0)
                lstBordxApproval.ListItems.Clear
                If ChkScan.Value = 0 Then
                    Call BordxApprovalListing(searchCriteria)
                End If
            End If
        End If
    
    Else
        MsgBox "Please select at least one record to be updated! ", vbCritical
    End If
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    CmdSearch.Enabled = True
    CmdToPayor.Enabled = True
End Sub
'**********************End Command Button *******************************

'**********************Start BordxDelivery Listing  *******************************
Private Sub BordxApprovalListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlCount As String
Dim ClaimAdmin As Boolean
Dim total As String
Dim Current As String
Dim TotalAmt As Variant

    
    total = 0
    Current = 0
    
    lstBordxApproval.ListItems.Clear

    sql = " Select BordxRefNo, BordxPAYCode, BordxDate, BordxClaimAmt, BordxAppStatus, " & _
          " BordxAppDate, BordxAppBy, BordxStatus " & _
          " From BORDX where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If
    
    rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdToPayor.Enabled = True
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
        'total = rst2.recordCount
        'lblTotal = total
        
            With rst2
                Set Record = lstBordxApproval.ListItems.Add(, , "")
                
                If Len(rst2!BordxPAYCode) Then
                    Record.SubItems(1) = rst2!BordxPAYCode
                End If
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(2) = rst2!BordxRefNo
                End If
                If Len(rst2!BordxDate) Then
                    Record.SubItems(3) = Format(rst2!BordxDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxClaimAmt) Then
                    Record.SubItems(4) = Format(rst2!BordxClaimAmt, "#,##0.00")
                End If
                If Len(rst2!Bordxstatus) Then
                    Record.SubItems(5) = rst2!Bordxstatus
                End If
                If Len(rst2!BordxAPPStatus) Then
                   Record.SubItems(6) = rst2!BordxAPPStatus
                End If
                If Len(rst2!BordxAppBy) Then
                    Record.SubItems(7) = Trim(rst2!BordxAppBy)
                End If
                If Len(rst2!BordxAppDate) Then
                    Record.SubItems(8) = Format(rst2!BordxAppDate, "dd/mm/yyyy")
                End If
                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!BordxClaimAmt) Then
                    TotalAmt = TotalAmt + rst2!BordxClaimAmt
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdToPayor.Enabled = True
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub BordxApprovalListing2(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlCount As String
Dim ClaimAdmin As Boolean

    sql = " Select BordxRefNo, BordxPAYCode, BordxDate, BordxClaimAmt, BordxAppStatus, " & _
          " BordxAppDate, BordxAppBy, BordxStatus " & _
          " From BORDX where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If
    
    rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdToPayor.Enabled = True
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
        'total = rst2.recordCount
        'lblTotal = total
        
            With rst2
                Set Record = lstBordxApproval.ListItems.Add(, , "")
                
                If Len(rst2!BordxPAYCode) Then
                    Record.SubItems(1) = rst2!BordxPAYCode
                End If
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(2) = rst2!BordxRefNo
                End If
                If Len(rst2!BordxDate) Then
                    Record.SubItems(3) = Format(rst2!BordxDate, "dd/mm/yyyy")
                End If
                If Len(rst2!BordxClaimAmt) Then
                    Record.SubItems(4) = Format(rst2!BordxClaimAmt, "#,##0.00")
                End If
                If Len(rst2!Bordxstatus) Then
                    Record.SubItems(5) = rst2!Bordxstatus
                End If
                If Len(rst2!BordxAPPStatus) Then
                   Record.SubItems(6) = rst2!BordxAPPStatus
                End If
                If Len(rst2!BordxAppBy) Then
                    Record.SubItems(7) = Trim(rst2!BordxAppBy)
                End If
                If Len(rst2!BordxAppDate) Then
                    Record.SubItems(8) = Format(rst2!BordxAppDate, "dd/mm/yyyy")
                End If
                
                Current2 = Current2 + 1
                lblTotal = Current2
                If Len(rst2!BordxClaimAmt) Then
                    TotalAmt2 = TotalAmt2 + rst2!BordxClaimAmt
                End If
                LblTotalAmt = Format(TotalAmt2, "#,##0.00")
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdToPayor.Enabled = True
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
End Sub
'**********************End BordxDelivery Listing *******************************


'**********************Start Check ListView Record*******************************
Private Sub lstBordxApproval_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i, newbound  As Integer
Dim tempArray() As String
Dim Found As Boolean
Dim temp
Static Arraycnt As Integer
    
        
    If Item.Checked = True Then
        If Item.SubItems(6) = False Then
            Arraycnt = IIf(UBound(BordxNoArray) = 0, 1, UBound(BordxNoArray) + 1)
            ReDim Preserve BordxNoArray(Arraycnt)
            BordxNoArray(Arraycnt) = Item.SubItems(2)
            Arraycnt = Arraycnt + 1
            ItemChecked = True
        Else
            MsgBox "Record has been Updated!", vbCritical
            Item.Checked = False
        End If
            
    Else
        ' Remove from the array
        newbound = 0
    
        If UBound(BordxNoArray) <> 0 Then
            For i = 0 To UBound(BordxNoArray)
                If (Item.SubItems(2)) <> BordxNoArray(i) Then
                    ReDim Preserve tempArray(newbound)
                    tempArray(newbound) = BordxNoArray(i)
                    newbound = newbound + 1
                End If
            Next
            Erase BordxNoArray
            If UBound(tempArray) = 0 Then
                Arraycnt = 1
            Else
                Arraycnt = UBound(tempArray) + 1
            End If
            ReDim BordxNoArray(UBound(tempArray))
            For i = 0 To UBound(tempArray)
                BordxNoArray(i) = tempArray(i)
            Next
            Erase tempArray
        Else
            Erase BordxNoArray
            ReDim BordxNoArray(0)
            Arraycnt = 0
        End If
    End If
End Sub
'**********************End Check ListView Record*******************************

'**********************Start Update all Record functions*******************************

Private Function UpdateApprovalRecords(tBordxNo As String) As Boolean
Dim StrSql As String
'Dim Approval As New ADODB.Recordset
Dim AppDate As String

  AppDate = Format(Date, "YYYY/MM/DD")
  
    StrSql = " update Bordx set " & _
             " BordxAppDate = '" & AppDate & "', " & _
             " BordxAppStatus = '1', " & _
             " BordxAppBy = '" & Logon & "' " & _
             " Where BordxRefNo ='" & Trim(tBordxNo) & "'"
    
    medixmasterconnWS.Execute StrSql
  
  
  'strsql = " Select BordxRefNo, BordxAppStatus, BordxAppBy, BordxAppDate " & _
           " From Bordx where BordxRefNo ='" & Trim(tBordxNo) & "'"
  
  'Approval.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
  
  'If Approval.recordCount > 0 Then
    'With Approval
      '!BordxAppDate = Date
      '!BordxAppStatus = True
      '!BordxAppBy = Logon
    'End With
    'Approval.Update
    'UpdateApprovalRecords = True
  'End If
  'Approval.Close
  'Set Approval = Nothing
End Function
'**********************End Update all Record functions*******************************

'**********************Start Sorted ListView *******************************
Private Sub lstBordxApproval_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstBordxApproval, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstBordxApproval, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstBordxApproval, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstBordxApproval, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
    Case 5
        SortListView lstBordxApproval, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstBordxApproval, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstBordxApproval, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstBordxApproval, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstBordxApproval, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    End Select
End Sub
'**********************End Sorted ListView *******************************

'**********************Start Combo Listing*******************************
Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim monthstart
Dim yearstart
Dim cmd As New ADODB.Command
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
        
    Do Until PAY.EOF
        With PAY
            CboPayor.AddItem !PAYCode & " - " & Trim(!PAYCompanyName)
        End With
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
        
 '   medixmasterconn.Close
 '   Set medixmasterconn = Nothing

    With CboStatus
        .AddItem "0 - Not Approved"
        .AddItem "1 - Approved"
        .AddItem "2 - Both"
        .Text = "0 - Not Approved"
    End With
    

End Sub
'**********************End Combo Listing*******************************

Private Sub TxtBordxAmt1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtBordxAmt2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtBordxDate1_LostFocus()
    If IsDate(txtBordxDate1) Then
    Else
        If txtBordxDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtBordxDate2_LostFocus()
    If IsDate(txtBordxDate2) Then
    Else
        If txtBordxDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDate2.SetFocus
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
    lvCount = lstBordxApproval.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstBordxApproval.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
End Sub

Public Sub SetCheck(hwnd As Long, lItemIndex As Long, bState As Boolean)
Dim LV As LVITEM
Dim r As Long
    With LV
    .mask = LVIF_STATE
    .state = IIf(bState, &H2000, &H1000)
    .stateMask = LVIS_STATEIMAGEMASK
    End With
    r = SendMessageAny(hwnd, LVM_SETITEMSTATE, lItemIndex, LV)
End Sub

Private Function BordxNoCheck()
    If Trim(ChkStr(TxtBordx1)) = ChkStr(Trim(lstBordxApproval.ListItems(listloop1).SubItems(2))) Then
        Check = True
    End If
End Function

