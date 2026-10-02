VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMainDC 
   Caption         =   "Main DCN Checking"
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
      Height          =   1455
      Left            =   0
      TabIndex        =   21
      Top             =   240
      Width           =   11775
      Begin VB.ComboBox cboPayor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   3600
      End
      Begin VB.TextBox TxtCaseNo2 
         Height          =   315
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   2
         Top             =   530
         Width           =   1575
      End
      Begin VB.TextBox TxtCaseNo1 
         Height          =   315
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   1
         Top             =   530
         Width           =   1575
      End
      Begin VB.TextBox TxtVersion2 
         Height          =   315
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   4
         Top             =   810
         Width           =   1575
      End
      Begin VB.TextBox TxtVersion1 
         Height          =   315
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   3
         Top             =   810
         Width           =   1575
      End
      Begin VB.TextBox TxtDCNo2 
         Height          =   315
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1095
         Width           =   1575
      End
      Begin VB.TextBox TxtDCNo1 
         Height          =   315
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1095
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtDCDate1 
         Height          =   315
         Left            =   6720
         TabIndex        =   7
         Top             =   180
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDCDate2 
         Height          =   315
         Left            =   8760
         TabIndex        =   8
         Top             =   180
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtHDNNo1 
         Height          =   315
         Left            =   6720
         MaxLength       =   15
         TabIndex        =   9
         Top             =   465
         Width           =   1575
      End
      Begin VB.TextBox TxtHDNNo2 
         Height          =   315
         Left            =   8760
         MaxLength       =   15
         TabIndex        =   10
         Top             =   465
         Width           =   1575
      End
      Begin VB.TextBox TxtHCNNo2 
         Height          =   315
         Left            =   8760
         MaxLength       =   15
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TxtHCNNo1 
         Height          =   315
         Left            =   6720
         MaxLength       =   15
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "To"
         Height          =   255
         Left            =   8400
         TabIndex        =   33
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label9 
         Caption         =   "HCN No"
         Height          =   255
         Left            =   5400
         TabIndex        =   32
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label7 
         Caption         =   "To"
         Height          =   255
         Left            =   8400
         TabIndex        =   31
         Top             =   525
         Width           =   360
      End
      Begin VB.Label Label4 
         Caption         =   "HDN No"
         Height          =   255
         Left            =   5400
         TabIndex        =   30
         Top             =   525
         Width           =   1080
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   8400
         TabIndex        =   29
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Main DC Date"
         Height          =   255
         Left            =   5400
         TabIndex        =   28
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   27
         Top             =   1150
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "Main DC No"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1150
         Width           =   1080
      End
      Begin VB.Label Label5 
         Caption         =   "Ver. No"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   915
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Wks. No"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   630
         Width           =   840
      End
      Begin VB.Label Label6 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   920
         Width           =   255
      End
      Begin VB.Label Label32 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   630
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   18
      Top             =   7320
      Width           =   11895
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   7920
         TabIndex        =   14
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print to File"
         Height          =   375
         Left            =   8760
         TabIndex        =   15
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   375
         Left            =   10560
         TabIndex        =   17
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9720
         TabIndex        =   16
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   7080
         TabIndex        =   13
         Top             =   160
         Width           =   855
      End
      Begin VB.Label Label64 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   5535
      Left            =   0
      TabIndex        =   20
      Top             =   1800
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Claim No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "CAS Pay Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "MDCN"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "HDN"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "HCN"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "MDCN Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "HDN Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "HCN Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Reimbursement Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Check Total"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Main DCN Checking"
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
      TabIndex        =   35
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmMainDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private intExit As Integer
Private m_blnDirection(1 To 11) As Boolean

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    ListView2.ListItems.Clear
    lblCurrent = 0
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
    Call PayorComboListing
End Sub

Private Sub cmdPrint_Click()
    'Call ExportToExcelAnalysis26(ListView2)
    Screen.MousePointer = vbHourglass
    
    If ListView2.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(ListView2)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    Screen.MousePointer = Default
End Sub

Private Sub CmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    CmdExit.Enabled = True
    CmdClear.Enabled = True
    cmdPrint.Enabled = True
End Sub

Private Sub cmdSearch_Click()
    cmdPrint.Enabled = False
    CmdClear.Enabled = False
    CmdExit.Enabled = False
    CmdSearch.Enabled = False
    Screen.MousePointer = vbHourglass
    medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        Call SearchRecords
    medixmasterconnWS.Close
    Set medixmasterconnWS = Nothing
    Screen.MousePointer = vbDefault
    CmdSearch.Enabled = True
End Sub

Private Sub PayorComboListing()
Dim PAY As New ADODB.Recordset
Dim cmd As New ADODB.Command
    
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
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
            PAY.MoveNext
        End With
    Loop
    PAY.Close
    Set PAY = Nothing
            
    medixmasterconn.Close
    Set medixmasterconn = Nothing
  
End Sub

Private Function SearchRecords()
Dim sql As String
Dim strAppend As String

    strAppend = ""
    
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And substring(CASNumber,1,2) = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If Len(Trim(TxtCaseNo1)) Then
        strAppend = strAppend + " And CASNumber >= '" & Trim(TxtCaseNo1) & "'"
    End If
    
    If Len(Trim(TxtCaseNo2)) Then
        strAppend = strAppend + " And CASNumber <= '" & Trim(TxtCaseNo2) & "'"
    End If
    
    If Len(Trim(TxtVersion1)) Then
        strAppend = strAppend + " And claimversion >= '" & Trim(TxtVersion1) & "'"
    End If
    
    If Len(Trim(txtVersion2)) Then
        strAppend = strAppend + " And claimversion <= '" & Trim(txtVersion2) & "'"
    End If
    
    If Len(Trim(TxtDCNo1)) Then
        strAppend = strAppend + " And DebitCreditRefNo >= '" & Trim(TxtDCNo1) & "'"
    End If
    
    If Len(Trim(TxtDCNo2)) Then
        strAppend = strAppend + " And DebitCreditRefNo <= '" & Trim(TxtDCNo2) & "'"
    End If
    
    If (txtDCDate1) <> "__/__/____" And IsDate(txtDCDate1) = True _
        And (txtDCDate2) <> "__/__/____" And IsDate(txtDCDate2) = True Then
        sql = ""
        sql = " AND DebitCreditDate >= '" & Format(Trim(txtDCDate1), "mm/dd/yyyy") & _
              "' And DebitCreditDate <= '" & Format(Trim(txtDCDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtHDNNo1)) Then
        strAppend = strAppend + " And DebitCreditHRefNo >= '" & Trim(TxtHDNNo1) & "'"
    End If
    
    If Len(Trim(TxtHDNNo2)) Then
        strAppend = strAppend + " And DebitCreditHRefNo <= '" & Trim(TxtHDNNo2) & "'"
    End If
    
    If Len(Trim(TxtHCNNo1)) Then
        strAppend = strAppend + " And DebitCreditHRefNo >= '" & Trim(TxtHCNNo1) & "'"
    End If
    
    If Len(Trim(TxtHCNNo2)) Then
        strAppend = strAppend + " And DebitCreditHRefNo <= '" & Trim(TxtHCNNo2) & "'"
    End If
    
    Call ClmCheckTotalListing(strAppend)
    
End Function

Private Sub ClmCheckTotalListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim CLM As New ADODB.Recordset
Dim Pay2MBM As Double
Dim strsql As String
    
    ListView2.ListItems.Clear
    Current = 0
    Total = 0
    TotalChkAmt = 0
            
    sql = " SELECT CASNumber, ClaimVersion, DebitCreditDate, CLMPayableByMedix2H, CLMTotalApproved, " & _
          " DebitCreditRefNo, DebitCreditHRefNo, CLMPayableByMedix2M, DebitCreditHAmt, CASPAYType " & _
          " From VIEW_MainDCN Where ((Substring(DebitCreditRefNo,1,1)= 'D' or Substring(DebitCreditRefNo,1,1) = 'C') or (Substring(DebitCreditHRefNo,2,1)= 'D' or Substring(DebitCreditHRefNo,2,1) = 'C'))"
    
    '(INVPayTo = 'H' OR INVPayTo = 'M') AND
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
          
    rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    
    Do
    
        Do Until rst2.EOF
        
                        
        With rst2
            
            If Len(rst2!CASNumber) And Len(rst2!ClaimVersion) Then
                Set Record = ListView2.ListItems.Add(, , rst2!CASNumber & "-" & rst2!ClaimVersion)
            End If
            
            If Len(rst2!CASPAYType) Then
                Record.SubItems(1) = Trim(rst2!CASPAYType)
            End If
                            
            If Len(rst2!DebitCreditDate) Then
                Record.SubItems(2) = Format(rst2!DebitCreditDate, "dd/mm/yyyy")
            End If
                            
            If Len(rst2!DebitCreditRefNo) Then
                Record.SubItems(3) = Trim(rst2!DebitCreditRefNo)
            End If
            
            If Len(rst2!DebitCreditHRefNo) Then
                If Mid(rst2!DebitCreditHRefNo, 1, 2) = "HD" Then
                    Record.SubItems(4) = Trim(rst2!DebitCreditHRefNo)
                    If Len(rst2!DebitCreditHAmt) Then
                        Record.SubItems(7) = Format(rst2!DebitCreditHAmt, "#,##0.00")
                        Record.SubItems(8) = "0.00"
                    End If
                ElseIf Mid(rst2!DebitCreditHRefNo, 1, 2) = "HC" Then
                    Record.SubItems(5) = Trim(rst2!DebitCreditHRefNo)
                    If Len(rst2!DebitCreditHAmt) Then
                        Record.SubItems(8) = Format(rst2!DebitCreditHAmt, "#,##0.00")
                        Record.SubItems(7) = "0.00"
                    End If
                End If
            Else
                Record.SubItems(7) = "0.00"
                Record.SubItems(8) = "0.00"
            End If
            

            strsql = " SELECT CLMPayableByMedix2H, CLMTotalApproved, CLMPayableByMedix2M, CLMReimburse " & _
            " From Claim Where CASNumber = '" & rst2!CASNumber & "' And claimversion = '" & rst2!ClaimVersion & "'"
            
            CLM.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
            If CLM.recordCount Then
                'If Len(CLM!CLMPayableByMedix2H) Then
                    If Len(CLM!CLMPayableByMedix2M) Then
                        Record.SubItems(6) = Format(CLM!CLMPayableByMedix2M, "#,##0.00")
                    Else
                        Record.SubItems(6) = "0.00"
                    End If
                'Else
                    'Record.SubItems(6) = "0.00"
                'End If
                
            If Len(CLM!CLMReimburse) Then
                Record.SubItems(9) = Format(CLM!CLMReimburse, "#,##0.00")
            Else
                Record.SubItems(9) = "0.00"
            End If

            CLM.Close
            Set CLM = Nothing

            End If
            TotalChkAmt = CDbl(Record.SubItems(6)) - CDbl(Record.SubItems(7)) - CDbl(Record.SubItems(8)) - CDbl(Record.SubItems(9))
            'LblMCO = Format(TotalMCO, "#,##0.00")
            Record.SubItems(10) = Format(TotalChkAmt, "#,##0.00")
            
            Current = Current + 1
            lblCurrent = Current
            
            'End If
      
        End With
        rst2.MoveNext
        DoEvents
    
        If intExit = 1 Then
            check = False
            Exit Do
        End If
        Loop
    Loop Until check = False
    'End If
    intExit = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdPrint.Enabled = True
        CmdClear.Enabled = True
        CmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    cmdPrint.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    Case 4
        SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
    End Select
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

Private Sub txtDCDate1_LostFocus()
    If IsDate(txtDCDate1) Then
    Else
        If txtDCDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDCDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDCDate2_LostFocus()
    If IsDate(txtDCDate2) Then
    Else
        If txtDCDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDCDate2.SetFocus
        End If
    End If
End Sub
