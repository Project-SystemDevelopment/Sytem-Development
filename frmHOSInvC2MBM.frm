VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHOSInvC2MBM 
   Caption         =   "Hosp Invoice Credit To Member"
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
      Left            =   120
      TabIndex        =   15
      Top             =   7320
      Width           =   11655
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   2400
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   6480
         TabIndex        =   9
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10080
         TabIndex        =   14
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10680
         TabIndex        =   16
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9120
         TabIndex        =   13
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   330
         Left            =   5760
         TabIndex        =   8
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   330
         Left            =   7080
         TabIndex        =   10
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   330
         Left            =   7680
         TabIndex        =   11
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview"
         Height          =   330
         Left            =   8280
         TabIndex        =   12
         Top             =   200
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   3840
         TabIndex        =   18
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   195
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstInvoiceAccList 
      Height          =   5820
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10266
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CN To"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "HCN No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "CN Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Worksheet No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "MainDCN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Member No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Member Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Patient Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "HCN  Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Previous Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "DR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "CR"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   20
      Top             =   240
      Width           =   11655
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   4920
      End
      Begin VB.ComboBox cboPrintStatus 
         Height          =   315
         ItemData        =   "frmHOSInvC2MBM.frx":0000
         Left            =   1080
         List            =   "frmHOSInvC2MBM.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   530
         Width           =   1695
      End
      Begin MSMask.MaskEdBox TxtHCNDateFR 
         Height          =   285
         Left            =   7320
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtHCNDateTo 
         Height          =   285
         Left            =   9120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtBatchDateFr 
         Height          =   285
         Left            =   7320
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtBatchDateTo 
         Height          =   285
         Left            =   9120
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtBatchNo1 
         Height          =   285
         Left            =   7320
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtBatchNo2 
         Height          =   285
         Left            =   9120
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "HCN Date Fr"
         Height          =   255
         Left            =   6120
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label41 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   27
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   26
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   25
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "Batch Seq Fr"
         Height          =   255
         Left            =   6120
         TabIndex        =   24
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label40 
         Caption         =   "Batch No"
         Height          =   255
         Left            =   6120
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label45 
         Caption         =   "Print Status"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   510
         Width           =   975
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Hosp Invoice Credit To Member"
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
      TabIndex        =   29
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmHOSInvC2MBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Public searchCriteria As String
Private Hospital As String
Dim InvoiceNo As String
'Sort Direction for each column in the ListView
Private m_blnDirection(1 To 13) As Boolean

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

Private Sub cmdPreview_Click()
    
    CmdSearch.Enabled = False
    cmdPreview.Enabled = False

    If cboPrintStatus = "" Then
        MsgBox "Please select Print Status", vbOKOnly + vbCritical
        cboPrintStatus.SetFocus
    ElseIf lstInvoiceAccList.ListItems.Count = 0 Then
        MsgBox "Please select a record(s) before proceeding!", vbOKOnly + vbCritical
    Else

            cnt = 0
            recordFound = False
            
        Screen.MousePointer = vbHourglass
            Call ExportToExcelInvoice(lstInvoiceAccList)
        Screen.MousePointer = Default
    End If
    
    CmdSearch.Enabled = True
    cmdPreview.Enabled = True

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

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim X As Boolean
End Sub

Private Sub Form_Load()
    intExit = 0
    lblCurrent = 0
    'lblTotal = 0
    LblTotalAmt = 0
    'ReDim CaseNoArray(0)
    Call ComboListing
    txtPath = DocPath & "HOSInvoice"
    txtLocalFileLocation = "C:\"
    X = CreatePath("c:\cardPrinting")
    If X = False Then
        MsgBox "Error Creating Folder c:\CardPrinting !" & vbNewLine & _
                " Please create that folder manually to store the output of leaflet in sofe copy!"
    End If

End Sub

'************************Start Command Button**************************
Private Sub cmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    'ReDim CaseNoArray(0)
    lstInvoiceAccList.ListItems.Clear
    lblCurrent = 0
    'lblTotal = 0
    LblTotalAmt = 0
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub

Public Sub CmdSearch_Click()
    
    CmdSearch.Enabled = False
    cmdPreview.Enabled = False
    If cboPrintStatus = "" Then
        MsgBox "Please select Print Status", vbOKOnly + vbCritical
        cboPrintStatus.SetFocus
    Else
        CmdClear.Enabled = False
        Screen.MousePointer = vbHourglass
            Arraycnt = 0
            lstInvoiceAccList.Height = 6900
            lstInvoiceAccList.Left = 120
            lstInvoiceAccList.Top = 480
            lstInvoiceAccList.Width = 11655
            Frame1.Visible = False
            medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                Call SearchRecord
            medixmasterconnWS.Close
            Set medixmasterconnWS = Nothing
        Screen.MousePointer = Default
    End If
    CmdSearch.Enabled = True
    cmdPreview.Enabled = True
End Sub

Private Sub CmdBack_Click()
    lstInvoiceAccList.Height = 6180
    lstInvoiceAccList.Left = 120
    lstInvoiceAccList.Top = 1080
    lstInvoiceAccList.Width = 11655
    Frame1.Visible = True
End Sub

Private Sub CmdView_Click()
    lstInvoiceAccList.Height = 6900
    lstInvoiceAccList.Left = 120
    lstInvoiceAccList.Top = 480
    lstInvoiceAccList.Width = 11655
    Frame1.Visible = False
End Sub

Private Sub CmdPrintFile_Click()
    If cboPrintStatus = "" Then
        MsgBox "Please select Print Status", vbOKOnly + vbCritical
        cboPrintStatus.SetFocus
    ElseIf lstInvoiceAccList.ListItems.Count = 0 Then
        MsgBox "Please select a record(s) before proceeding!", vbOKOnly + vbCritical
    Else

    Update = MsgBox("Do you want to print the record?", vbYesNo + vbExclamation)
        If Update = vbYes Then
            cnt = 0
            recordFound = False
            
            Screen.MousePointer = vbHourglass
                medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                    If Mid(cboPrintStatus, 1, 1) = "N" Then
                        InvoiceNo = GenerateINVHMSeqNo(Date, "C")
                    End If
                
                    If Mid(cboPrintStatus, 1, 1) = "N" Then
                        Call UpdatePrintRecord
                    End If
                medixmasterconnWS.Close
                Set medixmasterconnWS = Nothing
                
                    Call ExportToExcelInvoice(lstInvoiceAccList)
            
            Screen.MousePointer = Default
        End If
    End If
End Sub
'************************End Command Button**************************

'************************Start sorted ListView**************************

Private Sub lstInvoiceAccList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    Case 4
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
    Case 12
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    Case 13
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
    End Select
End Sub
'************************End sorted ListView**************************


'************************Start ComboListing**************************

Private Sub ComboListing()
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
        
    With cboPrintStatus
        .AddItem "N - New"
        .AddItem "R - Reprint"
        .Text = "N - New"
    End With
End Sub
'************************Start ComboListing**************************

'************************Start Search Record**************************

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
    
    
    strAppend = ""
    'Hospital = ""
    
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If Mid(cboPrintStatus, 1, 1) = "N" Then
       sql = ""
       sql = " AND INVDCMBMStatus = '0' "
       strAppend = strAppend + sql
    End If
    
    If Mid(cboPrintStatus, 1, 1) = "R" Then
       sql = ""
       sql = " AND INVDCMBMStatus = '1' "
       strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtBatchDateFr)) > 0 Then
        strAppend = strAppend + " And InvHosDCBatchDate >= '" & txtBatchDateFr & "'"
    End If
    
    If Len(Trim(txtBatchDateTo)) > 0 Then
        strAppend = strAppend + " And InvHosDCBatchDate <= '" & txtBatchDateTo & "'"
    End If
    
    If Len(Trim(txtBatchNo1)) Then
        strAppend = strAppend + " AND INVBatchNo >= '" & Trim(txtBatchNo1) & "'"
    End If
    
    If Len(Trim(txtBatchNo2)) Then
        strAppend = strAppend + " AND INVBatchNo <= '" & Trim(txtBatchNo2) & "'"
    End If
        
    If Len(Trim(TxtHCNDateFR)) And IsDate(TxtHCNDateFR) = True Then
        strAppend = strAppend + " And DebitCreditHDate >= '" & Format(TxtHCNDateFR, "mm/dd/YYYY") & "'"
    End If
    
    If Len(Trim(TxtHCNDateTo)) And IsDate(TxtHCNDateTo) = True Then
        strAppend = strAppend + " And DebitCreditHDate <= '" & Format(TxtHCNDateTo, "mm/dd/yyyy") & "'"
    End If
    
    searchCriteria = strAppend
    Call InvoiceAccountListing(strAppend)

End Function
'************************End Search Record**************************
'************************Start List Record**************************

'************************Start List Record**************************

Private Sub InvoiceAccountListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim HOS As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim strsqlHOS As String
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
Dim TotalInvNo As String
Dim DRCRAmt As String
    
    Total = 0
    Current = 0
    TotalAmt = 0
    LblTotalAmt = ""
    lstInvoiceAccList.ListItems.Clear
    DRCRAmt = 0
    sql = " Select CASNumber, INVWSVersion, MBMNumber, MBMName, PatientName, INSCode, " & _
          " CASDischargeDate, CaseRegDate, PLNCode, CLMPayableByMedix2M, " & _
          " INVHosDCPrevAmt, CASDateAdmitted, INVPayTo, INVPayee,  INVBordxRefNo,  " & _
          " INVDCMBMStatus, DebitCreditRefNo, DebitCreditHRefNo, DebitCreditHAmt, DebitCreditHDate " & _
          " From VIEW_InvoiceAccount where ClaimStatus = 'C' And INVPayTo = 'H' And Substring(DebitCreditHRefNo,2,1) = 'C'"
          
          
    'strCount = "select count(*) as totalRecord From VIEW_InvoiceAccount where 0=0 "
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
        strCount = strCount + strAppend
    End If

    'Set rstCount = medixmasterconnWS.Execute(strCount)
 
    
    rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
       
    'If rstCount!totalrecord = 0 Then
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
         
    Do
        Do Until rst2.EOF
        'Total = rst2.recordCount
        'lblTotal = Total
        
            
            With rst2
                                            
                                
                'If Len(!AccPeriodMonth) Then
                Set Record = lstInvoiceAccList.ListItems.Add(, , "MEMBER")
                'Else
                '    Set Record = lstInvoiceAccList.listitems.Add(, , "Not")
                'End If
                
                'If Len(rst2!INVPAYEE) Then
                '    Record.SubItems(1) = rst2!INVPAYEE
                'End If
                                                                       
                'strsqlHOS = "select HosCode, HosName from HOS where HosCode = '" & Trim(Record.SubItems(1)) & "'"
                'HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                
                'If HOS.recordCount Then
                '    Record.SubItems(1) = Trim(HOS!HOSName)
                'Else
                If Len(rst2!DebitCreditHRefNo) Then
                    Record.SubItems(1) = rst2!DebitCreditHRefNo
                End If
                
                If Len(rst2!DebitCreditHDate) Then
                    Record.SubItems(2) = rst2!DebitCreditHDate
                End If
                
                'If Len(rst2!INVWSVersion) Then
                    'Record.SubItems(3) = ""
                'End If
                
                If Len(rst2!CASNumber & "-" & rst2!INVWSVersion) Then
                    Record.SubItems(3) = rst2!CASNumber & "-" & rst2!INVWSVersion
                End If
                
                If Len(rst2!DebitCreditRefNo) Then
                    Record.SubItems(4) = rst2!DebitCreditRefNo
                End If
                
                If Len(rst2!INVBordxRefNo) Then
                    Record.SubItems(5) = Trim(rst2!INVBordxRefNo)
                End If
                                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(6) = Trim(rst2!MBMNumber)
                End If
                
                If Len(rst2!MBMName) Then
                    Record.SubItems(7) = Trim(rst2!MBMName)
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(8) = Trim(rst2!PatientName)
                End If
                
                
                'If Len(rst2!INVAmount) Then
                '    Record.SubItems(11) = Format(rst2!INVAmount, "#,##0.00;(#,##0.00)")
                'End If
                
                If Len(rst2!DebitCreditHAmt) Then
                    Record.SubItems(9) = Format(rst2!DebitCreditHAmt, "#,##0.00;(#,##0.00)")
                End If
                
                If Len(rst2!INVHosDCPrevAmt) Then
                    Record.SubItems(10) = Format(rst2!INVHosDCPrevAmt, "#,##0.00;(#,##0.00)")
                End If
                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!DebitCreditHAmt) Then
                    TotalAmt = TotalAmt + rst2!DebitCreditHAmt
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
                
                If rst2!INVHosDCPrevAmt >= 0 And rst2!DebitCreditHAmt >= 0 Then
                    DRCRAmt = CDbl(rst2!DebitCreditHAmt) - CDbl(rst2!INVHosDCPrevAmt)
                '    Record.SubItems(14) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                ElseIf rst2!INVHosDCPrevAmt < 0 And rst2!DebitCreditHAmt >= 0 Then
                    DRCRAmt = CDbl(rst2!DebitCreditHAmt) + CDbl(rst2!INVHosDCPrevAmt)
                '    Record.SubItems(14) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                ElseIf rst2!INVHosDCPrevAmt < 0 And rst2!DebitCreditHAmt < 0 Then
                    DRCRAmt = CDbl(rst2!DebitCreditHAmt) - CDbl(rst2!INVHosDCPrevAmt)
                '   Record.SubItems(14) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                Else
                    DRCRAmt = rst2!DebitCreditHAmt
                '    Record.SubItems(14) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                End If
                
                If DRCRAmt < 0 Then
                    Record.SubItems(11) = Format(DRCRAmt, "#,##0.00;#,##0.00")
                ElseIf DRCRAmt > 0 Then
                    Record.SubItems(12) = Format(DRCRAmt, "(#,##0.00)")
                ElseIf DRCRAmt = 0 Then
                    Record.SubItems(12) = Format(DRCRAmt, "#,##0.00;#,##0.00")
                End If

            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   check = False
                   Exit Do
                End If
            'HOS.Close
            'Set HOS = Nothing
        Loop
    Loop Until check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    
End Sub
'************************End List Record**************************


'************************Start ComboBox Change/KeyPress**************************

Private Sub cboPrintStatus_Change()
If cboPrintStatus.Text = "" Then
    If InStr(cboPrintStatus.Text, "null") Then
       cboPrintStatus.Enabled = False
    End If
End If
End Sub


Private Sub cboPrintStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer


If cboPrintStatus.Text = "" Then
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPrintStatus.Text, cboPrintStatus.SelStart) + Chr(KeyAscii) + Mid(cboPrintStatus.Text, cboPrintStatus.SelStart + cboPrintStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPrintStatus.ListCount - 1
 
        If NewText = LCase(Left(cboPrintStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPrintStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPrintStatus.Text = ValidValue
                cboPrintStatus.SelStart = 0
                cboPrintStatus.SelLength = Len(ValidValue)
            End If
        End If
End If
End Sub

'************************End ComboBox Change/KeyPress**************************





'**********************Start Update Record*******************************
Private Sub UpdatePrintRecord()
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim Update
Dim strsql As String
Dim AccPeriod As New ADODB.Recordset
             
        
    strsqlCLM = " Select CASNumber, INVWSVersion, DebitCreditHAmt " & _
    " From VIEW_InvoiceAccount where INVPayTo = 'H' And Substring(DebitCreditRefNo,1,1) = 'D'"
    
    CLM.Open strsqlCLM, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                
    Do Until CLM.EOF
    
        strsql = " Select CASNumber, INVWSVersion, INVDCMBMStatus, INVHosDCPrevAmt, " & _
        " InvHosDCBatchDate, InvHosDCBatchNo, InvHosDCPrintLogon, LastAccPeriodDate, INVPayTo " & _
        " From ClaimInvoice where CASNumber = '" & CLM!CASNumber & "' And INVWSVersion = '" & CLM!INVWSVersion & "' And INVPayTo = 'H'"
        
        AccPeriod.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
        Do Until AccPeriod.EOF
            With AccPeriod
                !InvHosDCBatchDate = Mid(InvoiceNo, 1, 10)
                !InvHosDCBatchNo = Trim(Mid(InvoiceNo, 12, 3))
                !InvHosDCPrintLogon = Logon
                !INVDCMBMStatus = "1"
                !INVHosDCPrevAmt = CLM!DebitCreditHAmt
            AccPeriod.Update
            AccPeriod.MoveNext
            End With
        Loop
        AccPeriod.Close
        Set AccPeriod = Nothing
        
    CLM.MoveNext
    Loop
    CLM.Close
    Set CLM = Nothing
End Sub
'**********************End Update Record*******************************

Private Sub MBMAccASCII()
Dim strsql As String
Dim AccPeriod As New ADODB.Recordset
Dim TotalSupplementary As Integer
Dim strsplOthers As String
Dim MBMOthers As New ADODB.Recordset
Dim MBMGroupCompany As String
Dim MBMCRELCode As String
Dim check As Boolean
Dim MBMDateJoin As String
        
        cnt = 0
        
        strsql = " Select CASNumber, INVWSVersion, MBMNumber, INSCode, " & _
            " INVNo, CASDischargeDate, CaseRegDate, PLNCode, " & _
            " CASDateAdmitted, INVPayTo, INVPayee, PatientName, AccPeriodMonth, AccPeriodYear " & _
            " MBMName, InvPrintStatus, INVPreviousAmt " & _
            " From VIEW_InvoiceAccount where ClaimStatus = 'C' And Substring(CASNumber,1,2) = '" & Mid(CboPayor, 1, 2) & "'"
               
        AccPeriod.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic

        
        Do Until AccPeriod.EOF
        For ii = 1 To AccPeriod.recordCount

        DataString = DataString & l("M", 1)
        DataString = DataString & l(AccPeriod!CASNumber & "-" & AccPeriod!INVWSVersion, 20)
        DataString = DataString & l(AccPeriod!MBMNumber, 25)
        DataString = DataString & l(AccPeriod!MBMName, 60)
        DataString = DataString & l(AccPeriod!PatientName, 60)
        DataString = DataString & l(AccPeriod!INSCode, 1)
        DataString = DataString & l(AccPeriod!PLNCode, 4)
        DataString = DataString & l(AccPeriod!INVNo, 10)
        'DataString = DataString & l(AccPeriod!INVAmount, 10)
        'DataString = DataString & l(AccPeriod!CLMTotalApproved, 10)
       
        AccPeriod.MoveNext
        
        'Debug.Print DataString
        Print #1, DataString
        Print #2, DataString
             'If intExit = 1 Then
             '    check = False
             '    Exit Do
             'End If
         
         'End If
        Next ii
 
        Loop
        'Loop Until check = False
End Sub

Public Function CreatePath(NewPath) As Boolean
    'Add a trailing slash if none
    If Right(NewPath, 1) <> "\" Then
        NewPath = NewPath & "\"
    End If
    
    'Call API
    If MakeSureDirectoryPathExists(NewPath) <> 0 Then
        'No errors, return True
        CreatePath = True
    End If

End Function

'**********************Start Update Record*******************************
Private Sub UpdateRecord()
Dim Update
Dim strsql As String
Dim AccPeriod As New ADODB.Recordset

         
    strsql = " Select CASNumber, INVWSVersion, AccPeriodStatus, " & _
    " AccPeriodMonth, AccPeriodYear, LastAccPeriodUser, LastAccPeriodDate " & _
    " From ClaimInvoice where CASNumber = '" & Trim(CASNo) & "' " & _
    " AND INVWSVersion = '" & Trim(CASVer) & "' AND INVNo = '" & Trim(INVNo) & "'"
    AccPeriod.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
      
    With AccPeriod
        !LastAccPeriodUser = Logon
        !LastAccPeriodDate = Date
        '!AccPeriodStatus = True
        '!InvPrintStatus = True
        !AccPeriodMonth = Trim(CboPeriodMth)
        !AccPeriodYear = Trim(CboPeriodYr)
    End With
    AccPeriod.Update
    AccPeriod.Close
    Set AccPeriod = Nothing
End Sub
'**********************End Update Record*******************************
