VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInvoiceAC 
   Caption         =   "Invoice Accounting"
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
   Begin VB.Frame Frame3 
      Caption         =   "For Update A/C Period"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   6700
      Width           =   5775
      Begin VB.ComboBox CboPeriodYr 
         Height          =   315
         ItemData        =   "FrmInvoiceAC.frx":0000
         Left            =   4080
         List            =   "FrmInvoiceAC.frx":0002
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox CboPeriodMth 
         Height          =   315
         ItemData        =   "FrmInvoiceAC.frx":0004
         Left            =   1320
         List            =   "FrmInvoiceAC.frx":0006
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "A/C Period Yr"
         Height          =   255
         Left            =   2880
         TabIndex        =   22
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "A/C Period Mth"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   23
      Top             =   7320
      Width           =   11655
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   3960
         TabIndex        =   55
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   1320
         TabIndex        =   54
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   6840
         TabIndex        =   31
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10320
         TabIndex        =   30
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10920
         TabIndex        =   29
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9360
         TabIndex        =   28
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   330
         Left            =   6120
         TabIndex        =   27
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   7440
         TabIndex        =   26
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   330
         Left            =   8160
         TabIndex        =   25
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   330
         Left            =   8760
         TabIndex        =   24
         Top             =   200
         Width           =   615
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   3120
         TabIndex        =   32
         Top             =   195
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   1000
         TabIndex        =   34
         Top             =   195
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   2280
         TabIndex        =   33
         Top             =   195
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstInvoiceAccList 
      Height          =   4620
      Left            =   120
      TabIndex        =   35
      Top             =   2040
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8149
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "A/C Period Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Pay To"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Case Number"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Version"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Worksheet No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Member No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Member Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "INS Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Plan Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "INV No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "INV Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Approved Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Previous Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1785
      Left            =   120
      TabIndex        =   36
      Top             =   240
      Width           =   11655
      Begin VB.CheckBox ChkBatchNo 
         Height          =   255
         Left            =   10680
         TabIndex        =   18
         Top             =   1440
         Width           =   375
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   255
         Width           =   4320
      End
      Begin VB.CheckBox chkDOR 
         Height          =   255
         Left            =   10680
         TabIndex        =   6
         Top             =   300
         Width           =   375
      End
      Begin VB.ComboBox CboPayTo 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   540
         Width           =   4320
      End
      Begin VB.CheckBox chkDOA 
         Height          =   255
         Left            =   10680
         TabIndex        =   9
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chkDOD 
         Height          =   255
         Left            =   10680
         TabIndex        =   12
         Top             =   855
         Width           =   375
      End
      Begin VB.CheckBox chkBatchDate 
         Height          =   255
         Left            =   10680
         TabIndex        =   15
         Top             =   1170
         Width           =   375
      End
      Begin MSMask.MaskEdBox txtDORDate2 
         Height          =   315
         Left            =   9120
         TabIndex        =   5
         Top             =   255
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDORDate1 
         Height          =   315
         Left            =   7080
         TabIndex        =   4
         Top             =   255
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOADate1 
         Height          =   315
         Left            =   7080
         TabIndex        =   7
         Top             =   540
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOADate2 
         Height          =   315
         Left            =   9120
         TabIndex        =   8
         Top             =   540
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODDate1 
         Height          =   315
         Left            =   7080
         TabIndex        =   10
         Top             =   825
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODDate2 
         Height          =   315
         Left            =   9120
         TabIndex        =   11
         Top             =   825
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBatchDate1 
         Height          =   315
         Left            =   7080
         TabIndex        =   13
         Top             =   1110
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBatchDate2 
         Height          =   315
         Left            =   9120
         TabIndex        =   14
         Top             =   1110
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtBatchNo1 
         Height          =   285
         Left            =   7080
         MaxLength       =   4
         TabIndex        =   16
         Top             =   1395
         Width           =   1335
      End
      Begin VB.TextBox txtBatchNo2 
         Height          =   285
         Left            =   9120
         TabIndex        =   17
         Top             =   1395
         Width           =   1455
      End
      Begin VB.ComboBox cboPrintStatus 
         Height          =   315
         ItemData        =   "FrmInvoiceAC.frx":0008
         Left            =   1320
         List            =   "FrmInvoiceAC.frx":000A
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   360
         TabIndex        =   50
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "DOR"
         Height          =   255
         Left            =   6120
         TabIndex        =   49
         ToolTipText     =   "Date of Register"
         Top             =   375
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "DOA"
         Height          =   255
         Left            =   6120
         TabIndex        =   48
         ToolTipText     =   "Date of Admitted"
         Top             =   645
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "DOD"
         Height          =   255
         Left            =   6120
         TabIndex        =   47
         ToolTipText     =   "Date of Discharge"
         Top             =   945
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   8640
         TabIndex        =   46
         Top             =   375
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   8640
         TabIndex        =   45
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   8640
         TabIndex        =   44
         Top             =   945
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "To"
         Height          =   255
         Left            =   8640
         TabIndex        =   43
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Batch Date"
         Height          =   255
         Left            =   6120
         TabIndex        =   42
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Pay To"
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label40 
         Caption         =   "Batch No"
         Height          =   180
         Left            =   6120
         TabIndex        =   40
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label41 
         Caption         =   "To"
         Height          =   255
         Left            =   8640
         TabIndex        =   39
         Top             =   1515
         Width           =   375
      End
      Begin VB.Label Label45 
         Caption         =   "Print Status"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   10680
         TabIndex        =   37
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Invoice  Accounting"
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
      Height          =   220
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   11805
   End
   Begin VB.Label Label39 
      Caption         =   "* - Search Empty Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6120
      TabIndex        =   51
      Top             =   7080
      Width           =   2295
   End
End
Attribute VB_Name = "FrmInvoiceAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim CaseNoArray() As String
Public searchCriteria As String
Dim CASNo As String
Dim CASVer As String
Dim PAYTo As String
Dim INVNo As String
'Sort Direction for each column in the ListView
Private m_blnDirection(1 To 15) As Boolean


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
    ReDim CaseNoArray(0)
    Call ComboListing
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
    ReDim CaseNoArray(0)
    lstInvoiceAccList.ListItems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub

Public Sub cmdSearch_Click()
    CmdClear.Enabled = False
    CmdSearch.Enabled = False
    CmdUpdate.Enabled = False
    Screen.MousePointer = vbHourglass
        Arraycnt = 0
        lstInvoiceAccList.Height = 6300
        lstInvoiceAccList.Left = 120
        lstInvoiceAccList.Top = 360
        lstInvoiceAccList.Width = 11655
        Frame1.Visible = False
        'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
            Call SearchRecord
       ' medixmasterconnWS.Close
       ' Set medixmasterconnWS = Nothing
    Screen.MousePointer = Default
    CmdSearch.Enabled = True
    CmdUpdate.Enabled = True

End Sub

Private Sub CmdBack_Click()
    lstInvoiceAccList.Height = 4380
    lstInvoiceAccList.Left = 120
    lstInvoiceAccList.Top = 2280
    lstInvoiceAccList.Width = 11655
    Frame1.Visible = True
End Sub


Private Sub CmdPrintFile_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelInvoice(lstInvoiceAccList)
    Screen.MousePointer = Default
End Sub

Private Sub CmdView_Click()
    lstInvoiceAccList.Height = 6300
    lstInvoiceAccList.Left = 0
    lstInvoiceAccList.Top = 360
    lstInvoiceAccList.Width = 11655
    Frame1.Visible = False
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
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
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
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    Case 13
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
    Case 14
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
    Case 15
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
    End Select
End Sub
'************************End sorted ListView**************************


'************************Start ComboListing**************************

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim sql As String
Dim YearStart
Dim MonthStart
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
            PAY.MoveNext
        End With
    Loop
    PAY.Close
    Set PAY = Nothing
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
       
    With cboPrintStatus
        .AddItem "N - New"
        .AddItem "R - Reprint"
        .AddItem "A - All"
    End With
    
    With CboPayTo
        .AddItem "H - Hospital"
        .AddItem "M - Member"
    End With
            
    For MonthStart = 1 To 12
        With CboPeriodMth
            .AddItem MonthStart
        End With
    Next MonthStart
    
    For YearStart = 1999 To 3000
        With CboPeriodYr
            .AddItem YearStart
        End With
    Next YearStart
End Sub
'************************Start ComboListing**************************

'************************Start Search Record**************************

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
    
    
    strAppend = ""
    
        
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
        
    If Len(Trim(CboPayTo)) Then
        strAppend = strAppend + " AND INVPayTo = '" & Trim(Mid(CboPayTo, 1, IIf(InStr(1, CboPayTo, "-") = 0, 4, InStr(1, CboPayTo, "-") - 1))) & "'"
    End If
           
    If chkDOA.Value = 1 Then
        sql = ""
        sql = "AND (CasDateAdmitted Is null OR CasDateAdmitted = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If chkDOD.Value = 1 Then
        sql = ""
        sql = "AND (CasDischargeDate Is null OR CasDischargeDate = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If chkDOR.Value = 1 Then
        sql = ""
        sql = "AND (CaseRegDate Is null OR CaseRegDate = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If chkBatchDate.Value = 1 Then
        sql = ""
        sql = "AND (INVBatchDate Is null OR INVBatchDate = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If ChkBatchNo.Value = 1 Then
        sql = ""
        sql = "AND (INVBatchNo Is null OR INVBatchNo = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If (txtDORDate1) <> "__/__/____" And IsDate(txtDORDate1) = True _
        And (txtDORDate2) <> "__/__/____" And IsDate(txtDORDate2) = True Then
        sql = ""
        sql = " AND CaseRegDate >= '" & Format(Trim(txtDORDate1), "mm/dd/yyyy") & _
              "' And CaseRegDate <= '" & Format(Trim(txtDORDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDORDate1)) > 0 And IsDate(txtDORDate1) = True Then
        strAppend = strAppend + " And CaseRegDate >= '" & Format(txtDORDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDORDate2)) > 0 And IsDate(txtDORDate2) = True Then
        strAppend = strAppend + " And CaseRegDate <= '" & Format(txtDORDate2, "mm/dd/yyyy") & "'"
    End If
    
    If (txtDOADate1) <> "__/__/____" And IsDate(txtDOADate1) = True _
        And (txtDOADate2) <> "__/__/____" And IsDate(txtDOADate2) = True Then
        sql = ""
        sql = " AND CasDateAdmitted >= '" & Format(Trim(txtDOADate1), "mm/dd/yyyy") & _
              "' And CasDateAdmitted <= '" & Format(Trim(txtDOADate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDOADate1)) > 0 And IsDate(txtDOADate1) = True Then
        strAppend = strAppend + " And CasDateAdmitted >= '" & Format(txtDOADate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOADate2)) > 0 And IsDate(txtDOADate2) = True Then
        strAppend = strAppend + " And CasDateAdmitted <= '" & Format(txtDOADate2, "mm/dd/yyyy") & "'"
    End If
    
    If (txtDODDate1) <> "__/__/____" And IsDate(txtDODDate1) = True _
        And (txtDODDate2) <> "__/__/____" And IsDate(txtDODDate2) = True Then
        sql = ""
        sql = " AND CasDischargeDate >= '" & Format(Trim(txtDODDate1), "mm/dd/yyyy") & _
              "' And CasDischargeDate <= '" & Format(Trim(txtDODDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDODDate1)) > 0 And IsDate(txtDODDate1) = True Then
        strAppend = strAppend + " And CasDischargeDate >= '" & Format(txtDODDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDODDate2)) > 0 And IsDate(txtDODDate2) = True Then
        strAppend = strAppend + " And CasDischargeDate <= '" & Format(txtDODDate2, "mm/dd/yyyy") & "'"
    End If
    
    If (txtBatchDate1) <> "__/__/____" And IsDate(txtBatchDate1) = True _
        And (txtBatchDate2) <> "__/__/____" And IsDate(txtBatchDate2) = True Then
        sql = ""
        sql = " AND INVBatchDate >= '" & Format(Trim(txtBatchDate1), "mm/dd/yyyy") & _
              "' And INVBatchDate <= '" & Format(Trim(txtBatchDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtBatchDate1)) > 0 And IsDate(txtBatchDate1) = True Then
        strAppend = strAppend + " And INVBatchDate >= '" & Format(txtBatchDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtBatchDate2)) > 0 And IsDate(txtBatchDate2) = True Then
        strAppend = strAppend + " And INVBatchDate <= '" & Format(txtBatchDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Mid(cboPrintStatus, 1, 1) = "N" Then
       sql = ""
       sql = " AND INVPrintStatus = '0' "
       strAppend = strAppend + sql
    End If
    
    If Mid(cboPrintStatus, 1, 1) = "R" Then
       sql = ""
       sql = " AND INVPrintStatus = '1' "
       strAppend = strAppend + sql
    End If
           
    If Len(Trim(txtBatchNo1)) Then
        strAppend = strAppend + " AND INVBatchNo >= '" & Trim(txtBatchNo1) & "'"
    End If
    
    If Len(Trim(txtBatchNo2)) Then
        strAppend = strAppend + " AND INVBatchNo <= '" & Trim(txtBatchNo2) & "'"
    End If
    
        
    searchCriteria = strAppend
    Call InvoiceAccountListing(strAppend)

End Function
'************************End Search Record**************************


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
    
    Total = 0
    Current = 0
    TotalAmt = 0
    LblTotalAmt = ""
    lstInvoiceAccList.ListItems.Clear
       
    sql = " Select CASNumber, INVWSVersion, MBMNumber, INSCode, " & _
          " INVNo, INVAmount, CASDischargeDate, CaseRegDate, PLNCode, " & _
          " CASDateAdmitted, INVPayTo, INVPayee, PatientName, " & _
          " MBMName, INVAmount, AccPeriodStatus, CLMTotalApproved, INVPreviousAmt " & _
          " From VIEW_InvoiceAccount where ClaimStatus = 'C' "
          
          
    strCount = "select count(*) as totalRecord From VIEW_InvoiceAccount where 0=0 "
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
        strCount = strCount + strAppend
    End If

    Set rstCount = medixmasterconnWS.Execute(strCount)
 
    
    rst2.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
       
    If rstCount!TotalRecord = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
         
    Do
        Do Until rst2.EOF
        Total = rstCount!TotalRecord
        lblTotal = Total
        
            
            With rst2
                                            
                                
                If !AccPeriodStatus = True Then
                    Set Record = lstInvoiceAccList.ListItems.Add(, , "Updated")
                Else
                    Set Record = lstInvoiceAccList.ListItems.Add(, , "Not")
                End If
                
                If Len(rst2!INVPayee) Then
                    Record.SubItems(1) = rst2!INVPayee
                End If
                                                                       
                strsqlHOS = "select HosCode, HosName from HOS where HosCode = '" & Trim(Record.SubItems(1)) & "'"
                HOS.Open strsqlHOS, medixmasterconn, adOpenKeyset, adLockOptimistic
                
                If HOS.recordCount Then
                    Record.SubItems(1) = Trim(HOS!HosName)
                Else
                    Record.SubItems(1) = "MEMBER"
                End If
                
                If Len(rst2!CASNumber) Then
                    Record.SubItems(2) = Trim(rst2!CASNumber)
                End If
                
                If Len(rst2!INVWSVersion) Then
                    Record.SubItems(3) = Trim(rst2!INVWSVersion)
                End If
                
                If Len(rst2!CASNumber & "-" & rst2!INVWSVersion) Then
                    Record.SubItems(4) = rst2!CASNumber & "-" & rst2!INVWSVersion
                End If
                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(5) = Trim(rst2!MBMNumber)
                End If
                
                If Len(rst2!MBMName) Then
                    Record.SubItems(6) = Trim(rst2!MBMName)
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(7) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!INSCode) Then
                   Record.SubItems(8) = Trim(rst2!INSCode)
                End If
                
                If Len(rst2!PLNCode) Then
                    Record.SubItems(9) = Trim(rst2!PLNCode)
                End If
                
                If Len(rst2!INVNo) Then
                    Record.SubItems(10) = Trim(rst2!INVNo)
                End If
                
                If Len(rst2!INVAmount) Then
                    Record.SubItems(11) = Format(rst2!INVAmount, "#,##0.00;(#,##0.00)")
                End If
                
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(12) = Format(rst2!CLMTotalApproved, "#,##0.00;(#,##0.00)")
                End If
                
                If Len(rst2!INVPreviousAmt) Then
                    Record.SubItems(13) = Format(rst2!INVPreviousAmt, "#,##0.00;(#,##0.00)")
                End If
                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!INVAmount) Then
                    TotalAmt = TotalAmt + rst2!INVAmount
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   check = False
                   Exit Do
                End If
            HOS.Close
            Set HOS = Nothing
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


If CboPayor.Text = "" Then
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
End If
        
End Sub

Private Sub CboPayTo_Change()

    If InStr(CboPayTo.Text, "null") Then
       CboPayTo.Enabled = False
    End If
End Sub

Private Sub CboPayTo_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayTo.Text, CboPayTo.SelStart) + Chr(KeyAscii) + Mid(CboPayTo.Text, CboPayTo.SelStart + CboPayTo.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayTo.ListCount - 1
 
        If NewText = LCase(Left(CboPayTo.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayTo.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPayTo.Text = ValidValue
                CboPayTo.SelStart = 0
                CboPayTo.SelLength = Len(ValidValue)
            End If
        End If
End Sub


Private Sub CboPeriodMth_Change()
If CboPeriodMth.Text = "" Then
    If InStr(CboPeriodMth.Text, "null") Then
       CboPeriodMth.Enabled = False
    End If
End If
End Sub


Private Sub CboPeriodMth_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer


If CboPeriodMth.Text = "" Then
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPeriodMth.Text, CboPeriodMth.SelStart) + Chr(KeyAscii) + Mid(CboPeriodMth.Text, CboPeriodMth.SelStart + CboPeriodMth.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPeriodMth.ListCount - 1
 
        If NewText = LCase(Left(CboPeriodMth.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPeriodMth.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPeriodMth.Text = ValidValue
                CboPeriodMth.SelStart = 0
                CboPeriodMth.SelLength = Len(ValidValue)
            End If
        End If
End If
End Sub

Private Sub CboPeriodYr_Change()
If CboPeriodYr.Text = "" Then
    If InStr(CboPeriodYr.Text, "null") Then
       CboPeriodYr.Enabled = False
    End If
End If
End Sub


Private Sub CboPeriodYr_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer


If CboPeriodYr.Text = "" Then
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPeriodYr.Text, CboPeriodYr.SelStart) + Chr(KeyAscii) + Mid(CboPeriodYr.Text, CboPeriodYr.SelStart + CboPeriodYr.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPeriodYr.ListCount - 1
 
        If NewText = LCase(Left(CboPeriodYr.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPeriodYr.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPeriodYr.Text = ValidValue
                CboPeriodYr.SelStart = 0
                CboPeriodYr.SelLength = Len(ValidValue)
            End If
        End If
End If
End Sub
'************************End ComboBox Change/KeyPress**************************


Private Sub txtDOADate1_LostFocus()
    If IsDate(txtDOADate1) Then
    Else
        If txtDOADate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOADate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDOADate2_LostFocus()
    If IsDate(txtDOADate2) Then
    Else
        If txtDOADate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOADate2.SetFocus
        End If
    End If
End Sub

Private Sub txtDODDate1_LostFocus()
    If IsDate(txtDODDate1) Then
    Else
        If txtDODDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDODDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDODDate2_LostFocus()
    If IsDate(txtDODDate2) Then
    Else
        If txtDODDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDODDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtDORDate1_LostFocus()
    If IsDate(txtDORDate1) Then
    Else
        If txtDORDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDORDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDORDate2_LostFocus()
    If IsDate(txtDORDate2) Then
    Else
        If txtDORDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDORDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtBatchDate1_LostFocus()
    If IsDate(txtBatchDate1) Then
    Else
        If txtBatchDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBatchDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtBatchDate2_LostFocus()
    If IsDate(txtBatchDate2) Then
    Else
        If txtBatchDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBatchDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtBatchNo1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtBatchNo2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub cmdUpdate_Click()
Dim ii As Integer
Dim cnt As Integer
Dim recordFound As Boolean
Dim delimiter1, delimiter2, delimiter3 As Integer
    
      '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        
        If lstInvoiceAccList.ListItems.Count = 0 Then
            MsgBox "Please select a record(s) before proceeding!", vbOKOnly + vbCritical
        ElseIf CboPeriodMth = "" Then
            MsgBox "Month cannot be Empty.", vbCritical
        ElseIf CboPeriodYr = "" Then
            MsgBox "Year cannot be empty!", vbCritical
        ElseIf UBound(CaseNoArray) = 0 Then
            MsgBox "Please select at least one record to be updated! ", vbCritical
        Else
    
        Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If Update = vbYes Then
                cnt = 0
           
                recordFound = False
                If lstInvoiceAccList.ListItems.Count <> 0 Then
                    For ii = 0 To UBound(CaseNoArray)
                        If Trim(CaseNoArray(ii)) <> "" Then
                                                                                
                            delimiter1 = InStr(1, CaseNoArray(ii), "~")
                            CASNo = Trim(Mid(CaseNoArray(ii), 1, IIf(delimiter1 = 0, 10, delimiter1 - 1)))
                            delimiter2 = InStr(1, CaseNoArray(ii), "-")
                            CASVer = Trim(Mid(CaseNoArray(ii), delimiter1 + 1, delimiter2 - delimiter1 - 1))
                            'delimiter3 = InStr(delimiter2 + 1, CaseNoArray(ii), "~")
                            'PAYTo = Trim(Mid(CaseNoArray(ii), delimiter2 + 1, delimiter3 - delimiter2 - 1))
                            INVNo = Trim(Mid(CaseNoArray(ii), delimiter2 + 1)) ', delimiter3 - 1))
        
                            Call UpdateRecord
                            recordFound = True
                        
                        End If
                    Next ii
                    Erase CaseNoArray
                    Arraycnt = 0
                    ReDim CaseNoArray(0)
                    MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    Call ClearForm(Me)
                    Call InvoiceAccountListing(searchCriteria)
                End If
            End If
        End If
      '  medixmasterconnWS.Close
      '  Set medixmasterconnWS = Nothing
End Sub

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
        !AccPeriodStatus = True
        !InvPrintStatus = True
        !AccPeriodMonth = Trim(CboPeriodMth)
        !AccPeriodYear = Trim(CboPeriodYr)
    End With
    AccPeriod.Update
    AccPeriod.Close
    Set AccPeriod = Nothing
End Sub
'**********************End Update Record*******************************

'****************** Start lstInvoiceAccList ItemCheck *****************************

Private Sub lstInvoiceAccList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i, newbound  As Integer
Dim tempArray() As String
Dim Found As Boolean
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Static Arraycnt As Integer
Dim temp

       

    If Item.Checked = True Then
        
                Arraycnt = IIf(UBound(CaseNoArray) = 0, 1, UBound(CaseNoArray) + 1)
                ReDim Preserve CaseNoArray(Arraycnt)
                CaseNoArray(Arraycnt) = Item.SubItems(2) & "~" & Item.SubItems(3) & "-" & Item.SubItems(10)
               
            'End If
    Else
    
        ' Remove from the array
        newbound = 0
    
        If UBound(CaseNoArray) <> 0 Then
            For i = 0 To UBound(CaseNoArray)
                If (Item.SubItems(2) & "~" & Item.SubItems(3) & "-" & Item.SubItems(10)) <> CaseNoArray(i) Then
                'If (Item.SubItems(1) & "-" & Item.SubItems(2) & "+" & Item.SubItems(27)) <> CaseNoArray(i) Then
                    ReDim Preserve tempArray(newbound)
                    tempArray(newbound) = CaseNoArray(i)
                    newbound = newbound + 1
                End If
            Next
            Erase CaseNoArray
            If UBound(tempArray) = 0 Then
                Arraycnt = 1
            Else
                Arraycnt = UBound(tempArray) + 1
            End If
            ReDim CaseNoArray(UBound(tempArray))
            For i = 0 To UBound(tempArray)
                CaseNoArray(i) = tempArray(i)
            Next
            Erase tempArray
        Else
            Erase CaseNoArray
            ReDim CaseNoArray(0)
            Arraycnt = 0
        End If
    End If
    
End Sub
'****************** End lstInvoiceAccList ItemCheck *****************************



