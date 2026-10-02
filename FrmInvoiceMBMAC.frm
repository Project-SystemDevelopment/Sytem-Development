VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInvoiceMBMAC 
   Caption         =   "Member Invoice Accounting"
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
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   8040
      Visible         =   0   'False
      Width           =   5655
      Begin VB.ComboBox CboPeriodMth 
         Height          =   315
         ItemData        =   "FrmInvoiceMBMAC.frx":0000
         Left            =   1320
         List            =   "FrmInvoiceMBMAC.frx":0002
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox CboPeriodYr 
         Height          =   315
         ItemData        =   "FrmInvoiceMBMAC.frx":0004
         Left            =   4080
         List            =   "FrmInvoiceMBMAC.frx":0006
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "A/C Period Mth"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "A/C Period Yr"
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   120
      TabIndex        =   27
      Top             =   240
      Width           =   11655
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   3000
      End
      Begin VB.ComboBox cboPrintStatus 
         Height          =   315
         ItemData        =   "FrmInvoiceMBMAC.frx":0008
         Left            =   1200
         List            =   "FrmInvoiceMBMAC.frx":000A
         TabIndex        =   1
         Text            =   "N - New"
         Top             =   480
         Width           =   3015
      End
      Begin MSMask.MaskEdBox txtBordxDateFr 
         Height          =   300
         Left            =   5400
         TabIndex        =   2
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBordxDateTo 
         Height          =   300
         Left            =   6840
         TabIndex        =   3
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtINVDateFr 
         Height          =   300
         Left            =   5400
         TabIndex        =   4
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtBatchDateFr 
         Height          =   285
         Left            =   9000
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtBatchNo1 
         Height          =   285
         Left            =   9000
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtINVDateTo 
         Height          =   300
         Left            =   6840
         TabIndex        =   5
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtBatchDateTo 
         Height          =   285
         Left            =   10440
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtBatchNo2 
         Height          =   285
         Left            =   10440
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "To"
         Height          =   255
         Left            =   6600
         TabIndex        =   44
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Bordx Date Fr"
         Height          =   255
         Left            =   4320
         TabIndex        =   43
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   10150
         TabIndex        =   39
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   10150
         TabIndex        =   38
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "Batch Seq Fr"
         Height          =   255
         Left            =   8040
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label40 
         Caption         =   "Batch No"
         Height          =   255
         Left            =   8040
         TabIndex        =   36
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Claim Date Fr"
         Height          =   255
         Left            =   4320
         TabIndex        =   35
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label41 
         Caption         =   "To"
         Height          =   255
         Left            =   6600
         TabIndex        =   34
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label45 
         Caption         =   "Print Status"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   7440
      Width           =   11655
      Begin VB.TextBox lblSQENo 
         Height          =   285
         Left            =   4680
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   2520
         TabIndex        =   46
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   840
         TabIndex        =   45
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview"
         Height          =   330
         Left            =   8520
         TabIndex        =   33
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   330
         Left            =   7920
         TabIndex        =   24
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   330
         Left            =   7320
         TabIndex        =   23
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   330
         Left            =   6000
         TabIndex        =   21
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   330
         Left            =   7920
         TabIndex        =   20
         Top             =   200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9360
         TabIndex        =   19
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10920
         TabIndex        =   18
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10320
         TabIndex        =   17
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   6720
         TabIndex        =   16
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   7440
         TabIndex        =   22
         Top             =   200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Seq No"
         Height          =   255
         Left            =   3960
         TabIndex        =   41
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   1560
         TabIndex        =   25
         Top             =   195
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView lstInvoiceAccList 
      Height          =   6180
      Left            =   120
      TabIndex        =   40
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10901
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
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pay To"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Wsk No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Claim Reg Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Main DCN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Mem No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Mem Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Pay2MBM"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Prev Amt"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "DR"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "CR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Date Seq"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Bat No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "PLN Code"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Member Invoice  Accounting"
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
      TabIndex        =   42
      Top             =   0
      Width           =   11805
   End
   Begin VB.Label txtLocalFileLocation 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   32
      Top             =   8160
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label txtPath 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   31
      Top             =   8160
      Visible         =   0   'False
      Width           =   705
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
      Left            =   7080
      TabIndex        =   30
      Top             =   8280
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "FrmInvoiceMBMAC"
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
Dim InvoiceNo As String
Dim strAppend As String
'Sort Direction for each column in the ListView
Private m_blnDirection(1 To 15) As Boolean

Private Sub cboPrintStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPrintStatus.Text, cboPrintStatus.SelStart) + Chr(KeyAscii) + Mid(txtHealthCode.Text, cboPrintStatus.SelStart + cboPrintStatus.SelLength + 1))
 
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

End Sub

Private Sub cmdPreview_Click()
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
    ReDim CaseNoArray(0)
    Call ComboListing
    txtPath = DocPath & "ACCInvoice"
    txtLocalFileLocation = "C:\"
    X = CreatePath("c:\cardPrinting")
    If X = False Then
        MsgBox "Error Creating Folder c:\CardPrinting !" & vbNewLine & _
                " Please create that folder manually to store the output of leaflet in sofe copy!"
    End If

End Sub

'************************Start Command Button**************************
Private Sub CmdExit_Click()
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
            lstInvoiceAccList.Height = 7000
            lstInvoiceAccList.Left = 120
            lstInvoiceAccList.Top = 360
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
    lstInvoiceAccList.Height = 6000
    lstInvoiceAccList.Left = 120
    lstInvoiceAccList.Top = 1450
    lstInvoiceAccList.Width = 11655
    Frame1.Visible = True
End Sub

Private Sub CmdPrint_Click()
Dim ii As Integer
Dim cnt As Integer
Dim recordFound As Boolean
Dim delimiter1, delimiter2, delimiter3 As Integer
Dim SystemDate As Date
        If CboPayor = "" Then
            MsgBox "Please select Payor", vbOKOnly + vbCritical
            CboPayor.SetFocus
        ElseIf cboPrintStatus = "" Then
            MsgBox "Please select Print Status", vbOKOnly + vbCritical
            cboPrintStatus.SetFocus
        ElseIf lstInvoiceAccList.ListItems.Count = 0 Then
            MsgBox "Please select a record(s) before proceeding!", vbOKOnly + vbCritical
        Else
    
        Update = MsgBox("Do you want to print the record?", vbYesNo + vbExclamation)
            If Update = vbYes Then
                cnt = 0
                
                recordFound = False
                'If lstInvoiceAccList.listitems.count <> 0 Then
                    If Mid(cboPrintStatus, 1, 1) = "N" Then
                        InvoiceNo = GenerateINVMSeqNo(Date)
                    Else
                        InvoiceNo = Mid(CboPayor, 1, 2) & Format(Date, "ddmmyyyy")
                    End If
                        Open txtPath & "\" & InvoiceNo For Output As #1
                        Open txtLocalFileLocation & InvoiceNo For Output As #2
                            'header = Trim(InvoiceNo)
                        'Print #1, Trim(header)
                        'Print #2, Trim(header)
                                                                            
                    'For ii = 0 To UBound(CaseNoArray)
                    '    If Trim(CaseNoArray(ii)) <> "" Then
                    '        delimiter1 = InStr(1, CaseNoArray(ii), "~")
                    '        CASNo = Trim(Mid(CaseNoArray(ii), 1, IIf(delimiter1 = 0, 10, delimiter1 - 1)))
                    '        delimiter2 = InStr(1, CaseNoArray(ii), "-")
                    '        CASVer = Trim(Mid(CaseNoArray(ii), delimiter1 + 1, delimiter2 - delimiter1 - 1))
                    '        INVNo = Trim(Mid(CaseNoArray(ii), delimiter2 + 1))
                            Call MBMAccASCII
                            'Call Print2File
                            If Mid(cboPrintStatus, 1, 1) = "N" Then
                                Call UpdatePrintRecord
                            End If
                            recordFound = True
                        
                        End If
                    'Next ii
                    'Erase CaseNoArray
                    'Arraycnt = 0
                    'ReDim CaseNoArray(0)
                    Close #1
                    Close #2
                    MsgBox "Record Printed...", vbOKOnly + vbInformation, DialogBoxTitle
                    Call ClearForm(Me)
                    Call InvoiceAccountListing(searchCriteria)
                End If
            'End If
        'End If
    
    
    'Screen.MousePointer = vbHourglass
        
    'If lstInvoiceAccList.listitems.count <> 0 Then
        'RptInvoiceTrack.show
    '    INVNo = GenerateINVMSeqNo(Mid(CboPayor, 1, 2), Date)
        
        
    'Else
    '    MsgBox "Report cannot be printed without any record.", vbInformation
    'End If
    'Screen.MousePointer = Default
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
            InvoiceNo = GenerateINVMSeqNo(Date)
            lblSQENo = InvoiceNo
        End If
        If Mid(cboPrintStatus, 1, 1) = "N" Then
            Call UpdatePrintRecord
        End If
        medixmasterconnWS.Close
        Set medixmasterconnWS = Nothing
        
        Call ExportListViewtoExcel(lstInvoiceAccList)
        Screen.MousePointer = Default
    End If
    End If
End Sub

Private Sub CmdView_Click()
    lstInvoiceAccList.Height = 7000
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
Dim sql As String
Dim YearStart
Dim MonthStart
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
    End With
    
End Sub
'************************Start ComboListing**************************

'************************Start Search Record**************************

Private Function SearchRecord()
Dim strsql As String
Dim sql As String
    
    strAppend = ""
    Hospital = ""
    
    If Len(Trim(txtINVDateFr)) And IsDate(txtINVDateFr) Then
        strAppend = strAppend + " And ClaimCloseDate >= '" & Format(txtINVDateFr, "mm/dd/YYYY") & "'"
    End If
    
    If Len(Trim(txtINVDateTo)) And IsDate(txtINVDateTo) Then
        strAppend = strAppend + " And ClaimCloseDate <= '" & Format(txtINVDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(Mid(CboPayor, 1, 2))) Then
        strAppend = strAppend + " And Substring(CASNumber,1 ,2)  = '" & Mid(CboPayor, 1, 2) & "'"
        'Hospital = Replace(cboHospital, "-", " ")
        'strAppend = strAppend + " And INVPayee = '" & Trim(Mid(Hospital, 1, 5)) & "'"
    End If
    
    If Mid(cboPrintStatus, 1, 1) = "N" Then
       sql = ""
       sql = " AND INVPrintStatus = '0'"
       strAppend = strAppend + sql
    End If
    
    If Mid(cboPrintStatus, 1, 1) = "R" Then
       sql = ""
       sql = " AND INVPrintStatus = '1' "
       strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtBatchDateFr)) > 0 And IsDate(txtBatchDateFr) = True Then
        strAppend = strAppend + " And INVBatchDate >= '" & txtBatchDateFr & "'"
    End If
    
    If Len(Trim(txtBatchDateTo)) > 0 And IsDate(txtBatchDateTo) = True Then
        strAppend = strAppend + " And INVBatchDate <= '" & txtBatchDateTo & "'"
    End If
    
    If Len(Trim(txtBatchNo1)) Then
        strAppend = strAppend + " AND INVBatchNo >= '" & Trim(txtBatchNo1) & "'"
    End If
    
    If Len(Trim(txtBatchNo2)) Then
        strAppend = strAppend + " AND INVBatchNo <= '" & Trim(txtBatchNo2) & "'"
    End If
        
    If Len(Trim(txtBordxDateFr)) And IsDate(txtBordxDateFr) Then
        strAppend = strAppend + " And BordxDate >= '" & Format(txtBordxDateFr, "mm/dd/YYYY") & "'"
    End If
    
    If Len(Trim(txtBordxDateTo)) And IsDate(txtBordxDateTo) Then
        strAppend = strAppend + " And BordxDate <= '" & Format(txtBordxDateTo, "mm/dd/yyyy") & "'"
    End If
        
    searchCriteria = strAppend
    Call InvoiceAccountListing(strAppend)

End Function
'************************End Search Record**************************



'************************Start List Record**************************

Private Sub InvoiceAccountListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
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
    DRCRAmt = 0
    LblTotalAmt = ""
    lblCurrent = ""
    lstInvoiceAccList.ListItems.Clear
       
    sql = " Select CASNumber, INVWSVersion, MBMNumber, " & _
          " MBMName, PatientName, PLNCode," & _
          " INVPayee, PatientName, DebitCreditRefNo, " & _
          " CLMTotalActual, INVPreviousAmt, BordxRefNo, CLMPayableByMedix2M, " & _
          " BordxDate, INVBatchDate, INVBatchNo, ClaimCloseDate " & _
          " From VIEW_MBMInvoice1 Where  0=0  "
    
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    If rst2.EOF = True Then
    'If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdPrintFile.Enabled = True
        'cmdPrint.Enabled = True
        'CmdUpdate.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
        'Total = rst2.recordCount
        'lblTotal = Total
        
            With rst2
                                
                If Len(!INVPayee) Then
                    Set Record = lstInvoiceAccList.ListItems.Add(, , !INVPayee)
                Else
                    Set Record = lstInvoiceAccList.ListItems.Add(, , "Empty")
                End If
                
                If Len(rst2!CASNumber & "-" & rst2!INVWSVersion) Then
                    Record.SubItems(1) = rst2!CASNumber & "-" & rst2!INVWSVersion
                End If
                
                If Len(rst2!ClaimCloseDate) Then
                    Record.SubItems(2) = rst2!ClaimCloseDate
                End If
                
                
                If Len(rst2!DebitCreditRefNo) Then
                    Record.SubItems(3) = Trim(rst2!DebitCreditRefNo)
                End If
                
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(4) = rst2!BordxRefNo
                End If
                
                'If Len(rst2!BordxDate) Then
                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(5) = Trim(rst2!MBMNumber)
                End If
                
                If Len(rst2!MBMName) Then
                    Record.SubItems(6) = Trim(rst2!MBMName)
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(7) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(8) = Format(rst2!CLMPayableByMedix2M, "#,##0.00")
                End If
                
                If Len(rst2!InvPreviousAmt) Then
                    Record.SubItems(9) = Format(rst2!InvPreviousAmt, "#,##0.00;(#,##0.00)")
                End If
                
                'If Len(rst2!CLMTotalActual) Then
                '    Record.SubItems(9) = Format(rst2!CLMTotalActual, "#,##0.00;(#,##0.00)")
                'End If
               '
               ' If Len(rst2!InvPreviousAmt) Then
               '     Record.SubItems(10) = Format(rst2!InvPreviousAmt, "#,##0.00;(#,##0.00)")
               ' End If
                                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!CLMPayableByMedix2M) Then
                    TotalAmt = TotalAmt + rst2!CLMPayableByMedix2M
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
                
                If rst2!InvPreviousAmt >= 0 And rst2!CLMPayableByMedix2M >= 0 Then
                    DRCRAmt = CDbl(rst2!CLMPayableByMedix2M) - CDbl(rst2!InvPreviousAmt)
                ElseIf rst2!InvPreviousAmt < 0 And rst2!CLMPayableByMedix2M >= 0 Then
                    DRCRAmt = CDbl(rst2!CLMPayableByMedix2M) + CDbl(rst2!InvPreviousAmt)
                ElseIf rst2!InvPreviousAmt < 0 And rst2!CLMPayableByMedix2M < 0 Then
                    DRCRAmt = CDbl(rst2!CLMPayableByMedix2M) - CDbl(rst2!InvPreviousAmt)
                Else
                    DRCRAmt = rst2!CLMPayableByMedix2M
                End If
                
                If DRCRAmt < 0 Then
                    Record.SubItems(10) = Format(DRCRAmt, "#,##0.00;#,##0.00")
                ElseIf DRCRAmt > 0 Then
                    Record.SubItems(11) = Format(DRCRAmt, "(#,##0.00)")
                ElseIf DRCRAmt = 0 Then
                    Record.SubItems(11) = Format(DRCRAmt, "#,##0.00;#,##0.00")
                End If
                
                If Len(rst2!INVBatchDate) Then
                    Record.SubItems(12) = rst2!INVBatchDate
                End If
                                
                If Len(rst2!INVBatchNo) Then
                    Record.SubItems(13) = rst2!INVBatchNo
                End If
                
                If Len(rst2!PLNCode) Then
                    Record.SubItems(14) = rst2!PLNCode
                End If
                
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   check = False
                   Exit Do
                End If
            
        Loop
    Loop Until check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdPrintFile.Enabled = True
    'cmdPrint.Enabled = True
    'CmdUpdate.Enabled = True
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


Private Sub cmdUpdate_Click()
Dim ii As Integer
Dim cnt As Integer
Dim recordFound As Boolean
Dim delimiter1, delimiter2, delimiter3 As Integer
    
        
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
End Sub

'**********************Start Update Record*******************************
Private Sub UpdatePrintRecord()
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim Update
Dim strsql As String
Dim AccPeriod As New ADODB.Recordset
             
    'strsql = " Select CASNumber, INVWSVersion, " & _
    '" InvBatchDate, InvBatchNo,InvPrintLogon,LastAccPeriodDate,INVPrintStatus " & _
    '" From ClaimInvoice where CASNumber = '" & Trim(CASNo) & "' " & _
    '" AND INVWSVersion = '" & Trim(CASVer) & "' AND INVNo = '" & Trim(INVNo) & "'"
    'AccPeriod.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    strsqlCLM = " Select CASNumber, INVWSVersion, CLMPayableByMedix2M, ClaimCloseDate " & _
    " From VIEW_MBMInvoice1 where INVPayTo = 'M'"
    
    strsqlCLM = strsqlCLM + strAppend
    
    CLM.Open strsqlCLM, medixmasterconnWS, adOpenKeyset, adLockOptimistic
            
    Do Until CLM.EOF
    
        strsql = " Select CASNumber, INVWSVersion, InvPrintStatus, InvPreviousAmt, " & _
        " InvBatchDate, InvBatchNo, InvPrintLogon, LastAccPeriodDate, INVPayTo " & _
        " From ClaimInvoice where CASNumber = '" & CLM!CASNumber & "' And INVWSVersion = '" & CLM!INVWSVersion & "' And INVPayTo = 'M'"
        
        AccPeriod.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
        Do Until AccPeriod.EOF
            With AccPeriod
                !INVBatchDate = Mid(InvoiceNo, 1, 9)
                !INVBatchNo = Trim(Mid(InvoiceNo, 11, 3))
                !InvPrintLogon = Logon
                !InvPrintStatus = "1"
                !InvPreviousAmt = CLM!CLMPayableByMedix2M
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

'Private Sub Print2File()
'Dim MBM As New ADODB.Recordset
'Dim strsql As String
'Dim TotalPrincipal As Integer
'Dim TotalSupplementary As Integer
'Dim RecordSaved
'Dim strsplOthers As String
'Dim MBMOthers As New ADODB.Recordset
'Dim strsqlMBMCover As String
'Dim MBMCoveredPersons As New ADODB.Recordset
'Dim strsqlMBMPrint As String
'Dim ASCII As New ADODB.Recordset

'            Open txtPath & "\" & txtFilName For Output As #1
'            Open txtLocalFileLocation & txtFilName For Output As #2
'            header = Trim("F" & Mid(cboPAYCode, 1, 2) & Mid(txtBordxFrom, 1, 2) & Mid(txtBordxFrom, 4, 2) & Mid(txtBordxFrom, 7, 4) & l("", 174))
'                Print #1, Trim(header)
'                Print #2, Trim(header)
'                'Debug.Print Header
'
'                Call MBMAccASCII
'                'Print #1, Trailer
'                'Print #2, Trailer
'                cmdGenerate.Enabled = False
'                Close #1
'                Close #2

'                MsgBox "File Generated Sucessfully...", vbOKOnly + vbInformation, DialogBoxTitle
'            End If
'            Screen.MousePointer = vbDefault
'        End If
'        End If
'End Sub

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
        
        'strsql = " Select CASNumber, INVWSVersion, " & _
        '" InvBatchDate, InvBatchNo,INVNo " & _
        '" From ClaimInvoice where CASNumber = '" & Trim(CASNo) & "' " & _
        '" AND INVWSVersion = '" & Trim(CASVer) & "' AND INVNo = '" & Trim(INVNo) & "'"
        'AccPeriod.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
        'strsql = " Select CASNumber, INVWSVersion, " & _
        '" InvBatchDate, InvBatchNo,INVNo " & _
        '" From ClaimInvoice where Substring(CASNumber,1,2) = '" & Mid(CboPayor, 1, 2) & "'"
        strsql = " Select CASNumber, INVWSVersion, MBMNumber, INSCode, " & _
            " INVNo, CASDischargeDate, CaseRegDate, PLNCode, " & _
            " CASDateAdmitted, INVPayTo, INVPayee, PatientName, AccPeriodMonth, AccPeriodYear " & _
            " MBMName, InvPrintStatus, INVPreviousAmt " & _
            " From VIEW_InvoiceAccount where ClaimStatus = 'C' And Substring(CASNumber,1,2) = '" & Mid(CboPayor, 1, 2) & "'"
               
        AccPeriod.Open strsql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic

        
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

