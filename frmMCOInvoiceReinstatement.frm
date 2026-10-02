VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMCOInvoiceReinstatement 
   Caption         =   "MCO Reinstatement Invoice "
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "MCO Reinstatement Invoice "
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   11160
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameMCO 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   60
      TabIndex        =   16
      Top             =   240
      Width           =   11055
      Begin VB.TextBox txtMCOBatchNo 
         Height          =   315
         Left            =   9600
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cboMCOCriteriaSel 
         Height          =   315
         ItemData        =   "frmMCOInvoiceReinstatement.frx":0000
         Left            =   1320
         List            =   "frmMCOInvoiceReinstatement.frx":000A
         TabIndex        =   0
         Top             =   240
         Width           =   4935
      End
      Begin VB.ComboBox cboMCOPayor 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   525
         Width           =   4935
      End
      Begin VB.TextBox TxtInvoiceNo 
         Height          =   315
         Left            =   7440
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox CboRptType 
         Height          =   315
         ItemData        =   "frmMCOInvoiceReinstatement.frx":0030
         Left            =   1320
         List            =   "frmMCOInvoiceReinstatement.frx":003A
         TabIndex        =   5
         Top             =   795
         Width           =   4935
      End
      Begin MSMask.MaskEdBox txtMCOBordxDate 
         Height          =   315
         Left            =   7440
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Crystal.CrystalReport cRptMCO 
         Left            =   7320
         Top             =   720
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         UserName        =   "sa"
         PrintFileLinesPerPage=   60
      End
      Begin MSMask.MaskEdBox txtMCOBordxDateTo 
         Height          =   315
         Left            =   9600
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMCOBordxDateFr 
         Height          =   315
         Left            =   7440
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog dlgCommonDialog_MCO 
         Left            =   7920
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblDisplay1 
         Height          =   255
         Left            =   6360
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Report Type"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   820
         Width           =   1815
      End
      Begin VB.Label lblDisplay2 
         Height          =   255
         Left            =   8760
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl20 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   520
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Single/Batch"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   60
      TabIndex        =   9
      Top             =   7200
      Width           =   11055
      Begin VB.CommandButton cmdMCOClose 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   300
         Left            =   10080
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOClear 
         Caption         =   "&Clear"
         Height          =   300
         Left            =   9240
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOPrintPreview 
         Caption         =   "Pre&view"
         Height          =   300
         Left            =   8400
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOPrint 
         Caption         =   "&Print"
         Height          =   300
         Left            =   7560
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   300
         Left            =   6720
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label200 
         Alignment       =   2  'Center
         Caption         =   "Total Principal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label lblTotPrinc 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lblTotSupp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3360
         TabIndex        =   13
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label201 
         Alignment       =   2  'Center
         Caption         =   "Total Supp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   2400
         TabIndex        =   12
         Top             =   240
         Width           =   960
      End
   End
   Begin MSComctlLib.ListView lvwMCO 
      Height          =   5610
      Left            =   60
      TabIndex        =   24
      Top             =   1560
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Member No"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cover ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Member Name"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Rel"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "BordxDate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Batch No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Pol No"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Eff Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Exp Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Adult/Child"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Invoice No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Invoice Date"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "MCOFee"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Reinstatement MCO Fees"
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
      TabIndex        =   23
      Top             =   0
      Width           =   11205
   End
End
Attribute VB_Name = "frmMCOInvoiceReinstatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MultiInvFound As Boolean
Dim tmpPayor As String
Dim ContactName As String
Dim PAYName As String
Dim PayAdd1 As String
Dim PayAdd2 As String
Dim PayAdd3 As String
Dim TelNo As String
Dim FaxNo As String

Private Sub cboMCOCriteriaSel_Click()
    Call cmdMCOClear_Click
    txtMCOBordxDate.Visible = False
    txtMCOBatchNo.Visible = False
    txtMCOBordxDateFr.Visible = False
    txtMCOBordxDateTo.Visible = False
    TxtInvoiceNo.Visible = False
    lblDisplay2.Visible = True
    If cboMCOCriteriaSel.ListIndex = 0 Then
        lblDisplay1.Caption = "Bordx Date"
        lblDisplay2.Caption = "Batch No"
        txtMCOBordxDate.Visible = True
        txtMCOBatchNo.Visible = True
    ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
        lblDisplay1.Caption = "Invoice No"
        lblDisplay2.Visible = False
        TxtInvoiceNo.Visible = True
    End If
End Sub

Private Sub cmdMCOClear_Click()
    TxtInvoiceNo = ""
    txtMCOBatchNo = ""
    cboMCOPayor = ""
    txtMCOBordxDate = "__/__/____"
    txtMCOBordxDateFr = "__/__/____"
    txtMCOBordxDateTo = "__/__/____"
    lvwMCO.listitems.Clear
    'cmdMCOPrint.Enabled = False
    cmdMCOPrintPreview.Enabled = False
    lblTotSupp = 0
    lblTotPrinc = 0
End Sub

Private Sub cmdMCOClose_Click()
    Unload Me
End Sub

Private Sub cmdMCOPrint_Click()
    If Trim(CboRptType) = "" Then
        MsgBox "Please Select Report Type !", vbOKOnly + vbCritical
        cmdMCOPrint.SetFocus
    Else
     '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    '    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        medixmasterconnPluto.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call PrintPreviewCmd(cmdMCOPrintPreview)
     '   medixmasterconn.Close
     '   Set medixmasterconn = Nothing
     '   medixmasterconnMaint.Close
     '   Set medixmasterconnMaint = Nothing
        medixmasterconnPluto.Close
        Set medixmasterconnPluto = Nothing
    End If
End Sub

Private Sub cmdMCOPrintPreview_Click()
    If Trim(CboRptType) = "" Then
        MsgBox "Please Select Report Type !", vbOKOnly + vbCritical
        cmdMCOPrint.SetFocus
    Else
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
      '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        medixmasterconnPluto.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call PrintPreviewCmd(cmdMCOPrintPreview)
     '   medixmasterconn.Close
      '  Set medixmasterconn = Nothing
     '   medixmasterconnMaint.Close
     '   Set medixmasterconnMaint = Nothing
        medixmasterconnPluto.Close
        Set medixmasterconnPluto = Nothing
    End If
End Sub

Private Sub cmdMCOSearch_Click()
    lblTotPrinc = "0"
    lblTotSupp = "0"
    cmdMCOSearch.Enabled = False
    
    If cboMCOCriteriaSel = "" Then
        MsgBox "Please Enter MCO Criteria!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdMCOSearch.Enabled = True
        Exit Sub
    End If
            
    If cboMCOCriteriaSel.ListIndex = 0 Then
        If cboMCOPayor = "" Then
            MsgBox "Please Enter Payor!", vbOKOnly + vbCritical, DialogBoxTitle
            cmdMCOSearch.Enabled = True
            Exit Sub
        ElseIf txtMCOBordxDate = "__/__/____" Then
            MsgBox "Please Enter Bordereaux Date!", vbOKOnly + vbCritical, DialogBoxTitle
            cmdMCOSearch.Enabled = True
            Exit Sub
        ElseIf Trim(txtMCOBatchNo) = "" Then
            MsgBox "Please Enter Batch No!", vbOKOnly + vbCritical, DialogBoxTitle
            cmdMCOSearch.Enabled = True
            Exit Sub
        End If
    ElseIf cboMCOCriteriaSel.ListIndex = 1 And TxtInvoiceNo = "" Then
            MsgBox "Please Enter Invoice No!", vbOKOnly + vbCritical, DialogBoxTitle
            cmdMCOSearch.Enabled = True
            Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    Call MCOListing
    Screen.MousePointer = Default
    cmdMCOSearch.Enabled = True

End Sub


Private Sub MCOListing()
Dim rsSearch As New ADODB.Recordset
Dim sql As String
Dim TotPrinc As Integer
Dim TotalSupp As Integer
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim ListMaster As ListItem
Dim AA As Integer
Dim PrevInvNO As String
Dim CurrInvNo As String
Dim SelField As String
Dim MCONull As Boolean
    
    lvwMCO.listitems.Clear
    'cmdMCOPrint.Enabled = False
    cmdMCOPrintPreview.Enabled = False
    sql = ""
    'sql = "SELECT * FROM VIEW_MCO_LISTING_ALL2"
    SelField = "Select MBMNumber, MBMCOVID, MBMPayorEffDate, MBMPayorExpDate, " & _
                " MBMPolicyNo, MBMName, RELCode, MBMIcBcPp, MBMGroupCompany, " & _
                " HLTCode, MBMDepartment, INSCode, MBMStatus, MBMMemberType, " & _
                " MBMCNumber, MBMCCoverID, MBMCName, MBMCICBCPPNo, MBMCRELCode, " & _
                " PLNCode, MBMEndtBordxDate, MBMLapsedDate, MBMReinstateDate, " & _
                " MBMInvoiceNo, MBMInvoiceDate, MBMInvoiceBy, MBMMCOFee, MBMPremium, " & _
                " MBMEndtBordxBatch, PAYCode, Type "

    Call SelectionCriterial(sql)

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Set cmd1.ActiveConnection = medixmasterconn
    cmd1.CommandText = "SP_MCOFeesReinstatement"
    cmd1.CommandTimeout = 300
    cmd1.CommandType = adCmdStoredProc
    Set Prm1 = cmd1.CreateParameter("SelFields", adVarChar, adParamInput, 2000, SelField)
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, sql)
    cmd1.Parameters.Append Prm2
    Set Prm3 = cmd1.CreateParameter("SortList", adVarChar, adParamInput, 2000, "")
    cmd1.Parameters.Append Prm3

    rsSearch.CacheSize = 100
    rsSearch.CursorLocation = adUseClient
    rsSearch.CursorType = adOpenForwardOnly
    Set rsSearch = cmd1.Execute
    Do Until rsSearch.EOF

    Set ListMaster = lvwMCO.listitems.Add(, , rsSearch!MBMNumber)
        If rsSearch!MBMCRELCode <> "" Then
            ListMaster.SubItems(1) = rsSearch!MBMCCoverID
            ListMaster.SubItems(2) = Trim(rsSearch!MBMCName)
            ListMaster.SubItems(3) = rsSearch!MBMCRELCode
            ListMaster.SubItems(9) = IIf(Len(rsSearch!Type), Trim(rsSearch!Type), "")
        Else
            ListMaster.SubItems(1) = rsSearch!MBMCOVID
            ListMaster.SubItems(2) = Trim(rsSearch!MBMName)
            ListMaster.SubItems(3) = "P"
            ListMaster.SubItems(9) = "Adult"
        End If
        
        ListMaster.SubItems(4) = rsSearch!MBMEndtBordxDate
        ListMaster.SubItems(5) = IIf(Len(rsSearch!MBMEndtBordxBatch), rsSearch!MBMEndtBordxBatch, "")
        ListMaster.SubItems(6) = IIf(Len(rsSearch!MBMPolicyNo), Trim(rsSearch!MBMPolicyNo), "")
        ListMaster.SubItems(7) = IIf(Len(rsSearch!MBMPayorEffDate), rsSearch!MBMPayorEffDate, "")
        ListMaster.SubItems(8) = IIf(Len(rsSearch!MBMPayorEXPDate), rsSearch!MBMPayorEXPDate, "")
        ListMaster.SubItems(10) = IIf(Len(rsSearch!MBMInvoiceNo), Trim(rsSearch!MBMInvoiceNo), "")
        ListMaster.SubItems(11) = IIf(Len(rsSearch!MBMInvoiceDate), Trim(rsSearch!MBMInvoiceDate), "")
        ListMaster.SubItems(12) = IIf(IsNumeric(rsSearch!MBMMCOFee), Round(rsSearch!MBMMCOFee), 0)
        
        If rsSearch!MBMCRELCode <> "" Then
           TotalSupp = TotalSupp + 1
           lblTotSupp = TotalSupp
        Else
            TotPrinc = TotPrinc + 1
            lblTotPrinc = TotPrinc
        End If
        
        rsSearch.MoveNext
    Loop
    
    rsSearch.Close
    Set rsSearch = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    If lvwMCO.listitems.Count = 0 Then
        MsgBox "Record not found !", vbInformation + vbOKOnly
        'cmdMCOPrint.Enabled = False
        Screen.MousePointer = Default
        Exit Sub
    Else
        cmdMCOPrintPreview.Enabled = True
    End If
End Sub

Private Sub SelectionCriterial(tmlSelCriterial As String)
Dim StrReprint As String
Dim strCriterial As String
    If cboMCOCriteriaSel.ListIndex = 1 Then
        strCriterial = " AND MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "'"
    ElseIf cboMCOCriteriaSel.ListIndex = 0 Then
        strCriterial = " AND (MBMEndtBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "'" & _
                       " AND MBMEndtBordxBatch = " & Trim(txtMCOBatchNo) & ")" & _
                       " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "'"
    End If
    tmlSelCriterial = strCriterial
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
            cboMCOPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            PAY.MoveNext
        End With
    Loop
    
    PAY.Close
    Set PAY = Nothing
End Sub

Private Sub Form_Load()
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call ComboListing
   ' medixmasterconn.Close
    'Set medixmasterconn = Nothing
End Sub

Private Sub PrintPreviewCmd(cboType As Boolean)
Dim RptType As Integer
    If MultiInvFound = True Then
        MsgBox "Different Invoice Found in Same Batch or Range, Please Search by Invoice No ! ", vbCritical + vbOKOnly
        Exit Sub
    Else
        Screen.MousePointer = vbHourglass
        If cboMCOCriteriaSel.ListIndex = 0 Then
            RptType = 1
        ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
            RptType = 3
        End If
        Call CrystalRptName(RptType)
        'cRptMCO.Reset
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub CrystalRptName(tmpCriSel As Integer)
Dim strFormula As String
Dim strsql As String
Dim rst As New ADODB.Recordset
Dim tmpMBMNumber As String
Dim tmpBatchNo As String
Dim tmpPayCode As String
Dim tmpEndtBordxDate As String
Dim tmpPolEffDate As String
Dim tmpPolExpDate As String
Dim tmpInvoiceDate As String
Dim strsqlPlutoMCO As String
Dim PlutoMCO As New ADODB.Recordset
Dim PLUTOStatus As Boolean
Dim sqlinsert As String
Dim tmpMBMName As String
Dim tmpMBMCName As String
Dim tmpMBMBasicMCO As Double
Dim tmpMBMMCO As Double
Dim tmpMBMMCOFee As Double
Dim tmpPremium As Double
Dim tmpAdultChild As String

'    tmpBatchNo = lvwMCO.SelectedItem.SubItems(5)
'    tmpPayCode = Mid(lvwMCO.SelectedItem.Text, 1, 2)
'    If tmpCriSel = 1 Then ' By Batch
'        strFormula = ""
'        strFormula = " AND MBMEndtBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "' " & _
'                       " AND MBMEndtBordxBatch = " & Trim(txtMCOBatchNo) & "" & _
'                       " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "'"
'    ElseIf tmpCriSel = 2 Then ' By Range
'        strFormula = " And MBMEndtBordxDate >= '" & Format(txtMCOBordxDateFr, "yyyy/mm/dd") & "'" & _
'                    " And MBMEndtBordxDate <= '" & Format(txtMCOBordxDateTo, "yyyy/mm/dd") & "'" & _
'                    " And PayCode = '" & tmpPayCode & "'"
'    ElseIf tmpCriSel = 3 Then ' By Invoice
'        strFormula = " And MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "'"
'    End If
'    strsqlPlutoMCO = "Select MBMNumber from MCO_Invoice_Reinstatement where 0 = 0  "
'    strsqlPlutoMCO = strsqlPlutoMCO + strFormula
'    PlutoMCO.Open strsqlPlutoMCO, medixmasterconnPluto, adOpenForwardOnly, adLockReadOnly
'    PLUTOStatus = False
'    Do While Not PlutoMCO.EOF
'        PLUTOStatus = True
'    PlutoMCO.MoveNext
'    Loop
'    PlutoMCO.Close
'    Set PlutoMCO = Nothing
'
'    If PLUTOStatus = False Then
'        strsql = ""
'        strsql = "Select * from VIEW_MBM_MCOReinstatement where 0 = 0"
'        strsql = strsql + strFormula
'        rst.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
'
'        Do While Not rst.EOF
'
'            tmpEndtBordxDate = Format(rst!MBMEndtBordxDate, "yyyy/mm/dd")
'            tmpPolEffDate = Format(rst!MBMPayorEffDate, "yyyy/mm/dd")
'            tmpPolExpDate = Format(rst!MBMPayorEXPDate, "yyyy/mm/dd")
'            tmpInvoiceDate = Format(rst!MBMInvoiceDate, "yyyy/mm/dd")
'            tmpMBMName = Replace(Trim(rst!MBMName), "'", "''")
'            If Len(rst!MBMCName) Then
'                tmpMBMCName = Replace(Trim(rst!MBMCName), "'", "''")
'            Else
'                tmpMBMCName = ""
'            End If
'            tmpMBMMCOFee = IIf(IsNumeric(rst!MBMMCOFee), rst!MBMMCOFee, 0)
'            tmpPremium = IIf(IsNumeric(rst!MBMPremium), rst!MBMPremium, 0)
'            tmpAdultChild = IIf(Len(rst!Type), rst!Type, "ADULT")
'            sqlinsert = ""
'            sqlinsert = "Insert into MCO_Invoice_Reinstatement(MBMNumber, MBMCoverID, MBMPayorEffDate, MBMPayorExpDate, " & _
'                        " MBMPolicyNo, MBMName, RELCode,MBMIcBcPp,MBMGroupCompany, " & _
'                        " HLTCode, MBMDepartment,INSCode,MBMPremium," & _
'                        " Type,PAYCode,MBMMemberType,MBMCNumber,MBMCCoverID, " & _
'                        " MBMCName, MBMCICBCPPNo,MBMCRELCode,MBMEndtBordxDate, MBMEndtBordxBatch, " & _
'                        " MBMMCOFee) " & _
'                        "Values('" & rst!MBMNumber & "', '00' , '" & tmpPolEffDate & "',  '" & tmpPolExpDate & "'," & _
'                        "'" & rst!MBMPolicyNo & "', '" & tmpMBMName & "', '" & rst!RELCode & "', '" & rst!MBMIcBcPp & "', '" & rst!MBMGroupCompany & "'," & _
'                        "'" & rst!HLTCode & "', '" & rst!MBMDepartment & "', '" & rst!INSCode & "', " & tmpPremium & ",  " & _
'                        "'" & tmpAdultChild & "', '" & rst!PAYCode & "', '" & rst!MBMMemberType & "','" & rst!MBMCNumber & "', '" & rst!MBMCCoverID & "', " & _
'                        "'" & tmpMBMCName & "' , '" & rst!MBMCICBCPPNo & "', '" & rst!MBMCRELCode & "', '" & tmpEndtBordxDate & "',  '" & rst!MBMEndtBordxBatch & "', " & tmpMBMMCO & ")"
'            medixmasterconnPluto.Execute sqlinsert
            
            
'        rst.MoveNext
'        Loop
'       rst.Close
'        Set rst = Nothing
'    End If
    Call PrintMCOInvoice
    If CboRptType.ListIndex = 0 Then
        Call MCOInvoiceExcel
        'cRptMCO.ReportFileName = crdPath & "\Report\MCOInvoice\MCOInvoiceTemplate.xslt"
    Else
        Call MCOListingExcel
        'If cboMCOListingVer = "MCO Listing" Then
        '    If OptByDept.value = True Then
        '        cRptMCO.ReportFileName = crdPath & "\Report\MCOInvoice\MCOListing_WithoutPremium_Dept.rpt"
        '    Else
        '        cRptMCO.ReportFileName = crdPath & "\Report\MCOInvoice\MCOListing_WithoutPremium.rpt"
        '    End If
        'ElseIf cboMCOListingVer = "MCO Listing With Premium" Then
        '    If OptByDept.value = True Then
        '        cRptMCO.ReportFileName = crdPath & "\Report\MCOInvoice\MCOListing_Premium_Dept.rpt"
        '    Else
        '        cRptMCO.ReportFileName = crdPath & "\Report\MCOInvoice\MCOListing_Premium.rpt"
        '    End If
        'End If
    End If
End Sub

Private Sub PrintMCOInvoice()
Dim AA As Integer
Dim tmpFormula As String
Dim tmpPrtGrp As Boolean
Dim InvoiceExist As Boolean
Dim MemberNo As String
Dim CovID As String
Dim tmpPln As String
Dim tmpIns As String
Dim TmpEffDate As String
Dim tmpExpDate As String
Dim tmpAdultChild As String
Dim tmpBMCO As Double
Dim tmpMCO As Double
Dim CompTrans As Boolean
Dim tmpPrefix As String

    'bPrintPass = False
    'CancelPrinting = False
    CompTrans = False
    InvoiceExist = False
    For AA = 1 To lvwMCO.listitems.Count
        tmpPayor = Mid(lvwMCO.listitems(AA).Text, 1, 2)
        If Trim(lvwMCO.listitems(AA).SubItems(10)) <> "" Then
            InvoiceExist = True
        End If
        Exit For
    Next AA
    
    tmpPrtGrp = False
    tmpPrefix = ""

    Call GetPayorInfo(tmpPayor, PAYName, PayAdd1, PayAdd2, PayAdd3, TelNo, FaxNo, ContactName, tmpPrefix)
    
    If cmdMCOPrint Then
        If InvoiceExist = False Then
            On Err GoTo ErrUpdate
            medixmasterconn.beginTrans
            medixmasterconnMaint.beginTrans
            
            Call IncInvoiceNo(tmpPayor, tmpPrefix)
            For AA = 1 To lvwMCO.listitems.Count
               ' If Trim(lvwMCO.listitems(AA).ListSubItems(10).Text) = "" Then
                    lvwMCO.listitems(AA).ListSubItems(10).Text = TxtInvoiceNo
                    MemberNo = Trim(lvwMCO.listitems(AA).Text)
                    CovID = Trim(lvwMCO.listitems(AA).ListSubItems(1))
                    'tmpPln = IIf(Len(lvwMCO.listitems(AA).ListSubItems(12)), Trim(lvwMCO.listitems(AA).ListSubItems(12)), "")
                    'tmpIns = IIf(Len(lvwMCO.listitems(AA).ListSubItems(13)), Trim(lvwMCO.listitems(AA).ListSubItems(13)), "")
                    TmpEffDate = IIf(Len(lvwMCO.listitems(AA).ListSubItems(7)), lvwMCO.listitems(AA).ListSubItems(7), "")
                    tmpExpDate = IIf(Len(lvwMCO.listitems(AA).ListSubItems(8)), lvwMCO.listitems(AA).ListSubItems(8), "")
                    tmpAdultChild = IIf(Len(lvwMCO.listitems(AA).ListSubItems(9)), Mid(lvwMCO.listitems(AA).ListSubItems(9), 1, 1), "")
                    tmpMCO = CDbl(IIf(IsNumeric(lvwMCO.listitems(AA).ListSubItems(12)), lvwMCO.listitems(AA).ListSubItems(12), 0))
                    Call UpdateINVPrintingFile(MemberNo, CovID)
             '   End If
            Next AA
            medixmasterconn.commitTrans
            medixmasterconnMaint.commitTrans
            CompTrans = True
        End If
        'If bPrintPass = False And CancelPrinting = False Then
        '    bPrintPass = PrintFunction
        '    If bPrintPass = False Then
        '        Screen.MousePointer = Default
        '        Exit Sub
        '    End If
        'End If
    Else
        cRptMCO.PrintReport
    End If
    CompTrans = True
ErrUpdate:
    If CompTrans = False And InvoiceExist = True Then
        medixmasterconn.RollbackTrans
        medixmasterconnMaint.RollbackTrans
    End If
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
End Sub

Private Sub IncInvoiceNo(tmpPayor As String, tmpPrefix As String)
Dim MBMMCOINVQ As New ADODB.Recordset
Dim strsql As String
Dim InvoiceNo As String

    If tmpPrefix = "S" Then
        strsql = "Select * from MBMMCOINVQ "
    Else
        strsql = "Select * from MBMMCOINVQ where PAYCode='" & Mid(tmpPayor, 1, 2) & "'"
    End If
    MBMMCOINVQ.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If MBMMCOINVQ.recordCount = 0 Then
        With MBMMCOINVQ
            .AddNew
                !PAYCode = tmpPayor
                !prefix = tmpPrefix
                !MCOInvoice_No = "00001"
                InvoiceNo = "00001"
            .Update
        End With
    Else
        With MBMMCOINVQ
            If Len(!prefix) Then
                tmpPrefix = !prefix
            Else
                !prefix = "S"
            End If
            !MCOInvoice_No = Format(!MCOInvoice_No + 1, "00000")
            InvoiceNo = Format(!MCOInvoice_No, "00000")
        .Update
        End With
    End If
    TxtInvoiceNo = tmpPrefix & InvoiceNo
    MBMMCOINVQ.Close
    Set MBMMCOINVQ = Nothing
End Sub


Private Sub UpdateINVPrintingFile(tmpMBMNo As String, TmpCovID As String)
Dim strsql As String
Dim InvoiceDate As String
    
    InvoiceDate = Format(Date, "YYYY/MM/DD")
    'If TmpCovID = "00" Then   'Update file for MBMOthers Table
    '    strsql = " update MBMOthers set " & _
    '             " MBMInvoiceDate = '" & InvoiceDate & "', " & _
    '             " MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "', " & _
    '             " MBMInvoiceBy = '" & Logon & "' " & _
    '             " Where MBMNumber = '" & Trim(tmpMBMNo) & "'"
    '             medixmasterconn.Execute strsql
                 
    'Else    'update file for MBMCoveredPersons Table
    '    strsql = " update MBMCoveredPersons set " & _
    '             " MBMInvoiceDate = '" & InvoiceDate & "', " & _
    '             " MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "', " & _
    '             " MBMInvoiceBy = '" & Logon & "' " & _
   '              " Where MBMCNumber = '" & Trim(tmpMBMNo) & "'" & _
    '             " and MBMCCoverID = '" & TmpCovID & "'"
    '             medixmasterconn.Execute strsql
    'End If
    If TmpCovID = "00" Then
        strsql = ""
        strsql = " update MBMReinstatement set " & _
             " MBMInvoiceDate = '" & InvoiceDate & "', " & _
             " MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "', " & _
             " MBMInvoiceBy = '" & Logon & "' " & _
             " Where MBMNumber = '" & Trim(tmpMBMNo) & "'"
             medixmasterconn.Execute strsql
    Else
        strsql = ""
        strsql = " update MBMReinstatement set " & _
             " MBMInvoiceDate = '" & InvoiceDate & "', " & _
             " MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "', " & _
             " MBMInvoiceBy = '" & Logon & "' " & _
             " Where MBMNumber = '" & Trim(tmpMBMNo) & "' and MBMCoverID = '" & TmpCovID & "'"
             medixmasterconn.Execute strsql
    End If
End Sub

Private Sub GetPayorInfo(tmpPayCode As String, PayorName As String, PayAdd1 As String, _
PayAdd2 As String, PayAdd3 As String, TelNo As String, FaxNo As String, _
ContactName As String, tmpPrefix As String)
Dim PayRec As New ADODB.Recordset
Dim strsqlPAY As String
    
    If Trim(tmpPayCode) <> "" Then
        strsqlPAY = "Select PAYCompanyName, PAYAddress1, PAYAddress2, PAYAddress3, PAYMAINContact, " & _
                    "PAYTELno, PAYFAXno,MCOInvoicePrefix from Pay where PAYCode='" & Trim(Mid(tmpPayCode, 1, 2)) & "'"
        PayRec.Open strsqlPAY, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do While Not PayRec.EOF
            PayorName = IIf(Len(PayRec!PAYCompanyName), Trim(PayRec!PAYCompanyName), "")
            PayAdd1 = IIf(Len(PayRec!PAYAddress1), Trim(PayRec!PAYAddress1), "")
            PayAdd2 = IIf(Len(PayRec!PAYAddress2), Trim(PayRec!PAYAddress2), "")
            PayAdd3 = IIf(Len(PayRec!PAYAddress3), Trim(PayRec!PAYAddress3), "")
            TelNo = IIf(Len(PayRec!PAYTELno), Trim(PayRec!PAYTELno), "")
            FaxNo = IIf(Len(PayRec!PAYFAXno), Trim(PayRec!PAYFAXno), "")
            ContactName = IIf(Len(PayRec!PAYMAINContact), Trim(PayRec!PAYMAINContact), "")
            tmpPrefix = IIf(Trim(PayRec!MCOInvoicePrefix) <> "", Trim(PayRec!MCOInvoicePrefix), "S")
            PayRec.MoveNext
        Loop
        PayRec.Close
        Set PayRec = Nothing
    End If
End Sub

Private Sub MCOInvoiceExcel()
'On Error Resume Next
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.WORKSHEET
Dim i As Integer
Dim ii As Integer
Dim RecordSelected As Integer
Dim MBMNote1 As New ADODB.Recordset
Dim strsql As String
Dim MBMNote2 As New ADODB.Recordset
Dim strsql2 As String
Dim MBMNote3 As New ADODB.Recordset
Dim strsql3 As String
Dim SuppRemark As New ADODB.Recordset
Dim strsqlSuppRemark As String
Dim SuppRemark2 As New ADODB.Recordset
Dim strsqlSuppRemark2 As String
Dim MBMBNF As New ADODB.Recordset
Dim strsql4 As String
Dim MBMCase1 As New ADODB.Recordset
Dim strsql5 As String
Dim MBMCase2 As New ADODB.Recordset
Dim strsql6 As String
Dim MBMCLM1 As New ADODB.Recordset
Dim strsqlCLM1 As String
Dim MBMCLM2 As New ADODB.Recordset
Dim strsqlCLM2 As String
Dim APPAmt As String
Dim APPAmt2 As String
Dim SuppRec As New ADODB.Recordset
Dim strsqlSuppRec As String
Dim SuppRec2 As New ADODB.Recordset
Dim strsqlSuppRec2 As String
Dim MBMRec As New ADODB.Recordset
Dim strsqlMBMRec As String
Dim MBMHistoryRec As New ADODB.Recordset
Dim strsqlMBMHistoryRec As String
Dim MBMAccHistoryRec As New ADODB.Recordset
Dim strsqlMBMAccHistoryRec As String
Dim CurrEffDate, PrevEffDate As String
Dim cnt As String


    cnt = 0
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
    End If
    
    'If Err.Number > 0 Then
    '    MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
    '    GoTo ExitFunction
    'End If
        
    On Error GoTo HANDLE_ERROR
    objExcel.Interactive = False
    
    
        If objExcel.Visible = False Then
            objExcel.Visible = True
        End If
    
        objExcel.WindowState = xlMaximized

        Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "MCOInvoice\MCOInvoiceTemplate.xls")
        'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMExclusion.xlt")
                
        
    '========================================================================================
        Set objWorksheet = objWorkbook.Sheets(1)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
    
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
           
            objWorksheet.Cells(6, "A") = PAYName
            objWorksheet.Cells(7, "A") = PayAdd1
            objWorksheet.Cells(8, "A") = PayAdd2
            objWorksheet.Cells(9, "A") = PayAdd3
            objWorksheet.Cells(13, "B") = ContactName
            objWorksheet.Cells(14, "B") = TelNo
            objWorksheet.Cells(14, "D") = FaxNo

            If cboMCOCriteriaSel.ListIndex = 0 Then
                strsqlMBMRec = " SELECT MBMEndtBordxDate,MBMEndtBordxBatch,PayCode,TotalPrincipal,INSCode,MBMAdultChild,MBMInvoiceNo,MBMInvoiceDate " & _
                "FROM View_MCOReinstatement_Invoice_P Where MBMEndtBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "' " & _
                " And MBMEndtBordxBatch = " & txtMCOBatchNo & " And PayCode = '" & Mid(cboMCOPayor, 1, 2) & "' " & _
                " Order By INSCode, MBMAdultChild "
           ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
                strsqlMBMRec = " SELECT MBMEndtBordxDate,MBMEndtBordxBatch,PayCode,TotalPrincipal,INSCode,MBMAdultChild,MBMInvoiceNo,MBMInvoiceDate " & _
                "FROM View_MCOReinstatement_Invoice_P Where MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "'" & _
                " Order By INSCode, MBMAdultChild "
            End If
            
            MBMRec.Open strsqlMBMRec, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            i = 0
            Do While Not MBMRec.EOF
                cnt = cnt + 1
                i = i + 1
                objWorksheet.Cells(18 + CDbl(i), "A") = cnt
                objWorksheet.Cells(18 + CDbl(i), "B") = MBMRec!MBMEndtBordxDate
                objWorksheet.Cells(18 + CDbl(i), "C") = MBMRec!MBMEndtBordxBatch
                objWorksheet.Cells(18 + CDbl(i), "D") = MBMRec!INSCode
                objWorksheet.Cells(18 + CDbl(i), "E") = MBMRec!MBMAdultChild
                objWorksheet.Cells(18 + CDbl(i), "F") = MBMRec!TotalPrincipal
                objWorksheet.Cells(9, "H") = MBMRec!MBMInvoiceNo
                objWorksheet.Cells(10, "H") = MBMRec!MBMInvoiceDate

            MBMRec.MoveNext
            Loop
            MBMRec.Close
            Set MBMRec = Nothing
            
            strsqlMBMRec = ""
            
            If cboMCOCriteriaSel.ListIndex = 0 Then
                strsqlMBMRec = " SELECT MBMEndtBordxDate,MBMEndtBordxBatch,PayCode,TotalCoveredPersons,INSCode,MBMAdultChild,MBMMCOFee " & _
                "FROM View_MCOReinstatement_Invoice_CP Where MBMEndtBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "' " & _
                " And MBMEndtBordxBatch = " & txtMCOBatchNo & " And PayCode = '" & Mid(cboMCOPayor, 1, 2) & "' " & _
                " Order By INSCode, MBMAdultChild "
           ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
                strsqlMBMRec = " SELECT MBMEndtBordxDate,MBMEndtBordxBatch,PayCode,TotalCoveredPersons,INSCode,MBMAdultChild,MBMMCOFee " & _
                "FROM View_MCOReinstatement_Invoice_CP Where MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "'" & _
                " Order By INSCode, MBMAdultChild "
            End If
            
            MBMRec.Open strsqlMBMRec, medixmasterconn, adOpenKeyset, adLockOptimistic
                i = 0

                Do While Not MBMRec.EOF
                    i = i + 1
                    objWorksheet.Cells(18 + CDbl(i), "G") = MBMRec!TotalCoveredPersons
                    objWorksheet.Cells(18 + CDbl(i), "H") = Round(MBMRec!MBMMCOFee)
                MBMRec.MoveNext
                Loop
                MBMRec.Close
                Set MBMRec = Nothing

        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        Exit Sub
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault
        Exit Sub
HANDLE_ERROR:
        MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
        Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
        Set objWorksheet = Nothing
        Set objWorkbook = Nothing
        GoTo ExitFunction
End Sub

Private Sub MCOListingExcel()
'On Error Resume Next
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.WORKSHEET
Dim i As Integer
Dim ii As Integer
Dim RecordSelected As Integer
Dim MBMNote1 As New ADODB.Recordset
Dim strsql As String
Dim MBMNote2 As New ADODB.Recordset
Dim strsql2 As String
Dim MBMNote3 As New ADODB.Recordset
Dim strsql3 As String
Dim SuppRemark As New ADODB.Recordset
Dim strsqlSuppRemark As String
Dim SuppRemark2 As New ADODB.Recordset
Dim strsqlSuppRemark2 As String
Dim MBMBNF As New ADODB.Recordset
Dim strsql4 As String
Dim MBMCase1 As New ADODB.Recordset
Dim strsql5 As String
Dim MBMCase2 As New ADODB.Recordset
Dim strsql6 As String
Dim MBMCLM1 As New ADODB.Recordset
Dim strsqlCLM1 As String
Dim MBMCLM2 As New ADODB.Recordset
Dim strsqlCLM2 As String
Dim APPAmt As String
Dim APPAmt2 As String
Dim SuppRec As New ADODB.Recordset
Dim strsqlSuppRec As String
Dim SuppRec2 As New ADODB.Recordset
Dim strsqlSuppRec2 As String
Dim MBMRec As New ADODB.Recordset
Dim strsqlMBMRec As String
Dim MBMHistoryRec As New ADODB.Recordset
Dim strsqlMBMHistoryRec As String
Dim MBMAccHistoryRec As New ADODB.Recordset
Dim strsqlMBMAccHistoryRec As String
Dim CurrEffDate, PrevEffDate As String
Dim cnt As String
Dim tmpTotalMCO As Double

    cnt = 0
    tmpTotalMCO = 0
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
    End If
    
    'If Err.Number > 0 Then
    '    MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
    '    GoTo ExitFunction
    'End If
        
    On Error GoTo HANDLE_ERROR
    objExcel.Interactive = False
    
    
        If objExcel.Visible = False Then
            objExcel.Visible = True
        End If
    
        objExcel.WindowState = xlMaximized

        Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "MCOInvoice\MCOListingTemplate.xls")
        'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\MBMExclusion.xlt")
                
        
    '========================================================================================
        Set objWorksheet = objWorkbook.Sheets(1)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
    
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
           

            If cboMCOCriteriaSel.ListIndex = 0 Then
                strsqlMBMRec = " SELECT * FROM View_MBM_MCOReinstatement Where MBMEndtBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "' " & _
                " And MBMEndtBordxBatch = " & txtMCOBatchNo & " And PayCode = '" & Mid(cboMCOPayor, 1, 2) & "' "
            ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
                strsqlMBMRec = " SELECT * FROM View_MBM_MCOReinstatement Where MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "'"
            End If
            MBMRec.Open strsqlMBMRec, medixmasterconn, adOpenKeyset, adLockOptimistic
            cnt = 0
            i = 0
            Do While Not MBMRec.EOF
                cnt = cnt + 1
                i = i + 1
                
                objWorksheet.Cells(2, "C") = MBMRec!PAYCode
                objWorksheet.Cells(3, "C") = MBMRec!MBMEndtBordxDate
                objWorksheet.Cells(3, "E") = Format(MBMRec!MBMEndtBordxBatch, "mm/dd/yyyy")
                
                objWorksheet.Cells(2, "H") = MBMRec!MBMInvoiceNo
                objWorksheet.Cells(3, "H") = Format(MBMRec!MBMInvoiceDate, "mm/dd/yyyy")

                objWorksheet.Cells(6 + CDbl(i), "A") = cnt
                objWorksheet.Cells(6 + CDbl(i), "B") = MBMRec!MBMNumber
                objWorksheet.Cells(6 + CDbl(i), "C") = MBMRec!MBMPolicyNo
                If Len(MBMRec!MBMCCoverID) Then
                    objWorksheet.Cells(6 + CDbl(i), "D") = MBMRec!MBMCName
                Else
                    objWorksheet.Cells(6 + CDbl(i), "D") = MBMRec!MBMName
                End If
                If Len(MBMRec!MBMCCoverID) Then
                    objWorksheet.Cells(6 + CDbl(i), "E") = MBMRec!MBMCRELCode
                Else
                    objWorksheet.Cells(6 + CDbl(i), "E") = MBMRec!RELCode
                End If
                If Len(MBMRec!MBMCCoverID) Then
                    objWorksheet.Cells(6 + CDbl(i), "F") = MBMRec!MBMCICBCPPNo
                Else
                    objWorksheet.Cells(6 + CDbl(i), "F") = MBMRec!MBMIcBcPp
                End If
                
                objWorksheet.Cells(6 + CDbl(i), "G") = MBMRec!MBMGroupCompany
                objWorksheet.Cells(6 + CDbl(i), "H") = MBMRec!DayInsured
                objWorksheet.Cells(6 + CDbl(i), "I") = Round(MBMRec!MBMMCOFee)
                tmpTotalMCO = CDbl(tmpTotalMCO) + CDbl(MBMRec!MBMMCOFee)
            MBMRec.MoveNext
            Loop
            MBMRec.Close
            Set MBMRec = Nothing
            
            i = i + 1
            objWorksheet.Cells(6 + CDbl(i), "I") = "---------------"
            
            i = i + 1
            objWorksheet.Cells(6 + CDbl(i), "I") = Round(tmpTotalMCO)
            
            i = i + 1
            objWorksheet.Cells(6 + CDbl(i), "I") = "---------------"
            
            If cboMCOCriteriaSel.ListIndex = 0 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMEndtBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "' " & _
                " And MBMEndtBordxBatch = " & txtMCOBatchNo & " And PayCode = '" & Mid(cboMCOPayor, 1, 2) & "' " & _
                " And INSCode = 'I' And Type = 'A' "
                
            ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "' " & _
                                " And INSCode = 'I' And Type = 'A' "
            End If
            MBMRec.Open strsqlMBMRec, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            i = i + 12
            objWorksheet.Cells(CDbl(i), "A") = "Individual Adult"
            objWorksheet.Cells(CDbl(i), "D") = MBMRec!TotalNo
            MBMRec.Close
            Set MBMRec = Nothing
            
            If cboMCOCriteriaSel.ListIndex = 0 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMEndtBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "' " & _
                " And MBMEndtBordxBatch = " & txtMCOBatchNo & " And PayCode = '" & Mid(cboMCOPayor, 1, 2) & "' " & _
                " And INSCode = 'I' And Type = 'C' "
                
            ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "' " & _
                                " And INSCode = 'I' And Type = 'C' "
            End If
            MBMRec.Open strsqlMBMRec, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            i = i + 1
            objWorksheet.Cells(CDbl(i), "A") = "Individual Child"
            objWorksheet.Cells(CDbl(i), "D") = MBMRec!TotalNo
            
            MBMRec.Close
            Set MBMRec = Nothing
            
            If cboMCOCriteriaSel.ListIndex = 0 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMEndtBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "' " & _
                " And MBMEndtBordxBatch = " & txtMCOBatchNo & " And PayCode = '" & Mid(cboMCOPayor, 1, 2) & "' " & _
                " And INSCode = 'F'  And MBMCCoverID Is Null "
                
            ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "' " & _
                                " And INSCode = 'F' And MBMCCoverID Is Null "
            End If
            MBMRec.Open strsqlMBMRec, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            i = i + 1
            objWorksheet.Cells(CDbl(i), "A") = "Family Principal"
            objWorksheet.Cells(CDbl(i), "D") = MBMRec!TotalNo
            
            MBMRec.Close
            Set MBMRec = Nothing
            
            If cboMCOCriteriaSel.ListIndex = 0 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMEndtBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "' " & _
                " And MBMEndtBordxBatch = " & txtMCOBatchNo & " And PayCode = '" & Mid(cboMCOPayor, 1, 2) & "' " & _
                " And INSCode = 'F' "
                
            ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "' " & _
                                " And INSCode = 'F' "
            End If
            
            MBMRec.Open strsqlMBMRec, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            objWorksheet.Cells(CDbl(i), "E") = "Total Family"
            objWorksheet.Cells(CDbl(i), "H") = MBMRec!TotalNo
            
            MBMRec.Close
            Set MBMRec = Nothing
            
            
            If cboMCOCriteriaSel.ListIndex = 0 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMEndtBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "' " & _
                " And MBMEndtBordxBatch = " & txtMCOBatchNo & " And PayCode = '" & Mid(cboMCOPayor, 1, 2) & "' " & _
                " And INSCode = 'G' And Type = 'A' "
                
            ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "' " & _
                                " And INSCode = 'G' And Type = 'A' "
            End If
            
            MBMRec.Open strsqlMBMRec, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            i = i + 1
            objWorksheet.Cells(CDbl(i), "A") = "Group Individual Adult"
            objWorksheet.Cells(CDbl(i), "D") = MBMRec!TotalNo
            
            MBMRec.Close
            Set MBMRec = Nothing
            
            
            If cboMCOCriteriaSel.ListIndex = 0 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMEndtBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "' " & _
                " And MBMEndtBordxBatch = " & txtMCOBatchNo & " And PayCode = '" & Mid(cboMCOPayor, 1, 2) & "' " & _
                " And INSCode = 'G' And Type = 'C' "
                
            ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "' " & _
                                " And INSCode = 'G' And Type = 'C' "
            End If
            
            MBMRec.Open strsqlMBMRec, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            i = i + 1
            objWorksheet.Cells(CDbl(i), "A") = "Group Individual Child"
            objWorksheet.Cells(CDbl(i), "D") = MBMRec!TotalNo
            
            MBMRec.Close
            Set MBMRec = Nothing
            
            If cboMCOCriteriaSel.ListIndex = 0 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMEndtBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "' " & _
                " And MBMEndtBordxBatch = " & txtMCOBatchNo & " And PayCode = '" & Mid(cboMCOPayor, 1, 2) & "' " & _
                " And INSCode = 'H' And MBMCCoverID Is Null "
                
            ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "' " & _
                                " And INSCode = 'H' And MBMCCoverID Is Null "
            End If
            
            MBMRec.Open strsqlMBMRec, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            i = i + 1
            objWorksheet.Cells(CDbl(i), "A") = "Group Family Principal"
            objWorksheet.Cells(CDbl(i), "D") = MBMRec!TotalNo
            
            MBMRec.Close
            Set MBMRec = Nothing
            
            If cboMCOCriteriaSel.ListIndex = 0 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMEndtBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "' " & _
                " And MBMEndtBordxBatch = " & txtMCOBatchNo & " And PayCode = '" & Mid(cboMCOPayor, 1, 2) & "' " & _
                " And INSCode = 'H' "
                
            ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "' " & _
                                " And INSCode = 'H' "
            End If
            
            MBMRec.Open strsqlMBMRec, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            objWorksheet.Cells(CDbl(i), "E") = "Total Group Family"
            objWorksheet.Cells(CDbl(i), "H") = MBMRec!TotalNo
            
            MBMRec.Close
            Set MBMRec = Nothing
            
            If cboMCOCriteriaSel.ListIndex = 0 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMEndtBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "' " & _
                " And MBMEndtBordxBatch = " & txtMCOBatchNo & " And PayCode = '" & Mid(cboMCOPayor, 1, 2) & "' And MBMCCoverID Is Null  "

            ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "' " & _
                                " And MBMCCoverID Is Null "
            End If
            
            MBMRec.Open strsqlMBMRec, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            i = i + 1
            objWorksheet.Cells(CDbl(i), "A") = "Total No Of Policy"
            objWorksheet.Cells(CDbl(i), "D") = MBMRec!TotalNo
            
            MBMRec.Close
            Set MBMRec = Nothing
            
            If cboMCOCriteriaSel.ListIndex = 0 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMEndtBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "' " & _
                " And MBMEndtBordxBatch = " & txtMCOBatchNo & " And PayCode = '" & Mid(cboMCOPayor, 1, 2) & "'"

            ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
                strsqlMBMRec = " SELECT Count(*) As TotalNo FROM View_MBM_MCOReinstatement Where MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "'"
            End If
            
            MBMRec.Open strsqlMBMRec, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            objWorksheet.Cells(CDbl(i), "E") = "Total Covered Persons"
            objWorksheet.Cells(CDbl(i), "H") = MBMRec!TotalNo
            
            MBMRec.Close
            Set MBMRec = Nothing
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        Exit Sub
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault
        Exit Sub
HANDLE_ERROR:
        MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
        Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
        Set objWorksheet = Nothing
        Set objWorkbook = Nothing
        GoTo ExitFunction
End Sub
Private Sub txtMCOBordxDate_LostFocus()
    If IsDate(txtMCOBordxDate) Then
    Else
        If txtMCOBordxDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtMCOBordxDate.SetFocus
        End If
    End If
End Sub
