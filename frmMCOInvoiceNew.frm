VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMCOInvoiceNew 
   Caption         =   "MCO Invoice"
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
   Begin VB.Frame FrameINV 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11055
      Begin VB.TextBox txtINVInvoiceNo 
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
         Left            =   6000
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton OptByGrpCom 
         Caption         =   "By Group Company"
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
         Left            =   2160
         TabIndex        =   32
         Top             =   250
         Width           =   2055
      End
      Begin VB.OptionButton OptByPayor 
         Caption         =   "By Payor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   250
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox txtMCOBatchNo 
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
         Left            =   8520
         MaxLength       =   2
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboINVCriteriaSel 
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
         ItemData        =   "frmMCOInvoiceNew.frx":0000
         Left            =   2160
         List            =   "frmMCOInvoiceNew.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   2775
      End
      Begin MSMask.MaskEdBox txtINVBordxDate 
         Height          =   315
         Left            =   6000
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Crystal.CrystalReport cRptINV 
         Left            =   9720
         Top             =   1680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         UserName        =   "sa"
         PrintFileLinesPerPage=   60
         WindowShowRefreshBtn=   -1  'True
      End
      Begin MSMask.MaskEdBox txtINVBordxDateTo 
         Height          =   315
         Left            =   8520
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtINVBordxDateFr 
         Height          =   315
         Left            =   6000
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog dlgCommonDialog_INV 
         Left            =   10440
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox cboINVPayor 
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
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   885
         Width           =   3855
      End
      Begin VB.ComboBox cboINVPrintOpt 
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
         ItemData        =   "frmMCOInvoiceNew.frx":0041
         Left            =   2160
         List            =   "frmMCOInvoiceNew.frx":004B
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1170
         Width           =   2415
      End
      Begin VB.ComboBox CboGrpCompany 
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
         Left            =   7200
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   885
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Grp Company"
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
         Left            =   6120
         TabIndex        =   29
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Batch No"
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
         Left            =   7680
         TabIndex        =   28
         Top             =   645
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lbl30 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   7560
         TabIndex        =   26
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl29 
         Alignment       =   2  'Center
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
         Left            =   5040
         TabIndex        =   17
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl36 
         Caption         =   "Print Option"
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
         TabIndex        =   16
         Top             =   1245
         Width           =   855
      End
      Begin VB.Label lbl34 
         Caption         =   "Payor"
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
         TabIndex        =   15
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label lbl27 
         Caption         =   "Single/Batch/Range"
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
         TabIndex        =   14
         Top             =   645
         Width           =   1695
      End
   End
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5400
      Left            =   120
      TabIndex        =   24
      Top             =   1800
      Width           =   11055
      Begin MSComctlLib.ListView lvwInvoice 
         Height          =   4995
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8811
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   19
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Payor"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Membership No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Cov ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Rel"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Bordx Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Policy No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Hlt Type"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Batch No"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Group Company"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Printed"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   "Invoice No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "PlnCode"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "InsType"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "EffDate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "ExpDate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "AdultChild"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "BasicMCo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "McoAmt"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   7200
      Width           =   11055
      Begin VB.CommandButton cmdINVClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9240
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdINVPrintPreview 
         Caption         =   "Pre&view"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8400
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdINVPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7560
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdINVSearch 
         Caption         =   "S&earch"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6720
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtInvoiceNo 
         Height          =   285
         Left            =   5880
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdINVClose 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10080
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label203 
         Alignment       =   2  'Center
         Caption         =   "Total Supp"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   2880
         TabIndex        =   23
         Top             =   240
         Width           =   960
      End
      Begin VB.Label lblTotSupp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3840
         TabIndex        =   22
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label lblTotPrinc 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1440
         TabIndex        =   21
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label202 
         Alignment       =   2  'Center
         Caption         =   "Total Principal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "MCO Invoice"
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
      TabIndex        =   27
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmMCOInvoiceNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SuppIDList() As String
Dim PAYCode As String
Dim TotSuppCnt As Integer
Dim MultiInvFound As Boolean
Dim GetSubSupp As String
Dim GetSuppMCO As String


Private Sub cboINVPayor_Click()
Dim PAYGRP As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strsql As String
    
    If OptByGrpCom.value = True Then
        CboGrpCompany.Clear
        strsql = "AND PAYCode= '" & Mid(cboINVPayor, 1, 2) & "'"
     '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "SearchComboListing"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("Selection", adVarChar, adParamInput, 255, "1")
        cmd.Parameters.Append Prm2
        PAYGRP.CursorLocation = adUseClient
        PAYGRP.CursorType = adOpenForwardOnly
        PAYGRP.CacheSize = 100
        Set PAYGRP = cmd.Execute
        
        Do Until PAYGRP.EOF
            With PAYGRP
                CboGrpCompany.AddItem Trim(PAYGRP!MBMGroupCompany)
                PAYGRP.MoveNext
            End With
        Loop
        PAYGRP.Close
        Set PAYGRP = Nothing
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
    End If
End Sub

Private Sub cmdINVClear_Click()
    txtMCOBatchNo = ""
    txtINVBordxDate = "__/__/____"
    txtINVBordxDateFr = "__/__/____"
    txtINVBordxDateTo = "__/__/____"
    txtINVInvoiceNo = ""
    lvwInvoice.listitems.Clear
    cmdINVPrint.Enabled = False
    cmdINVPrintPreview.Enabled = False
    lblTotPrinc = 0
    lblTotSupp = 0
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
 
Private Sub INVPrincipalListing()
Dim rsSearch As New ADODB.Recordset
Dim sql As String
Dim TotalPrincipal As Integer
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim ListMaster As ListItem
Dim AA As Integer
Dim tmpSuppID As String
Dim suppExist As Boolean
Dim PrevInvNO As String
Dim CurrInvNo As String
Dim SelField As String
    MultiInvFound = False
    PrevInvNO = ""
    ReDim SuppIDList(0)

    ReDim SuppIDList(0)
    TotSuppCnt = 0
    lvwInvoice.listitems.Clear
    cmdINVPrint.Enabled = False
    cmdINVPrintPreview.Enabled = False
    sql = ""
    'sql = "SELECT * FROM VIEW_Invoice_ReportSearch2"
    SelField = "Select PayCode, MBMNumber, MBMMemberType, MBMBordxDate, MBMPolicyNo, " & _
               "HLTCode, MBMBatchNo, MBMName, MBMGroupCompany, MBMInvoiceNo, " & _
               "PLNCode, INSCode, MBMPayorEffDate, MBMPayorEXPDate, MBMAdultChild, " & _
               "MBMBasicMCO, MBMMCOFee "

    Call SelectionCriterial(sql)

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Set cmd1.ActiveConnection = medixmasterconn
    cmd1.CommandText = "SP_MCOFees"
    cmd1.CommandTimeout = 300
    cmd1.CommandType = adCmdStoredProc
    Set Prm1 = cmd1.CreateParameter("SelType", adVarChar, adParamInput, 50, "PrincList")
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("SelFields", adVarChar, adParamInput, 2000, SelField)
    cmd1.Parameters.Append Prm2
    Set Prm3 = cmd1.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, sql)
    cmd1.Parameters.Append Prm3
    Set prm4 = cmd1.CreateParameter("SortList", adVarChar, adParamInput, 2000, "")
    cmd1.Parameters.Append prm4

    rsSearch.CacheSize = 100
    rsSearch.CursorLocation = adUseClient
    rsSearch.CursorType = adOpenForwardOnly
    Set rsSearch = cmd1.Execute
    Do While Not rsSearch.EOF
        cmdINVPrint.Enabled = True
        cmdINVPrintPreview.Enabled = True
        Set ListMaster = lvwInvoice.listitems.Add(, , rsSearch!PAYCode)
        ListMaster.SubItems(1) = rsSearch!MBMNumber
        ListMaster.SubItems(2) = "00"
        ListMaster.SubItems(3) = rsSearch!MBMMemberType
        ListMaster.SubItems(4) = rsSearch!MBMBordxDate
        ListMaster.SubItems(5) = IIf(Len(rsSearch!MBMPolicyNo), Trim(rsSearch!MBMPolicyNo), "")
        ListMaster.SubItems(6) = Trim(rsSearch!HLTCode)
        ListMaster.SubItems(7) = IIf(Len(rsSearch!MBMBatchNo), Trim(rsSearch!MBMBatchNo), "")
        ListMaster.SubItems(8) = Trim(rsSearch!MBMName)
        ListMaster.SubItems(9) = IIf(Len(rsSearch!MBMGroupCompany), Trim(rsSearch!MBMGroupCompany), "")
        ListMaster.SubItems(10) = "NO"
        ListMaster.SubItems(11) = ""
        If Len(rsSearch!MBMInvoiceNo) Then
            ListMaster.SubItems(10) = "YES"
            ListMaster.SubItems(11) = Trim(rsSearch!MBMInvoiceNo)
        End If
        ListMaster.SubItems(12) = IIf(Len(rsSearch!PLNCode), Trim(rsSearch!PLNCode), "")
        ListMaster.SubItems(13) = IIf(Len(rsSearch!INSCode), Trim(rsSearch!INSCode), "")
        ListMaster.SubItems(14) = IIf(Len(rsSearch!MBMPayorEffDate), rsSearch!MBMPayorEffDate, "")
        ListMaster.SubItems(15) = IIf(Len(rsSearch!MBMPayorEXPDate), rsSearch!MBMPayorEXPDate, "")
        ListMaster.SubItems(16) = IIf(Len(rsSearch!MBMAdultChild), Trim(rsSearch!MBMAdultChild), "")
        ListMaster.SubItems(17) = IIf(IsNumeric(rsSearch!MBMBasicMCO), rsSearch!MBMBasicMCO, 0)
        ListMaster.SubItems(18) = IIf(IsNumeric(rsSearch!MBMMCOFee), rsSearch!MBMMCOFee, 0)
        CurrInvNo = IIf(Len(rsSearch!MBMInvoiceNo), Trim(rsSearch!MBMInvoiceNo), "")
        If PrevInvNO = "" Then
            PrevInvNO = IIf(Len(rsSearch!MBMInvoiceNo), Trim(rsSearch!MBMInvoiceNo), "")
        ElseIf CurrInvNo <> PrevInvNO Then
            MultiInvFound = True
        End If

        TotalPrincipal = TotalPrincipal + 1
        lblTotPrinc = TotalPrincipal
        PAYCode = rsSearch!PAYCode
        If cboINVCriteriaSel.ListIndex <> 2 Then
            If Trim(rsSearch!INSCode) = "F" Or Trim(rsSearch!INSCode) = "H" Then
                suppExist = False
                tmpSuppID = Trim(rsSearch!MBMNumber)
                If UBound(SuppIDList) > 0 Then
                    For AA = LBound(SuppIDList) To UBound(SuppIDList)
                         If tmpSuppID = SuppIDList(AA) Then
                            suppExist = True
                            Exit For
                         End If
                     Next AA
                End If
                If suppExist = False Then
                    ReDim Preserve SuppIDList(AA + 1)
                    SuppIDList(AA + 1) = tmpSuppID
                End If
                Call INVSuppListing(True, rsSearch!MBMNumber)
            End If
        End If
        rsSearch.MoveNext
    Loop
    rsSearch.Close
    Set rsSearch = Nothing
    
    Call INVSuppListing(False, "")
  '  medixmasterconn.Close
 '   Set medixmasterconn = Nothing
    If lvwInvoice.listitems.Count = 0 Then
        MsgBox "Record not found !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub cboINVCriteriaSel_Change()
    Call InvCriteriaSel
End Sub

Private Sub cboINVCriteriaSel_Click()
    Call InvCriteriaSel
End Sub

Private Sub cboINVCriteriaSel_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboINVCriteriaSel, KeyAscii)
End Sub

Private Sub cboINVPayor_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboINVPayor, KeyAscii)
End Sub

Private Sub cboINVPrintOpt_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboINVPrintOpt, KeyAscii)
End Sub

Private Sub cmdINVClose_Click()
    Unload Me
End Sub

Private Sub cmdINVPrint_Click()
    Call PrintOrPreview(True)
End Sub

Private Sub cmdINVPrintPreview_Click()
    Call PrintOrPreview(False)
End Sub

Private Sub cmdINVSearch_Click()
Dim ErrFound As Boolean
    ErrFound = True
    lblTotPrinc = 0
    lblTotSupp = 0
    
    If cboINVCriteriaSel.ListIndex = 0 Then
        If txtINVBordxDate = "__/__/____" Then
            MsgBox "Please Enter Bordereaux Date!", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf txtMCOBatchNo = "" Then
            MsgBox "Please Enter Batch Number!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            ErrFound = False
        End If
    ElseIf cboINVCriteriaSel.ListIndex = 1 Then
        If txtINVBordxDateFr = "__/__/____" Or txtINVBordxDateTo = "__/__/____" Then
            MsgBox "Please enter valid Bordereaux Range !", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            ErrFound = False
        End If
    ElseIf cboINVCriteriaSel.ListIndex = 2 Then
        If txtINVInvoiceNo = "" Then
            MsgBox "Please enter Invoice Number!", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            ErrFound = False
        End If
    End If
    cmdINVSearch.Enabled = False

    If ErrFound = False Then
        TotSuppCnt = 0
        lblTotPrinc = 0
        cmdINVClose.Enabled = False
        Screen.MousePointer = vbHourglass
        Call INVPrincipalListing
        Screen.MousePointer = Default
        cmdINVClose.Enabled = True
    End If
    cmdINVSearch.Enabled = True
End Sub

Private Sub INVSuppListing(tmpFrMBMNo As Boolean, tmpMBMNo As String)
Dim strsql As String
Dim MBMCov As New ADODB.Recordset
Dim ListMaster As ListItem
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim tmpMBMID As String
Dim AA As Integer
Dim suppExist As Boolean
Dim SelField As String

'    strsql = "SELECT * FROM VIEW_Invoice_SuppReportSearch2 "
    strsql = ""
    Call SelectionCriterial(strsql)
    If tmpFrMBMNo = True Then
        strsql = strsql & " AND MBMCNumber='" & tmpMBMNo & "'"
    End If
    strsql = strsql & " ORDER BY MBMCNumber"
    SelField = "Select MBMCNumber, MBMCCoverID, MBMCRELCode,  MBMPolicyNo,  " & _
               "MBMBordxDate, MBMBatchNo, MBMCName, MBMGroupCompany, MBMInvoiceNo, " & _
               "MBMCPLNCode, INSCode, MBMCEffDate, MBMCExpDate, MBMCAdultChild, " & _
               "MBMCBasicMCO, MBMCMCO "

    Set cmd1.ActiveConnection = medixmasterconn
    cmd1.CommandText = "SP_MCOFees"
    cmd1.CommandTimeout = 300
    cmd1.CommandType = adCmdStoredProc
    Set Prm1 = cmd1.CreateParameter("SelType", adVarChar, adParamInput, 50, "SuppList")
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("SelFields", adVarChar, adParamInput, 2000, SelField)
    cmd1.Parameters.Append Prm2
    Set Prm3 = cmd1.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, strsql)
    cmd1.Parameters.Append Prm3
    Set prm4 = cmd1.CreateParameter("SortList", adVarChar, adParamInput, 2000, "")
    cmd1.Parameters.Append prm4
    MBMCov.CacheSize = 100
    MBMCov.CursorLocation = adUseClient
    MBMCov.CursorType = adOpenForwardOnly
    Set MBMCov = cmd1.Execute
    Do While Not MBMCov.EOF
        suppExist = False
        If tmpFrMBMNo = False Then
            tmpMBMID = Trim(MBMCov!MBMCNumber)
            If UBound(SuppIDList) > 0 Then
                For AA = LBound(SuppIDList) To UBound(SuppIDList)
                     If tmpMBMID = Trim(SuppIDList(AA)) Then
                        suppExist = True
                        Exit For
                     End If
                 Next AA
            End If
        End If
        If suppExist = False Then
           Set ListMaster = lvwInvoice.listitems.Add(, , PAYCode)
            ListMaster.SubItems(1) = MBMCov!MBMCNumber
            ListMaster.SubItems(2) = MBMCov!MBMCCoverID
            ListMaster.SubItems(3) = MBMCov!MBMCRELCode
            ListMaster.SubItems(4) = IIf(Len(MBMCov!MBMBordxDate), MBMCov!MBMBordxDate, lvwInvoice.SelectedItem.SubItems(4))
            ListMaster.SubItems(5) = IIf(Len(MBMCov!MBMPolicyNo), Trim(MBMCov!MBMPolicyNo), "")
            ListMaster.SubItems(6) = lvwInvoice.SelectedItem.SubItems(6)
            ListMaster.SubItems(7) = IIf(Len(MBMCov!MBMBatchNo), MBMCov!MBMBatchNo, lvwInvoice.SelectedItem.SubItems(7))
            ListMaster.SubItems(8) = Trim(MBMCov!MBMCName)
            ListMaster.SubItems(9) = IIf(Len(MBMCov!MBMGroupCompany), Trim(MBMCov!MBMGroupCompany), "")
            If Len(MBMCov!MBMInvoiceNo) Then
                ListMaster.SubItems(10) = "YES"
                ListMaster.SubItems(11) = Trim(MBMCov!MBMInvoiceNo)
            Else
                ListMaster.SubItems(10) = "NO"
                ListMaster.SubItems(11) = ""
            End If
            ListMaster.SubItems(12) = IIf(Len(MBMCov!MBMCPLNCode), Trim(MBMCov!MBMCPLNCode), "")
            ListMaster.SubItems(13) = lvwInvoice.SelectedItem.SubItems(13)
            ListMaster.SubItems(14) = IIf(Len(MBMCov!MBMCEffDate), MBMCov!MBMCEffDate, lvwInvoice.SelectedItem.SubItems(14))
            ListMaster.SubItems(15) = IIf(Len(MBMCov!MBMCExpDate), MBMCov!MBMCExpDate, lvwInvoice.SelectedItem.SubItems(15))
            ListMaster.SubItems(16) = IIf(Len(MBMCov!MBMCAdultChild), Trim(MBMCov!MBMCAdultChild), "")
            ListMaster.SubItems(17) = IIf(IsNumeric(MBMCov!MBMCBasicMCO), MBMCov!MBMCBasicMCO, 0)
            ListMaster.SubItems(18) = IIf(IsNumeric(MBMCov!MBMCMCO), MBMCov!MBMCMCO, 0)

            TotSuppCnt = TotSuppCnt + 1
            lblTotSupp = TotSuppCnt
            cmdINVPrint.Enabled = True
            cmdINVPrintPreview.Enabled = True
        End If
        MBMCov.MoveNext
    Loop

    MBMCov.Close
    Set MBMCov = Nothing

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
            cboINVPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            PAY.MoveNext
        End With
    Loop
    
    PAY.Close
    Set PAY = Nothing
End Sub

Private Sub Form_Load()
Dim rsSearch As New ADODB.Recordset
Dim sql As String
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    Call ComboListing
    
    sql = " Select * From MCOPending Where Status = 'R' "
    rsSearch.Open sql, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
        
    If Not rsSearch.EOF Then
        MsgBox "Released bordx to be billed!", vbInformation + vbOKOnly
    End If
    
    rsSearch.Close
    Set rsSearch = Nothing
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    
End Sub

Private Sub lvwInvoice_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwInvoice.SortKey = ColumnHeader.Index - 1
    lvwInvoice.Sorted = True
    If lvwInvoice.SortOrder = lvwAscending Then
       lvwInvoice.SortOrder = lvwDescending
    Else
       lvwInvoice.SortOrder = lvwAscending
    End If
End Sub

Private Sub UpdateINVPrintingFile(tmpMBMNo As String, TmpCovID As String)
Dim strsql As String
Dim InvoiceDate As String
    
    InvoiceDate = Format(Date, "YYYY/MM/DD")
    If TmpCovID = "00" Then   'Update file for MBMOthers Table
        strsql = " update MBMOthers set " & _
                 " MBMInvoiceDate = '" & InvoiceDate & "', " & _
                 " MBMInvoiceNo = '" & Trim(txtInvoiceNo) & "', " & _
                 " MBMInvoiceBy = '" & Logon & "' " & _
                 " Where MBMNumber = '" & Trim(tmpMBMNo) & "'"
                 medixmasterconn.Execute strsql
    Else    'update file for MBMCoveredPersons Table
        strsql = " update MBMCoveredPersons set " & _
                 " MBMInvoiceDate = '" & InvoiceDate & "', " & _
                 " MBMInvoiceNo = '" & Trim(txtInvoiceNo) & "', " & _
                 " MBMInvoiceBy = '" & Logon & "' " & _
                 " Where MBMCNumber = '" & Trim(tmpMBMNo) & "'" & _
                 " and MBMCCoverID = '" & TmpCovID & "'"
                 medixmasterconn.Execute strsql
    End If
End Sub

Private Sub UpdateMCOPending(MemberNo As String)
Dim sql As String
Dim LastUpdateDate As String

    'Update file for MCO Pending Table
    LastUpdateDate = Format(Date, "YYYY/MM/DD")
    
    sql = " update MCOPending set " & _
          " Status = 'B', " & _
          " LastUpdateDate = '" & LastUpdateDate & "', " & _
          " LastUpdateUser = '" & Logon & "' " & _
          " Where MBMNumber = '" & Trim(MemberNo) & "'"
    medixmasterconnMaint.Execute sql
End Sub

Private Sub IncInvoiceNo()
Dim MBMMCOINVQ As New ADODB.Recordset
Dim strsql As String
Dim InvoiceNo As String
    
    strsql = "Select * from MBMMCOINVQ "
    MBMMCOINVQ.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If MBMMCOINVQ.recordCount = 0 Then
        With MBMMCOINVQ
            .AddNew
                !MCOInvoice_No = "00001"
                InvoiceNo = "00001"
            .Update
        End With
    Else
        With MBMMCOINVQ
            !MCOInvoice_No = Format(!MCOInvoice_No + 1, "00000")
            InvoiceNo = Format(!MCOInvoice_No, "00000")
        .Update
        End With
    End If
    txtInvoiceNo = "S" & InvoiceNo
    MBMMCOINVQ.Close
    Set MBMMCOINVQ = Nothing
End Sub

Private Sub OptByPayor_Click()
    Label1.Visible = False
    CboGrpCompany.Visible = False
    CboGrpCompany.Clear
End Sub

Private Sub OptByGrpCom_Click()
Dim PAYGRP As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strsql As String
Dim strAppendCase As String
    
    Label1.Visible = True
    CboGrpCompany.Visible = True
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   '
    CboGrpCompany.Clear
    
    strAppendCase = "1"
    
    strsql = "AND PAYCode= '" & Mid(cboINVPayor, 1, 2) & "'"
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchComboListing"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("Selection", adVarChar, adParamInput, 255, strAppendCase)
    cmd.Parameters.Append Prm2
    PAYGRP.CursorLocation = adUseClient
    PAYGRP.CursorType = adOpenForwardOnly
    PAYGRP.CacheSize = 100
    Set PAYGRP = cmd.Execute
    
    Do Until PAYGRP.EOF
        With PAYGRP
            CboGrpCompany.AddItem Trim(PAYGRP!MBMGroupCompany)
            PAYGRP.MoveNext
        End With
    Loop
    PAYGRP.Close
    Set PAYGRP = Nothing
    
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing

End Sub

Private Sub txtINVBordxDate_LostFocus()
    If IsDate(txtINVBordxDate) Then
    Else
        If txtINVBordxDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtINVBordxDate.SetFocus
        End If
    End If
End Sub

Private Sub txtINVBordxDateFr_LostFocus()
    If IsDate(txtINVBordxDateFr) Then
    Else
        If txtINVBordxDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtINVBordxDateFr.SetFocus
        End If
    End If
End Sub

Private Sub txtINVBordxDateTo_LostFocus()
    If IsDate(txtINVBordxDateTo) Then
    Else
        If txtINVBordxDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtINVBordxDateTo.SetFocus
        End If
    End If
End Sub

Private Sub txtINVInvoiceNo_LostFocus()
    txtINVInvoiceNo = ChkStr(txtINVInvoiceNo)
End Sub

Private Sub txtMCOBatchNo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub InvCriteriaSel()
    Call cmdINVClear_Click
    lbl30.Caption = ""
    Label4.Visible = False
    txtMCOBatchNo.Visible = False
    txtINVBordxDateFr.Visible = False
    txtINVBordxDateTo.Visible = False
    txtINVInvoiceNo.Visible = False
    cboINVPayor.Enabled = True
    CboGrpCompany.Visible = False
    txtINVBordxDate.Visible = False
    txtINVBordxDateFr.Visible = False
    txtINVBordxDateTo.Visible = False
    cboINVPrintOpt.Enabled = False
    Select Case cboINVCriteriaSel.ListIndex
        Case 0
            lbl29.Caption = "Bordx Date"
            Label4.Visible = True
            txtMCOBatchNo.Visible = True
            txtINVBordxDate.Visible = True
            If OptByGrpCom.value = True Then CboGrpCompany.Visible = True
            cboINVPrintOpt.ListIndex = 0
            cboINVPrintOpt.Enabled = True
        Case 1
            lbl29.Caption = "Bordx From"
            lbl30.Caption = "To"
            txtINVBordxDateFr.Visible = True
            txtINVBordxDateTo.Visible = True
            If OptByGrpCom.value = True Then CboGrpCompany.Visible = True
            cboINVPrintOpt.ListIndex = 0
            cboINVPrintOpt.Enabled = False
        Case 2
            lbl29.Caption = "Invoice Num"
            txtINVInvoiceNo.Visible = True
            cboINVPayor.Enabled = False
            CboGrpCompany.Visible = False
            cboINVPrintOpt.ListIndex = 1
        End Select
End Sub

Private Sub ComboKeyPress(tmpComboBox As ComboBox, KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
'    NewText = LCase(Left(tmpComboBox.Text, tmpComboBox.SelStart) + Chr(KeyAscii) + Mid(tmpComboBox.Text, tmpComboBox.SelStart + tmpComboBox.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To tmpComboBox.ListCount - 1
        If NewText = LCase(Left(tmpComboBox.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = tmpComboBox.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
         If ValidCount = 1 Then
           tmpComboBox.Text = ValidValue
           tmpComboBox.SelStart = 0
           tmpComboBox.SelLength = Len(ValidValue)
         End If
    End If
End Sub


Private Sub PrintOrPreview(SaveFlag As Boolean)
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.WORKSHEET
Dim strsql As String
Dim StrCri As String
Dim strsqlPay As String
Dim PayRec As New ADODB.Recordset
Dim MCO As New ADODB.Recordset
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim TotalPrin As String
Dim SubSupp As String
Dim MemberNo As String
Dim CovID As String
Dim InsDesc As String
Dim CompTrans As Boolean
Dim UPdTrans As Boolean
Dim tmpPln As String
Dim tmpIns As String
Dim TmpEffDate As String
Dim TmpExpDate As String
Dim tmpAdultChild As String
Dim tmpBMco As Double
Dim tmpMCO As Double
Dim Sortlist As String
Dim RecCnt As Integer
Dim tmpPrincMco As Double
Dim AA As Integer
Dim tmpGrpCom As String
    TotalPrin = 0
    SubSupp = 0

    If MultiInvFound = True Then
        MsgBox "Different Invoice Found in Same Batch or Range, Please Search by Invoice No ! ", vbCritical + vbOKOnly
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    
    If objExcel Is Nothing Then
        Set objExcel = CreateObject("Excel.application")
            objExcel.Interactive = False
        If objExcel.Visible = False Then
            objExcel.Visible = True
        End If
        objExcel.WindowState = xlMaximized
'        Set objWorkbook = objExcel.Workbooks.Add
        Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "MCOInvoice\MCOInvoice.xlt")
        
        Set objWorksheet = objWorkbook.Sheets(1)
        objExcel.DisplayAlerts = False
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
    End If
        
    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    CompTrans = False
    UPdTrans = False
    On Error GoTo HANDLE_ERROR
    
    'sql = "SELECT * FROM VIEW_Invoice_ReportSearch2"
    StrCri = ""
    strsql = "SELECT INSCode, MBMAdultChild, MBMBordxDate, MBMBatchNo, MBMInvoiceNo, MBMInvoiceDate, " & _
             "COUNT(*) AS TotalPrincipal, SUM(BillAmt) AS BillAmt, PAYCode "
    Sortlist = " GROUP BY INSCode, MBMAdultChild, MBMBordxDate, MBMBatchNo, MBMInvoiceNo, MBMInvoiceDate, PAYCode "
                  
    If OptByGrpCom.value = True Then
        strsql = strsql & " , MBMGroupCompany "
        Sortlist = Sortlist & " , MBMGroupCompany "
    End If
    Call SelectionCriterial(StrCri)
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    Set cmd1.ActiveConnection = medixmasterconn
    cmd1.CommandText = "SP_MCOFees"
    cmd1.CommandTimeout = 300
    cmd1.CommandType = adCmdStoredProc
    Set Prm1 = cmd1.CreateParameter("SelType", adVarChar, adParamInput, 50, "PrincList")
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("SelFields", adVarChar, adParamInput, 2000, strsql)
    cmd1.Parameters.Append Prm2
    Set Prm3 = cmd1.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, StrCri)
    cmd1.Parameters.Append Prm3
    Set prm4 = cmd1.CreateParameter("SortList", adVarChar, adParamInput, 2000, Sortlist)
    cmd1.Parameters.Append prm4

    MCO.CacheSize = 100
    MCO.CursorLocation = adUseClient
    MCO.CursorType = adOpenForwardOnly
    Set MCO = cmd1.Execute
    RecCnt = 0
    Do While Not MCO.EOF
        RecCnt = RecCnt + 1
        If RecCnt = 1 Then
            strsqlPay = "SELECT PAYCompanyName , PAYAddress1, PAYAddress2, PAYAddress3, " & _
                        "PAYMAINContact, PAYTELno, PAYFAXNo " & _
                        "From PAY WHERE PAYCODE = '" & MCO!PAYCode & "'"
            PayRec.Open strsqlPay, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            Do While Not PayRec.EOF
                objWorksheet.Cells(8, "A") = IIf(Len(PayRec!PAYCompanyName), Trim(PayRec!PAYCompanyName), "")
                objWorksheet.Cells(9, "A") = IIf(Len(PayRec!PAYAddress1), Trim(PayRec!PAYAddress1), "")
                objWorksheet.Cells(10, "A") = IIf(Len(PayRec!PAYAddress2), Trim(PayRec!PAYAddress2), "")
                objWorksheet.Cells(11, "A") = IIf(Len(PayRec!PAYAddress3), Trim(PayRec!PAYAddress3), "")
                objWorksheet.Cells(13, "A") = "ATTENTION: " & " " & IIf(Len(PayRec!PAYMAINContact), Trim(PayRec!PAYMAINContact), "")
                objWorksheet.Cells(15, "A") = "TEL :" & " " & IIf(Len(PayRec!PAYTELno), Trim(PayRec!PAYTELno), "")
                objWorksheet.Cells(15, "C") = "FAX :" & " " & IIf(Len(PayRec!PAYFAXno), Trim(PayRec!PAYFAXno), "")
                PayRec.MoveNext
            Loop
            PayRec.Close
            Set PayRec = Nothing
            If Trim(MCO!MBMInvoiceNo) <> "" Then
                objWorksheet.Cells(8, "H") = Trim(MCO!MBMInvoiceNo)
            End If
            If OptByGrpCom.value = True Then
                objWorksheet.Cells(12, "A") = IIf(Len(MCO!MBMGroupCompany), Trim(MCO!MBMGroupCompany), "")
            End If
            If Len(MCO!MBMInvoiceDate) Then
                objWorksheet.Cells(10, "H") = MCO!MBMInvoiceDate
            Else
                objWorksheet.Cells(10, "H") = Date
            End If
        End If
        InsDesc = ""
        If Trim(MCO!INSCode) = "I" Then
            InsDesc = "I - INDIVIDUAL"
        ElseIf Trim(MCO!INSCode) = "F" Then
            InsDesc = "F - FAMILY"
        ElseIf Trim(MCO!INSCode) = "G" Then
            InsDesc = "G - GROUP INDIVIDUAL"
        ElseIf Trim(MCO!INSCode) = "H" Then
            InsDesc = "H - GROUP FAMILY"
        End If
        objWorksheet.Cells(21 + CDbl(RecCnt), "A") = RecCnt
        objWorksheet.Cells(21 + CDbl(RecCnt), "B") = Format(MCO!MBMBordxDate, "mm/dd/yyyy")
        objWorksheet.Cells(21 + CDbl(RecCnt), "C") = InsDesc
            
        If Trim(MCO!MBMAdultChild) = "A" Then
            objWorksheet.Cells(21 + CDbl(RecCnt), "E") = "Adult"
        ElseIf Trim(MCO!MBMAdultChild) = "C" Then
            objWorksheet.Cells(21 + CDbl(RecCnt), "E") = "Child"
        End If
        objWorksheet.Cells(21 + CDbl(RecCnt), "F") = IIf(IsNumeric(MCO!TotalPrincipal) = True, MCO!TotalPrincipal, 0)
        TotalPrin = CDbl(TotalPrin) + IIf(IsNumeric(MCO!TotalPrincipal) = True, MCO!TotalPrincipal, 0)
        objWorksheet.Cells(21 + CDbl(RecCnt), "H").NumberFormat = "#,##0.00_);(#,##0.00)"
        tmpPrincMco = IIf(IsNumeric(MCO!BillAmt), Format(MCO!BillAmt, "#,##0.00"), 0)
        objWorksheet.Cells(21 + CDbl(RecCnt), "H") = tmpPrincMco
        
        If Trim(MCO!INSCode) = "F" Or Trim(MCO!INSCode) = "H" Then
        tmpGrpCom = ""
            If OptByGrpCom.value = True Then
                tmpGrpCom = IIf(Len(MCO!MBMGroupCompany), Trim(MCO!MBMGroupCompany), "")
            End If
            If Trim(MCO!MBMAdultChild) = "A" Then
                Call GetSuppMCOFee(TotalPrin, tmpPrincMco, MCO!INSCode, MCO!MBMBordxDate, MCO!MBMBatchNo, tmpGrpCom)
                objWorksheet.Cells(21 + CDbl(RecCnt), "G") = GetSubSupp
                objWorksheet.Cells(21 + CDbl(RecCnt), "H") = GetSuppMCO
            Else
                objWorksheet.Cells(21 + CDbl(RecCnt), "G") = IIf(IsNumeric(MCO!TotalPrincipal) = True, MCO!TotalPrincipal, 0)
            End If
        Else
            SubSupp = CDbl(SubSupp) + CDbl(TotalPrin)
            objWorksheet.Cells(21 + CDbl(RecCnt), "G") = SubSupp
        End If
        MCO.MoveNext
        TotalPrin = 0
        SubSupp = 0
    Loop
    MCO.Close
    Set MCO = Nothing
    
    objWorksheet.Range("F39").Formula = "=SUM(F22:F37)"
    objWorksheet.Range("G39").Formula = "=SUM(G22:G37)"
    objWorksheet.Range("H39").Formula = "=SUM(H22:H37)"
    If SaveFlag = True Then

        UPdTrans = True
        medixmasterconn.beginTrans
        medixmasterconnMaint.beginTrans
        If cboINVPrintOpt.ListIndex = 0 Then
            Call IncInvoiceNo
            objWorksheet.Cells(8, "H") = Trim(txtInvoiceNo)
        End If
        
        For AA = 1 To lvwInvoice.listitems.Count
            If Trim(lvwInvoice.listitems(AA).ListSubItems(10).Text) = "NO" Then
                lvwInvoice.listitems(AA).ListSubItems(10).Text = "YES"
                lvwInvoice.listitems(AA).ListSubItems(11).Text = txtInvoiceNo
                MemberNo = Trim(lvwInvoice.listitems(AA).ListSubItems(1).Text)
                CovID = Trim(lvwInvoice.listitems(AA).ListSubItems(2).Text)
                tmpPln = IIf(Len(lvwInvoice.listitems(AA).ListSubItems(12)), Trim(lvwInvoice.listitems(AA).ListSubItems(12)), "")
                tmpIns = IIf(Len(lvwInvoice.listitems(AA).ListSubItems(13)), Trim(lvwInvoice.listitems(AA).ListSubItems(13)), "")
                TmpEffDate = IIf(Len(lvwInvoice.listitems(AA).ListSubItems(14)), lvwInvoice.listitems(AA).ListSubItems(14), "")
                TmpExpDate = IIf(Len(lvwInvoice.listitems(AA).ListSubItems(15)), lvwInvoice.listitems(AA).ListSubItems(15), "")
                tmpAdultChild = IIf(Len(lvwInvoice.listitems(AA).ListSubItems(16)), Trim(lvwInvoice.listitems(AA).ListSubItems(16)), "")
                tmpBMco = CDbl(IIf(IsNumeric(lvwInvoice.listitems(AA).ListSubItems(17)), lvwInvoice.listitems(AA).ListSubItems(17), 0))
                tmpMCO = CDbl(IIf(IsNumeric(lvwInvoice.listitems(AA).ListSubItems(18)), lvwInvoice.listitems(AA).ListSubItems(18), 0))
                Call UpdateINVPrintingFile(MemberNo, CovID)
                Call UpdateMCOPending(MemberNo)
                Call UPdateMBMInvoice(MemberNo, CovID, tmpPln, tmpIns, TmpEffDate, TmpExpDate, tmpAdultChild, tmpBMco, tmpMCO)
            End If
        Next AA
        medixmasterconn.commitTrans
        medixmasterconnMaint.commitTrans
    End If
    CompTrans = True
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    Screen.MousePointer = vbDefault
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
    Exit Sub
HANDLE_ERROR:
        MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
        Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
        Set objWorksheet = Nothing
        Set objWorkbook = Nothing
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        If CompTrans = False And UPdTrans = True Then
            medixmasterconn.RollbackTrans
            medixmasterconnMaint.RollbackTrans
        End If
      '  medixmasterconn.Close
       ' Set medixmasterconn = Nothing
      '  medixmasterconnMaint.Close
      '  Set medixmasterconnMaint = Nothing
        Screen.MousePointer = vbDefault
End Sub

Private Sub SelectionCriterial(tmlSelCriterial As String)
Dim StrReprint As String
Dim strCriterial As String

    If cboINVCriteriaSel.ListIndex = 2 Then
        strCriterial = " AND MBMInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
    ElseIf cboINVCriteriaSel.ListIndex = 0 Then
        strCriterial = " AND (MBMBordxDate = '" & Format(txtINVBordxDate, "mm/dd/yyyy") & "'" & _
                       " AND MBMBatchNo = " & Trim(txtMCOBatchNo) & ")" & _
                       " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "'"
    ElseIf cboINVCriteriaSel.ListIndex = 1 Then
        strCriterial = " AND MBMBordxDate >= '" & Format(txtINVBordxDateFr, "mm/dd/yyyy") & "'" & _
                       " AND MBMBordxDate <= '" & Format(txtINVBordxDateTo, "mm/dd/yyyy") & "'" & _
                       " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "'"
    End If
    If OptByGrpCom.value = True Then
        strCriterial = strCriterial & " AND MBMGroupCompany = '" & RemStr(CboGrpCompany) & "'"
    End If
    If cboINVPrintOpt.ListIndex = 0 Then
        StrReprint = " AND (MBMInvoiceNo IS NULL OR MBMInvoiceNo = '')"
    ElseIf cboINVPrintOpt.ListIndex = 1 Then
        StrReprint = " AND (MBMInvoiceNo IS NOT NULL OR MBMInvoiceNo <> '')"
    End If
    tmlSelCriterial = strCriterial & StrReprint

End Sub

Private Sub GetSuppMCOFee(TotalPrin, MBMCLNMCOFee, tmpInsType, BordxDate, BatchNo, tmpGrp)
Dim cmd As New ADODB.Command
Dim MCO2 As New ADODB.Recordset
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim Prm3 As New ADODB.Parameter
Dim prm4 As New ADODB.Parameter
Dim tmpMCO As Double
Dim tmpPrinc As String
Dim strSelFields As String
Dim strSortlist As String
Dim StrCri As String
    StrCri = ""
    '    strsql = "SELECT * FROM VIEW_Invoice_SuppReportSearch2 "
    strSelFields = "SELECT COUNT(*) AS CQTY,  SUM(BillAmt) AS BillAmt," & _
            " MBMCAdultChild, MBMBordxDate, MBMBatchNo "
            
    'Call SelectionCriterial(StrCri)
    If cboINVCriteriaSel.ListIndex = 2 Then
        StrCri = " AND MBMInvoiceNo = '" & Trim(txtINVInvoiceNo) & "'"
    Else
        StrCri = " AND MBMBordxDate='" & Format(BordxDate, "YYYY/MM/DD") & _
                 "' AND MBMBatchNo='" & Trim(BatchNo) & "'" & _
                 " AND INSCode='" & Trim(tmpInsType) & "' " & _
                 " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "'"
    End If
    strSortlist = " GROUP BY MBMCAdultChild, MBMBordxDate, MBMBatchNo "
    If OptByGrpCom.value = True Then
        StrCri = StrCri & " AND MBMGroupCompany ='" & tmpGrp & "'"
        strSortlist = strSortlist & " , MBMGroupCompany "
    End If
    GetSubSupp = 0
    GetSuppMCO = 0
    tmpMCO = MBMCLNMCOFee
    tmpPrinc = CInt(TotalPrin)
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SP_MCOFees"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("SelType", adVarChar, adParamInput, 50, "SuppList")
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, strSelFields)
    cmd.Parameters.Append Prm2
    Set Prm3 = cmd.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, StrCri)
    cmd.Parameters.Append Prm3
    Set prm4 = cmd.CreateParameter("SortList", adVarChar, adParamInput, 2000, strSortlist)
    cmd.Parameters.Append prm4

    MCO2.CursorLocation = adUseClient
    MCO2.CursorType = adOpenForwardOnly
    MCO2.CacheSize = 100
    Set MCO2 = cmd.Execute
    
    Do While Not MCO2.EOF
        tmpPrinc = tmpPrinc + IIf(IsNumeric(MCO2!CQty) = True, MCO2!CQty, 0)
        tmpMCO = tmpMCO + IIf(IsNumeric(MCO2!BillAmt) = True, MCO2!BillAmt, 0)
        MCO2.MoveNext
    Loop
    MCO2.Close
    Set MCO2 = Nothing
    GetSubSupp = tmpPrinc
    GetSuppMCO = tmpMCO
End Sub

Private Sub UPdateMBMInvoice(MemberNo, CovID, tmpPln, tmpIns, TmpEffDate, TmpExpDate, tmpAdultChild As String, tmpBMco As Double, tmpMCO As Double)
Dim strsql As String
    strsql = "Insert Into MBMMCOInvoice(MBMNumber, MBMCovID, MCOINvNo, MCOInvDate, " & _
             "MCOInvBy, PlanCode, InsCode, MBMEffDate, MBMExpDate, " & _
             "AdultChild, BasicMCO, BillMCOAmt)" & _
             " Values ('" & MemberNo & "','" & CovID & "','" & Trim(txtInvoiceNo) & _
             "','" & Format(Date, "yyyy/mm/dd") & "','" & Logon & "','" & tmpPln & "','" & tmpIns & _
             "','" & Format(TmpEffDate, "yyyy/mm/dd") & "','" & Format(TmpExpDate, "yyyy/mm/dd") & _
             "','" & tmpAdultChild & "'," & tmpBMco & "," & tmpMCO & ")"
    medixmasterconn.Execute strsql
End Sub
