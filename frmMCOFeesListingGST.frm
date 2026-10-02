VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMCOFeesListingGST 
   Caption         =   "MCO Fees (GST)"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   11145
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5280
      Left            =   0
      TabIndex        =   31
      Top             =   1800
      Width           =   11055
      Begin MSComctlLib.ListView lvwMCO 
         Height          =   4890
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8625
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   17
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Payor"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Membership No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "CovID"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Rel"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "BordxDate"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Policy No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Batch"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Group Company"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Invoice No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "PlnCode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "EffDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "ExpDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Adult/Child"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Basic Mco"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Mco"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "GST Fee"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame FrameMCO 
      Height          =   1335
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   11055
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
         Left            =   7560
         TabIndex        =   21
         Top             =   885
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox TxtInvoiceNo 
         Height          =   315
         Left            =   7560
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboMCOCriteriaSel 
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
         ItemData        =   "frmMCOFeesListingGST.frx":0000
         Left            =   2160
         List            =   "frmMCOFeesListingGST.frx":000D
         TabIndex        =   15
         Top             =   600
         Width           =   4215
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
         Left            =   10200
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton OptPayor 
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
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton OptGrpCom 
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
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton OptByDept 
         Caption         =   "By Department"
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
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtMCOBordxDate 
         Height          =   315
         Left            =   7560
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMCOBordxDateTo 
         Height          =   315
         Left            =   9720
         TabIndex        =   17
         Top             =   600
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
         Left            =   7560
         TabIndex        =   18
         Top             =   600
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
         Left            =   9120
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComDlg.CommonDialog dlgCommonDialog_MCO 
         Left            =   6720
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox cboMCOPayor 
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
         ItemData        =   "frmMCOFeesListingGST.frx":0045
         Left            =   2160
         List            =   "frmMCOFeesListingGST.frx":0047
         TabIndex        =   19
         Top             =   885
         Width           =   4215
      End
      Begin VB.ComboBox CboRptType 
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
         ItemData        =   "frmMCOFeesListingGST.frx":0049
         Left            =   2280
         List            =   "frmMCOFeesListingGST.frx":0053
         TabIndex        =   22
         Top             =   1320
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.ComboBox cboMCOListingVer 
         Enabled         =   0   'False
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
         ItemData        =   "frmMCOFeesListingGST.frx":006D
         Left            =   2160
         List            =   "frmMCOFeesListingGST.frx":0077
         TabIndex        =   23
         Top             =   1410
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label Label2 
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
         TabIndex        =   30
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lbl20 
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
         TabIndex        =   29
         Top             =   915
         Width           =   1095
      End
      Begin VB.Label lbl22 
         Caption         =   "Listing Versions"
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
         TabIndex        =   28
         Top             =   1480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label3 
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
         Left            =   6480
         TabIndex        =   27
         Top             =   915
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblDisplay1 
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
         Left            =   6480
         TabIndex        =   26
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblDisplay2 
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
         Left            =   9000
         TabIndex        =   25
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Report Type"
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
         Left            =   480
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   7080
      Width           =   11055
      Begin VB.CommandButton cmdMCOSearch 
         Caption         =   "&Search"
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
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOPrint 
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
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOPrintPreview 
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
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOClear 
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
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOClose 
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
         TabIndex        =   1
         Top             =   240
         Width           =   855
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
         TabIndex        =   9
         Top             =   240
         Width           =   960
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
         TabIndex        =   8
         Top             =   240
         Width           =   1080
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
         TabIndex        =   7
         Top             =   240
         Width           =   1065
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
         TabIndex        =   6
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "MCO Fees (GST)"
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
      TabIndex        =   33
      Top             =   0
      Width           =   11205
   End
End
Attribute VB_Name = "frmMCOFeesListingGST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CancelPrinting As Boolean
Dim bPrintPass As Boolean
Dim MultiInvFound As Boolean
Dim Billed As Boolean
Dim tmpPartyStatus As String
Private m_blnDirection(1 To 11) As Boolean

Private Sub cboMCOPayor_Click()
Dim PAYGRP As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strsql As String
Dim sqlPAY As String
Dim PAY As New ADODB.Recordset
    
    tmpPartyStatus = ""
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    sqlPAY = "Select PAYCode, BillingPartyStatus from PAY where PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "'"
    PAY.Open sqlPAY, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        If PAY.recordCount Then
            tmpPartyStatus = IIf(Len(PAY!BillingPartyStatus), PAY!BillingPartyStatus, "")
        End If
    PAY.Close
    Set PAY = Nothing
    
    If OptGrpCom.Value = True Or OptByDept.Value = True Then
        CboGrpCompany.Clear
        strsql = "AND PAYCode= '" & Mid(cboMCOPayor, 1, 2) & "'"
      
        Set cmd.ActiveConnection = medixmasterconn
        cmd.CommandText = "SearchComboListing"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
        cmd.Parameters.Append Prm1
        Set Prm2 = cmd.CreateParameter("Selection", adVarChar, adParamInput, 255, 1)
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
    End If
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    
End Sub

'Private Sub CboRptType_Click()
'    cboMCOListingVer.Enabled = False
'    If CboRptType.ListIndex = 1 Then
'        cboMCOListingVer.Enabled = True
'    End If
'End Sub

'Private Sub CboRptType_KeyPress(KeyAscii As Integer)
'    Call ComboKeyPress(CboRptType, KeyAscii)
'End Sub

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

Private Sub cboMCOCriteriaSel_Click()
    Call cmdMCOClear_Click
    txtMCOBordxDate.Visible = False
    txtMCOBatchNo.Visible = False
    txtMCOBordxDateFr.Visible = False
    txtMCOBordxDateTo.Visible = False
    TxtInvoiceNo.Visible = False
    lblDisplay2.Visible = True
    If cboMCOCriteriaSel.ListIndex = 0 Or cboMCOCriteriaSel.ListIndex = 3 Then
        lblDisplay1.Caption = "Bordx Date"
        lblDisplay2.Caption = "Batch No"
        txtMCOBordxDate.Visible = True
        txtMCOBatchNo.Visible = True
    ElseIf cboMCOCriteriaSel.ListIndex = 1 Or cboMCOCriteriaSel.ListIndex = 4 Then
        lblDisplay1.Caption = "Bordx Date"
        lblDisplay2.Caption = "To"
        txtMCOBordxDateFr.Visible = True
        txtMCOBordxDateTo.Visible = True
    ElseIf cboMCOCriteriaSel.ListIndex = 2 Or cboMCOCriteriaSel.ListIndex = 5 Then
        lblDisplay1.Caption = "Invoice No"
        lblDisplay2.Visible = False
        TxtInvoiceNo.Visible = True
    End If
End Sub

'Private Sub cboMCOListingVer_KeyPress(KeyAscii As Integer)
'    Call ComboKeyPress(cboMCOListingVer, KeyAscii)
'End Sub

Private Sub cboMCOPayor_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboMCOPayor, KeyAscii)
End Sub

Private Sub cmdMCOClear_Click()
    TxtInvoiceNo = ""
    txtMCOBatchNo = ""
    cboMCOPayor = ""
    CboGrpCompany = ""
    txtMCOBordxDate = "__/__/____"
    txtMCOBordxDateFr = "__/__/____"
    txtMCOBordxDateTo = "__/__/____"
    'cboMCOListingVer.Enabled = False
    'cboMCOListingVer.ListIndex = 0
    lvwMCO.ListItems.Clear
    cmdMCOPrint.Enabled = False
    cmdMCOPrintPreview.Enabled = False
    lblTotSupp = 0
    lblTotPrinc = 0
End Sub

Private Sub cmdMCOClose_Click()
    Unload Me
End Sub

Private Sub cmdMCOPrint_Click()
Dim tempPath As String
    Call PrintPreviewCmd(cmdMCOPrintPreview)
End Sub

Private Sub cmdMCOPrintPreview_Click()
    Dim SqlQry As String
    Dim Dept As String
    Dim Isfrst As Boolean
    Dim IsHeader As Boolean
    Dim Invgenerated As Boolean
    Dim objApp As excel.Application
    Dim objWs As excel.WORKSHEET
    Dim objWb As excel.Workbook
    Dim Rcnt As Integer
    Dim recordcnt As Integer
    Dim ColCnt As Integer
    Dim Parm1 As ADODB.Parameter
    
    Dim tmpPayor As String
    Dim ContactName As String
    Dim PAYName As String
    Dim PayAdd1 As String
    Dim PayAdd2 As String
    Dim PayAdd3 As String
    Dim TelNo As String
    Dim faxno As String
    Dim tmpPrefix As String
    
    Dim tmpPrtGrp As Boolean
    Isfrst = True
    IsHeader = True
    Rcnt = 17
    recordcnt = 1
    ColCnt = 1
    Invgenerated = False
    Dim TotMCOFees As Double
    Dim TotGSTFees As Double
    Dim TotQty As Integer
    Dim TotMBMdep As Integer
    
    TotMCOFees = 0
    TotGSTFees = 0
    TotQty = 0
    TotMBMdep = 0
    
    Dim sql As String
    Dim SelField As String
    
    Dim cmd1 As New ADODB.Command
    Dim Prm1 As ADODB.Parameter
    Dim Prm2 As ADODB.Parameter
    Dim Prm3 As ADODB.Parameter
    Dim rst As New ADODB.Recordset
    
    
    sql = ""
    SelField = "Select PayCode, MBMNumber, MBMCRELCode, MBMBordxDate, MBMPolicyNo, " & _
               "HLTCode, MBMBatchNo, MBMName, MBMGroupCompany, MBMInvoiceNo, " & _
               "PLNCode, INSCode, MBMPayorEffDate, MBMPayorEXPDate, Type, " & _
               "MBMBasicMCO, MBMMCOFee, MBMCRELCode, MBMCCoverID "
    Call SelectionCriterial(sql)
    Set cmd1.ActiveConnection = medixmasterconn
    cmd1.CommandText = "SP_GSTMCOFeesNewPreview"
    cmd1.CommandTimeout = 300
    cmd1.CommandType = adCmdStoredProc
    Set Prm1 = cmd1.CreateParameter("SelFields", adVarChar, adParamInput, 2000, SelField)
    cmd1.Parameters.Append Prm1
    Set Prm2 = cmd1.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, sql)
    cmd1.Parameters.Append Prm2
    Set Prm3 = cmd1.CreateParameter("SortList", adVarChar, adParamInput, 2000, "")
    cmd1.Parameters.Append Prm3
    rst.CacheSize = 100
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    Set rst = cmd1.Execute
    
        Do While Not rst.EOF
            With rst
                If Isfrst Then
                    Dept = rst!MBMDepartment
                    Isfrst = False
                End If
                  If IsHeader Then
                            Set objApp = CreateObject("Excel.Application")
                            objApp.Interactive = False
                            objApp.Visible = False
                            objApp.WindowState = xlMaximized
                            objApp.DisplayAlerts = False
                            Set objWb = objApp.Workbooks.Add("\\svr-super\HISReports\Report\GSTMCOReports\MCOInvoice_DeptGst.xlt")
                            'Set objWb = objApp.Workbooks.Add(crdPath & "\Report\GSTMCOReports\MCOInvoice_DeptGst.xlt")
                            '"C:\Documents and Settings\administrator.MEDIX\Desktop\MCOInvoice\MCOInvoice_DeptGst.xlt"
                            'Dim tmpPayor As String
                            tmpPayor = Mid(cboMCOPayor, 1, 2)
                            Call GetPayorInfo(tmpPayor, PAYName, PayAdd1, PayAdd2, PayAdd3, TelNo, faxno, ContactName, tmpPrefix)
                            
                            Set objWs = objWb.Sheets(1)
                            Set objWs = objWb.Worksheets.Item(1)
                            objWs.Cells(5, "C") = PAYName
                            objWs.Cells(7, "C") = PayAdd1
                            objWs.Cells(8, "C") = PayAdd2
                            objWs.Cells(9, "C") = PayAdd3
                            objWs.Cells(11, "C") = tmpPrtGrp
                            objWs.Cells(13, "C") = ContactName
                            objWs.Cells(14, "C") = TelNo
                            objWs.Cells(14, "E") = faxno
                            
                            objWs.Cells(5, "G") = CStr(recordcnt)
                            objWs.Cells(8, "G") = IIf(Len(rst!MBMInvoiceNo), rst!MBMInvoiceNo, "")
                            objWs.Cells(9, "G") = IIf(Len(rst!MBMInvoiceDate), Format(rst!MBMInvoiceDate, "dd/MM/YYYY"), "")
                            objWs.Cells(10, "G") = "30 Days"
                            IsHeader = False
                            Invgenerated = True
                    End If
                            objWs.Cells(Rcnt, "B") = ColCnt
                            objWs.Cells(Rcnt, "C") = IIf(Len(rst!MBMBordxDate), Format(rst!MBMBordxDate, "DD/MM/YYYY"), "")
                            objWs.Cells(Rcnt, "D") = IIf(Len(rst!MBMBatchNo), rst!MBMBatchNo, "")
                            objWs.Cells(Rcnt, "E") = IIf(Len(rst!EmpType), rst!EmpType, "")
                            If Len(rst!EmpCnt) Then
                                objWs.Cells(Rcnt, "F") = IIf(Len(rst!EmpCnt), rst!EmpCnt, 0)   'TotQty
                                TotQty = TotQty + CInt(rst!EmpCnt)
                            Else
                                objWs.Cells(Rcnt, "F") = 0
                                TotQty = TotQty
                            End If
                            If Len(rst!DepCnt) Then
                                objWs.Cells(Rcnt, "G") = IIf(Len(rst!DepCnt), rst!DepCnt, 0)   'TotMBMdep
                                TotMBMdep = TotMBMdep + CInt(rst!DepCnt)
                            Else
                                objWs.Cells(Rcnt, "G") = IIf(Len(rst!DepCnt), rst!DepCnt, 0)   'TotMBMdep
                                TotMBMdep = TotMBMdep
                            End If
                            If Len(rst!MCO) Then
                                objWs.Cells(Rcnt, "H") = IIf(Len(rst!MCO), rst!MCO, 0)         'TotMCOFees
                                TotMCOFees = TotMCOFees + CDbl(rst!MCO)
                            Else
                                objWs.Cells(Rcnt, "H") = 0        'TotMCOFees
                                TotMCOFees = TotMCOFees
                            End If
                            If Len(rst!GST) Then
                                objWs.Cells(Rcnt, "I") = IIf(Len(rst!GST), rst!GST, 0)         'TotGSTFees
                                TotGSTFees = TotGSTFees + CDbl(rst!GST)
                            Else
                                objWs.Cells(Rcnt, "I") = IIf(Len(rst!GST), rst!GST, 0)         'TotGSTFees
                                TotGSTFees = TotGSTFees
                            End If
                            ColCnt = ColCnt + 1
                            Rcnt = Rcnt + 1
            End With
        rst.MoveNext
        Loop
        rst.Close
        Set rst = Nothing
                    
                    Rcnt = Rcnt + 3
                    objWs.Cells(Rcnt, "E") = "Total"
                    If (TotQty) > 0 Then
                        objWs.Cells(Rcnt, "F") = TotQty    'TotQty
                    Else
                        objWs.Cells(Rcnt, "F") = 0
                    End If
                    If (TotMBMdep) > 0 Then
                        objWs.Cells(Rcnt, "G") = TotMBMdep  'TotMBMdep
                    Else
                        objWs.Cells(Rcnt, "G") = 0   'TotMBMdep
                    End If
                    If (TotMCOFees) > 0 Then
                        objWs.Cells(Rcnt, "H") = TotMCOFees         'TotMCOFees
                    Else
                        objWs.Cells(Rcnt, "H") = 0        'TotMCOFees
                    End If
                    If (TotGSTFees) > 0 Then
                        objWs.Cells(Rcnt, "I") = TotGSTFees         'TotGSTFees
                    Else
                        objWs.Cells(Rcnt, "I") = 0        'TotGSTFees
                    End If
                    
        TotGSTFees = 0
        TotMCOFees = 0
        TotQty = 0
        TotMBMdep = 0
        
        
        Dim LstRs As New ADODB.Recordset
        Dim LstCmd As New ADODB.Command
        Dim LstParm1 As ADODB.Parameter
        Dim LstParm2 As ADODB.Parameter
        Dim LstParm3 As ADODB.Parameter
        Dim IsLsting As Boolean
        Dim GST As Double
        Dim MCO As Double
        
        IsLsting = True
        MCO = 0
        GST = 0
        
         
        Set LstCmd.ActiveConnection = medixmasterconn
        LstCmd.CommandTimeout = 300
        LstCmd.CommandType = adCmdStoredProc
        LstCmd.CommandText = "SP_GSTMCODetailsPreview"
        LstRs.CursorLocation = adUseClient
        LstRs.CursorType = adOpenForwardOnly
        LstRs.CacheSize = 100
        'Set LstParm1 = LstCmd.CreateParameter("SelFields", adVarChar, adParamInput, 2000, SelField)
        'cmd1.Parameters.Append LstParm1
        Set LstParm1 = LstCmd.CreateParameter("SelCriterial", adVarChar, adParamInput, 2000, sql)
        LstCmd.Parameters.Append LstParm1
        'Set LstParm1 = LstCmd.CreateParameter("SortList", adVarChar, adParamInput, 2000, "")
        'cmd1.Parameters.Append LstParm1
        Set LstRs = LstCmd.Execute
        Rcnt = 5
         Do While Not LstRs.EOF
            With LstRs
                    If IsLsting Then
                        Set objWs = objWb.Sheets(2)
                        Set objWs = objWb.Worksheets.Item(2)
                        IsLsting = False
                    End If
                    objWs.Cells(Rcnt, "B") = PAYName 'tmpPayor 'tmpPayor
                    objWs.Cells(Rcnt, "C") = IIf(Len(!MBMNumber), !MBMNumber, "")
                    objWs.Cells(Rcnt, "D") = IIf(Len(!MBMPolicyNo), !MBMPolicyNo, "")
                    objWs.Cells(Rcnt, "E") = IIf(Len(!MBMName), !MBMName, "")
                    objWs.Cells(Rcnt, "F") = IIf(Len(!Relation), !Relation, "")
                    objWs.Cells(Rcnt, "G") = IIf(Len(!ICNumber), !ICNumber, "")
                    objWs.Cells(Rcnt, "I") = IIf(Len(!MBMGroupCompany), !MBMGroupCompany, "")
                    objWs.Cells(Rcnt, "J") = IIf(Len(!Premiume), !Premiume, "")
                    If Len(!MCO) Then
                        objWs.Cells(Rcnt, "K") = IIf(Len(!MCO), !MCO, "") '--MCO
                        MCO = MCO + CDbl(!MCO)
                    Else
                        objWs.Cells(Rcnt, "K") = IIf(Len(!MCO), !MCO, "") '--MCO
                        MCO = MCO + 0
                    End If
                    If Len(!GST) Then
                        objWs.Cells(Rcnt, "L") = IIf(Len(!GST), !GST, "") '--GST
                        GST = GST + CDbl(!GST)
                    Else
                        objWs.Cells(Rcnt, "L") = IIf(Len(!GST), !GST, "") '--GST
                        GST = GST + 0
                    End If
                    
            End With
            Rcnt = Rcnt + 1
            LstRs.MoveNext
         Loop
         Rcnt = Rcnt + 3
            objWs.Cells(Rcnt, "J") = "Total"
            If Len(MCO) Then
                objWs.Cells(Rcnt, "K") = MCO
            Else
                objWs.Cells(Rcnt, "K") = 0
            End If
            If Len(GST) Then
                objWs.Cells(Rcnt, "L") = IIf(Len(GST), GST, "") '--GST
            Else
                objWs.Cells(Rcnt, "L") = 0
            End If
         objWs.Activate
         Dim fPDF As String
         Dim DirPath As String
            If Len(Dir("C:\GSTMCOReports\", vbDirectory)) = 0 Then
               MkDir "C:\GSTMCOReports\"
            End If
         DirPath = "C:\GSTMCOReports\"
         fPDF = DirPath + TxtInvoiceNo + "_" + Replace(CStr(DateTime.Time), ":", "_") + ")" + ".xls"
         objWs.SaveAs fPDF
         objWb.Saved = True
         objWb.Close
         objApp.Quit
         Set objWs = Nothing
         Set objApp = Nothing
        If Invgenerated Then
            MsgBox ("MCO Invoice generated Successfully & Saved in this path 'C:\GSTMCOReports\'. ")
            Shell "C:\WINDOWS\explorer.exe """ & DirPath & "", vbNormalFocus
        Else
            MsgBox ("No Invoice number to Generate Report!")
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
        
    If (cboMCOCriteriaSel.ListIndex = 0) Then
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
    ElseIf (cboMCOCriteriaSel.ListIndex = 1) Then
        If cboMCOPayor = "" Then
            MsgBox "Please Enter Payor!", vbOKOnly + vbCritical, DialogBoxTitle
            cmdMCOSearch.Enabled = True
            Exit Sub
        ElseIf txtMCOBordxDateFr = "__/__/____" Then
            MsgBox "Please Enter Bordereaux Date From!", vbOKOnly + vbCritical, DialogBoxTitle
            cmdMCOSearch.Enabled = True
            Exit Sub
        ElseIf txtMCOBordxDateTo = "__/__/____" Then
            MsgBox "Please Enter Bordereaux Date To!", vbOKOnly + vbCritical, DialogBoxTitle
            cmdMCOSearch.Enabled = True
            Exit Sub
        End If
    ElseIf (cboMCOCriteriaSel.ListIndex = 2) And TxtInvoiceNo = "" Then
            MsgBox "Please Enter Invoice No!", vbOKOnly + vbCritical, DialogBoxTitle
            cmdMCOSearch.Enabled = True
            Exit Sub
    End If
    If OptGrpCom.Value = True And CboGrpCompany = "" Then
        If (cboMCOCriteriaSel.ListIndex <> 2 And cboMCOCriteriaSel.ListIndex <> 5) Then
            MsgBox "Please Enter Group Company!", vbOKOnly + vbCritical, DialogBoxTitle
            cmdMCOSearch.Enabled = True
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    Call MCOListing
    Screen.MousePointer = Default
    cmdMCOSearch.Enabled = True
    
End Sub

Private Sub MCOListing()
Dim rsSearch As New ADODB.Recordset
Dim sql As String
Dim TotPrinc As Long
Dim TotalSupp As Long
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
    Billed = False
    MultiInvFound = False
    PrevInvNO = ""
    lvwMCO.ListItems.Clear
    cmdMCOPrint.Enabled = False
    cmdMCOPrintPreview.Enabled = False
    sql = ""
    SelField = "Select PayCode, MBMNumber, MBMCRELCode, MBMBordxDate, MBMPolicyNo, " & _
               "HLTCode, MBMBatchNo, MBMName, MBMGroupCompany, MBMInvoiceNo, " & _
               "PLNCode, INSCode, MBMPayorEffDate, MBMPayorEXPDate, Type, " & _
               "MBMBasicMCO, MBMMCOFee, MBMCRELCode, MBMCCoverID "
    Call SelectionCriterial(sql)
    Set cmd1.ActiveConnection = medixmasterconn
    cmd1.CommandText = "SP_GSTMCOFeesNew"
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

    Set ListMaster = lvwMCO.ListItems.Add(, , rsSearch!PAYCode)
        ListMaster.SubItems(1) = rsSearch!MBMNumber
        ListMaster.SubItems(2) = rsSearch!CoverID
        ListMaster.SubItems(3) = IIf(Len(rsSearch!Relation), rsSearch!Relation, "")
        ListMaster.SubItems(4) = IIf(Len(rsSearch!MBMBordxDate), Format(rsSearch!MBMBordxDate, "DD/MM/YYYY"), "")
        ListMaster.SubItems(5) = IIf(Len(rsSearch!MBMPolicyNo), Trim(rsSearch!MBMPolicyNo), "") ' need to add
        ListMaster.SubItems(6) = IIf(Len(rsSearch!MBMBatchNo), rsSearch!MBMBatchNo, "")
        ListMaster.SubItems(7) = Trim(rsSearch!MBMName)
        ListMaster.SubItems(8) = IIf(Len(rsSearch!MBMGroupCompany), Trim(rsSearch!MBMGroupCompany), "")
        ListMaster.SubItems(9) = IIf(Len(rsSearch!MBMInvoiceNo), Trim(rsSearch!MBMInvoiceNo), "")
        ListMaster.SubItems(10) = IIf(Len(rsSearch!PLNCode), Trim(rsSearch!PLNCode), "")
        ListMaster.SubItems(11) = IIf(Len(rsSearch!MBMPayorEffDate), rsSearch!MBMPayorEffDate, "")
        ListMaster.SubItems(12) = IIf(Len(rsSearch!MBMPayorEXPDate), rsSearch!MBMPayorEXPDate, "")
        ListMaster.SubItems(13) = IIf(Len(rsSearch!Type), Trim(rsSearch!Type), "")
        ListMaster.SubItems(14) = IIf(IsNumeric(rsSearch!MBMBasicMCO), rsSearch!MBMBasicMCO, 0)
        If Len((rsSearch!MCoFee)) Then
                If CInt(rsSearch!MCoFee) > 0 Then
                    ListMaster.SubItems(15) = Format(rsSearch!MCoFee, "#,#.00")
                Else
                 If CInt(rsSearch!MBMBasicMCO) = 0 Then
                    ListMaster.SubItems(15) = Format(rsSearch!MBMBasicMCO, "#,#.00")
                 Else
                        ListMaster.SubItems(15) = Format(rsSearch!MBMBasicMCO, "#,#.00")
                 End If
                    
                End If
        Else
                'If CInt(rsSearch!MCOFee) > 0 Then
                '    ListMaster.SubItems(15) = Format(rsSearch!MCOFee, "#,#.00")
                'Else
                If IsNumeric(rsSearch!MBMBasicMCO) Then
                 If CInt(rsSearch!MBMBasicMCO) = 0 Then
                    ListMaster.SubItems(15) = Format(rsSearch!MBMBasicMCO, "#,#.00")
                 Else
                        ListMaster.SubItems(15) = Format(rsSearch!MBMBasicMCO, "#,#.00")
                 End If
                End If
                'End If
        
        End If
        
        ListMaster.SubItems(16) = IIf(IsNumeric(rsSearch!MBMGSTAmt), rsSearch!MBMGSTAmt, 0)
        
        If rsSearch!Relation = "P" Then
            TotPrinc = TotPrinc + 1
        Else
           TotalSupp = TotalSupp + 1
           lblTotSupp = TotalSupp
        End If
        lblTotPrinc = TotPrinc
        If Len(rsSearch!MBMInvoiceNo) Then
            Billed = True
            CurrInvNo = Trim(rsSearch!MBMInvoiceNo)
            If PrevInvNO = "" Then
                PrevInvNO = Trim(rsSearch!MBMInvoiceNo)
            ElseIf CurrInvNo <> PrevInvNO Then
                MultiInvFound = True
            End If
        End If
        
        'If Len(rsSearch!MCOFee) Then
        'Else
        '    MCONull = True
        'End If
        
        rsSearch.MoveNext
    Loop
    
    rsSearch.Close
    Set rsSearch = Nothing
    If lvwMCO.ListItems.Count = 0 Then
        MsgBox "Record not found !", vbInformation + vbOKOnly
        cmdMCOPrint.Enabled = False
        Screen.MousePointer = Default
        Exit Sub
    Else
        cmdMCOPrintPreview.Enabled = True
        If Billed = True Then
            cmdMCOPrint.Enabled = False
        Else
            cmdMCOPrint.Enabled = True
        End If
    End If
    If MCONull = True Then
        MsgBox "MCO Fee is Empty!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Sub CrystalRptName(tmpCriSel As Integer)
Dim strFormula As String
Dim strsql As String
Dim rst As New ADODB.Recordset
Dim tmpBatchNo As String
Dim tmpPayCode As String

    tmpBatchNo = lvwMCO.SelectedItem.SubItems(7)
    tmpPayCode = lvwMCO.SelectedItem.Text
    If tmpCriSel = 1 Then ' By Batch
        strFormula = ""
        strFormula = " And MBMBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "'" & _
                     " And PayCode = '" & tmpPayCode & "'" & _
                     " And MBMBatchNo = '" & tmpBatchNo & "' " & _
                     " And ((SUBSTRING(MBMInvoiceNo, 1, 1) = '') or (MBMInvoiceNo is null)) "
    ElseIf tmpCriSel = 2 Then ' By Range
        strFormula = " And MBMBordxDate >= '" & Format(txtMCOBordxDateFr, "mm/dd/yyyy") & "'" & _
                    " And MBMBordxDate <= '" & Format(txtMCOBordxDateTo, "mm/dd/yyyy") & "'" & _
                    " And PayCode = '" & tmpPayCode & "' " & _
                    " And ((SUBSTRING(MBMInvoiceNo, 1, 1) = '') or (MBMInvoiceNo is null)) "
    ElseIf tmpCriSel = 3 Then ' By Invoice
        strFormula = " And MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "'"
    End If
    'Call CrystalRptFormula(tmpCriSel, strFormula)
    Call PrintMCOInvoice(tmpCriSel)
End Sub

Private Sub PrintMCOListing(tmpCriSel As Integer)
Dim PrmString As String
Dim tmpPayCode As String
    
    tmpPayCode = cboMCOPayor
    If tmpCriSel = 1 Then ' By Batch
        PrmString = "BORDEREAUX DATE :        " & txtMCOBordxDate & "      Batch No :    " & txtMCOBatchNo
    ElseIf tmpCriSel = 2 Then 'By Range
        PrmString = "BORDEREAUX DATE :       FROM     " & txtMCOBordxDateFr & "    TO    " & txtMCOBordxDateTo
    ElseIf tmpCriSel = 3 Then 'By Range'invoice
        tmpPayCode = lvwMCO.SelectedItem.Text
        PrmString = "INVOICE NUMBER :           " & TxtInvoiceNo
    End If
    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All2.MBMPolicyNo}"
    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_ALL2.MBMNumber}"
    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL2.MBMCCoverID}"
    cRptMCO.ParameterFields(0) = "PayorName;" & tmpPayCode & ";true"
    cRptMCO.ParameterFields(1) = "ReportType;" & PrmString & ";true"
    If cmdMCOPrint Then
        If bPrintPass = False And CancelPrinting = False Then
            bPrintPass = PrintFunction
            If bPrintPass = False Then
                Screen.MousePointer = Default
                Exit Sub
            End If
        End If
    End If

    cRptMCO.PrintReport
    cRptMCO.PageShow (1)

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
Dim rsSearch As New ADODB.Recordset
Dim sql As String
    
    Billed = False
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    Call ComboListing
    
    sql = " Select * From MCOPending Where Status = 'R' "
    rsSearch.Open sql, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
        
    If Not rsSearch.EOF Then
        MsgBox "Released bordx to be billed!", vbInformation + vbOKOnly
    End If
    
    rsSearch.Close
    Set rsSearch = Nothing
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
End Sub

Private Function PrintFunction() As Boolean
On Error GoTo CancelPrint
    
    If cRptMCO Is Nothing Then Exit Function

    With dlgCommonDialog_MCO
        .DialogTitle = "Print Crystal Report "
        .CancelError = True
        .Flags = cdlPDNoSelection + cdlPDNoPageNums
        .ShowPrinter
        
        If Err <> MSComDlg.cdlCancel Then
            cRptMCO.PrintReport
        End If
    End With
    PrintFunction = True
    Exit Function
CancelPrint:
        PrintFunction = False
        CancelPrinting = True
        Exit Function

End Function



Private Sub lvwMCO_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    Case 4
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lvwMCO, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
    Case 6
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lvwMCO, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lvwMCO, ColumnHeader.Index, ldtString, m_blnDirection(11)
    End Select
End Sub

Private Sub OptByDept_Click()
    Call SearchGrpCompany
End Sub

Private Sub OptPayor_Click()
    Label3.Visible = False
    CboGrpCompany.Visible = False
    CboGrpCompany.Clear
End Sub

Private Sub OptGrpCom_Click()
    Call SearchGrpCompany
End Sub

Private Sub txtInvoiceNo_LostFocus()
    TxtInvoiceNo = ChkStr(TxtInvoiceNo)
End Sub

Private Sub txtMCOBatchNo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
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

Private Sub txtMCOBordxDateFr_LostFocus()
    If IsDate(txtMCOBordxDateFr) Then
    Else
        If txtMCOBordxDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtMCOBordxDateFr.SetFocus
        End If
    End If
End Sub

Private Sub txtMCOBordxDateTo_LostFocus()
    If IsDate(txtMCOBordxDateTo) Then
    Else
        If txtMCOBordxDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtMCOBordxDateTo.SetFocus
        End If
    End If
End Sub

Private Sub SelectionCriterial(tmlSelCriterial As String)
Dim StrReprint As String
Dim strCriterial As String
    If cboMCOCriteriaSel.ListIndex = 2 Then
        strCriterial = " AND MBMInvoiceNo  = '" & Trim(TxtInvoiceNo) & "'"
    ElseIf cboMCOCriteriaSel.ListIndex = 0 Then
       strCriterial = " AND (MBMBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "'" & _
                       " AND MBMBatchNo = " & Trim(txtMCOBatchNo) & ")" & _
                       " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "' " & _
                       " And ((SUBSTRING(MBMInvoiceNo, 1, 1) = '') or (MBMInvoiceNo is null)) "
    ElseIf cboMCOCriteriaSel.ListIndex = 1 Then
        strCriterial = " AND MBMBordxDate >= '" & Format(txtMCOBordxDateFr, "mm/dd/yyyy") & "'" & _
                       " AND MBMBordxDate <= '" & Format(txtMCOBordxDateTo, "mm/dd/yyyy") & "'" & _
                       " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "' " & _
                       " And ((SUBSTRING(MBMInvoiceNo, 1, 1) = '') or (MBMInvoiceNo is null)) "
    End If
    If OptGrpCom.Value = True Then
        strCriterial = strCriterial & " AND MBMGroupCompany = '" & RemStr(CboGrpCompany) & "'"
    End If
    tmlSelCriterial = strCriterial '& StrReprint
End Sub

Public Sub BindExcelwithDept(SqlRS As ADODB.Recordset)

End Sub

Public Sub BindExcel(SqlRS As ADODB.Recordset)

End Sub

Private Sub PrintMCOInvoice(RptType As Integer)
Dim AA As Long
Dim tmpFormula As String
Dim tmpPayor As String
Dim ContactName As String
Dim PAYName As String
Dim PayAdd1 As String
Dim PayAdd2 As String
Dim PayAdd3 As String
Dim TelNo As String
Dim faxno As String
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
Dim strsqlPlutoMCO As String
Dim PlutoMCO As New ADODB.Recordset
Dim PLUTOStatus As Boolean
Dim sqlinsert As String
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim rst As New ADODB.Recordset
Dim strFormula As String
Dim strsql As String
Dim TmpBordxDate As String
Dim tmpPolEffDate As String
Dim tmpPolExpDate As String
Dim tmpInvoiceDate As String
Dim tmpMBMName As String
Dim tmpMBMCName As String
Dim tmpMBMBasicMCO As Double
Dim tmpMBMMCO As Double
Dim tmpMBMMCOFee As Double
Dim tmpPremium As Double
Dim tmpGRPCompany As String
Dim tmpDepartment As String
Dim tmpMBMNumber As String
Dim tmpMBMGSTAmt As Double
    bPrintPass = False
    CancelPrinting = False
    CompTrans = False
    InvoiceExist = False
    If RptType = 1 Then ' By Batch
        strFormula = ""
        strFormula = " And MBMBordxDate = '" & Format(txtMCOBordxDate, "mm/dd/yyyy") & "'" & _
                     " And PayCode = '" & Mid(cboMCOPayor, 1, 2) & "'" & _
                     " And MBMBatchNo = '" & txtMCOBatchNo & "' " & _
                      " And ((SUBSTRING(MBMInvoiceNo, 1, 1) <> 'V') or (MBMInvoiceNo is null)) "
    ElseIf RptType = 2 Then ' By Range
        strFormula = " And MBMBordxDate >= '" & Format(txtMCOBordxDateFr, "mm/dd/yyyy") & "'" & _
                    " And MBMBordxDate <= '" & Format(txtMCOBordxDateTo, "mm/dd/yyyy") & "'" & _
                    " And PayCode = '" & Mid(cboMCOPayor, 1, 2) & "' " & _
                     " And ((SUBSTRING(MBMInvoiceNo, 1, 1) <> 'V') or (MBMInvoiceNo is null)) "
    ElseIf RptType = 3 Then ' By Invoice
        strFormula = " And MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "'"
    End If
    
    For AA = 1 To lvwMCO.ListItems.Count
        tmpPayor = lvwMCO.ListItems(AA).Text
        If Trim(lvwMCO.ListItems(AA).SubItems(9)) <> "" Then
            InvoiceExist = True
        End If
        Exit For
    Next AA
    
    tmpPrtGrp = False
    If OptGrpCom.Value = True Then
        tmpPrtGrp = True
    End If
    tmpPrefix = ""
    If Trim(tmpPartyStatus) = "C" Then
        Call GetThirdPartyInfo(tmpPayor, PAYName, PayAdd1, PayAdd2, PayAdd3, TelNo, faxno, ContactName, tmpPrefix)
    Else
        Call GetPayorInfo(tmpPayor, PAYName, PayAdd1, PayAdd2, PayAdd3, TelNo, faxno, ContactName, tmpPrefix)
    End If
    
    If cmdMCOPrint Then
        If InvoiceExist = False Then
            On Err GoTo ErrUpdate
            medixmasterconn.beginTrans
            medixmasterconnMaint.beginTrans

            Call IncInvoiceNo(tmpPayor, tmpPrefix)
            For AA = 1 To lvwMCO.ListItems.Count
                If Trim(lvwMCO.ListItems(AA).ListSubItems(9).Text) = "" Then
                    lvwMCO.ListItems(AA).ListSubItems(9).Text = TxtInvoiceNo
                    MemberNo = Trim(lvwMCO.ListItems(AA).ListSubItems(1))
                    CovID = Trim(lvwMCO.ListItems(AA).ListSubItems(2))
                    tmpPln = IIf(Len(lvwMCO.ListItems(AA).ListSubItems(10)), Trim(lvwMCO.ListItems(AA).ListSubItems(10)), "")
                    tmpIns = "" 'IIf(Len(lvwMCO.ListItems(AA).ListSubItems(13)), Trim(lvwMCO.ListItems(AA).ListSubItems(13)), "")
                    TmpEffDate = IIf(Len(lvwMCO.ListItems(AA).ListSubItems(11)), lvwMCO.ListItems(AA).ListSubItems(11), "")
                    tmpExpDate = IIf(Len(lvwMCO.ListItems(AA).ListSubItems(12)), lvwMCO.ListItems(AA).ListSubItems(12), "")
                    tmpAdultChild = IIf(Len(lvwMCO.ListItems(AA).ListSubItems(13)), Mid(lvwMCO.ListItems(AA).ListSubItems(13), 1, 1), "")
                    tmpBMCO = CDbl(IIf(IsNumeric(lvwMCO.ListItems(AA).ListSubItems(14)), lvwMCO.ListItems(AA).ListSubItems(14), 0))
                    tmpMCO = CDbl(IIf(IsNumeric(lvwMCO.ListItems(AA).ListSubItems(15)), lvwMCO.ListItems(AA).ListSubItems(15), 0))
                    'tmpMCO = CDbl(IIf(IsNumeric(lvwMCO.ListItems(AA).ListSubItems(16)), lvwMCO.ListItems(AA).ListSubItems(16), 0))
                    Call UpdateINVPrintingFile(MemberNo, CovID, tmpMCO)
                    Call UpdateMCOPending(MemberNo)
                    Call UPdateMBMInvoice(MemberNo, CovID, tmpPln, tmpIns, TmpEffDate, tmpExpDate, tmpAdultChild, tmpBMCO, tmpMCO)
                End If
            Next AA
            medixmasterconn.commitTrans
            medixmasterconnMaint.commitTrans
           CompTrans = True
        End If
    End If
    
    'Directly Call For Invoice From Svr alpha No Pluto
    
    Dim SqlQry As String
    Dim Dept As String
    Dim Isfrst As Boolean
    Dim IsHeader As Boolean
    Dim Invgenerated As Boolean
    Dim objApp As excel.Application
    Dim objWs As excel.WORKSHEET
    Dim objWb As excel.Workbook
    Dim Rcnt As Integer
    Dim recordcnt As Integer
    Dim ColCnt As Integer
    Dim Parm1 As ADODB.Parameter
    Isfrst = True
    IsHeader = True
    Rcnt = 17
    recordcnt = 1
    ColCnt = 1
    Invgenerated = False
    
    Dim TotMCOFees As Double
    Dim TotGSTFees As Double
    Dim TotQty As Integer
    Dim TotMBMdep As Integer
    
    TotMCOFees = 0
    TotGSTFees = 0
    TotQty = 0
    TotMBMdep = 0
    
    If RptType = 1 Or RptType = 3 Then ' By Invoice
    
    Set cmd1.ActiveConnection = medixmasterconn
    cmd1.CommandTimeout = 300
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandText = "SP_Invoice_DeptGST"
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    rst.CacheSize = 100
    Set Parm1 = cmd1.CreateParameter("InvNumber", adVarChar, adParamInput, 200, TxtInvoiceNo)
    cmd1.Parameters.Append Parm1
    Set rst = cmd1.Execute
        Do While Not rst.EOF
            With rst
                If Isfrst Then
                    Dept = rst!MBMDepartment
                    Isfrst = False
                End If
                  If IsHeader Then
                            Set objApp = CreateObject("Excel.Application")
                            objApp.Interactive = False
                            objApp.Visible = False
                            objApp.WindowState = xlMaximized
                            objApp.DisplayAlerts = False
                            Set objWb = objApp.Workbooks.Add(crdPath & "\Report\GSTMCOReports\MCOInvoice_DeptGst.xlt")
                            '"C:\Documents and Settings\administrator.MEDIX\Desktop\MCOInvoice\MCOInvoice_DeptGst.xlt"
                            Set objWs = objWb.Sheets(1)
                            Set objWs = objWb.Worksheets.Item(1)
                            'objApp.Open
                            objWs.Cells(5, "C") = PAYName
                            objWs.Cells(7, "C") = PayAdd1
                            objWs.Cells(8, "C") = PayAdd2
                            objWs.Cells(9, "C") = PayAdd3
                            objWs.Cells(11, "C") = tmpPrtGrp
                            objWs.Cells(13, "C") = ContactName
                            objWs.Cells(14, "C") = TelNo
                            objWs.Cells(14, "E") = faxno
                            
                            objWs.Cells(5, "G") = CStr(recordcnt)
                            objWs.Cells(8, "G") = IIf(Len(rst!MBMInvoiceNo), rst!MBMInvoiceNo, "")
                            objWs.Cells(9, "G") = IIf(Len(rst!MBMInvoiceDate), Format(rst!MBMInvoiceDate, "dd/MM/YYYY"), "")
                            objWs.Cells(10, "G") = "30 Days"
                            IsHeader = False
                            Invgenerated = True
                    End If
                            objWs.Cells(Rcnt, "B") = ColCnt
                            objWs.Cells(Rcnt, "C") = IIf(Len(rst!MBMBordxDate), Format(rst!MBMBordxDate, "DD/MM/YYYY"), "")
                            objWs.Cells(Rcnt, "D") = IIf(Len(rst!MBMBatchNo), rst!MBMBatchNo, "")
                            objWs.Cells(Rcnt, "E") = IIf(Len(rst!EmpType), rst!EmpType, "")
                            If Len(rst!EmpCnt) Then
                                objWs.Cells(Rcnt, "F") = IIf(Len(rst!EmpCnt), rst!EmpCnt, 0)   'TotQty
                                TotQty = TotQty + CInt(rst!EmpCnt)
                            Else
                                objWs.Cells(Rcnt, "F") = 0
                                TotQty = TotQty
                            End If
                            If Len(rst!DepCnt) Then
                                objWs.Cells(Rcnt, "G") = IIf(Len(rst!DepCnt), rst!DepCnt, 0)   'TotMBMdep
                                TotMBMdep = TotMBMdep + CInt(rst!DepCnt)
                            Else
                                objWs.Cells(Rcnt, "G") = IIf(Len(rst!DepCnt), rst!DepCnt, 0)   'TotMBMdep
                                TotMBMdep = TotMBMdep
                            End If
                            If Len(rst!MCO) Then
                                objWs.Cells(Rcnt, "H") = IIf(Len(rst!MCO), rst!MCO, 0)         'TotMCOFees
                                TotMCOFees = TotMCOFees + CDbl(rst!MCO)
                            Else
                                objWs.Cells(Rcnt, "H") = 0        'TotMCOFees
                                TotMCOFees = TotMCOFees
                            End If
                            If Len(rst!GST) Then
                                objWs.Cells(Rcnt, "I") = IIf(Len(rst!GST), rst!GST, 0)         'TotGSTFees
                                TotGSTFees = TotGSTFees + CDbl(rst!GST)
                            Else
                                objWs.Cells(Rcnt, "I") = IIf(Len(rst!GST), rst!GST, 0)         'TotGSTFees
                                TotGSTFees = TotGSTFees
                            End If
                            ColCnt = ColCnt + 1
                            Rcnt = Rcnt + 1
            End With
        rst.MoveNext
        Loop
        rst.Close
        Set rst = Nothing
                    
                    Rcnt = Rcnt + 3
                    objWs.Cells(Rcnt, "E") = "Total"
                    If (TotQty) > 0 Then
                        objWs.Cells(Rcnt, "F") = TotQty    'TotQty
                    Else
                        objWs.Cells(Rcnt, "F") = 0
                    End If
                    If (TotMBMdep) > 0 Then
                        objWs.Cells(Rcnt, "G") = TotMBMdep  'TotMBMdep
                    Else
                        objWs.Cells(Rcnt, "G") = 0   'TotMBMdep
                    End If
                    If (TotMCOFees) > 0 Then
                        objWs.Cells(Rcnt, "H") = TotMCOFees         'TotMCOFees
                    Else
                        objWs.Cells(Rcnt, "H") = 0        'TotMCOFees
                    End If
                    If (TotGSTFees) > 0 Then
                        objWs.Cells(Rcnt, "I") = TotGSTFees         'TotGSTFees
                    Else
                        objWs.Cells(Rcnt, "I") = 0        'TotGSTFees
                    End If
                    
        TotGSTFees = 0
        TotMCOFees = 0
        TotQty = 0
        TotMBMdep = 0
        
        
        Dim LstRs As New ADODB.Recordset
        Dim LstCmd As New ADODB.Command
        Dim LstParm As ADODB.Parameter
        Dim IsLsting As Boolean
        Dim GST As Double
        Dim MCO As Double
        
        IsLsting = True
        MCO = 0
        GST = 0
        
         
        Set LstCmd.ActiveConnection = medixmasterconn
        LstCmd.CommandTimeout = 300
        LstCmd.CommandType = adCmdStoredProc
        LstCmd.CommandText = "SP_Invoice_DeptGSTDetails"
        LstRs.CursorLocation = adUseClient
        LstRs.CursorType = adOpenForwardOnly
        LstRs.CacheSize = 100
        Set LstParm = LstCmd.CreateParameter("InvNumber", adVarChar, adParamInput, 200, TxtInvoiceNo)
        LstCmd.Parameters.Append LstParm
        Set LstRs = LstCmd.Execute
        Rcnt = 5
         Do While Not LstRs.EOF
            With LstRs
                    If IsLsting Then
                        Set objWs = objWb.Sheets(2)
                        Set objWs = objWb.Worksheets.Item(2)
                        IsLsting = False
                    End If
                    objWs.Cells(Rcnt, "B") = PAYName 'tmpPayor 'tmpPayor
                    objWs.Cells(Rcnt, "C") = IIf(Len(!MBMNumber), !MBMNumber, "")
                    objWs.Cells(Rcnt, "D") = IIf(Len(!MBMPolicyNo), !MBMPolicyNo, "")
                    objWs.Cells(Rcnt, "E") = IIf(Len(!MBMName), !MBMName, "")
                    objWs.Cells(Rcnt, "F") = IIf(Len(!Relation), !Relation, "")
                    objWs.Cells(Rcnt, "G") = IIf(Len(!ICNumber), !ICNumber, "")
                    objWs.Cells(Rcnt, "I") = IIf(Len(!MBMGroupCompany), !MBMGroupCompany, "")
                    objWs.Cells(Rcnt, "J") = IIf(Len(!Premiume), !Premiume, "")
                    If Len(!MCO) Then
                        objWs.Cells(Rcnt, "K") = IIf(Len(!MCO), !MCO, "") '--MCO
                        MCO = MCO + CDbl(!MCO)
                    Else
                        objWs.Cells(Rcnt, "K") = IIf(Len(!MCO), !MCO, "") '--MCO
                        MCO = MCO + 0
                    End If
                    If Len(!GST) Then
                        objWs.Cells(Rcnt, "L") = IIf(Len(!GST), !GST, "") '--GST
                        GST = GST + CDbl(!GST)
                    Else
                        objWs.Cells(Rcnt, "L") = IIf(Len(!GST), !GST, "") '--GST
                        GST = GST + 0
                    End If
                    
            End With
            Rcnt = Rcnt + 1
            LstRs.MoveNext
         Loop
         Rcnt = Rcnt + 3
            objWs.Cells(Rcnt, "J") = "Total"
            If Len(MCO) Then
                objWs.Cells(Rcnt, "K") = MCO
            Else
                objWs.Cells(Rcnt, "K") = 0
            End If
            If Len(GST) Then
                objWs.Cells(Rcnt, "L") = IIf(Len(GST), GST, "") '--GST
            Else
                objWs.Cells(Rcnt, "L") = 0
            End If
         objWs.Activate
         Dim fPDF As String
         Dim DirPath As String
            If Len(Dir("C:\GSTMCOReports\", vbDirectory)) = 0 Then
               MkDir "C:\GSTMCOReports\"
            End If
         DirPath = "C:\GSTMCOReports\"
         fPDF = DirPath + TxtInvoiceNo + "_" + Replace(CStr(DateTime.Time), ":", "_") + ")" + ".xls"
         objWs.SaveAs fPDF
         objWb.Saved = True
         objWb.Close
         objApp.Quit
         Set objWs = Nothing
         Set objApp = Nothing
        If Invgenerated Then
            MsgBox ("MCO Invoice generated Successfully & Saved in this path 'C:\GSTMCOReports\'. ")
            Shell "C:\WINDOWS\explorer.exe """ & DirPath & "", vbNormalFocus
        Else
            MsgBox ("No Invoice number to Generate Report!")
        End If
    End If
    CompTrans = True
ErrUpdate:
    If CompTrans = False And InvoiceExist = True Then
        medixmasterconn.RollbackTrans
        medixmasterconnMaint.RollbackTrans
    End If
End Sub
Private Sub GetThirdPartyInfo(tmpPayCode As String, PayorName As String, PayAdd1 As String, _
PayAdd2 As String, PayAdd3 As String, TelNo As String, faxno As String, _
ContactName As String, tmpPrefix As String)
Dim PayRec As New ADODB.Recordset
Dim strsqlPay As String
    
    If Trim(tmpPayCode) <> "" Then
        strsqlPay = "Select BrokerCompanyName, BrokerAddress1, BrokerAddress2, BrokerAddress3, BrokerMAINContact, " & _
                    "BrokerTELno, BrokerFAXno from Broker where PAYCode='" & Trim(tmpPayCode) & "'"
        PayRec.Open strsqlPay, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do While Not PayRec.EOF
            PayorName = IIf(Len(PayRec!BrokerCompanyName), Trim(PayRec!BrokerCompanyName), "")
            PayAdd1 = IIf(Len(PayRec!BrokerAddress1), Trim(PayRec!BrokerAddress1), "")
            PayAdd2 = IIf(Len(PayRec!BrokerAddress2), Trim(PayRec!BrokerAddress2), "")
            PayAdd3 = IIf(Len(PayRec!BrokerAddress3), Trim(PayRec!BrokerAddress3), "")
            TelNo = IIf(Len(PayRec!BrokerTELno), Trim(PayRec!BrokerTELno), "")
            faxno = IIf(Len(PayRec!BrokerFAXno), Trim(PayRec!BrokerFAXno), "")
            ContactName = IIf(Len(PayRec!BrokerMAINContact), Trim(PayRec!BrokerMAINContact), "")
            tmpPrefix = "S"
            PayRec.MoveNext
        Loop
        PayRec.Close
        Set PayRec = Nothing
    End If
End Sub
Private Sub GetPayorInfo(tmpPayCode As String, PayorName As String, PayAdd1 As String, _
PayAdd2 As String, PayAdd3 As String, TelNo As String, faxno As String, _
ContactName As String, tmpPrefix As String)
Dim PayRec As New ADODB.Recordset
Dim strsqlPay As String
    
    If Trim(tmpPayCode) <> "" Then
        strsqlPay = "Select PAYCompanyName, PAYAddress1, PAYAddress2, PAYAddress3, PAYMAINContact, " & _
                    "PAYTELno, PAYFAXno,MCOInvoicePrefix from Pay where PAYCode='" & Trim(tmpPayCode) & "'"
        PayRec.Open strsqlPay, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do While Not PayRec.EOF
            PayorName = IIf(Len(PayRec!PAYCompanyName), Trim(PayRec!PAYCompanyName), "")
            PayAdd1 = IIf(Len(PayRec!PAYAddress1), Trim(PayRec!PAYAddress1), "")
            PayAdd2 = IIf(Len(PayRec!PAYAddress2), Trim(PayRec!PAYAddress2), "")
            PayAdd3 = IIf(Len(PayRec!PAYAddress3), Trim(PayRec!PAYAddress3), "")
            TelNo = IIf(Len(PayRec!PAYTELno), Trim(PayRec!PAYTELno), "")
            faxno = IIf(Len(PayRec!PAYFAXno), Trim(PayRec!PAYFAXno), "")
            ContactName = IIf(Len(PayRec!PAYMAINContact), Trim(PayRec!PAYMAINContact), "")
            tmpPrefix = IIf(Trim(PayRec!MCOInvoicePrefix) <> "", Trim(PayRec!MCOInvoicePrefix), "S")
            PayRec.MoveNext
        Loop
        PayRec.Close
        Set PayRec = Nothing
    End If
End Sub

Private Sub PrintPreviewCmd(cboType As Boolean)
Dim RptType As Integer
    If MultiInvFound = True Then
        MsgBox "Different Invoice Found in Same Batch or Range, Please Search by Invoice No ! ", vbCritical + vbOKOnly
        Exit Sub
    Else
        Screen.MousePointer = vbHourglass
        RptType = 1
         Call PrintMCOInvoice(RptType)
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub ComboKeyPress(tmpComboBox As ComboBox, KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(tmpComboBox.Text, tmpComboBox.SelStart) + Chr(KeyAscii) + Mid(tmpComboBox.Text, tmpComboBox.SelStart + tmpComboBox.SelLength + 1))
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

Private Sub IncInvoiceNo(tmpPayor As String, tmpPrefix As String)
Dim MBMMCOINVQ As New ADODB.Recordset
Dim strsql As String
Dim InvoiceNo As String


'If Len(tmpPayor) Then
'    strsql = "Select * from MBMMCOINVQ where PAYCode='" & tmpPayor & "'and Prefix='" & tmpPrefix & "'"
'Else
'    strsql = "Select * from MBMMCOINVQ where Prefix='" & tmpPrefix & "'"
'End If

    If tmpPrefix = "S" Then
        strsql = "Select * from MBMMCOINVQ "
    Else
        strsql = "Select * from MBMMCOINVQ where PAYCode='" & tmpPayor & "'"
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

Private Sub UPdateMBMInvoice(MemberNo, CovID, tmpPln, tmpIns, TmpEffDate, tmpExpDate, tmpAdultChild As String, tmpBMCO As Double, tmpMCO As Double)
Dim strsql As String
Dim tmpMBMGSTAmt As Double
Dim tmpINVDate As String

    tmpMBMGSTAmt = 0
    tmpINVDate = Format(GetServerTime, "yyyy/mm/dd")
    'tmpINVDate = "01/04/2015"
    If CDate(tmpINVDate) >= "01/04/2015" Then
        If IsNumeric(tmpMCO) Then
            If tmpMCO <> 0 Then
                tmpMBMGSTAmt = tmpMCO * 0.06
            End If
        End If
    End If
    
    strsql = "Insert Into MBMMCOInvoice(MBMNumber, MBMCovID, MCOINvNo, MCOInvDate, " & _
             "MCOInvBy, PlanCode, InsCode, MBMEffDate, MBMExpDate, " & _
             "AdultChild, BasicMCO, BillMCOAmt,GSTAmt)" & _
             " Values ('" & MemberNo & "','" & CovID & "','" & Trim(TxtInvoiceNo) & _
             "','" & Format(GetServerTime, "yyyy/mm/dd") & "','" & Logon & "','" & tmpPln & "','" & tmpIns & _
             "','" & Format(TmpEffDate, "yyyy/mm/dd") & "','" & Format(tmpExpDate, "yyyy/mm/dd") & _
             "','" & tmpAdultChild & "'," & tmpBMCO & "," & tmpMCO & "," & tmpMBMGSTAmt & ")"
    medixmasterconn.Execute strsql
End Sub

Private Sub UpdateINVPrintingFile(tmpMBMNo As String, TmpCovID As String, tmpMCO As Double)
Dim strsql As String
Dim InvoiceDate As String
Dim tmpMBMGSTAmt As Double
    'MBMGSTAmt
    
    tmpMBMGSTAmt = 0
    InvoiceDate = Format(GetServerTime, "YYYY/MM/DD")
    'InvoiceDate = "01/04/2015"
    If CDate(InvoiceDate) >= "01/04/2015" Then
        If IsNumeric(tmpMCO) Then
            If tmpMCO <> 0 Then
                tmpMBMGSTAmt = tmpMCO * 0.06
            End If
        End If
    End If
    
    If TmpCovID = "00" Then   'Update file for MBMOthers Table
        strsql = " update MBMOthers set " & _
                 " MBMInvoiceDate = '" & InvoiceDate & "', " & _
                 " MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "', " & _
                 " MBMInvoiceBy = '" & Logon & "', " & _
                 " MBMGSTAmt = '" & tmpMBMGSTAmt & "' " & _
                 " Where MBMNumber = '" & Trim(tmpMBMNo) & "'"
                 medixmasterconn.Execute strsql
    Else    'update file for MBMCoveredPersons Table
        strsql = " update MBMCoveredPersons set " & _
                 " MBMInvoiceDate = '" & InvoiceDate & "', " & _
                 " MBMInvoiceNo = '" & Trim(TxtInvoiceNo) & "', " & _
                 " MBMInvoiceBy = '" & Logon & "', " & _
                 " MBMCGSTAmt = '" & tmpMBMGSTAmt & "' " & _
                 " Where MBMCNumber = '" & Trim(tmpMBMNo) & "'" & _
                 " and MBMCCoverID = '" & TmpCovID & "'"
                 medixmasterconn.Execute strsql
    End If
End Sub
Private Sub UpdateMCOPending(MemberNo As String)
Dim sql As String
Dim LastUpdateDate As String

    'Update file for MCO Pending Table
    LastUpdateDate = Format(GetServerTime, "YYYY/MM/DD")
    
    sql = " update MCOPending set " & _
          " Status = 'B', " & _
          " LastUpdateDate = '" & LastUpdateDate & "', " & _
          " LastUpdateUser = '" & Logon & "' " & _
          " Where MBMNumber = '" & Trim(MemberNo) & "'"
    medixmasterconnMaint.Execute sql
End Sub

Private Sub CrystalRptFormula(tmpCriSel As Integer, strFormula As String)
Dim tmpBatchNo As String
Dim tmpPayCode As String
    tmpBatchNo = lvwMCO.SelectedItem.SubItems(7)
    tmpPayCode = lvwMCO.SelectedItem.Text
    strFormula = ""
    'If tmpCriSel = 1 Then ' By Batch
    '    strFormula = " {VIEW_MCO_LISTING_All2.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
    '                 " {VIEW_MCO_LISTING_All2.PayCode} = '" & tmpPayCode & "'" & " And " & _
    '                 " {VIEW_MCO_LISTING_All2.MBMBatchNo} = " & tmpBatchNo & ""
    'ElseIf tmpCriSel = 2 Then ' By Range
    '    strFormula = " {VIEW_MCO_LISTING_All2.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
    '                " {VIEW_MCO_LISTING_All2.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & " AND " & _
    '                " {VIEW_MCO_LISTING_All2.PayCode} = '" & tmpPayCode & "'"
    'ElseIf tmpCriSel = 3 Then ' By Invoice
    '    strFormula = " {VIEW_MCO_LISTING_All2.MBMInvoiceNo} = '" & Trim(TxtInvoiceNo) & "'"
    'End If
    
    'If cboMCOCriteriaSel.ListIndex >= 3 Then
    '    strFormula = strFormula & " AND {VIEW_MCO_LISTING_All2.MBMCRelCode} = 'P' "
    'End If
    'If OptGrpCom.Value = True Then
    '    strFormula = strFormula & " AND {VIEW_MCO_LISTING_All2.MBMGroupCompany} = '" & RemStr(CboGrpCompany) & "'"
    'End If
End Sub

Private Sub SearchGrpCompany()
Dim PAYGRP As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim strsql As String
    
    Label3.Visible = True
    CboGrpCompany.Visible = True
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    CboGrpCompany.Clear
    
    strsql = "AND PAYCode= '" & Mid(cboMCOPayor, 1, 2) & "'"
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchComboListing"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strsql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("Selection", adVarChar, adParamInput, 255, 1)
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








