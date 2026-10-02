VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMCOFeesListing 
   Caption         =   "Fee Listing"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   11205
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameMCO 
      Caption         =   "Enter Criteria For Searching"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   11055
      Begin VB.ComboBox cboMCOCriteriaSel 
         Height          =   315
         ItemData        =   "frmMCOFeesListing.frx":0000
         Left            =   2160
         List            =   "frmMCOFeesListing.frx":000A
         TabIndex        =   0
         Top             =   360
         Width           =   4215
      End
      Begin MSMask.MaskEdBox txtMCOBordxDate 
         Height          =   315
         Left            =   7560
         TabIndex        =   9
         Top             =   285
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
         Left            =   7080
         Top             =   1080
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
         Left            =   7560
         TabIndex        =   10
         Top             =   560
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMCOBordxDateFr 
         Height          =   315
         Left            =   7560
         TabIndex        =   11
         Top             =   285
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
         Left            =   7560
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox cboMCOPayor 
         Height          =   315
         Left            =   2160
         TabIndex        =   8
         Top             =   650
         Width           =   4215
      End
      Begin VB.ComboBox cboMCOListingVer 
         Height          =   315
         ItemData        =   "frmMCOFeesListing.frx":0032
         Left            =   2160
         List            =   "frmMCOFeesListing.frx":0042
         TabIndex        =   7
         Top             =   930
         Width           =   4215
      End
      Begin VB.ComboBox cboMCOPrintOpt 
         Height          =   315
         ItemData        =   "frmMCOFeesListing.frx":00C8
         Left            =   2160
         List            =   "frmMCOFeesListing.frx":00D5
         TabIndex        =   6
         Top             =   1220
         Width           =   1695
      End
      Begin VB.ComboBox cboMCOPrintSeq 
         Height          =   315
         ItemData        =   "frmMCOFeesListing.frx":00F8
         Left            =   4800
         List            =   "frmMCOFeesListing.frx":0108
         TabIndex        =   5
         Top             =   1220
         Width           =   1575
      End
      Begin VB.TextBox txtMCOBatchNo 
         Height          =   315
         Left            =   7560
         TabIndex        =   4
         Top             =   560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Single/Batch/Range"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lbl26 
         Caption         =   "Sequence"
         Height          =   255
         Left            =   3960
         TabIndex        =   20
         Top             =   1280
         Width           =   735
      End
      Begin VB.Label lbl20 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   680
         Width           =   1095
      End
      Begin VB.Label lbl22 
         Caption         =   "Listing Versions"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbl21 
         Caption         =   "Print Option"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1280
         Width           =   855
      End
      Begin VB.Label lbl19 
         Caption         =   "To"
         Height          =   255
         Left            =   6720
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lbl18 
         Caption         =   "From"
         Height          =   255
         Left            =   6720
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbl15 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   6600
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Batch No"
         Height          =   255
         Left            =   6600
         TabIndex        =   13
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lbl16 
         Caption         =   "Date "
         Height          =   255
         Left            =   6960
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Fee Listing"
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
         TabIndex        =   2
         Top             =   60
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   7200
      Width           =   11055
      Begin VB.CommandButton cmdMCOSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   300
         Left            =   6720
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOPrint 
         Caption         =   "&Print"
         Height          =   300
         Left            =   7560
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOPrintPreview 
         Caption         =   "Pre&view"
         Height          =   300
         Left            =   8400
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOClear 
         Caption         =   "&Clear"
         Height          =   300
         Left            =   9240
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOClose 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   300
         Left            =   10080
         TabIndex        =   25
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
         TabIndex        =   33
         Top             =   240
         Width           =   960
      End
      Begin VB.Label txtMCOTotalSupplementary1 
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
         TabIndex        =   32
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label txtMCOTotalPrincipal 
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
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Listing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5040
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Width           =   11055
      Begin MSComctlLib.ListView lvwMCO 
         Height          =   4650
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8202
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Payor Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Membership No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cover ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Relationship"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Bordereaux Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Policy No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Health Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Batch No"
            Object.Width           =   2540
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
            SubItemIndex    =   10
            Text            =   "Printed"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmMCOFeesListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const LEAF_PAYCODE = 0
Const LEAF_MEMBERNO = 1
Const LEAF_COVERID = 2
Const LEAF_RELATION = 3
Const LEAF_BORDXDATE = 4
Const LEAF_POLICYNO = 5
Const LEAF_HLTCODE = 6
Const LEAF_BATCHNO = 7
Const LEAF_NAME = 8
Const LEAF_GRPCO = 9
Const LEAF_PRINTED = 10

Const MCO_PAYCODE = 0
Const MCO_MEMBERNO = 1
Const MCO_COVERID = 2
Const MCO_RELATION = 3
Const MCO_BORDXDATE = 4
Const MCO_POLICYNO = 5
Const MCO_HLTCODE = 6
Const MCO_BATCHNO = 7
Const MCO_NAME = 8
Const MCO_GRPCO = 9
Const MCO_PRINTED = 10

Const INV_PAYCODE = 0
Const INV_MEMBERNO = 1
Const INV_COVERID = 2
Const INV_RELATION = 3
Const INV_BORDXDATE = 4
Const INV_POLICYNO = 5
Const INV_HLTCODE = 6
Const INV_BATCHNO = 7
Const INV_NAME = 8
Const INV_GRPCO = 9
Const INV_PRINTED = 10
Const INV_INVOICE = 11

Const CRD_PAYCODE = 0
Const CRD_MEMBERNO = 1
Const CRD_COVERID = 2
Const CRD_RELATION = 3
Const CRD_BORDXDATE = 4
Const CRD_POLICYNO = 5
Const CRD_HLTCODE = 6
Const CRD_BATCHNO = 7
Const CRD_NAME = 8
Const CRD_GRPCO = 9
Const CRD_PRINTED = 10

Dim SDate As String
Dim DateFrom As String
Dim DateTo As String
Dim Exportsql As String
Dim PAYCode As String
Dim BordxDat As String
Dim MBMNo As String
Dim Printed As String
Dim CancelPrinting As Boolean
Dim bPrintPass As Boolean
Dim TotalSupplementary As Integer
Dim INVNo As String

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
Select Case cboMCOCriteriaSel.ListIndex
    Case 0
        lbl15.Visible = True
        lbl16.Visible = True
        'lbl17.Visible = False
        lbl18.Visible = False
        Label4.Visible = True
        lbl19.Visible = False
        txtMCOBordxDate.Visible = True
        txtMCOBatchNo.Visible = True
        txtMCOBordxDateFr.Visible = False
        txtMCOBordxDateTo.Visible = False
        txtMCOBatchNo = ""
        cboMCOPayor = ""
        txtMCOBordxDate = "__/__/____"
        txtMCOBordxDateFr = "__/__/____"
        txtMCOBordxDateTo = "__/__/____"
        cboMCOPrintOpt.ListIndex = 0
        cboMCOPrintSeq.ListIndex = 0
        cboMCOListingVer.ListIndex = 0
        lvwMCO.listitems.Clear
        cmdMCOPrint.Enabled = False
        cmdMCOPrintPreview.Enabled = False
        txtMCOTotalSupplementary1 = 0
        txtMCOTotalPrincipal = 0
    
    Case 1
        lbl15.Visible = False
        lbl16.Visible = False
        'lbl17.Visible = True
        lbl18.Visible = True
        lbl19.Visible = True
        Label4.Visible = False
        txtMCOBordxDate.Visible = False
        txtMCOBatchNo.Visible = False
        txtMCOBordxDateFr.Visible = True
        txtMCOBordxDateTo.Visible = True
        cboMCOPayor = ""
        txtMCOBordxDate = "__/__/____"
        txtMCOBordxDateFr = "__/__/____"
        txtMCOBordxDateTo = "__/__/____"
        cboMCOPrintOpt.ListIndex = 0
        cboMCOPrintSeq.ListIndex = 0
        cboMCOListingVer.ListIndex = 0
        lvwMCO.listitems.Clear
        cmdMCOPrint.Enabled = False
        cmdMCOPrintPreview.Enabled = False
        txtMCOTotalSupplementary1 = 0
        txtMCOTotalPrincipal = 0
    End Select

End Sub

Private Sub cboMCOListingVer_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboMCOListingVer.Text, cboMCOListingVer.SelStart) + Chr(KeyAscii) + Mid(cboMCOListingVer.Text, cboMCOListingVer.SelStart + cboMCOListingVer.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboMCOListingVer.ListCount - 1
        
        If NewText = LCase(Left(cboMCOListingVer.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboMCOListingVer.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboMCOListingVer.Text = ValidValue
               cboMCOListingVer.SelStart = 0
               cboMCOListingVer.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboMCOPayor_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboMCOPayor.Text, cboMCOPayor.SelStart) + Chr(KeyAscii) + Mid(cboMCOPayor.Text, cboMCOPayor.SelStart + cboMCOPayor.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboMCOPayor.ListCount - 1
        
        If NewText = LCase(Left(cboMCOPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboMCOPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboMCOPayor.Text = ValidValue
               cboMCOPayor.SelStart = 0
               cboMCOPayor.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboMCOPrintOpt_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboMCOPrintOpt.Text, cboMCOPrintOpt.SelStart) + Chr(KeyAscii) + Mid(cboMCOPrintOpt.Text, cboMCOPrintOpt.SelStart + cboMCOPrintOpt.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboMCOPrintOpt.ListCount - 1
        
        If NewText = LCase(Left(cboMCOPrintOpt.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboMCOPrintOpt.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboMCOPrintOpt.Text = ValidValue
               cboMCOPrintOpt.SelStart = 0
               cboMCOPrintOpt.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboMCOPrintSeq_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboMCOPrintSeq.Text, cboMCOPrintSeq.SelStart) + Chr(KeyAscii) + Mid(cboMCOPrintSeq.Text, cboMCOPrintSeq.SelStart + cboMCOPrintSeq.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboMCOPrintSeq.ListCount - 1
        
        If NewText = LCase(Left(cboMCOPrintSeq.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboMCOPrintSeq.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboMCOPrintSeq.Text = ValidValue
               cboMCOPrintSeq.SelStart = 0
               cboMCOPrintSeq.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cmdMCOClear_Click()
    txtMCOBatchNo = ""
    cboMCOPayor = ""
    txtMCOBordxDate = "__/__/____"
    txtMCOBordxDateFr = "__/__/____"
    txtMCOBordxDateTo = "__/__/____"
    cboMCOPrintOpt.ListIndex = 0
    cboMCOPrintSeq.ListIndex = 0
    cboMCOListingVer.ListIndex = 0
    lvwMCO.listitems.Clear
    cmdMCOPrint.Enabled = False
    cmdMCOPrintPreview.Enabled = False
    txtMCOTotalSupplementary1 = 0
    txtMCOTotalPrincipal = 0
End Sub

Private Sub cmdMCOClose_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub cmdMCOPrint_Click()
    If cboMCOPrintOpt.ListIndex = 0 Then
        MCOPrint "NO"
    ElseIf cboMCOPrintOpt.ListIndex = 1 Then
        MCOPrint "YES"
    ElseIf cboMCOPrintOpt.ListIndex = 2 Then
        MCOPrint "ALL"
    End If
End Sub

Private Sub cmdMCOPrintPreview_Click()
    If cboMCOPrintOpt.ListIndex = 0 Then
        MCOPrint "NO"
    ElseIf cboMCOPrintOpt.ListIndex = 1 Then
        MCOPrint "YES"
    ElseIf cboMCOPrintOpt.ListIndex = 2 Then
        MCOPrint "ALL"
    End If

End Sub

Private Sub cmdMCOSearch_Click()

    If cboMCOCriteriaSel = "" Then
            MsgBox "Please Enter MCO Criteria!", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If
    
    If cboMCOCriteriaSel.ListIndex = 0 And txtMCOBordxDate = "__/__/____" And txtMCOBatchNo = "" And cboMCOPayor = "" Then
        MsgBox "Please Enter Bordereaux Date, Batch No and Payor !", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        If cboMCOCriteriaSel.ListIndex = 1 And cboMCOPayor = "" And txtMCOBordxDateFr = "__/__/____" And txtMCOBordxDateTo = "__/__/____" Then
            MsgBox "Please Enter Bordereaux Range and Payor!", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        End If
    End If
    
    If cboMCOPrintOpt = "" Then
        MsgBox "Please Enter Print Option !", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        If cboMCOPrintSeq = "" Then
            MsgBox "Please Enter Sequence !", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        Else
            If cboMCOListingVer = "" Then
                MsgBox "Please Enter Leaflet Version !", vbOKOnly + vbCritical, DialogBoxTitle
                Exit Sub
            End If
        End If
    End If
    
    
    'Select Case SSTab1.Tab
    '   Case 1  'MCO Listing
            RecordSearching_MCO lvwMCO
    'End Select
End Sub

Private Sub RecordSearching_MCO(objListView As Object)
'On Error Resume Next
    Screen.MousePointer = vbHourglass
    Call MCOPrincipalListing
    Screen.MousePointer = Default
End Sub

Private Sub MCOPrincipalListing()
Dim rsSearch As New ADODB.Recordset
Dim rsPrint As New ADODB.Recordset
Dim ColumnItem As ListItem
Dim sql As String
Dim itemToPrint As Integer
Dim isPrinted As Boolean
Dim CoverQty As Integer
Dim MCOCriteria As String
Dim TotalPrincipal As Integer

TotalSupplementary = 0
Select Case cboMCOCriteriaSel.ListIndex
    Case 0 'Criteria Selection for MCO on Bordx Batch
            SDate = CStr(Month(txtMCOBordxDate)) & "/" & CStr(Day(txtMCOBordxDate)) & "/" & CStr(Year(txtMCOBordxDate))
            lvwMCO.listitems.Clear
            
            sql = "SELECT * FROM VIEW_MCO_ReportSearch"
            sql = sql & " WHERE MBMBordxDate = '" & SDate & "'"
            sql = sql & " AND MBMBatchNo = '" & txtMCOBatchNo & "'"
            sql = sql & " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "'"
            
            Select Case cboMCOPrintSeq.ListIndex
                Case 0 'MCO Print Seq for Group Company
                    sql = sql & " ORDER BY MBMGroupCompany, MBMNumber"
                    MCOCriteria = " ORDER BY MBMGroupCompany, MBMNumber"
                    rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
                Case 1 'MCO Print Seq for Membership No
                    sql = sql & " ORDER BY MBMNumber"
                    rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
                Case 2 'MCO Print Seq for Policy No
                    sql = sql & " ORDER BY MBMPolicyNo"
                    rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
                Case 3 'MCO Print Seq for Members Name
                    sql = sql & " ORDER BY MBMName"
                    rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
            End Select
     
    Case 1 'Criteria Selection for MCO on Bordx Range
            DateFrom = CStr(Month(txtMCOBordxDateFr)) & "/" & CStr(Day(txtMCOBordxDateFr)) & "/" & CStr(Year(txtMCOBordxDateFr))
            DateTo = CStr(Month(txtMCOBordxDateTo)) & "/" & CStr(Day(txtMCOBordxDateTo)) & "/" & CStr(Year(txtMCOBordxDateTo))
            lvwMCO.listitems.Clear
            
            sql = "SELECT * FROM VIEW_MCO_ReportSearch "
            sql = sql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
            sql = sql & " AND MBMBordxDate <= '" & DateTo & "'"
            sql = sql & " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "'"
            
            Select Case cboMCOPrintSeq.ListIndex
                Case 0 'MCO Print Seq for Group Company
                    sql = sql & " ORDER BY MBMGroupCompany, MBMNumber"
                    rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
                Case 1 'MCO Print Seq for Membership No
                    sql = sql & " ORDER BY MBMNumber"
                    rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
                Case 2 'MCO Print Seq for Policy No
                    sql = sql & " ORDER BY MBMPolicyNo"
                    rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
                Case 3 'MCO Print Seq for Members Name
                    sql = sql & " ORDER BY MBMName"
                    rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
            End Select
End Select
'End Select
    If rsSearch.recordCount = 0 Then
        MsgBox "Record not found !", vbInformation + vbOKOnly
        cmdMCOPrint.Enabled = False
        Screen.MousePointer = Default
        Exit Sub
    End If
              
    'TotalPrincipal = 0
    Do Until rsSearch.EOF
    Total = rsSearch.recordCount
    lblTotal = CInt(Total)
        Set ListMaster = lvwMCO.listitems.Add(, , rsSearch!PAYCode)
            ListMaster.SubItems(1) = rsSearch!MBMNumber
            ListMaster.SubItems(2) = "00"
            ListMaster.SubItems(3) = rsSearch!MBMMemberType
            ListMaster.SubItems(4) = rsSearch!MBMBordxDate
            ListMaster.SubItems(5) = rsSearch!MBMPolicyNo
            ListMaster.SubItems(6) = Trim(rsSearch!HLTCode)
            If Len(rsSearch!MBMBatchNo) Then
                ListMaster.SubItems(7) = rsSearch!MBMBatchNo
            End If
            ListMaster.SubItems(8) = Trim(rsSearch!MBMName)
                                
            If Len(Trim(rsSearch!MBMGroupCompany)) = 0 Or IsNull(rsSearch!MBMGroupCompany) = True Then
                ListMaster.SubItems(9) = ""
            Else
                ListMaster.SubItems(9) = rsSearch!MBMGroupCompany
            End If
                            
                TotalPrincipal = TotalPrincipal + 1
                txtMCOTotalPrincipal = TotalPrincipal
                PAYCode = rsSearch!PAYCode
                MBMNo = rsSearch!MBMNumber
                BordxDat = rsSearch!MBMBordxDate
                'Select Case SSTab1.Tab
                '    Case 1 'MCO Listing Printing
                        itemToPrint = MCO_PRINTED
                        cmdMCOPrint.Enabled = True
                        cmdMCOPrintPreview.Enabled = True
                '    End Select
                    
                    If rsSearch.recordCount > 0 Then
                '        Select Case SSTab1.Tab
                '            Case 1 'MCO Listing Printing
                                If Len(rsSearch!MBMMCOListPrinted) Then
                                    isPrinted = rsSearch!MBMMCOListPrinted
                                Else
                                    isPrinted = False
                                End If
                '        End Select
                                                
                        If isPrinted = True Then
                            ListMaster.SubItems(itemToPrint) = "YES"
                        Else
                            ListMaster.SubItems(itemToPrint) = "NO"
                        End If
                        Printed = Trim(ListMaster.SubItems(itemToPrint))
                    Else
                        ListMaster.SubItems(itemToPrint) = "NO"
                        Printed = ListMaster.SubItems(itemToPrint)
                    End If
                                                
                    If Trim(rsSearch!INSCode) = "F" Or Trim(rsSearch!INSCode) = "H" Then
                        Call MCOSuppListing
                    End If
                    
                    
                    
                    
    rsSearch.MoveNext
    Loop
    'LLcurrent = lvwLeaflet.listitems.count
    'LLtotal = lvwLeaflet.listitems.count
    rsSearch.Close
    Set rsSearch = Nothing

End Sub

Private Sub MCOSuppListing()
'On Error Resume Next
Dim strsql As String
Dim sql As String
Dim strsqls As String
Dim rsPrint As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim ListMaster As ListItem
Dim isPrinted As Boolean
Dim Total As Integer
Dim Current As Integer

'Select Case SSTab1.Tab
'    Case 1
        Select Case cboMCOCriteriaSel.ListIndex
            Case 0
                    SDate = CStr(Month(txtMCOBordxDate)) & "/" & CStr(Day(txtMCOBordxDate)) & "/" & CStr(Year(txtMCOBordxDate))
                    strsql = "SELECT * FROM VIEW_Card_Export_Search "
                    strsql = strsql & " WHERE MBMBordxDate = '" & SDate & "'"
                    strsql = strsql & " AND MBMBatchNo = '" & Trim(txtMCOBatchNo) & "'"
                    strsql = strsql & " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    
                    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            Case 1
                    DateFrom = CStr(Month(txtMCOBordxDateFr)) & "/" & CStr(Day(txtMCOBordxDateFr)) & "/" & CStr(Year(txtMCOBordxDateFr))
                    DateTo = CStr(Month(txtMCOBordxDateTo)) & "/" & CStr(Day(txtMCOBordxDateTo)) & "/" & CStr(Year(txtMCOBordxDateTo))
                    strsql = "SELECT * FROM VIEW_Card_Export_Search "
                    strsql = strsql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
                    strsql = strsql & " AND MBMBordxDate <= '" & DateTo & "'"
                    strsql = strsql & " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    
                    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        End Select
'End Select
        
               Do Until MBMCov.EOF
                   Set ListMaster = lvwMCO.listitems.Add(, , PAYCode)
                        ListMaster.SubItems(1) = MBMCov!MBMCNumber
                        ListMaster.SubItems(2) = MBMCov!MBMCCoverID
                        ListMaster.SubItems(3) = MBMCov!MBMCRELCode
                        ListMaster.SubItems(4) = lvwMCO.SelectedItem.SubItems(4)
                        ListMaster.SubItems(5) = lvwMCO.SelectedItem.SubItems(5)
                        ListMaster.SubItems(6) = lvwMCO.SelectedItem.SubItems(6)
                        ListMaster.SubItems(7) = lvwMCO.SelectedItem.SubItems(7)
                        ListMaster.SubItems(8) = MBMCov!MBMCName
                    
                       If MBMCov!MBMGroupCompany = "" Then
                            ListMaster.SubItems(9) = ""
                        Else
                            If Len(MBMCov!MBMGroupCompany) Then
                                ListMaster.SubItems(9) = MBMCov!MBMGroupCompany
                            End If
                        End If
                    
                       TotalSupplementary = TotalSupplementary + 1
                       txtMCOTotalSupplementary1 = TotalSupplementary
                       sql = "SELECT * FROM MBMPrinting"
                       BordxDat = CStr(Month(MBMCov!MBMBordxDate)) & "/" & CStr(Day(MBMCov!MBMBordxDate)) & "/" & CStr(Year(MBMCov!MBMBordxDate))
                  
                       sql = sql & " WHERE MBMBordxDate = '" & BordxDat & "'"
                       sql = sql & " AND PAYCode = '" & PAYCode & "'"
                       sql = sql & " AND HLTCode = '" & lvwMCO.SelectedItem.SubItems(6) & "'"
                       sql = sql & " AND MBMBatchno = '" & lvwMCO.SelectedItem.SubItems(7) & "'"
                       sql = sql & " AND MBMNumber = '" & MBMCov!MBMCNumber & "'"
                       
                       rsPrint.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
                       'Select Case SSTab1.Tab
                       'Case 1 'Leaflet Printing
                           itemToPrint = MCO_PRINTED
                           cmdMCOPrint.Enabled = True
                           cmdMCOPrintPreview.Enabled = True
                       'End Select
                       
                       If Printed = "YES" Then
                           ListMaster.SubItems(itemToPrint) = "YES"
                       Else
                           ListMaster.SubItems(itemToPrint) = "NO"
                       End If
                       
                       rsPrint.Close
                       MBMCov.MoveNext
                   Loop
            
                MBMCov.Close
                Set rsPrint = Nothing
                Set MBMCov = Nothing

End Sub

Private Sub MCOPrint(PrintMode As String)
Dim ii As Integer
Dim MemberNo As String

Dim tmpBDate, tmpBatchNo, TMPPayCode, tmpHLTCode, tmpMemberNo As String
Dim tmpPrintedBDate, tmpPrintedBatchNo, tmpPrintedPaycode, tempPrintedHLTcode, tmpPrintedMemberNo As String
Dim BDateArray() As String
Dim BatchNoArray() As String
Dim PaycodeArray() As String
Dim HLTcodeArray() As String
Dim MemberNoArray() As String

Dim cnt As Integer
Dim cont As Integer
Dim cout As Integer
Dim Arraycnt As Integer
Dim Arraycount As Integer
Dim PrintM As String
Dim i As Integer
Dim A As Integer
Dim B As Integer
Dim rstInv As New ADODB.Recordset
Dim strsql2 As String
Dim INVNo As String
Dim InvNoFr, InvNoTo As String

 Screen.MousePointer = vbHourglass
Select Case cboMCOCriteriaSel.ListIndex
    Case 0 'MCO Criteria Selection for Bordx Batch
            cnt = 0
            If PrintMode = "ALL" Then
                For ii = 1 To lvwMCO.listitems.count
                    Set lvwMCO.SelectedItem = lvwMCO.listitems(ii)
                        tmpMemberNo = lvwMCO.SelectedItem.SubItems(MCO_MEMBERNO)
                        tmpBDate = lvwMCO.SelectedItem.SubItems(MCO_BORDXDATE)
                        tmpHLTCode = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                        tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                        TMPPayCode = lvwMCO.SelectedItem
                        
                        If tmpBDate = tmpPrintedBDate And tmpBatchNo = tmpPrintedBatchNo And TMPPayCode = tmpPrintedPaycode And _
                           tmpHLTCode = tempPrintedHLTcode And tmpMemberNo = tmpPrintedMemberNo Then
                        Else
                           tmpPrintedBDate = tmpBDate
                           tmpPrintedBatchNo = tmpBatchNo
                           tmpPrintedPaycode = TMPPayCode
                           tempPrintedHLTcode = tmpHLTCode
                           tmpPrintedMemberNo = tmpMemberNo
                           cnt = cnt + 1
                           ReDim Preserve BDateArray(cnt)
                           ReDim Preserve BatchNoArray(cnt)
                           ReDim Preserve PaycodeArray(cnt)
                           ReDim Preserve HLTcodeArray(cnt)
                           ReDim Preserve MemberNoArray(cnt)
                           BDateArray(cnt) = tmpBDate
                           BatchNoArray(cnt) = tmpBatchNo
                           PaycodeArray(cnt) = TMPPayCode
                           HLTcodeArray(cnt) = tmpHLTCode
                           MemberNoArray(cnt) = tmpMemberNo
                        End If
                    Next ii
            Else
               For ii = 1 To lvwMCO.listitems.count
                    Set lvwMCO.SelectedItem = lvwMCO.listitems(ii)   'lvwLeaflet.ListItems(ii)
                    If lvwMCO.SelectedItem.SubItems(MCO_PRINTED) = PrintMode Then
                        tmpMemberNo = lvwMCO.SelectedItem.SubItems(MCO_MEMBERNO)
                        tmpBDate = lvwMCO.SelectedItem.SubItems(MCO_BORDXDATE)
                        tmpHLTCode = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                        tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                        TMPPayCode = lvwMCO.SelectedItem
                        
                        If tmpBDate = tmpPrintedBDate And tmpBatchNo = tmpPrintedBatchNo And TMPPayCode = tmpPrintedPaycode And _
                            tmpHLTCode = tempPrintedHLTcode And tmpMemberNo = tmpPrintedMemberNo Then
                        Else
                            tmpPrintedBDate = tmpBDate
                            tmpPrintedBatchNo = tmpBatchNo
                            tmpPrintedPaycode = TMPPayCode
                            tempPrintedHLTcode = tmpHLTCode
                            tmpPrintedMemberNo = tmpMemberNo
                            cnt = cnt + 1
                            ReDim Preserve BDateArray(cnt)
                            ReDim Preserve BatchNoArray(cnt)
                            ReDim Preserve PaycodeArray(cnt)
                            ReDim Preserve HLTcodeArray(cnt)
                            ReDim Preserve MemberNoArray(cnt)
                            BDateArray(cnt) = tmpBDate
                            BatchNoArray(cnt) = tmpBatchNo
                            PaycodeArray(cnt) = TMPPayCode
                            HLTcodeArray(cnt) = tmpHLTCode
                            MemberNoArray(cnt) = tmpMemberNo
                        End If
                    End If
               Next ii
            End If
    
    Case 1 'MCO Criteria Selection for Bordx Range
            cnt = 0
            If PrintMode = "ALL" Then
                    For ii = 1 To lvwMCO.listitems.count
                        Set lvwMCO.SelectedItem = lvwMCO.listitems(ii)
                        tmpMemberNo = lvwMCO.SelectedItem.SubItems(MCO_MEMBERNO)
                        tmpBDate = lvwMCO.SelectedItem.SubItems(MCO_BORDXDATE)
                        tmpHLTCode = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                        tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                        TMPPayCode = lvwMCO.SelectedItem
                        
                        If tmpBDate = tmpPrintedBDate And tmpBatchNo = tmpPrintedBatchNo And TMPPayCode = tmpPrintedPaycode And _
                           tmpHLTCode = tempPrintedHLTcode And tmpMemberNo = tmpPrintedMemberNo Then
                        Else
                           tmpPrintedBDate = tmpBDate
                           tmpPrintedBatchNo = tmpBatchNo
                           tmpPrintedPaycode = TMPPayCode
                           tempPrintedHLTcode = tmpHLTCode
                           tmpPrintedMemberNo = tmpMemberNo
                           cnt = cnt + 1
                           ReDim Preserve BDateArray(cnt)
                           ReDim Preserve BatchNoArray(cnt)
                           ReDim Preserve PaycodeArray(cnt)
                           ReDim Preserve HLTcodeArray(cnt)
                           ReDim Preserve MemberNoArray(cnt)
                           BDateArray(cnt) = tmpBDate
                           BatchNoArray(cnt) = tmpBatchNo
                           PaycodeArray(cnt) = TMPPayCode
                           HLTcodeArray(cnt) = tmpHLTCode
                           MemberNoArray(cnt) = tmpMemberNo
                        End If
                    Next ii
            Else
                  For ii = 1 To lvwMCO.listitems.count
                        Set lvwMCO.SelectedItem = lvwMCO.listitems(ii)   'lvwLeaflet.ListItems(ii)
                        If lvwMCO.SelectedItem.SubItems(MCO_PRINTED) = PrintMode Then
                            tmpMemberNo = lvwMCO.SelectedItem.SubItems(MCO_MEMBERNO)
                            tmpBDate = lvwMCO.SelectedItem.SubItems(MCO_BORDXDATE)
                            tmpHLTCode = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                            TMPPayCode = lvwMCO.SelectedItem
                            
                            If tmpBDate = tmpPrintedBDate And tmpBatchNo = tmpPrintedBatchNo And TMPPayCode = tmpPrintedPaycode And _
                                tmpHLTCode = tempPrintedHLTcode And tmpMemberNo = tmpPrintedMemberNo Then
                            Else
                                tmpPrintedBDate = tmpBDate
                                tmpPrintedBatchNo = tmpBatchNo
                                tmpPrintedPaycode = TMPPayCode
                                tempPrintedHLTcode = tmpHLTCode
                                tmpPrintedMemberNo = tmpMemberNo
                                cnt = cnt + 1
                                ReDim Preserve BDateArray(cnt)
                                ReDim Preserve BatchNoArray(cnt)
                                ReDim Preserve PaycodeArray(cnt)
                                ReDim Preserve HLTcodeArray(cnt)
                                ReDim Preserve MemberNoArray(cnt)
                                BDateArray(cnt) = tmpBDate
                                BatchNoArray(cnt) = tmpBatchNo
                                PaycodeArray(cnt) = TMPPayCode
                                HLTcodeArray(cnt) = tmpHLTCode
                                MemberNoArray(cnt) = tmpMemberNo
                            End If
                        End If
                   Next ii
            End If
End Select


If cmdMCOPrintPreview Then
    Call MCOListingVer
Else
    If cmdMCOPrint Then
            A = 0
            B = 0
            For i = 1 To lvwMCO.listitems.count
                Set lvwMCO.SelectedItem = lvwMCO.listitems(i)   'lvwMCO.ListItems(ii)
                    PrintM = lvwMCO.SelectedItem.SubItems(10)
                    If PrintM = "YES" Then
                        A = A + 1
                    Else
                        B = B + 1
                    End If
            Next i
        
        Select Case cboMCOCriteriaSel.ListIndex
            Case 0 'MCO Criteria Selection for Bordx Batch Print
                    If (A = 0 And B <> 0) And cboMCOPrintOpt.ListIndex = 1 Then
                        MsgBox "Please select 'Print New' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
                    Else
                        If (A <> 0 And B = 0) And cboMCOPrintOpt.ListIndex = 0 Then
                            MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
                        Else
                            Call MCOListingVer
                            
                            If bPrintPass = True And CancelPrinting = False Then
                                For ii = 1 To UBound(BDateArray)
                                    UpdateMCOPrintingFile BDateArray(ii), BatchNoArray(ii), PaycodeArray(ii), HLTcodeArray(ii), MemberNoArray(ii), "MCO"
                                Next ii
                                Call cmdMCOSearch_Click
                            End If
                         End If
                    End If
        
            
            Case 1 'MCO Criteria Selection for Bordx Range Print
                    If (A = 0 And B <> 0) And cboMCOPrintOpt.ListIndex = 1 Then
                        MsgBox "Please select 'Print New' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
                    Else
                        If (A <> 0 And B = 0) And cboMCOPrintOpt.ListIndex = 0 Then
                            MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
                        Else
                            Call MCOListingVer
                        End If
                    End If
        End Select
    End If
End If
Screen.MousePointer = Default
End Sub

Private Sub MCOListingVer()
'On Error Resume Next
Dim i As Integer
Dim rsSearch As New ADODB.Recordset

bPrintPass = False
CancelPrinting = False
If cmdMCOPrintPreview Then
    Select Case cboMCOCriteriaSel.ListIndex
        Case 0 ' MCO Criteria Selection for Bordx Batch Print Preview button
                Call MCOListingVer_BBPrintPrev
            
        Case 1 ' MCO Criteria Selection for Bordx Range Print Preview
                Call MCOListingVer_BRPrintPreV
    End Select
Else
    If cmdMCOPrint Then
        Select Case cboMCOCriteriaSel.ListIndex
            Case 0 'MCO Criteria Selection for Bordx Batch Print button
                    Call MCOListingVer_BBPrint
            
            Case 1 ' MCO Criteia Selection for Bordx Range
                    Call MCOListingVer_BRPrint
           
        End Select
    End If
End If

End Sub

Private Sub MCOListingVer_BBPrintPrev()
Select Case cboMCOPrintOpt.ListIndex
    Case 0  'MCO Print Option for Print New
            Select Case cboMCOListingVer
                Case "MCO Listing"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                         
                    ss = "{VIEW_MCO_LISTING.MBMBordxDate} = '" & SDate & "'"
                    'cRptMCO.Connect = "Server=HIS;UID=sa;Pwd;DB=HISDB"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WithoutPremium.rpt"
                    'cRptMCO.ReportFileName = App.path & "\Report\MCOListing_WithoutPremium.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING.PayCode} = '" & Mid(cboMCOPayor, 1, 2) & "'" & " And " & _
                                               " {VIEW_MCO_LISTING.MBMBatchNo} = " & txtMCOBatchNo & ""
                                                            
                    Call MCOListingBBPrintSeqPrev
                             
                Case "MCO Listing - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                         
                    ss = "{VIEW_MCO_LISTING.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WP_GrpCo_SubTotal.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING.MBMBatchNo} = " & tmpBatchNo & ""
                                 
                    Call MCOListingBBPrintSeqPrev
                                 
                Case "MCO Listing With Premium"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                         
                    ss = "{VIEW_MCO_LISTING.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING.MBMBatchNo} = " & tmpBatchNo & ""
                                 
                    Call MCOListingBBPrintSeqPrev
                                 
                Case "MCO Listing With Premium - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                         
                    ss = "{VIEW_MCO_LISTING.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_GrpCo_SubTotal.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING.MBMBatchNo} = " & tmpBatchNo & ""
                    Call MCOListingBBPrintSeqPrev
                                 
            End Select
    
    Case 1 'MCO Print Option for Reprint
            Select Case cboMCOListingVer
                Case "MCO Listing"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                        
                    ss = "{VIEW_MCO_Listing_Printed.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WithoutPremium_Printed.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_Listing_Printed.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_Listing_Printed.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_Listing_Printed.MBMBatchNo} = " & tmpBatchNo & ""
                                
                    Call MCOListingBBPrintSeqPrev
                            
                Case "MCO Listing - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                        
                    ss = "{VIEW_MCO_Listing_Printed.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WP_GrpCo_SubTotal_Printed.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_Listing_Printed.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_Listing_Printed.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_Listing_Printed.MBMBatchNo} = " & tmpBatchNo & ""
                                
                    Call MCOListingBBPrintSeqPrev
                                
                Case "MCO Listing With Premium"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                        
                    ss = "{VIEW_MCO_Listing_Printed.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_Printed.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_Listing_Printed.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_Listing_Printed.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_Listing_Printed.MBMBatchNo} = " & tmpBatchNo & ""
                                
                    Call MCOListingBBPrintSeqPrev
                                
                Case "MCO Listing With Premium - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                        
                    ss = "{VIEW_MCO_Listing_Printed.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_GrpCo_SubTotal_Printed.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_Listing_Printed.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_Listing_Printed.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_Listing_Printed.MBMBatchNo} = " & tmpBatchNo & ""
                    Call MCOListingBBPrintSeqPrev
                                
            End Select
    
    Case 2 'MCO Print Option for Print All
            Select Case cboMCOListingVer
                Case "MCO Listing"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                         
                    ss = "{VIEW_MCO_LISTING_All.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WithoutPremium_All.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_All.PayCode} = '" & Mid(cboMCOPayor, 1, 2) & "'" & " And " & _
                                               " {VIEW_MCO_LISTING_All.MBMBatchNo} = " & txtMCOBatchNo & ""
                                                            
                    Call MCOListingBBPrintSeqPrev
                             
                Case "MCO Listing - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                         
                    ss = "{VIEW_MCO_LISTING_All.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WP_GrpCo_SubTotal_All.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_All.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING_All.MBMBatchNo} = " & tmpBatchNo & ""
                                 
                    Call MCOListingBBPrintSeqPrev
                                 
                Case "MCO Listing With Premium"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                         
                    ss = "{VIEW_MCO_LISTING_All.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_All.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_All.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING_All.MBMBatchNo} = " & tmpBatchNo & ""
                                 
                    Call MCOListingBBPrintSeqPrev
                                 
                Case "MCO Listing With Premium - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                         
                    ss = "{VIEW_MCO_LISTING_All.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_GrpCo_SubTotal_All.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_All.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING_All.MBMBatchNo} = " & tmpBatchNo & ""
                    Call MCOListingBBPrintSeqPrev
                                 
            End Select
End Select
End Sub

Private Sub MCOListingVer_BRPrintPreV()
Select Case cboMCOPrintOpt.ListIndex
    Case 0 ' MCO Print Option for Print New
            Select Case cboMCOListingVer
                Case "MCO Listing"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                                            
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WithoutPremium.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_LISTING.PayCode} = '" & TMPPayCode & "'"
                                                               
                    Call MCOListingBRPrintSeqPrev
                                        
                Case "MCO Listing - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                    
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WP_GrpCo_SubTotal.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_LISTING.PayCode} = '" & TMPPayCode & "'"
                                    
                   Call MCOListingBRPrintSeqPrev
                                    
                Case "MCO Listing With Premium"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                    
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_LISTING.PayCode} = '" & TMPPayCode & "'"
                    Call MCOListingBRPrintSeqPrev
                                    
                Case "MCO Listing With Premium - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                                    
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_GrpCo_SubTotal.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_LISTING.PayCode} = '" & TMPPayCode & "'"
                                    
                    Call MCOListingBRPrintSeqPrev
                                    
            End Select
                    
    Case 1 'MCO Print Option for Reprint
            Select Case cboMCOListingVer
                Case "MCO Listing"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                                        
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WithoutPremium_Printed.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_Listing_Printed.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_Listing_Printed.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_Listing_Printed.PayCode} = '" & TMPPayCode & "'"
                    'cRptMCO.SelectionFormula = " {VIEW_MCO_Listing_Printed.MBMBordxDate} >= '" & Format(txtMCOBordxDate, "dd/mm/yyyy") And "' & _
                    '                           " {VIEW_MCO_Listing_Printed.MBMBordxDate} <= '" & Format(txtMCOBordxDateTo, "dd/mm/yyyy") AND "' & _
                    '                           " {VIEW_MCO_Listing_Printed.PayCode} = '" & TMPPayCode & "'"
                                                           
                    Call MCOListingBRPrintSeqPrev
                                    
                Case "MCO Listing - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                                        
                    ss = "{VIEW_MCO_Listing_Printed.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WP_GrpCo_SubTotal_Printed.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_Listing_Printed.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_Listing_Printed.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_Listing_Printed.PayCode} = '" & TMPPayCode & "'"
                    Call MCOListingBRPrintSeqPrev
                                
                Case "MCO Listing With Premium"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                                        
                    ss = "{VIEW_MCO_Listing_Printed.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_Printed.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_Listing_Printed.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_Listing_Printed.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_Listing_Printed.PayCode} = '" & TMPPayCode & "'"
                    Call MCOListingBRPrintSeqPrev
                                
                Case "MCO Listing With Premium - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                                        
                    ss = "{VIEW_MCO_Listing_Printed.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_GrpCo_SubTotal_Printed.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_Listing_Printed.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_Listing_Printed.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_Listing_Printed.PayCode} = '" & TMPPayCode & "'"
                    Call MCOListingBRPrintSeqPrev
                                
            End Select
        

    Case 2 'MCO Print Option for Print All
            Select Case cboMCOListingVer
                Case "MCO Listing"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                                                
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WithoutPremium_All.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_Listing_All.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_Listing_All.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_Listing_All.PayCode} = '" & TMPPayCode & "'"
                                                                   
                    Call MCOListingBRPrintSeqPrev
                                            
                Case "MCO Listing - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                                                
                    ss = "{VIEW_MCO_Listing.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WP_GrpCo_SubTotal_All.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_Listing_All.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_Listing_All.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_Listing_All.PayCode} = '" & TMPPayCode & "'"
                    Call MCOListingBRPrintSeqPrev
                                        
                Case "MCO Listing With Premium"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                                                
                    ss = "{VIEW_MCO_Listing_All.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_All.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_Listing_All.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_Listing_All.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_Listing_All.PayCode} = '" & TMPPayCode & "'"
                    Call MCOListingBRPrintSeqPrev
                                        
                Case "MCO Listing With Premium - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                                                
                    ss = "{VIEW_MCO_Listing_All.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_GrpCo_SubTotal_All.rpt"
                    cRptMCO.Destination = crptToWindow
                    cRptMCO.WindowState = crptMaximized
                    cRptMCO.SelectionFormula = " {VIEW_MCO_Listing_All.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_Listing_All.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_Listing_All.PayCode} = '" & TMPPayCode & "'"
                    Call MCOListingBRPrintSeqPrev
                                        
            End Select
End Select
End Sub
Private Sub MCOListingVer_BBPrint()
Select Case cboMCOPrintOpt.ListIndex
    Case 0  ' MCO Print Option for Print New
            Select Case cboMCOListingVer
                Case "MCO Listing"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                                
                    ss = "{VIEW_MCO_LISTING.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WithoutPremium.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING.MBMBatchNo} = " & tmpBatchNo & ""
                                        
                    Call MCOListingBBPrintSeq
                                            
                Case "MCO Listing - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                                
                    ss = "{VIEW_MCO_LISTING.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WP_GrpCo_SubTotal.rpt"
                                         
                    Call MCOListingBBPrintSeq
                    
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING.MBMBatchNo} = " & tmpBatchNo & ""
                                            
                Case "MCO Listing With Premium"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                                
                    ss = "{VIEW_MCO_LISTING.MBMBordxDate} = '" & SDate & "'"
                                        
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING.MBMBatchNo} = " & tmpBatchNo & ""
                                            
                    Call MCOListingBBPrintSeq
                                            
                Case "MCO Listing With Premium - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                    
                    ss = "{VIEW_MCO_LISTING.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_GrpCo_SubTotal.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING.MBMBatchNo} = " & tmpBatchNo & ""
                    Call MCOListingBBPrintSeq
                                            
            End Select
                        
    Case 1 'MCO Print Option for Reprint
            Select Case cboMCOListingVer
                Case "MCO Listing"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                                    
                    ss = "{VIEW_MCO_LISTING_Printed.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WithoutPremium_Printed.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_Printed.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_Printed.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING_Printed.MBMBatchNo} = " & tmpBatchNo & ""
                                            
                    Call MCOListingBBPrintSeq
                                        
                Case "MCO Listing - Group Company Sub-Total"
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                                    
                    ss = "{VIEW_MCO_LISTING_Printed.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WP_GrpCo_SubTotal_Printed.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_Printed.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_Printed.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING_Printed.MBMBatchNo} = " & tmpBatchNo & ""
                                            
                    Call MCOListingBBPrintSeq
                                            
                Case "MCO Listing With Premium"
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                                    
                    ss = "{VIEW_MCO_LISTING_Printed.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_Printed.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_Printed.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_Printed.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING_Printed.MBMBatchNo} = " & tmpBatchNo & ""
                                            
                    Call MCOListingBBPrintSeq
                                            
                Case "MCO Listing With Premium - Group Company Sub-Total"
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                                    
                    ss = "{VIEW_MCO_LISTING_Printed.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_GrpCo_SubTotal_Printed.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_Printed.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_Printed.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING_Printed.MBMBatchNo} = " & tmpBatchNo & ""
                    Call MCOListingBBPrintSeq
                                            
            End Select
    
    Case 2 ' MCO Print Option for Print All
            Select Case cboMCOListingVer
                Case "MCO Listing"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                                
                    ss = "{VIEW_MCO_LISTING_All.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WithoutPremium_All.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_All.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING_All.MBMBatchNo} = " & tmpBatchNo & ""
                                        
                    Call MCOListingBBPrintSeq
                                            
                Case "MCO Listing - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                                
                   ss = "{VIEW_MCO_LISTING_All.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WP_GrpCo_SubTotal_All.rpt"
                                         
                    Call MCOListingBBPrintSeq
                    
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_All.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING_All.MBMBatchNo} = " & tmpBatchNo & ""
                                            
                Case "MCO Listing With Premium"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                                                
                    ss = "{VIEW_MCO_LISTING_All.MBMBordxDate} = '" & SDate & "'"
                                        
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_All.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_All.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING_All.MBMBatchNo} = " & tmpBatchNo & ""
                                            
                    Call MCOListingBBPrintSeq
                                            
                Case "MCO Listing With Premium - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    TMPPayCode = lvwMCO.SelectedItem
                    SDate = Trim(txtMCOBordxDate)
                    
                    ss = "{VIEW_MCO_LISTING_All.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_GrpCo_SubTotal_All.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_All.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                               " {VIEW_MCO_LISTING_All.MBMBatchNo} = " & tmpBatchNo & ""
                    Call MCOListingBBPrintSeq
                                            
            End Select
End Select
End Sub
Private Sub MCOListingVer_BRPrint()
Select Case cboMCOPrintOpt.ListIndex
    Case 0 'MCO Print Option for Print New
            Select Case cboMCOListingVer
                Case "MCO Listing"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                                                
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WithoutPremium.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_LISTING.PayCode} = '" & TMPPayCode & "'"
                                                               
                    Call MCOListingBRPrintSeq
                                            
                Case "MCO Listing - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                                    
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WP_GrpCo_SubTotal.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_LISTING.PayCode} = '" & TMPPayCode & "'"
                                                              
                    Call MCOListingBRPrintSeq
                                        
                Case "MCO Listing With Premium"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                    
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_LISTING.PayCode} = '" & TMPPayCode & "'"
                                                              
                    If Trim(strsql2) <> "" Then
                        cRptMCO.SelectionFormula = cRptMCO.SelectionFormula & "And (" & strsql2 & ") "
                    End If
                    
                    Call MCOListingBRPrintSeq
                                        
                Case "MCO Listing With Premium - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                                    
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_GrpCo_SubTotal.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_LISTING.PayCode} = '" & TMPPayCode & "'"
                     
                     Call MCOListingBRPrintSeq
                                        
            End Select
    
    Case 1 'MCO Print Option for Reprint
            Select Case cboMCOListingVer
                Case "MCO Listing"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                                            
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WithoutPremium_Printed.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_Printed.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_Printed.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_LISTING_Printed.PayCode} = '" & TMPPayCode & "'"
                                                           
                    Call MCOListingBRPrintSeq
                                        
                Case "MCO Listing - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                                            
                    ss = "{VIEW_MCO_LISTING_Printed.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WP_GrpCo_SubTotal_Printed.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_Printed.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_Printed.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_LISTING_Printed.PayCode} = '" & TMPPayCode & "'"
                                                          
                    Call MCOListingBRPrintSeq
                                    
                Case "MCO Listing With Premium"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                    
                    ss = "{VIEW_MCO_LISTING_Printed.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_Printed.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_Printed.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_Printed.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_LISTING_Printed.PayCode} = '" & TMPPayCode & "'"
                                                          
                    Call MCOListingBRPrintSeq
                                    
                Case "MCO Listing With Premium - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                    
                    ss = "{VIEW_MCO_LISTING_Printed.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_GrpCo_SubTotal_Printed.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_Printed.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_Printed.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_LISTING_Printed.PayCode} = '" & TMPPayCode & "'"
                                                           
                    Call MCOListingBRPrintSeq
                                    
            End Select
    
    Case 2
            Select Case cboMCOListingVer
                Case "MCO Listing"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                                                
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WithoutPremium_All.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_All.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_LISTING_All.PayCode} = '" & TMPPayCode & "'"
                                                               
                    Call MCOListingBRPrintSeq
                                            
                Case "MCO Listing - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                    
                    ss = "{VIEW_MCO_LISTING_All.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WP_GrpCo_SubTotal_All.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_All.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_LISTING_All.PayCode} = '" & TMPPayCode & "'"
                                                              
                    Call MCOListingBRPrintSeq
                                        
                Case "MCO Listing With Premium"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                    
                    ss = "{VIEW_MCO_LISTING_All.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_All.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_All.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_LISTING_All.PayCode} = '" & TMPPayCode & "'"
                                                              
                    Call MCOListingBRPrintSeq
                                        
                Case "MCO Listing With Premium - Group Company Sub-Total"
                    TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
                    TMPPayCode = lvwMCO.SelectedItem
                    tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
                    DateFrom = Trim(txtMCOBordxDateFr)
                    DateTo = Trim(txtMCOBordxDateTo)
                    
                    ss = "{VIEW_MCO_LISTING_All.MBMBordxDate} = '" & SDate & "'"
                    cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_GrpCo_SubTotal_All.rpt"
                    cRptMCO.Destination = crptToPrinter
                    cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                               " {VIEW_MCO_LISTING_All.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                               " {VIEW_MCO_LISTING_All.PayCode} = '" & TMPPayCode & "'"
                                                               
                    Call MCOListingBRPrintSeq
                                        
            End Select
End Select
End Sub

Private Sub MCOListingBBPrintSeqPrev()
'On Error Resume Next
Select Case cboMCOPrintOpt.ListIndex
    Case 0
            Select Case cboMCOPrintSeq.ListIndex
                Case 0
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING.MBMGroupCompany}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    
                Case 1
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING.MBMNumber}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                
                Case 2
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING.MBMPolicyNo}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                    
                Case 3
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING.MBMName}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.PrintReport
                    cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
            End Select
    
    
    Case 1
            Select Case cboMCOPrintSeq.ListIndex
                Case 0
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing_Printed.MBMGroupCompany}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_Printed.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_Printed.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                
                Case 1
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing_Printed.MBMNumber}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING_Printed.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                            
                Case 2
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing_Printed.MBMPolicyNo}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_Printed.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_Printed.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                        
                Case 3
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing_Printed.MBMName}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_Printed.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_Printed.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
            End Select
    
    Case 2
            Select Case cboMCOPrintSeq.ListIndex
                Case 0
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All.MBMGroupCompany}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING_All.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                    
                Case 1
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All.MBMNumber}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING_ALL.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                        
                Case 2
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All.MBMPolicyNo}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_ALL.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                            
                Case 3
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All.MBMName}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_ALL.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
            End Select
End Select
End Sub

Private Sub MCOListingBRPrintSeqPrev()
'On Error Resume Next
Select Case cboMCOPrintOpt.ListIndex
    Case 0
            Select Case cboMCOPrintSeq.ListIndex
                Case 0
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing.MBMGroupCompany}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                    
                Case 1
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING.MBMNumber}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                    
                Case 2
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING.MBMPolicyNo}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                    
                Case 3
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING.MBMName}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
            End Select
    
    Case 1
            Select Case cboMCOPrintSeq.ListIndex
                Case 0
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing_Printed.MBMGroupCompany}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_Printed.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_Printed.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                
                Case 1
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing_Printed.MBMNumber}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING_Printed.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                            
                Case 2
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing_Printed.MBMPolicyNo}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_Printed.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_Printed.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                        
                Case 3
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing_Printed.MBMName}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_Printed.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_Printed.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
            End Select
    
    Case 2
            Select Case cboMCOPrintSeq.ListIndex
                Case 0
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All.MBMGroupCompany}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING_All.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                    
                Case 1
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All.MBMNumber}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING_ALL.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                        
                Case 2
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All.MBMPolicyNo}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_ALL.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                            
                Case 3
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All.MBMName}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_ALL.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
            End Select
End Select
End Sub

Private Sub MCOListingBBPrintSeq()
'On Error Resume Next
Select Case cboMCOPrintOpt.ListIndex
    Case 0
            Select Case cboMCOPrintSeq.ListIndex
                Case 0
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING.MBMGroupCompany}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    'Call cmdMCOSearch_Click

                Case 1
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING.MBMNumber}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
                    'Call cmdMCOSearch_Click
                
                Case 2
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING.MBMPolicyNo}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    '
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    'cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
                    
                Case 3
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING.MBMName}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    'Call cmdMCOSearch_Click
            End Select
    
    
    Case 1
            Select Case cboMCOPrintSeq.ListIndex
                Case 0
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing_Printed.MBMGroupCompany}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_Printed.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_Printed.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    'Call cmdMCOSearch_Click
                
                Case 1
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing_Printed.MBMNumber}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING_Printed.MBMCCoverID}"
                    'cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    'cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    'cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    'cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    'cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                        'cRptMCO.Action = 1
                        'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
                Case 2
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing_Printed.MBMPolicyNo}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_Printed.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_Printed.MBMCCoverID}"
                    'cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    'cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    'cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    'cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    'cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                        'cRptMCO.Action = 1
                        'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
                Case 3
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing_Printed.MBMName}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_Printed.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_Printed.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
            End Select
    
    Case 2
            Select Case cboMCOPrintSeq.ListIndex
                Case 0
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All.MBMGroupCompany}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING_All.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
                Case 1
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All.MBMNumber}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING_ALL.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
                Case 2
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All.MBMPolicyNo}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_ALL.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
                Case 3
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All.MBMName}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_ALL.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & SDate & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
            End Select
End Select
End Sub

Private Sub MCOListingBRPrintSeq()
'On Error Resume Next
Select Case cboMCOPrintOpt.ListIndex
    Case 0
            Select Case cboMCOPrintSeq.ListIndex
                Case 0
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing.MBMGroupCompany}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
                Case 1
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING.MBMNumber}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.PrintReport
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                Case 2
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING.MBMPolicyNo}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
                Case 3
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING.MBMName}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
            End Select
    
    Case 1
            Select Case cboMCOPrintSeq.ListIndex
                Case 0
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing_Printed.MBMGroupCompany}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_Printed.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_Printed.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
                Case 1
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing_Printed.MBMNumber}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING_Printed.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
                Case 2
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing_Printed.MBMPolicyNo}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_Printed.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_Printed.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
                Case 3
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_Listing_Printed.MBMName}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_Printed.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_Printed.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
            End Select
    
    Case 2
            Select Case cboMCOPrintSeq.ListIndex
                Case 0
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All.MBMGroupCompany}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING_All.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
                Case 1
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All.MBMNumber}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_LISTING_ALL.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
                Case 2
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All.MBMPolicyNo}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_ALL.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
                Case 3
                    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All.MBMName}"
                    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_ALL.MBMNumber}"
                    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL.MBMCCoverID}"
                    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
                    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
                    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
                    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
                    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
                    'cRptMCO.Action = 1
                    'Call cmdMCOSearch_Click
                    If bPrintPass = False And CancelPrinting = False Then
                        bPrintPass = PrintFunction
                        If bPrintPass = False Then
                            Screen.MousePointer = Default
                            Exit Sub
                        End If
                    Else
                        cRptMCO.PrintReport
                    End If
                    
            End Select
End Select
End Sub

Sub ComboListing()
    Dim PAY As New ADODB.Recordset
    Dim sqlstr As String
   
    sqlstr = "select PAYCode , PAYCompanyName from PAY "
    
    PAY.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do While Not PAY.EOF
    '    If SSTab1.Tab = 0 Then
    '        SSTab1.TabEnabled(0) = True
    '        cboLLPayor.AddItem (PAY!PAYCode) & " - " & (PAY!PAYCompanyName)
    '        PAY.MoveNext
    '    Else
    '        If SSTab1.Tab = 1 Then
    '            SSTab1.TabEnabled(1) = True
                cboMCOPayor.AddItem (PAY!PAYCode) & " - " & (PAY!PAYCompanyName)
                PAY.MoveNext
    '        Else
    '            If SSTab1.Tab = 2 Then
    '                SSTab1.TabEnabled(2) = True
    '                cboINVPayor.AddItem (PAY!PAYCode) & " - " & (PAY!PAYCompanyName)
    '                PAY.MoveNext
    '            Else
    '                If SSTab1.Tab = 3 Then
    '                    SSTab1.TabEnabled(3) = True
    '                    cboExpPayor.AddItem (PAY!PAYCode) & " - " & (PAY!PAYCompanyName)
    '                    PAY.MoveNext
    '                End If
    '            End If
    '        End If
    '    End If
    Loop
    
    PAY.Close
    Set PAY = Nothing
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
 
Private Sub Form_Load()
    Call ComboListing
    'Call FormTabView
       X = CreatePath("c:\cardPrinting")
        If X = False Then
            MsgBox "Error Creating Folder c:\CardPrinting !" & vbNewLine & _
                    " Please create that folder manually to store the output of leaflet in sofe copy!"
        End If
End Sub

Private Function PrintFunction() As Boolean
On Error GoTo cancelPrint
'Select Case SSTab1.Tab
    'Case 0
    '    If cRptLL Is Nothing Then Exit Function
    '
    '    With dlgCommonDialog
    '        .DialogTitle = "Print Crystal Report "
    '        .CancelError = True
    '        .Flags = cdlPDNoSelection + cdlPDNoPageNums
    '        .ShowPrinter
    '
    '        If ERR <> MSComDlg.cdlCancel Then
    '            cRptLL.PrintReport
    '        End If
    '    End With
    '    PrintFunction = True
    '    Exit Function
'cancelPrint:
'        PrintFunction = False
'        CancelPrinting = True
    
'    Case 1
        If cRptMCO Is Nothing Then Exit Function
    
            With dlgCommonDialog_MCO
                .DialogTitle = "Print Crystal Report "
                .CancelError = True
                .Flags = cdlPDNoSelection + cdlPDNoPageNums
                .ShowPrinter
                
                If ERR <> MSComDlg.cdlCancel Then
                    cRptMCO.PrintReport
                End If
            End With
            PrintFunction = True
            Exit Function
cancelPrint:
                PrintFunction = False
                CancelPrinting = True
                Exit Function

'    Case 2
'        If cRptINV Is Nothing Then Exit Function
'
'            With dlgCommonDialog_INV
'                .DialogTitle = "Print Crystal Report "
'                .CancelError = True
'                 'For ii = 1 To UBound(MemberNoArray)
'                    .Flags = cdlPDNoSelection + cdlPDNoPageNums
'                    'If ActiveForm.rtfText.SelLength = 0 Then
'                     '   .Flags = .Flags + cdlPDAllPages
'                    'Else
'                     '   .Flags = .Flags + cdlPDSelection
'                    'End If
'
'                .ShowPrinter
'
'                If ERR <> MSComDlg.cdlCancel Then
'                    If cboINVPrintOpt.ListIndex = 0 Then
'                        Call IncInvoiceNo
'                        cRptINV.ParameterFields(0) = "MCOInvoiceNo;" & (TxtInvoiceYear & txtInvoiceNo) & ";true"
'                        cRptINV.PrintReport
'                    Else
'                        If cboINVPrintOpt.ListIndex = 1 Then
'                            cRptINV.PrintReport
'                        End If
'                    End If
'                End If
'            ' Next ii
'            End With
'            PrintFunction = True
'            Exit Function
'cancelPrint2:
 '               PrintFunction = False
 '               CancelPrinting = True
 '               Exit Function
                
'    Case 3
'        If cRptLL Is Nothing Then Exit Function
'
'            With dlgCommonDialog
'                .DialogTitle = "Print Crystal Report "
'                .CancelError = True
'                 'For ii = 1 To UBound(MemberNoArray)
'                    .Flags = cdlPDNoSelection + cdlPDNoPageNums
'                    'If ActiveForm.rtfText.SelLength = 0 Then
'                     '   .Flags = .Flags + cdlPDAllPages
'                    'Else
'                     '   .Flags = .Flags + cdlPDSelection
'                    'End If
'
'                .ShowPrinter
'
'                If ERR <> MSComDlg.cdlCancel Then
'                    cRptLL.PrintReport
'                End If
'            ' Next ii
'            End With
'            PrintFunction = True
'            Exit Function
'cancelPrint3:
'                PrintFunction = False
'                CancelPrinting = True
'                Exit Function

'End Select
    
End Function

Private Sub IncInvoiceNo()
'On Error Resume Next
    Dim MBMMCOINVQ As New ADODB.Recordset
    Dim strsql As String
    Dim InvoiceNo As String
    Dim Code As String
    
    
    strsql = "Select * from MBMMCOINVQ "
    MBMMCOINVQ.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If MBMMCOINVQ.recordCount = 0 Then
        With MBMMCOINVQ
            .AddNew
                !MCOInvoice_No = "000001"
                InvoiceNo = "000001"
                '!MCOYear = Year(Date)
            .Update
        End With
    Else
    
        With MBMMCOINVQ
                !MCOInvoice_No = Format(!MCOInvoice_No + 1, "000000")
                InvoiceNo = Format(!MCOInvoice_No, "000000")
        .Update
        End With
    End If

    txtInvoiceNo = InvoiceNo
    'TxtInvoiceYear = Year(Date)
    
    MBMMCOINVQ.Close
    Set MBMMCOINVQ = Nothing
            
End Sub

Private Sub lvwMCO_BeforeLabelEdit(Cancel As Integer)
    'lvwMCO.SortKey = ColumnHeader.Index - 1
    'lvwMCO.Sorted = True
    'If lvwMCO.SortOrder = lvwAscending Then
    '   lvwMCO.SortOrder = lvwDescending
    'Else
    '   lvwMCO.SortOrder = lvwAscending
    'End If
End Sub

Private Sub UpdateMCOPrintingFile(BordxDate As String, BatchNo As String, PAYCode As String, HLTCode As String, MemberNo As String, WhichField As String)
Dim rsUpdate As New ADODB.Recordset
Dim rsMBMUpdate As New ADODB.Recordset
Dim strsql As String
Dim sql As String
    
'Update file to BordxPrinting
    sql = "SELECT * FROM BordxPrinting"
    sql = sql & " WHERE MBMBordxDate = '" & Format(CDate(BordxDate), "MM/DD/YYYY") & "'"
    sql = sql & " AND MBMBatchNo = " & BatchNo & ""
    sql = sql & " AND PAYCode = '" & PAYCode & "'"
    sql = sql & " AND HLTCode = '" & Trim(HLTCode) & "'"

    rsUpdate.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
 
 'Update file to MBMPrinting
    strsql = "SELECT * FROM MBMPrinting"
    strsql = strsql & " WHERE MBMBordxDate = '" & Format(CDate(BordxDate), "MM/DD/YYYY") & "'"
    strsql = strsql & " AND MBMBatchNo = " & BatchNo & ""
    strsql = strsql & " AND PAYCode = '" & PAYCode & "'"
    strsql = strsql & " AND HLTCode = '" & Trim(HLTCode) & "'"
    strsql = strsql & " AND MBMNumber = '" & MemberNo & "'"
 
    rsMBMUpdate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
 
    If rsMBMUpdate.recordCount = 0 Then 'AddNew
        rsMBMUpdate.AddNew
        rsMBMUpdate!MBMBordxDate = BordxDate
        rsMBMUpdate!MBMBatchNo = BatchNo
        rsMBMUpdate!PAYCode = PAYCode
        rsMBMUpdate!HLTCode = Trim(HLTCode)
        rsMBMUpdate!MBMNumber = MemberNo
 
        Select Case WhichField
            Case "MCO"
                 rsMBMUpdate!MBMMCOListPrinted = 1
                 rsMBMUpdate!MBMMCOListDate = Date
        End Select
        rsMBMUpdate.Update
    Else 'Update
        Select Case WhichField
            Case "MCO"
                 rsMBMUpdate!MBMMCOListPrinted = 1
                 rsMBMUpdate!MBMMCOListDate = Date
        End Select
        rsMBMUpdate.Update
    End If
             
    If rsUpdate.recordCount = 0 Then 'AddNew
        rsUpdate.AddNew
        rsUpdate!MBMBordxDate = BordxDate
        rsUpdate!MBMBatchNo = BatchNo
        rsUpdate!PAYCode = PAYCode
        rsUpdate!HLTCode = Trim(HLTCode)
    
        Select Case WhichField
            Case "MCO"
                 rsUpdate!MBMMCOListPrinted = "1"
                 rsUpdate!MBMMCOListDate = Date
        End Select
        rsUpdate.Update
    Else 'Update
        Select Case WhichField
            Case "MCO"
                 rsUpdate!MBMMCOListPrinted = "1"
                 rsUpdate!MBMMCOListDate = Date
        End Select
        rsUpdate.Update
    End If
    rsUpdate.Close
    rsMBMUpdate.Close

End Sub

