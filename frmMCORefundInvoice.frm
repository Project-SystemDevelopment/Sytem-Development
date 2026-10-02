VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMCORefundInvoice 
   Caption         =   "MCO Refund Invoice"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   7200
      Width           =   11055
      Begin VB.CommandButton cmdINVClear 
         Caption         =   "&Clear"
         Height          =   300
         Left            =   9240
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdINVPrintPreview 
         Caption         =   "Pre&view"
         Height          =   300
         Left            =   8400
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdINVPrint 
         Caption         =   "&Print"
         Height          =   300
         Left            =   7560
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdINVSearch 
         Caption         =   "&Search"
         Height          =   300
         Left            =   6720
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtInvoiceYear 
         Height          =   285
         Left            =   5400
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtInvoiceNo 
         Height          =   285
         Left            =   5880
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   5640
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdINVClose 
         Caption         =   "&Exit"
         Height          =   300
         Left            =   10080
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label203 
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
         Left            =   2880
         TabIndex        =   36
         Top             =   240
         Width           =   960
      End
      Begin VB.Label txtINVTotalSupplementary1 
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
         Left            =   3840
         TabIndex        =   35
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label txtINVTotalPrincipal 
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
         Left            =   1440
         TabIndex        =   34
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label202 
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
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "MCO Refund Invoice"
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
         TabIndex        =   23
         Top             =   60
         Width           =   3285
      End
   End
   Begin VB.Frame FrameINV 
      Caption         =   "Enter Criteria For Searching"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   11055
      Begin VB.TextBox txtINVInvoiceNoTo 
         Height          =   315
         Left            =   8520
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtINVBatchNo 
         Height          =   315
         Left            =   8520
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cboINVCriteriaSel 
         Height          =   315
         ItemData        =   "frmMCORefundInvoice.frx":0000
         Left            =   2160
         List            =   "frmMCORefundInvoice.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtINVInvoiceNo 
         Height          =   315
         Left            =   6000
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtINVInvoiceNoFr 
         Height          =   315
         Left            =   6000
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtINVBordxDate 
         Height          =   315
         Left            =   6000
         TabIndex        =   11
         Top             =   360
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
         TabIndex        =   12
         Top             =   360
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
         TabIndex        =   13
         Top             =   360
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
         Left            =   1560
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox cboINVPayor 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   650
         Width           =   5415
      End
      Begin VB.ComboBox cboINVListingVer 
         Height          =   315
         ItemData        =   "frmMCORefundInvoice.frx":0053
         Left            =   2160
         List            =   "frmMCORefundInvoice.frx":005A
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   930
         Width           =   5415
      End
      Begin VB.ComboBox cboINVPrintOpt 
         Height          =   315
         ItemData        =   "frmMCORefundInvoice.frx":0067
         Left            =   2160
         List            =   "frmMCORefundInvoice.frx":0071
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cboINVPrintSeq 
         Height          =   315
         ItemData        =   "frmMCORefundInvoice.frx":0089
         Left            =   5400
         List            =   "frmMCORefundInvoice.frx":0099
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Width           =   2175
      End
      Begin Crystal.CrystalReport cRptInvoice 
         Left            =   1560
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
      Begin VB.Label lbl29 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   5040
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbl30 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   7560
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbl36 
         Caption         =   "Print Option"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lbl35 
         Caption         =   "Listing Versions"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbl34 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lbl37 
         Caption         =   "Sequence"
         Height          =   255
         Left            =   4640
         TabIndex        =   16
         Top             =   1250
         Width           =   735
      End
      Begin VB.Label lbl27 
         Caption         =   "Single/Batch/Range"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame7 
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
      Height          =   5160
      Left            =   120
      TabIndex        =   0
      Top             =   2055
      Width           =   11055
      Begin MSComctlLib.ListView lvwInvoice 
         Height          =   4755
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8387
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   12
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
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Invoice No"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmMCORefundInvoice"
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
'On Error Resume Next
Dim errorFound As Boolean

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
 

Private Sub RecordSearching_INV(objListView As Object)
    Screen.MousePointer = vbHourglass
    Call INVPrincipalListing
    Screen.MousePointer = Default
End Sub

Private Sub INVPrincipalListing()
Dim rsSearch As New ADODB.Recordset
Dim rsPrint As New ADODB.Recordset
Dim sql As String
Dim itemToPrint As Integer
Dim isPrinted As Boolean
Dim CoverQty As Integer
Dim INVCriteria As String
Dim TotalPrincipal As Integer


TotalSupplementary = 0
lvwInvoice.listitems.Clear
sql = "SELECT * FROM VIEW_InvoiceRef_SearchRpt"
Select Case cboINVCriteriaSel.ListIndex
    Case 0 'Criteria Selection for Invoice on Bordx Batch
            SDate = CStr(Month(txtINVBordxDate)) & "/" & CStr(Day(txtINVBordxDate)) & "/" & CStr(Year(txtINVBordxDate))
            sql = sql & " WHERE MBMBordxDate = '" & SDate & "'"
            sql = sql & " AND MBMBatchNo = '" & txtINVBatchNo & "'"
            sql = sql & " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "'"
    Case 1   'Criteria Selection for Invoice on Bordx Range
             DateFrom = CStr(Month(txtINVBordxDateFr)) & "/" & CStr(Day(txtINVBordxDateFr)) & "/" & CStr(Year(txtINVBordxDateFr))
             DateTo = CStr(Month(txtINVBordxDateTo)) & "/" & CStr(Day(txtINVBordxDateTo)) & "/" & CStr(Year(txtINVBordxDateTo))
            sql = sql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
            sql = sql & " AND MBMBordxDate <= '" & DateTo & "'"
            sql = sql & " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "'"
    Case 2  'Criteria Selection for Invoice on Invoice Number
            sql = sql & " WHERE MBMInvoiceNo = '" & txtINVInvoiceNo & "'"
    Case 3  'Criteria Selection for Invoice on Invoice Range
            sql = sql & " WHERE MBMInvoiceNo >= '" & txtINVInvoiceNoFr & "'"
            sql = sql & " AND MBMInvoiceNo <= '" & txtINVInvoiceNoTo & "'"
End Select

If cboINVPrintOpt.ListIndex = 0 Then
    sql = sql & " AND (MBMRefInvoiced is NULL or  MBMRefInvoiced  = 0 )"
Else
    sql = sql & " AND MBMRefInvoiced = 1 "
End If

Select Case cboINVPrintSeq.ListIndex
    Case 0 'Invoice Print Seq for Group Company
        sql = sql & " ORDER BY MBMGroupCompany, MBMNumber"
    Case 1 'Invoice Print Seq for Membership Number
        sql = sql & " ORDER BY MBMNumber"
    Case 2 'Invoice Print Seq for Policy No
        sql = sql & " ORDER BY MBMPolicyNo"
    Case 3 'Invoice Print Seq for Member's Name
        sql = sql & " ORDER BY MBMName"
End Select


rsSearch.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText

'End Select
    If rsSearch.recordCount = 0 Then
        MsgBox "Record not found !", vbInformation + vbOKOnly
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
        
    Else
        cmdINVPrint.Enabled = True
        cmdINVPrintPreview.Enabled = True

            Do Until rsSearch.EOF
                Set ListMaster = lvwInvoice.listitems.Add(, , rsSearch!PAYCode)
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
                                        
                    If Len(Trim(rsSearch!MBMGroupCompany)) Then
                        ListMaster.SubItems(9) = ""
                    Else
                        If Len(Trim(rsSearch!MBMGroupCompany)) Then
                            ListMaster.SubItems(9) = rsSearch!MBMGroupCompany
                        End If
                    End If
                     
                    'If Len(rsSearch!MBMInvoiceNo) Then
                    '    ListMaster.SubItems(11) = rsSearch!MBMInvoiceNo
                    'End If
                    
                        TotalPrincipal = TotalPrincipal + 1
                        txtINVTotalPrincipal = TotalPrincipal
                        PAYCode = rsSearch!PAYCode
                        MBMNo = rsSearch!MBMNumber
                        BordxDat = rsSearch!MBMBordxDate
                        'If Len(rsSearch!MBMInvoiceNo) Then
                        '    INVNo = rsSearch!MBMInvoiceNo
                        'End If
                        
                        'Select Case SSTab1.Tab
                        '    Case 2 'Invoice Printing
                                itemToPrint = INV_PRINTED
                                cmdINVPrint.Enabled = True
                                cmdINVPrintPreview.Enabled = True
                        '    End Select
                            
                            If rsSearch.recordCount > 0 Then
                        '        Select Case SSTab1.Tab
                        '            Case 2 'Invoice Printing
                                        If Len(rsSearch!MBMRefInvoiced) Then
                                            isPrinted = rsSearch!MBMRefInvoiced
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
                                Call INVSuppListing
                            End If
            rsSearch.MoveNext
            Loop
    End If
    rsSearch.Close
    Set rsSearch = Nothing
End Sub

Private Sub cboINVCriteriaSel_Change()
Select Case cboINVCriteriaSel.ListIndex
    Case 0
        lbl29.Caption = "Bordx Date"
        lbl30.Caption = "Batch Num"
        txtINVBordxDate.Visible = True
        txtINVBatchNo.Visible = True
        txtINVBordxDateFr.Visible = False
        txtINVBordxDateTo.Visible = False
        txtINVInvoiceNo.Visible = False
        txtINVInvoiceNoFr.Visible = False
        txtINVInvoiceNoTo.Visible = False
        cboINVPayor.Enabled = True
        txtINVBatchNo = ""
       ' cboINVPayor = ""
        txtINVBordxDate = "__/__/____"
        txtINVBordxDateFr = "__/__/____"
        txtINVBordxDateTo = "__/__/____"
        cboINVPrintOpt.ListIndex = 0
        cboINVPrintOpt.Enabled = True
        cboINVPrintSeq.ListIndex = 0
        cboINVListingVer.ListIndex = 0
        txtINVInvoiceNo = ""
        txtINVInvoiceNoFr = ""
        txtINVInvoiceNoTo = ""
        lvwInvoice.listitems.Clear
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
        txtINVTotalPrincipal = 0
        txtINVTotalSupplementary1 = 0
        
    Case 1
        lbl29.Caption = "Bordx From"
        lbl30.Caption = "To"
        txtINVBordxDate.Visible = False
        txtINVBatchNo.Visible = False
        txtINVBordxDateFr.Visible = True
        txtINVBordxDateTo.Visible = True
        txtINVInvoiceNo.Visible = False
        txtINVInvoiceNoFr.Visible = False
        txtINVInvoiceNoTo.Visible = False
        cboINVPayor.Enabled = True
        txtINVBatchNo = ""
    '    cboINVPayor = ""
        txtINVBordxDate = "__/__/____"
        txtINVBordxDateFr = "__/__/____"
        txtINVBordxDateTo = "__/__/____"
        cboINVPrintOpt.ListIndex = 0
        cboINVPrintOpt.Enabled = True
        cboINVPrintSeq.ListIndex = 0
        cboINVListingVer.ListIndex = 0
        txtINVInvoiceNo = ""
        txtINVInvoiceNoFr = ""
        txtINVInvoiceNoTo = ""
        lvwInvoice.listitems.Clear
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
        txtINVTotalPrincipal = 0
        txtINVTotalSupplementary1 = 0
        
    Case 2
        lbl29.Caption = "Invoice Num"
        lbl30.Caption = ""
        txtINVBordxDate.Visible = False
        txtINVBatchNo.Visible = False
        txtINVBordxDateFr.Visible = False
        txtINVBordxDateTo.Visible = False
        txtINVInvoiceNo.Visible = True
        txtINVInvoiceNoFr.Visible = False
        txtINVInvoiceNoTo.Visible = False
        cboINVPayor.Enabled = False
        txtINVBatchNo = ""
        txtINVBordxDate = "__/__/____"
        txtINVBordxDateFr = "__/__/____"
        txtINVBordxDateTo = "__/__/____"
        cboINVPrintOpt.ListIndex = 1
        cboINVPrintOpt.Enabled = False
        cboINVPrintSeq.ListIndex = 0
        cboINVListingVer.ListIndex = 0
        txtINVInvoiceNo = ""
        txtINVInvoiceNoFr = ""
        txtINVInvoiceNoTo = ""
        lvwInvoice.listitems.Clear
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
        txtINVTotalPrincipal = 0
        txtINVTotalSupplementary1 = 0

    Case 3
        lbl29.Caption = "Invoice From"
        lbl30.Caption = "To"
        txtINVBordxDate.Visible = False
        txtINVBatchNo.Visible = False
        txtINVBordxDateFr.Visible = False
        txtINVBordxDateTo.Visible = False
        txtINVInvoiceNo.Visible = False
        txtINVInvoiceNoFr.Visible = True
        txtINVInvoiceNoTo.Visible = True
        cboINVPayor.Enabled = False
        txtINVBatchNo = ""
        'cboINVPayor = ""
        txtINVBordxDate = "__/__/____"
        txtINVBordxDateFr = "__/__/____"
        txtINVBordxDateTo = "__/__/____"
        cboINVPrintOpt.ListIndex = 1
        cboINVPrintOpt.Enabled = False
        cboINVPrintSeq.ListIndex = 0
        cboINVListingVer.ListIndex = 0
        txtINVInvoiceNo = ""
        txtINVInvoiceNoFr = ""
        txtINVInvoiceNoTo = ""
        lvwInvoice.listitems.Clear
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
        txtINVTotalPrincipal = 0
        txtINVTotalSupplementary1 = 0
End Select
End Sub

Private Sub cboINVCriteriaSel_Click()
Select Case cboINVCriteriaSel.ListIndex
    Case 0
        lbl29.Caption = "Bordx Date"
        lbl30.Caption = "Batch Num"
        txtINVBordxDate.Visible = True
        txtINVBatchNo.Visible = True
        txtINVBordxDateFr.Visible = False
        txtINVBordxDateTo.Visible = False
        txtINVInvoiceNo.Visible = False
        txtINVInvoiceNoFr.Visible = False
        txtINVInvoiceNoTo.Visible = False
        cboINVPayor.Enabled = True
        txtINVBatchNo = ""
       ' cboINVPayor = ""
        txtINVBordxDate = "__/__/____"
        txtINVBordxDateFr = "__/__/____"
        txtINVBordxDateTo = "__/__/____"
        cboINVPrintOpt.ListIndex = 0
        cboINVPrintOpt.Enabled = True
        cboINVPrintSeq.ListIndex = 0
        cboINVListingVer.ListIndex = 0
        txtINVInvoiceNo = ""
        txtINVInvoiceNoFr = ""
        txtINVInvoiceNoTo = ""
        lvwInvoice.listitems.Clear
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
        txtINVTotalPrincipal = 0
        txtINVTotalSupplementary1 = 0
        
    Case 1
        lbl29.Caption = "Bordx From"
        lbl30.Caption = "To"
        txtINVBordxDate.Visible = False
        txtINVBatchNo.Visible = False
        txtINVBordxDateFr.Visible = True
        txtINVBordxDateTo.Visible = True
        txtINVInvoiceNo.Visible = False
        txtINVInvoiceNoFr.Visible = False
        txtINVInvoiceNoTo.Visible = False
        cboINVPayor.Enabled = True
        txtINVBatchNo = ""
    '    cboINVPayor = ""
        txtINVBordxDate = "__/__/____"
        txtINVBordxDateFr = "__/__/____"
        txtINVBordxDateTo = "__/__/____"
        cboINVPrintOpt.ListIndex = 0
        cboINVPrintOpt.Enabled = True
        cboINVPrintSeq.ListIndex = 0
        cboINVListingVer.ListIndex = 0
        txtINVInvoiceNo = ""
        txtINVInvoiceNoFr = ""
        txtINVInvoiceNoTo = ""
        lvwInvoice.listitems.Clear
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
        txtINVTotalPrincipal = 0
        txtINVTotalSupplementary1 = 0
        
    Case 2
        lbl29.Caption = "Invoice Num"
        lbl30.Caption = ""
        txtINVBordxDate.Visible = False
        txtINVBatchNo.Visible = False
        txtINVBordxDateFr.Visible = False
        txtINVBordxDateTo.Visible = False
        txtINVInvoiceNo.Visible = True
        txtINVInvoiceNoFr.Visible = False
        txtINVInvoiceNoTo.Visible = False
        cboINVPayor.Enabled = False
        txtINVBatchNo = ""
        txtINVBordxDate = "__/__/____"
        txtINVBordxDateFr = "__/__/____"
        txtINVBordxDateTo = "__/__/____"
        cboINVPrintOpt.ListIndex = 1
        cboINVPrintOpt.Enabled = False
        cboINVPrintSeq.ListIndex = 0
        cboINVListingVer.ListIndex = 0
        txtINVInvoiceNo = ""
        txtINVInvoiceNoFr = ""
        txtINVInvoiceNoTo = ""
        lvwInvoice.listitems.Clear
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
        txtINVTotalPrincipal = 0
        txtINVTotalSupplementary1 = 0

    Case 3
        lbl29.Caption = "Invoice From"
        lbl30.Caption = "To"
        txtINVBordxDate.Visible = False
        txtINVBatchNo.Visible = False
        txtINVBordxDateFr.Visible = False
        txtINVBordxDateTo.Visible = False
        txtINVInvoiceNo.Visible = False
        txtINVInvoiceNoFr.Visible = True
        txtINVInvoiceNoTo.Visible = True
        cboINVPayor.Enabled = False
        txtINVBatchNo = ""
        'cboINVPayor = ""
        txtINVBordxDate = "__/__/____"
        txtINVBordxDateFr = "__/__/____"
        txtINVBordxDateTo = "__/__/____"
        cboINVPrintOpt.ListIndex = 1
        cboINVPrintOpt.Enabled = False
        cboINVPrintSeq.ListIndex = 0
        cboINVListingVer.ListIndex = 0
        txtINVInvoiceNo = ""
        txtINVInvoiceNoFr = ""
        txtINVInvoiceNoTo = ""
        lvwInvoice.listitems.Clear
        cmdINVPrint.Enabled = False
        cmdINVPrintPreview.Enabled = False
        txtINVTotalPrincipal = 0
        txtINVTotalSupplementary1 = 0
End Select
End Sub

Private Sub cboINVCriteriaSel_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboINVCriteriaSel.Text, cboINVCriteriaSel.SelStart) + Chr(KeyAscii) + Mid(cboINVCriteriaSel.Text, cboINVCriteriaSel.SelStart + cboINVCriteriaSel.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboINVCriteriaSel.ListCount - 1
        
        If NewText = LCase(Left(cboINVCriteriaSel.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboINVCriteriaSel.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboINVCriteriaSel.Text = ValidValue
               cboINVCriteriaSel.SelStart = 0
               cboINVCriteriaSel.SelLength = Len(ValidValue)
             End If
        End If

End Sub


Private Sub cboINVListingVer_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboINVListingVer.Text, cboINVListingVer.SelStart) + Chr(KeyAscii) + Mid(cboINVListingVer.Text, cboINVListingVer.SelStart + cboINVListingVer.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboINVListingVer.ListCount - 1
        
        If NewText = LCase(Left(cboINVListingVer.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboINVListingVer.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboINVListingVer.Text = ValidValue
               cboINVListingVer.SelStart = 0
               cboINVListingVer.SelLength = Len(ValidValue)
             End If
        End If
End Sub


Private Sub cboINVPayor_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboINVPayor.Text, cboINVPayor.SelStart) + Chr(KeyAscii) + Mid(cboINVPayor.Text, cboINVPayor.SelStart + cboINVPayor.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboINVPayor.ListCount - 1
        
        If NewText = LCase(Left(cboINVPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboINVPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboINVPayor.Text = ValidValue
               cboINVPayor.SelStart = 0
               cboINVPayor.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboINVPrintOpt_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboINVPrintOpt.Text, cboINVPrintOpt.SelStart) + Chr(KeyAscii) + Mid(cboINVPrintOpt.Text, cboINVPrintOpt.SelStart + cboINVPrintOpt.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboINVPrintOpt.ListCount - 1
        
        If NewText = LCase(Left(cboINVPrintOpt.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboINVPrintOpt.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboINVPrintOpt.Text = ValidValue
               cboINVPrintOpt.SelStart = 0
               cboINVPrintOpt.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboINVPrintSeq_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboINVPrintSeq.Text, cboINVPrintSeq.SelStart) + Chr(KeyAscii) + Mid(cboINVPrintSeq.Text, cboINVPrintSeq.SelStart + cboINVPrintSeq.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboINVPrintSeq.ListCount - 1
        
        If NewText = LCase(Left(cboINVPrintSeq.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboINVPrintSeq.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboINVPrintSeq.Text = ValidValue
               cboINVPrintSeq.SelStart = 0
               cboINVPrintSeq.SelLength = Len(ValidValue)
             End If
        End If

End Sub

Private Sub cmdINVClear_Click()
    txtINVBatchNo = ""
  '  cboINVPayor = ""
    txtINVBordxDate = "__/__/____"
    txtINVBordxDateFr = "__/__/____"
    txtINVBordxDateTo = "__/__/____"
    cboINVPrintOpt.ListIndex = 0
    cboINVPrintSeq.ListIndex = 0
    cboINVListingVer.ListIndex = 0
    txtINVInvoiceNo = ""
    txtINVInvoiceNoFr = ""
    txtINVInvoiceNoTo = ""
    lvwInvoice.listitems.Clear
    cmdINVPrint.Enabled = False
    cmdINVPrintPreview.Enabled = False
    txtINVTotalPrincipal = 0
    txtINVTotalSupplementary1 = 0

End Sub

Private Sub cmdINVClose_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'   End If
End Sub

Private Sub cmdINVPrint_Click()
    If cboINVPrintOpt.ListIndex = 0 Then
        InvoicePrint "NO"
    ElseIf cboINVPrintOpt.ListIndex = 1 Then
        InvoicePrint "YES"
    End If
End Sub

Private Sub InvoicePrint(PrintMode As String)
'On Error Resume Next
Dim ii As Integer
Dim MemberNo, strsql As String
Dim Found As Boolean
Dim tmpBDate, tmpBatchNo, TMPPayCode, tmpHLTCode, tmpMemberNo, tmpInvNo As String
Dim tmpPrintedBDate, tmpPrintedBatchNo, tmpPrintedPaycode, tempPrintedHLTcode, tmpPrintedMemberNo As String
Dim tmpPrintedInvoice As String
Dim BDateArray() As String
Dim CInvoiceNo() As String
Dim CDateArray() As String
Dim BatchNoArray() As String
Dim PaycodeArray() As String
Dim HLTcodeArray() As String
Dim MemberNoArray() As String
Dim InvoiceNoArray() As String
Dim pos1, pos2, selectedBatchNo   As Integer
Dim SelectedBordxDate As String

Dim cnt As Integer
Dim cont As Integer
Dim count As Integer
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

Dim INV As String
Dim intOK  As Boolean
 
 bPrintPass = False
 CancelPrinting = False
 Arraycnt = 1
 Screen.MousePointer = vbHourglass

    
     txtInvoiceNo = lvwInvoice.SelectedItem.SubItems(INV_INVOICE)
'Select Case cboINVCriteriaSel.ListIndex
 '   Case 0
            cnt = 0
            For ii = 1 To lvwInvoice.listitems.count
                Set lvwInvoice.SelectedItem = lvwInvoice.listitems(ii)   'lvwLeaflet.ListItems(ii)
                    If lvwInvoice.SelectedItem.SubItems(INV_PRINTED) = PrintMode Then
                        tmpInvNo = lvwInvoice.SelectedItem.SubItems(INV_INVOICE)
                        tmpMemberNo = lvwInvoice.SelectedItem.SubItems(INV_MEMBERNO)
                        tmpBDate = lvwInvoice.SelectedItem.SubItems(INV_BORDXDATE)
                        tmpHLTCode = lvwInvoice.SelectedItem.SubItems(INV_HLTCODE)
                        tmpBatchNo = lvwInvoice.SelectedItem.SubItems(INV_BATCHNO)
                        TMPPayCode = lvwInvoice.SelectedItem
                            
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
    
If cmdINVPrintPreview Then
 
    tmpBDate = lvwInvoice.SelectedItem.SubItems(INV_BORDXDATE)
    TmpHLT = lvwInvoice.SelectedItem.SubItems(INV_HLTCODE)
    tmpBatchNo = lvwInvoice.SelectedItem.SubItems(INV_BATCHNO)
    TMPPayCode = lvwInvoice.SelectedItem
    tmpInvNo = lvwInvoice.SelectedItem.SubItems(INV_INVOICE)
    SDate = Trim(txtINVBordxDate)
    DateFrom = Trim(txtINVBordxDateFr)
    DateTo = Trim(txtINVBordxDateTo)
    InvNoFr = Trim(txtINVInvoiceNoFr)
    InvNoTo = Trim(txtINVInvoiceNoTo)
    
    Select Case cboINVCriteriaSel.ListIndex
        Case 0 'Invoice Criteria Selection for Bordx Batch
                Select Case cboINVPrintOpt.ListIndex
                    Case 0 ' Invoice Print Option for Print New
                        ss0 = "{VIEW_Invoice_ReportRefund_PrintNew.HLTCode} = '" & TmpHLT & "'"
                        ss = "{VIEW_Invoice_ReportRefund_PrintNew.MBMBordxDate} = '" & SDate & "'"
                        cRptInvoice.ReportFileName = crdPath & "Report\InvoicesRefundReport.rpt"
                        cRptInvoice.Destination = crptToWindow
                        cRptInvoice.WindowState = crptMaximized
                        cRptInvoice.SelectionFormula = " {VIEW_Invoice_ReportRefund_PrintNew.MBMBordxDate} = date(" & Year(txtINVBordxDate) & "," & Month(txtINVBordxDate) & "," & Day(txtINVBordxDate) & ") " & " And " & _
                                                   " {VIEW_Invoice_ReportRefund_PrintNew.HLTCode} = '" & TmpHLT & "'" & " And " & _
                                                   " {VIEW_Invoice_ReportRefund_PrintNew.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                                   " {VIEW_Invoice_ReportRefund_PrintNew.MBMBatchNo} = " & tmpBatchNo & " "
                        txtInvoiceNo = ""
                        cRptInvoice.ParameterFields(0) = "MCOInvoiceNo;" & txtInvoiceNo & ";true"
                       ' cRptINV.PrintReport
                        cRptInvoice.Action = 1
                        txtInvoiceNo = ""
        
                    Case 1 ' Invoice Print Option for Reprint
                        ss0 = "{VIEW_Invoice_ReportRefund_View.HLTCode} = '" & TmpHLT & "'"
                        ss = "{VIEW_Invoice_ReportRefund_View.MBMBordxDate} = '" & SDate & "'"
                        cRptInvoice.ReportFileName = crdPath & "\Report\InvoicesRefundReport_Reprint.rpt"
                        'cRptINV.ReportFileName = crdPath & "Report\InvoicesReport.rpt"
                        cRptInvoice.Destination = crptToWindow
                        cRptInvoice.WindowState = crptMaximized
                        cRptInvoice.SelectionFormula = " {VIEW_InvoiceRefund_Report_View.MBMBordxDate} = date(" & Year(txtINVBordxDate) & "," & Month(txtINVBordxDate) & "," & Day(txtINVBordxDate) & ") " & " And " & _
                                                    " {VIEW_InvoiceRefund_Report_View.HLTCode} = '" & TmpHLT & "'" & " And " & _
                                                   " {VIEW_InvoiceRefund_Report_View.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                                   " {VIEW_InvoiceRefund_Report_View.MBMBatchNo} = " & tmpBatchNo
                        'cRptINV.SelectionFormula = " {VIEW_Invoice_Report_PrintNew.MBMBordxDate} = date(" & Year(txtINVBordxDate) & "," & Month(txtINVBordxDate) & "," & Day(txtINVBordxDate) & ") " & " And " & _
                        '" {VIEW_Invoice_Report_PrintNew.HLTCode} = '" & TmpHLT & "'" & " And " & _
                        '" {VIEW_Invoice_Report_PrintNew.PayCode} = '" & TMPPayCode & "'" & " And " & _
                        '" {VIEW_Invoice_Report_PrintNew.MBMBatchNo} = " & tmpBatchNo & " "
                        cRptInvoice.ParameterFields(0) = "MCOInvoiceNo;" & txtInvoiceNo & ";true"
'                        cRptINV.PrintReport
                        cRptInvoice.Action = 1
                End Select
    
        Case 1 'Invoice Criteria Selection for Bordx Range
                Select Case cboINVPrintOpt.ListIndex
                    Case 0 ' Invoice Print Option for Print New
                        cRptInvoice.ReportFileName = crdPath & "\Report\InvoicesRefundReport_Range.rpt"
                        cRptInvoice.Destination = crptToWindow
                        cRptInvoice.WindowState = crptMaximized
                        
                        cRptInvoice.SelectionFormula = " {VIEW_Invoice_ReportRefund_PrintNew.MBMBordxDate} >= date(" & Year(txtINVBordxDateFr) & "," & Month(txtINVBordxDateFr) & "," & Day(txtINVBordxDateFr) & ") " & " And " & _
                                                   " {VIEW_Invoice_ReportRefund_PrintNew.MBMBordxDate} <= date(" & Year(txtINVBordxDateTo) & "," & Month(txtINVBordxDateTo) & "," & Day(txtINVBordxDateTo) & ") And " & _
                                                   " {VIEW_Invoice_ReportRefund_PrintNew.PayCode} = '" & TMPPayCode & "' "
                        
                        txtInvoiceNo = ""
                        cRptInvoice.ParameterFields(0) = "MCOInvoiceNo;" & txtInvoiceNo & ";true"
                        cRptInvoice.Action = 1
                        txtInvoiceNo = ""
        
                    Case 1 ' Invoice Print Option for Reprint
                       cRptInvoice.ReportFileName = crdPath & "\Report\InvoicesRefundReport_Range.rpt"
                       cRptInvoice.Destination = crptToWindow
                       cRptInvoice.WindowState = crptMaximized
                            
                       cRptInvoice.SelectionFormula = " {VIEW_Invoice_ReportRefund_View_Range.MBMBordxDate} >= date(" & Year(txtINVBordxDateFr) & "," & Month(txtINVBordxDateFr) & "," & Day(txtINVBordxDateFr) & ") " & " And " & _
                                                  " {VIEW_Invoice_ReportRefund_View_Range.MBMBordxDate} <= date(" & Year(txtINVBordxDateTo) & "," & Month(txtINVBordxDateTo) & "," & Day(txtINVBordxDateTo) & ") " & " AND " & _
                                                  " {VIEW_Invoice_ReportRefund_View_Range.PayCode} = '" & TMPPayCode & "' " '& _

                       cRptInvoice.ParameterFields(0) = " MCOInvoiceNo;" & txtInvoiceNo & ";true"
                       'cRptInvoice.PrintReport
                       cRptInvoice.Action = 1
                End Select
              
        Case 2 'Invoice Criteria Selection for Single Invoice No
                Select Case cboINVPrintOpt.ListIndex
                    Case 0 ' Invoice Print Option for Print New
                       MsgBox "Please select 'Reprint' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
                    
                    Case 1 ' Invoice Print Option for Reprint
                       
                       'cRptINV.ReportFileName = crdPath & "\Report\InvoicesReport_View.rpt"
                       cRptInvoice.ReportFileName = crdPath & "Report\InvoicesRefundReport.rpt"
                       cRptInvoice.Destination = crptToWindow
                       cRptInvoice.WindowState = crptMaximized
                       cRptInvoice.SelectionFormula = "{VIEW_Invoice_ReportRefund_PrintNew.MBMInvoiceNo} = '" & Trim(txtINVInvoiceNo) & "'"
                       'INV = " {VIEW_Invoice_Report_View.MBMInvoiceNo} "
                       cRptInvoice.ParameterFields(0) = "MCOInvoiceNo;" & txtInvoiceNo & ";true"
                       'cRptInvoice.PrintReport
                        cRptInvoice.Action = 1
                End Select
              
        Case 3 'Invoice Criteria Selection for Invoice No in Range
                Select Case cboINVPrintOpt.ListIndex
                    Case 0 ' Invoice Print Option for Print New
                       MsgBox "Please select 'Reprint' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
                    
                    Case 1
'                       tmpBDate = lvwInvoice.SelectedItem.SubItems(INV_BORDXDATE)
 '                      TmpHLT = lvwInvoice.SelectedItem.SubItems(INV_HLTCODE)
  '                     tmpBatchNo = lvwInvoice.SelectedItem.SubItems(INV_BATCHNO)
  '                     TMPPayCode = lvwInvoice.SelectedItem
'                       InvNoFr = Trim(txtINVInvoiceNoFr)
 '                      InvNoTo = Trim(txtINVInvoiceNoTo)
                        
                       cRptInvoice.ReportFileName = crdPath & "\Report\InvoicesRefundReport_Reprint.rpt"
                       cRptInvoice.Destination = crptToWindow
                       cRptInvoice.WindowState = crptMaximized
                           
                       cRptInvoice.SelectionFormula = " {VIEW_InvoiceRefund_Report_View.MBMInvoiceNo} >= '" & txtINVInvoiceNoFr & "'" & " And " & _
                                                  " {VIEW_InvoiceRefund_Report_View.MBMInvoiceNo} <= '" & txtINVInvoiceNoTo & "'"
                       cRptInvoice.ParameterFields(0) = "MCOInvoiceNo;" & txtInvoiceNo & ";true"
                       'cRptInvoice.PrintReport
                        cRptInvoice.Action = 1
                End Select
    End Select

Else
    If cmdINVPrint Then
        Select Case cboINVCriteriaSel.ListIndex
            Case 0 ' Invoice Print Option for Print New
                    A = 0
                    B = 0
                    For i = 1 To lvwInvoice.listitems.count
                        Set lvwInvoice.SelectedItem = lvwInvoice.listitems(i)   'lvwMCO.ListItems(ii)
                            PrintM = lvwInvoice.SelectedItem.SubItems(10)
                            If PrintM = "YES" Then
                                A = A + 1
                            Else
                                B = B + 1
                            End If
                    Next i
                        
                    If (A = 0 And B <> 0) And cboINVPrintOpt.ListIndex = 1 Then
                        MsgBox "Please select 'Print New' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
                    Else
                        If (A <> 0 And B = 0) And cboINVPrintOpt.ListIndex = 0 Then
                            MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
                        Else
                            If (A = 0 And B <> 0) Then
                                Select Case cboINVPrintOpt.ListIndex
                                        Case 0 ' Invoice Print Option for Print New
                                            
                                            If lvwInvoice.SelectedItem.SubItems(INV_PRINTED) = "NO" Then
                                                TmpHLT = lvwInvoice.SelectedItem.SubItems(INV_HLTCODE)
                                                tmpBatchNo = lvwInvoice.SelectedItem.SubItems(INV_BATCHNO)
                                                TMPPayCode = lvwInvoice.SelectedItem
                                                SDate = Trim(txtINVBordxDate)
                                                                    
                                                ss0 = "{VIEW_Invoice_ReportRefund_PrintNew.HLTCode} = '" & TmpHLT & "'"
                                                ss = "{VIEW_Invoice_ReportRefund_PrintNew.MBMBordxDate} = '" & SDate & "'"
                                                cRptInvoice.ReportFileName = crdPath & "\Report\InvoicesRefundReport.rpt"
                                                cRptInvoice.Destination = crptToPrinter
                                                cRptInvoice.SelectionFormula = " {VIEW_Invoice_ReportRefund_PrintNew.MBMBordxDate} = date(" & Year(txtINVBordxDate) & "," & Month(txtINVBordxDate) & "," & Day(txtINVBordxDate) & ") " & " And " & _
                                                                           " {VIEW_Invoice_ReportRefund_PrintNew.HLTCode} = '" & TmpHLT & "'" & "And " & _
                                                                           " {VIEW_Invoice_ReportRefund_PrintNew.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                                                           " {VIEW_Invoice_ReportRefund_PrintNew.MBMBatchNo} = " & tmpBatchNo & ""
                                                                           
                                                cRptInvoice.ParameterFields(0) = "MCOInvoiceNo;" & (TxtInvoiceYear & txtInvoiceNo) & ";true"
                                               ' cRptINV.Action = 1
                                                If bPrintPass = False And CancelPrinting = False Then
                                                    bPrintPass = PrintFunction
                                                    If bPrintPass = False Then
                                                        Screen.MousePointer = Default
                                                        Exit Sub
                                                    End If
                                                Else
                                                    cRptInvoice.PrintReport
                                                End If
                                                
                                                
                                                If bPrintPass = True And CancelPrinting = False Then
                                                    'Update MBMPrinting
                                                    For ii = 1 To UBound(BDateArray)
                                                        UpdateINVPrintingFile BDateArray(ii), BatchNoArray(ii), PaycodeArray(ii), HLTcodeArray(ii), MemberNoArray(ii), "INV"
                                                    Next ii
                                                    Call cmdINVSearch_Click
                                                End If
                                            End If
                                        
                                        Case 1
                                            
                                            If lvwInvoice.SelectedItem.SubItems(INV_PRINTED) = "YES" Then
                                                TmpHLT = lvwInvoice.SelectedItem.SubItems(INV_HLTCODE)
                                                tmpBatchNo = lvwInvoice.SelectedItem.SubItems(INV_BATCHNO)
                                                TMPPayCode = lvwInvoice.SelectedItem
                                                SDate = Trim(txtINVBordxDate)
                                                                        
                                                ss0 = "{VIEW_Invoice_ReportRefund_View.HLTCode} = '" & TmpHLT & "'"
                                                ss = "{VIEW_Invoice_ReportRefund_View.MBMBordxDate} = '" & SDate & "'"
                                                cRptInvoice.ReportFileName = crdPath & "\Report\InvoicesRefundReport_Reprint.rpt"
                                                cRptInvoice.Destination = crptToPrinter
                                                cRptInvoice.SelectionFormula = " {VIEW_InvoiceRefund_Report_View.MBMBordxDate} = date(" & Year(txtINVBordxDate) & "," & Month(txtINVBordxDate) & "," & Day(txtINVBordxDate) & ") " & " And " & _
                                                                           " {VIEW_InvoiceRefund_Report_View.HLTCode} = '" & TmpHLT & "'" & "And " & _
                                                                           " {VIEW_InvoiceRefund_Report_View.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                                                           " {VIEW_InvoiceRefund_Report_View.MBMBatchNo} = " & tmpBatchNo & ""
                                                
                                                cRptInvoice.ParameterFields(0) = "MCOInvoiceNo;" & txtInvoiceNo & ";true"
                                                If bPrintPass = False And CancelPrinting = False Then
                                                    bPrintPass = PrintFunction
                                                    If bPrintPass = False Then
                                                        Screen.MousePointer = Default
                                                        Exit Sub
                                                    End If
                                                Else
                                                    cRptInvoice.PrintReport
                                                End If
                                                
                                                'cRptINV.Action = 1
                                            End If
                                End Select
                            Else
                                TmpHLT = lvwInvoice.SelectedItem.SubItems(INV_HLTCODE)
                                tmpBatchNo = lvwInvoice.SelectedItem.SubItems(INV_BATCHNO)
                                TMPPayCode = lvwInvoice.SelectedItem
                                SDate = Trim(txtINVBordxDate)
                                            
                                ss0 = "{VIEW_InvoiceRefund_Report_View.HLTCode} = '" & TmpHLT & "'"
                                ss = "{VIEW_InvoiceRefund_Report_View.MBMBordxDate} = '" & SDate & "'"
                                cRptInvoice.ReportFileName = crdPath & "\Report\InvoicesRefundReport_Reprint.rpt"
                                cRptInvoice.Destination = crptToPrinter
                                cRptInvoice.SelectionFormula = " {VIEW_InvoiceRefund_Report_View.MBMBordxDate} = date(" & Year(txtINVBordxDate) & "," & Month(txtINVBordxDate) & "," & Day(txtINVBordxDate) & ") " & " And " & _
                                                           " {VIEW_InvoiceRefund_Report_View.HLTCode} = '" & TmpHLT & "'" & "And " & _
                                                           " {VIEW_InvoiceRefund_Report_View.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                                           " {VIEW_InvoiceRefund_Report_View.MBMBatchNo} = " & tmpBatchNo & ""
                    
                                cRptInvoice.ParameterFields(0) = "MCOInvoiceNo;" & (TxtInvoiceYear & txtInvoiceNo) & ";true"
                                If bPrintPass = False And CancelPrinting = False Then
                                    bPrintPass = PrintFunction
                                    If bPrintPass = False Then
                                        Screen.MousePointer = Default
                                        Exit Sub
                                    Else
                                        cRptInvoice.Action = 1
                                    End If
                                Else
                                    cRptInvoice.PrintReport
                                End If
                                'cRptINV.Action = 1
                            
                            End If
                        End If
                    End If
                   
            Case 1
                A = 0
                B = 0
                For i = 1 To lvwInvoice.listitems.count
                    Set lvwInvoice.SelectedItem = lvwInvoice.listitems(i)   'lvwMCO.ListItems(ii)
                        PrintM = lvwInvoice.SelectedItem.SubItems(10)
                        'If PrintM = "YES" Or PrintM = True Then
                        '    A = A + 1
                        'Else
                        '    B = B + 1
                        'End If
                        If PrintM = "No" Then
                            B = B + 1
                        Else
                            A = A + 1
                        End If
                Next i
                
                If (A = 0 And B <> 0) And cboINVPrintOpt.ListIndex = 1 Then
                    MsgBox "Please select 'Print New' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
                Else
                    If (A <> 0 And B = 0) And cboINVPrintOpt.ListIndex = 0 Then
                        MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
                    Else
                        If (A <> 0 And B <> 0) Or ((A <> 0 And B = 0) And cboINVPrintOpt.ListIndex = 1) Then
                            Select Case cboINVPrintOpt.ListIndex
                                Case 0
                                    Call IncInvoiceNo
                            
                                    tmpBDate = lvwInvoice.SelectedItem.SubItems(INV_BORDXDATE)
                                    TmpHLT = lvwInvoice.SelectedItem.SubItems(INV_HLTCODE)
                                    tmpBatchNo = lvwInvoice.SelectedItem.SubItems(INV_BATCHNO)
                                    TMPPayCode = lvwInvoice.SelectedItem
                                    DateFrom = Trim(txtINVBordxDateFr)
                                    DateTo = Trim(txtINVBordxDateTo)
                                    
                                    
                                    
                                    cRptInvoice.ReportFileName = crdPath & "\Report\InvoicesRefundReport_Range.rpt"
                                    cRptInvoice.Destination = crptToPrinter
                                    cRptInvoice.WindowState = crptMaximized
                                    
                                    cRptInvoice.SelectionFormula = " {VIEW_Invoice_ReportRefund_PrintNew.MBMBordxDate} >= date(" & Year(txtINVBordxDateFr) & "," & Month(txtINVBordxDateFr) & "," & Day(txtINVBordxDateFr) & ") " & " And " & _
                                                               " {VIEW_Invoice_ReportRefund_PrintNew.MBMBordxDate} <= date(" & Year(txtINVBordxDateTo) & "," & Month(txtINVBordxDateTo) & "," & Day(txtINVBordxDateTo) & ") And " & _
                                                               " {VIEW_Invoice_ReportRefund_PrintNew.PayCode} = '" & TMPPayCode & "' "
                                    
                                    
                                    cRptInvoice.ParameterFields(0) = "MCOInvoiceNo;" & TxtInvoiceYear & txtInvoiceNo & ";true"
                                    'cRptINV.Action = 1
                                    If bPrintPass = False And CancelPrinting = False Then
                                        bPrintPass = PrintFunction
                                        If bPrintPass = False Then
                                            Screen.MousePointer = Default
                                            Exit Sub
                                        End If
                                    Else
                                        cRptInvoice.PrintReport
                                    End If
                                    
                                    'Update BordxPrinting
                                    For ii = 1 To UBound(BDateArray)
                                        UpdateINVPrintingFile BDateArray(ii), BatchNoArray(ii), PaycodeArray(ii), HLTcodeArray(ii), MemberNoArray(ii), "INV"
                                    Next ii
                                    
                                    Call cmdINVSearch_Click
                                
                                Case 1
                            
                                    tmpBDate = lvwInvoice.SelectedItem.SubItems(INV_BORDXDATE)
                                    TmpHLT = lvwInvoice.SelectedItem.SubItems(INV_HLTCODE)
                                    tmpBatchNo = lvwInvoice.SelectedItem.SubItems(INV_BATCHNO)
                                    TMPPayCode = lvwInvoice.SelectedItem
                                    DateFrom = Trim(txtINVBordxDateFr)
                                    DateTo = Trim(txtINVBordxDateTo)
                                
                                    cRptInvoice.ReportFileName = crdPath & "\Report\InvoicesRefundReport_View_Range.rpt"
                                    cRptInvoice.Destination = crptToPrinter
                                    
                                    cRptInvoice.SelectionFormula = " {VIEW_Invoice_ReportRefund_View.MBMBordxDate} >= date(" & Year(txtINVBordxDateFr) & "," & Month(txtINVBordxDateFr) & "," & Day(txtINVBordxDateFr) & ") " & " And " & _
                                                               " {VIEW_Invoice_ReportRefund_View.MBMBordxDate} <= date(" & Year(txtINVBordxDateTo) & "," & Month(txtINVBordxDateTo) & "," & Day(txtINVBordxDateTo) & ") AND " & _
                                                               " {VIEW_Invoice_ReportRefund_View.PayCode} = '" & TMPPayCode & "' "
                                    
                                    cRptInvoice.ParameterFields(0) = "MCOInvoiceNo;" & txtInvoiceNo & ";true"
                                    'cRptINV.Action = 1
                                    If bPrintPass = False And CancelPrinting = False Then
                                        bPrintPass = PrintFunction
                                        If bPrintPass = False Then
                                            Screen.MousePointer = Default
                                            Exit Sub
                                        Else
                                            cRptInvoice.PrintReport
                                        End If
                                    Else
                                        cRptInvoice.PrintReport
                                    End If
                                    
                                End Select
                        End If
                    End If
                End If
                
            Case 2
                    Select Case cboINVPrintOpt.ListIndex
                        Case 0
                            MsgBox "Please select 'Reprint' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
                        
                        Case 1
                            
                            tmpBDate = lvwInvoice.SelectedItem.SubItems(INV_BORDXDATE)
                            TmpHLT = lvwInvoice.SelectedItem.SubItems(INV_HLTCODE)
                            tmpBatchNo = lvwInvoice.SelectedItem.SubItems(INV_BATCHNO)
                            TMPPayCode = lvwInvoice.SelectedItem
                            INVNo = Trim(txtINVInvoiceNo)
                            
                            cRptInvoice.ReportFileName = crdPath & "\Report\InvoicesRefundReport_Reprint.rpt"
                            'cRptINV.RetrieveSQLQuery = txtINVInvoiceNo
                            cRptInvoice.Destination = crptToPrinter
                               
                            cRptInvoice.SelectionFormula = " {VIEW_InvoiceRefund_Report_View.MBMInvoiceNo} = '" & txtINVInvoiceNo & "'"
                            
                            cRptInvoice.ParameterFields(0) = "MCOInvoiceNo;" & txtInvoiceNo & ";true"
                            'cRptINV.Action = 1
                            If bPrintPass = False And CancelPrinting = False Then
                                bPrintPass = PrintFunction
                                If bPrintPass = False Then
                                    Screen.MousePointer = Default
                                    Exit Sub
                                End If
                            Else
                                cRptInvoice.PrintReport
                            End If
                    End Select
            Case 3
                    Select Case cboINVPrintOpt.ListIndex
                        Case 0
                            MsgBox "Please select 'Reprint' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
                        Case 1
                            tmpBDate = lvwInvoice.SelectedItem.SubItems(INV_BORDXDATE)
                            TmpHLT = lvwInvoice.SelectedItem.SubItems(INV_HLTCODE)
                            tmpBatchNo = lvwInvoice.SelectedItem.SubItems(INV_BATCHNO)
                            TMPPayCode = lvwInvoice.SelectedItem
                            InvNoFr = Trim(txtINVInvoiceNoFr)
                            InvNoTo = Trim(txtINVInvoiceNoTo)
                            
                            cRptInvoice.ReportFileName = crdPath & "\Report\InvoicesRefundReport_Reprint.rpt"
                            cRptInvoice.Destination = crptToPrinter
                               
                            cRptInvoice.SelectionFormula = " {VIEW_InvoiceRefund_Report_View.MBMInvoiceNo} >= '" & txtINVInvoiceNoFr & "'" & " And " & _
                                                       " {VIEW_InvoiceRefund_Report_View.MBMInvoiceNo} <= '" & txtINVInvoiceNoTo & "'"
                            cRptInvoice.ParameterFields(0) = "MCOInvoiceNo;" & txtInvoiceNo & ";true"
                            'cRptINV.Action = 1
                            If bPrintPass = False And CancelPrinting = False Then
                                bPrintPass = PrintFunction
                                If bPrintPass = False Then
                                    Screen.MousePointer = Default
                                    Exit Sub
                                End If
                            Else
                                cRptInvoice.PrintReport
                            End If
                    End Select
        End Select
    End If
End If
Screen.MousePointer = Default
End Sub

Private Sub cmdINVPrintPreview_Click()
   ' On Error Resume Next
    If cboINVPrintOpt.ListIndex = 0 Then
        InvoicePrint "NO"
    Else
        InvoicePrint "YES"
    End If
End Sub

Private Sub cmdINVSearch_Click()
'On Error Resume Next
Dim errorFound As Boolean
errorFound = True
If cboINVCriteriaSel.ListIndex = 0 Then
    If txtINVBordxDate = "__/__/____" Then
        MsgBox "Please Enter Bordereaux Date!", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf txtINVBatchNo = "" Then
        MsgBox "Please Enter Batch Number!", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboINVPayor = "" Then
        MsgBox "Please Select a Payor!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        errorFound = False
    End If
ElseIf cboINVCriteriaSel.ListIndex = 1 Then
    If txtINVBordxDateFr = "__/__/____" Or txtINVBordxDateTo = "__/__/____" Then
        MsgBox "Please enter valid Bordereaux Range !", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboINVPayor = "" Then
        MsgBox "Please Select a Payor!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        errorFound = False
    End If
ElseIf cboINVCriteriaSel.ListIndex = 2 Then
    If txtINVInvoiceNo = "" Then
        MsgBox "Please enter Invoice Number!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        errorFound = False
    End If
ElseIf cboINVCriteriaSel.ListIndex = 3 Then
    If txtINVInvoiceNoFr = "" Or txtINVInvoiceNoTo = "" Then
        MsgBox "Please enter valid Invoice Range!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        errorFound = False
    End If
End If
If errorFound = False Then
    TotalSupplementary = 0
    txtINVTotalPrincipal = 0
    cmdINVClose.Enabled = False
    RecordSearching_INV lvwInvoice
    cmdINVClose.Enabled = True
End If

'If cboINVPrintOpt = "" Then
 '   MsgBox "Please Enter Print Option !", vbOKOnly + vbCritical, DialogBoxTitle
  '  Exit Sub
'Else
 '   If cboINVPrintSeq = "" Then
  '      MsgBox "Please Enter Sequence !", vbOKOnly + vbCritical, DialogBoxTitle
   '     Exit Sub
'    Else
 '       If cboINVListingVer = "" Then
  '          MsgBox "Please Enter Leaflet Version !", vbOKOnly + vbCritical, DialogBoxTitle
   ''         Exit Sub
     '   End If
    'End If
'End If
  '
'Select Case SSTab1.Tab
 '   Case 2  'Invoice
        'cmdINVClose.Enabled = False
        'RecordSearching_INV lvwInvoice
        'cmdINVClose.Enabled = True
  '  End Select
End Sub

Private Sub INVSuppListing()
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
Dim tempMBMCNumber, tempCoverID
'Select Case SSTab1.Tab
'    Case 2 ' Invoice Tab
        Select Case cboINVCriteriaSel.ListIndex
            Case 0
                    SDate = CStr(Month(txtINVBordxDate)) & "/" & CStr(Day(txtINVBordxDate)) & "/" & CStr(Year(txtINVBordxDate))
                    strsql = "SELECT * FROM VIEW_MBMCovRef_RptSearch "
                    strsql = strsql & " WHERE MBMCBordxDate = '" & SDate & "'"
                    strsql = strsql & " AND MBMCBatchNo = '" & Trim(txtINVBatchNo) & "'"
                    strsql = strsql & " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    strsql = strsql & " order by  MBMCCoverID"
                    
                   ' strsql = strsql & " or MBMInvoiceNo = '" & Trim(INVNo) & "'"
                    
                    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            Case 1
                    DateFrom = CStr(Month(txtINVBordxDateFr)) & "/" & CStr(Day(txtINVBordxDateFr)) & "/" & CStr(Year(txtINVBordxDateFr))
                    DateTo = CStr(Month(txtINVBordxDateTo)) & "/" & CStr(Day(txtINVBordxDateTo)) & "/" & CStr(Year(txtINVBordxDateTo))
                    strsql = "SELECT * FROM VIEW_MBMCovRef_RptSearch "
                    strsql = strsql & " WHERE MBMCBordxDate >= '" & DateFrom & "'"
                    strsql = strsql & " AND MBMCBordxDate <= '" & DateTo & "'"
                    strsql = strsql & " AND PAYCode = '" & Mid(cboINVPayor, 1, 2) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    strsql = strsql & " AND MBMRefInvoiceNo = '" & Trim(INVNo) & "'"
                    
                    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        
            Case 2
                    strsql = "SELECT * FROM VIEW_Card_Export_Search "
                    strsql = strsql & " WHERE MBMInvoiceNo = '" & Trim(INVNo) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    
                    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                    
            Case 3
                    strsql = "SELECT * FROM VIEW_MBMCovRef_RptSearch "
                    strsql = strsql & " WHERE MBMRefInvoiceNo = '" & Trim(INVNo) & "'"
                    strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                    
                    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                    
        End Select
'End Select
               Do Until MBMCov.EOF
                   If tempMBMCNumber <> MBMCov!MBMCNumber Then
                        Set ListMaster = lvwInvoice.listitems.Add(, , PAYCode)
                         ListMaster.SubItems(1) = MBMCov!MBMCNumber
                         ListMaster.SubItems(2) = MBMCov!MBMCCoverID
                         ListMaster.SubItems(3) = MBMCov!MBMCRELCode
                         ListMaster.SubItems(4) = lvwInvoice.SelectedItem.SubItems(4)
                         ListMaster.SubItems(5) = lvwInvoice.SelectedItem.SubItems(5)
                         ListMaster.SubItems(6) = lvwInvoice.SelectedItem.SubItems(6)
                         ListMaster.SubItems(7) = lvwInvoice.SelectedItem.SubItems(7)
                         ListMaster.SubItems(8) = MBMCov!MBMCName
                     
                        If MBMCov!MBMGroupCompany = "" Then
                             ListMaster.SubItems(9) = ""
                         Else
                             If Len(MBMCov!MBMGroupCompany) Then
                                 ListMaster.SubItems(9) = MBMCov!MBMGroupCompany
                             End If
                         End If
                     
                        If Len(MBMCov!MBMRefInvoiceNo) Then
                             ListMaster.SubItems(11) = MBMCov!MBMRefInvoiceNo
                         End If
                     
                        TotalSupplementary = TotalSupplementary + 1
                        txtINVTotalSupplementary1 = TotalSupplementary
 '                       Select Case SSTab1.Tab
 '                       Case 2 'Leaflet Printing
                            itemToPrint = INV_PRINTED
                            cmdINVPrint.Enabled = True
                            cmdINVPrintPreview.Enabled = True
                        'End Select
                        
                        If Printed = "YES" Then
                            ListMaster.SubItems(itemToPrint) = "YES"
                        Else
                            ListMaster.SubItems(itemToPrint) = "NO"
                        End If
                        
                        'rsPrint.Close
                        tempMBMCNumber = MBMCov!MBMCNumber
                        tempCoverID = MBMCov!MBMCCoverID
                     Else
                        If tempCoverID <> MBMCov!MBMCCoverID Then
                         Set ListMaster = lvwInvoice.listitems.Add(, , PAYCode)
                         ListMaster.SubItems(1) = MBMCov!MBMCNumber
                         ListMaster.SubItems(2) = MBMCov!MBMCCoverID
                         ListMaster.SubItems(3) = MBMCov!MBMCRELCode
                         ListMaster.SubItems(4) = lvwInvoice.SelectedItem.SubItems(4)
                         ListMaster.SubItems(5) = lvwInvoice.SelectedItem.SubItems(5)
                         ListMaster.SubItems(6) = lvwInvoice.SelectedItem.SubItems(6)
                         ListMaster.SubItems(7) = lvwInvoice.SelectedItem.SubItems(7)
                         ListMaster.SubItems(8) = MBMCov!MBMCName
                     
                        If MBMCov!MBMGroupCompany = "" Then
                             ListMaster.SubItems(9) = ""
                         Else
                             If Len(MBMCov!MBMGroupCompany) Then
                                 ListMaster.SubItems(9) = MBMCov!MBMGroupCompany
                             End If
                         End If
                     
                        If Len(MBMCov!MBMRefInvoiceNo) Then
                             ListMaster.SubItems(11) = MBMCov!MBMRefInvoiceNo
                         End If
                     
                        TotalSupplementary = TotalSupplementary + 1
                        txtINVTotalSupplementary1 = TotalSupplementary
                        'Select Case SSTab1.Tab
                        'Case 2 'Leaflet Printing
                            itemToPrint = INV_PRINTED
                            cmdINVPrint.Enabled = True
                            cmdINVPrintPreview.Enabled = True
                        'End Select
                        
                        If Printed = "YES" Then
                            ListMaster.SubItems(itemToPrint) = "YES"
                        Else
                            ListMaster.SubItems(itemToPrint) = "NO"
                        End If
                        
                        'rsPrint.Close
                        tempMBMCNumber = MBMCov!MBMCNumber
                        tempCoverID = MBMCov!MBMCCoverID
                        End If
                    End If
                    MBMCov.MoveNext
                   Loop
            
                MBMCov.Close
                'Set rsPrint = Nothing
                Set MBMCov = Nothing

End Sub

Sub ComboListing()
    Dim PAY As New ADODB.Recordset
    Dim sqlstr As String
   
    sqlstr = "select PAYCode , PAYCompanyName from PAY "
    
    PAY.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do While Not PAY.EOF
        'If SSTab1.Tab = 0 Then
        '    SSTab1.TabEnabled(0) = True
        '    cboLLPayor.AddItem (PAY!PAYCode) & " - " & (PAY!PAYCompanyName)
        '    PAY.MoveNext
        'Else
        '    If SSTab1.Tab = 1 Then
        '        SSTab1.TabEnabled(1) = True
        '        cboMCOPayor.AddItem (PAY!PAYCode) & " - " & (PAY!PAYCompanyName)
        '        PAY.MoveNext
        '    Else
        '        If SSTab1.Tab = 2 Then
        '           SSTab1.TabEnabled(2) = True
                    cboINVPayor.AddItem (PAY!PAYCode) & " - " & (PAY!PAYCompanyName)
                    PAY.MoveNext
        '        Else
        '            If SSTab1.Tab = 3 Then
        '                SSTab1.TabEnabled(3) = True
        '                cboExpPayor.AddItem (PAY!PAYCode) & " - " & (PAY!PAYCompanyName)
        '                PAY.MoveNext
        '            End If
        '        End If
        '    End If
        'End If
    Loop
    
    PAY.Close
    Set PAY = Nothing
End Sub


Private Sub cmdMCOClose_Click()
Dim RecordExit
    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        bExit = True
        Unload Me
    End If
End Sub

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
'    Case 0
'        If cRptLL Is Nothing Then Exit Function
'
'        With dlgCommonDialog
'            .DialogTitle = "Print Crystal Report "
'            .CancelError = True
'            .Flags = cdlPDNoSelection + cdlPDNoPageNums
'            .ShowPrinter
'
'            If ERR <> MSComDlg.cdlCancel Then
'                cRptLL.PrintReport
'            End If
'        End With
'        PrintFunction = True
'        Exit Function
'cancelPrint:
'        PrintFunction = False
'        CancelPrinting = True
'
'    Case 1
'        If cRptMCO Is Nothing Then Exit Function
'
'            With dlgCommonDialog_MCO
'                .DialogTitle = "Print Crystal Report "
'                .CancelError = True
'                .Flags = cdlPDNoSelection + cdlPDNoPageNums
'                .ShowPrinter
'
'                If ERR <> MSComDlg.cdlCancel Then
'                    cRptMCO.PrintReport
'                End If
'            End With
'            PrintFunction = True
'            Exit Function
'cancelPrint1:
'                PrintFunction = False
'                CancelPrinting = True
'                Exit Function

'    Case 2
        If cRptInvoice Is Nothing Then Exit Function
    
            With dlgCommonDialog_INV
                .DialogTitle = "Print Crystal Report "
                .CancelError = True
                 'For ii = 1 To UBound(MemberNoArray)
                    .Flags = cdlPDNoSelection + cdlPDNoPageNums
                    'If ActiveForm.rtfText.SelLength = 0 Then
                     '   .Flags = .Flags + cdlPDAllPages
                    'Else
                     '   .Flags = .Flags + cdlPDSelection
                    'End If
                            
                .ShowPrinter
                
                If ERR <> MSComDlg.cdlCancel Then
                    If cboINVPrintOpt.ListIndex = 0 Then
                        Call IncInvoiceNo
                        cRptInvoice.ParameterFields(0) = "MCOInvoiceNo;" & (TxtInvoiceYear & txtInvoiceNo) & ";true"
                        cRptInvoice.PrintReport
                    Else
                        If cboINVPrintOpt.ListIndex = 1 Then
                            cRptInvoice.PrintReport
                        End If
                    End If
                End If
            ' Next ii
            End With
            PrintFunction = True
            Exit Function
cancelPrint:
                PrintFunction = False
                CancelPrinting = True
                Exit Function
                
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

Private Sub lvwInvoice_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwInvoice.SortKey = ColumnHeader.Index - 1
    lvwInvoice.Sorted = True
    If lvwInvoice.SortOrder = lvwAscending Then
       lvwInvoice.SortOrder = lvwDescending
    Else
       lvwInvoice.SortOrder = lvwAscending
    End If
End Sub

Private Sub UpdateINVPrintingFile(BordxDate As String, BatchNo As String, PAYCode As String, HLTCode As String, MemberNo As String, WhichField As String)
Dim rsUpdate As New ADODB.Recordset
Dim rsMBMUpdate As New ADODB.Recordset
Dim sql As String
Dim strsql As String

    
  'Update file for BordxPrinting
    sql = "SELECT * FROM MBMRefund"
    sql = sql & " WHERE MBMBordxDate = '" & Format(CDate(BordxDate), "MM/DD/YYYY") & "'"
    sql = sql & " AND MBMBatchNo = " & BatchNo & ""
    sql = sql & " AND PAYCode = '" & PAYCode & "'"
    sql = sql & " AND HLTCode = '" & Trim(HLTCode) & "'"
    sql = sql & " AND MBMNumber = '" & MemberNo & "'"
    rsUpdate.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
 'rsUpdate.Close
    'update file for MBMPrinting
   'strsql = "SELECT * FROM MBMPrinting"
   ' strsql = strsql & " WHERE MBMBordxDate = '" & Format(CDate(BordxDate), "MM/DD/YYYY") & "'"
   ' strsql = strsql & " AND MBMBatchNo = " & BatchNo & ""
   ' strsql = strsql & " AND PAYCode = '" & PAYCode & "'"
   ' strsql = strsql & " AND HLTCode = '" & Trim(HLTCode) & "'"
    'strsql = strsql & " AND MBMNumber = '" & MemberNo & "'"
 
    'rsMBMUpdate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic, adCmdText
 
    'AddNew file for MBMPrinting
    If rsUpdate.recordCount = 0 Then
        rsUpdate.AddNew
        rsUpdate!MBMBordxDate = BordxDate
        rsUpdate!MBMBatchNo = BatchNo
        rsUpdate!PAYCode = PAYCode
        rsUpdate!HLTCode = Trim(HLTCode)
        rsUpdate!MBMNumber = MemberNo

        Select Case WhichField
            Case "INV"
                rsUpdate!MBMRefInvoiced = "1"
                rsUpdate!MBMRefInvoiceDate = Date
                rsUpdate!MBMRefInvoiceNo = TxtInvoiceYear & txtInvoiceNo
 
        End Select
        rsUpdate.Update
    Else 'Update file for MBMPrinting
        Select Case WhichField
            Case "INV"
                rsUpdate!MBMRefInvoiced = "1"
                rsUpdate!MBMRefInvoiceDate = Date
                rsUpdate!MBMRefInvoiceNo = TxtInvoiceYear & txtInvoiceNo
           
        End Select
        rsUpdate.Update
    End If
             
    'AddNew file for BordxPrinting
    If rsUpdate.recordCount = 0 Then
        rsUpdate.AddNew
        rsUpdate!MBMBordxDate = BordxDate
        rsUpdate!MBMBatchNo = BatchNo
        rsUpdate!PAYCode = PAYCode
        rsUpdate!HLTCode = Trim(HLTCode)
    
        Select Case WhichField
            Case "INV"
                 rsUpdate!MBMInvoiced = "1"
                 rsUpdate!MBMInvoiceDate = Date
                 rsUpdate!MBMInvoiceNo = TxtInvoiceYear & txtInvoiceNo
        End Select
        rsUpdate.Update
    Else 'Update file for BordxPrinting
        Select Case WhichField
            Case "INV"
                rsUpdate!MBMRefInvoiced = "1"
                rsUpdate!MBMRefInvoiceDate = Date
                rsUpdate!MBMRefInvoiceNo = TxtInvoiceYear & txtInvoiceNo
        End Select
        rsUpdate.Update
    End If
    rsUpdate.Close
    'rsMBMUpdate.Close

End Sub

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



