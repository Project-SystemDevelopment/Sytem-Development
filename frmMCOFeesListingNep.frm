VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMCOFeesListingNep 
   Caption         =   "Fee Listing (Neptune)"
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
   Begin VB.Frame FrameMCO 
      Caption         =   "Enter Criteria For Searching"
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   240
      Width           =   11055
      Begin VB.ComboBox cboMCOCriteriaSel 
         Height          =   315
         ItemData        =   "frmMCOFeesListingNep.frx":0000
         Left            =   2160
         List            =   "frmMCOFeesListingNep.frx":000A
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
      Begin MSMask.MaskEdBox txtMCOBordxDate 
         Height          =   315
         Left            =   7560
         TabIndex        =   3
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
         TabIndex        =   6
         Top             =   555
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
         TabIndex        =   5
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
         TabIndex        =   2
         Top             =   650
         Width           =   4215
      End
      Begin VB.ComboBox cboMCOListingVer 
         Height          =   315
         ItemData        =   "frmMCOFeesListingNep.frx":0032
         Left            =   2160
         List            =   "frmMCOFeesListingNep.frx":003C
         TabIndex        =   19
         Top             =   930
         Width           =   4215
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
         TabIndex        =   27
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lbl20 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   680
         Width           =   1095
      End
      Begin VB.Label lbl22 
         Caption         =   "Listing Versions"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbl19 
         Caption         =   "To"
         Height          =   255
         Left            =   6720
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lbl18 
         Caption         =   "From"
         Height          =   255
         Left            =   6720
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbl15 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   6600
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Batch No"
         Height          =   255
         Left            =   6600
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lbl16 
         Caption         =   "Date "
         Height          =   255
         Left            =   6960
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   7200
      Width           =   11055
      Begin VB.CommandButton cmdMCOSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   300
         Left            =   6720
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOPrint 
         Caption         =   "&Print"
         Height          =   300
         Left            =   7560
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMCOPrintPreview 
         Caption         =   "Pre&view"
         Height          =   300
         Left            =   8400
         TabIndex        =   9
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
      Begin VB.CommandButton cmdMCOClose 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   300
         Left            =   10080
         TabIndex        =   11
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
      Height          =   5640
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   11055
      Begin MSComctlLib.ListView lvwMCO 
         Height          =   5250
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   9260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Payor Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Membership No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Cover ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Relationship"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Bordereaux Date"
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
            Text            =   "Health Type"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Batch No"
            Object.Width           =   1764
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
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Fee Listing (Neptune)"
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
      TabIndex        =   28
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmMCOFeesListingNep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim sDate As String
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
Dim MCONull As Boolean
Private m_blnDirection(1 To 10) As Boolean

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
        cboMCOListingVer.ListIndex = 0
        lvwMCO.listitems.Clear
        cmdMCOPrint.Enabled = False
        cmdMCOPrintPreview.Enabled = False
        txtMCOTotalSupplementary1 = 0
        txtMCOTotalPrincipal = 0
    
    Case 1
        lbl15.Visible = False
        lbl16.Visible = False
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

Private Sub cmdMCOClear_Click()
    txtMCOBatchNo = ""
    cboMCOPayor = ""
    txtMCOBordxDate = "__/__/____"
    txtMCOBordxDateFr = "__/__/____"
    txtMCOBordxDateTo = "__/__/____"
    cboMCOListingVer.ListIndex = 0
    lvwMCO.listitems.Clear
    cmdMCOPrint.Enabled = False
    cmdMCOPrintPreview.Enabled = False
    txtMCOTotalSupplementary1 = 0
    txtMCOTotalPrincipal = 0
End Sub

Private Sub cmdMCOClose_Click()
    Unload Me
End Sub

Private Sub cmdMCOPrint_Click()
    MCOPrint "ALL"
End Sub

Private Sub cmdMCOPrintPreview_Click()
    MCOPrint "ALL"
End Sub

Private Sub cmdMCOSearch_Click()
    
    txtMCOTotalPrincipal = "0"
    txtMCOTotalSupplementary1 = "0"
    
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
    
    If cboMCOListingVer = "" Then
        MsgBox "Please Enter Leaflet Version !", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If

    RecordSearching_MCO lvwMCO
    
End Sub

Private Sub RecordSearching_MCO(objListView As Object)
    Screen.MousePointer = vbHourglass
        Call MCOPrincipalListing
    Screen.MousePointer = Default
End Sub

Private Sub MCOPrincipalListing()
Dim rsSearch As New ADODB.Recordset
Dim rsPrint As New ADODB.Recordset
Dim ColumnItem As ListItem
Dim Sql As String
Dim itemToPrint As Integer
Dim isPrinted As Boolean
Dim CoverQty As Integer
Dim MCOCriteria As String
Dim TotalPrincipal As Integer

TotalSupplementary = 0
MCONull = False

Select Case cboMCOCriteriaSel.ListIndex
    Case 0 'Criteria Selection for MCO on Bordx Batch
            sDate = CStr(Month(txtMCOBordxDate)) & "/" & CStr(Day(txtMCOBordxDate)) & "/" & CStr(Year(txtMCOBordxDate))
            lvwMCO.listitems.Clear
            
            Sql = "SELECT * FROM VIEW_MCO_ReportSearch2 "
            Sql = Sql & " WHERE MBMBordxDate = '" & sDate & "'"
            Sql = Sql & " AND MBMBatchNo = '" & txtMCOBatchNo & "'"
            Sql = Sql & " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "'"
            Sql = Sql & " ORDER BY MBMPolicyNo"
            rsSearch.Open Sql, medixneptuneconn, adOpenKeyset, adLockOptimistic, adCmdText
     
    Case 1 'Criteria Selection for MCO on Bordx Range
            DateFrom = CStr(Month(txtMCOBordxDateFr)) & "/" & CStr(Day(txtMCOBordxDateFr)) & "/" & CStr(Year(txtMCOBordxDateFr))
            DateTo = CStr(Month(txtMCOBordxDateTo)) & "/" & CStr(Day(txtMCOBordxDateTo)) & "/" & CStr(Year(txtMCOBordxDateTo))
            lvwMCO.listitems.Clear
            
            Sql = "SELECT * FROM VIEW_MCO_ReportSearch2 "
            Sql = Sql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
            Sql = Sql & " AND MBMBordxDate <= '" & DateTo & "'"
            Sql = Sql & " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "'"
            Sql = Sql & " ORDER BY MBMPolicyNo"
            rsSearch.Open Sql, medixneptuneconn, adOpenKeyset, adLockOptimistic, adCmdText
            
End Select

    If rsSearch.recordCount = 0 Then
        MsgBox "Record not found !", vbInformation + vbOKOnly
        cmdMCOPrint.Enabled = False
        Screen.MousePointer = Default
        Exit Sub
    End If
    
    Do Until rsSearch.EOF
    total = rsSearch.recordCount
    lblTotal = CInt(total)
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
                
                If Len(rsSearch!MBMMCOFee) Then
                    cmdMCOPrint.Enabled = True
                    cmdMCOPrintPreview.Enabled = True
                Else
                    MCONull = True
                    cmdMCOPrint.Enabled = False
                    cmdMCOPrintPreview.Enabled = False
                End If
                                                            
                If Trim(rsSearch!INSCode) = "F" Or Trim(rsSearch!INSCode) = "H" Then
                    Call MCOSuppListing
                End If
                    
    rsSearch.MoveNext
    Loop
    
    rsSearch.Close
    Set rsSearch = Nothing
    
    If MCONull = True Then
        MsgBox "MCO Fee is Empty!", vbOKOnly + vbCritical, DialogBoxTitle
    End If
    
End Sub

Private Sub MCOSuppListing()
Dim strsql As String
Dim Sql As String
Dim strsqls As String
Dim rsPrint As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim ListMaster As ListItem
Dim isPrinted As Boolean
Dim total As Integer
Dim Current As Integer


    Select Case cboMCOCriteriaSel.ListIndex
        Case 0
                sDate = CStr(Month(txtMCOBordxDate)) & "/" & CStr(Day(txtMCOBordxDate)) & "/" & CStr(Year(txtMCOBordxDate))
                strsql = "SELECT * FROM VIEW_MCO_SuppReportSearch "
                strsql = strsql & " WHERE MBMBordxDate = '" & sDate & "'"
                strsql = strsql & " AND MBMBatchNo = '" & Trim(txtMCOBatchNo) & "'"
                strsql = strsql & " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "'"
                strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                
                MBMCov.Open strsql, medixneptuneconn, adOpenKeyset, adLockOptimistic
        
        Case 1
                DateFrom = CStr(Month(txtMCOBordxDateFr)) & "/" & CStr(Day(txtMCOBordxDateFr)) & "/" & CStr(Year(txtMCOBordxDateFr))
                DateTo = CStr(Month(txtMCOBordxDateTo)) & "/" & CStr(Day(txtMCOBordxDateTo)) & "/" & CStr(Year(txtMCOBordxDateTo))
                strsql = "SELECT * FROM VIEW_MCO_SuppReportSearch "
                strsql = strsql & " WHERE MBMBordxDate >= '" & DateFrom & "'"
                strsql = strsql & " AND MBMBordxDate <= '" & DateTo & "'"
                strsql = strsql & " AND PAYCode = '" & Mid(cboMCOPayor, 1, 2) & "'"
                strsql = strsql & " AND MBMCNumber = '" & Trim(MBMNo) & "'"
                
                MBMCov.Open strsql, medixneptuneconn, adOpenKeyset, adLockOptimistic
    End Select

        
               Do Until MBMCov.EOF
                   Set ListMaster = lvwMCO.listitems.Add(, , PAYCode)
                        ListMaster.SubItems(1) = MBMCov!MBMCNumber
                        ListMaster.SubItems(2) = MBMCov!MBMCCoverID
                        ListMaster.SubItems(3) = MBMCov!MBMCRELCode
                        ListMaster.SubItems(4) = MBMCov!MBMCBordxDate
                        ListMaster.SubItems(5) = MBMCov!MBMPolicyNo
                        ListMaster.SubItems(6) = lvwMCO.SelectedItem.SubItems(6)
                        ListMaster.SubItems(7) = IIf(IsNumeric(MBMCov!MBMBatchNo) = True, MBMCov!MBMBatchNo, 0)
                        ListMaster.SubItems(8) = Trim(MBMCov!MBMCName)
                    
                        If MBMCov!MBMGroupCompany = "" Then
                            ListMaster.SubItems(9) = ""
                        Else
                            If Len(MBMCov!MBMGroupCompany) Then
                                ListMaster.SubItems(9) = MBMCov!MBMGroupCompany
                            End If
                        End If
                    
                       TotalSupplementary = TotalSupplementary + 1
                       txtMCOTotalSupplementary1 = TotalSupplementary
                       
                       MBMCov.MoveNext
                   Loop
            
                MBMCov.Close
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
Dim Cnt As Integer
Dim cont As Integer
Dim cout As Integer
Dim Arraycnt As Integer
Dim Arraycount As Integer
Dim PrintM As String
Dim i As Integer
Dim rstInv As New ADODB.Recordset
Dim strsql2 As String
Dim InvNoFr, InvNoTo As String

Screen.MousePointer = vbHourglass
Select Case cboMCOCriteriaSel.ListIndex
    Case 0 'MCO Criteria Selection for Bordx Batch
            Cnt = 0
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
                           Cnt = Cnt + 1
                           ReDim Preserve BDateArray(Cnt)
                           ReDim Preserve BatchNoArray(Cnt)
                           ReDim Preserve PaycodeArray(Cnt)
                           ReDim Preserve HLTcodeArray(Cnt)
                           ReDim Preserve MemberNoArray(Cnt)
                           BDateArray(Cnt) = tmpBDate
                           BatchNoArray(Cnt) = tmpBatchNo
                           PaycodeArray(Cnt) = TMPPayCode
                           HLTcodeArray(Cnt) = tmpHLTCode
                           MemberNoArray(Cnt) = tmpMemberNo
                        End If
                    Next ii
            End If
    
    Case 1 'MCO Criteria Selection for Bordx Range
            Cnt = 0
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
                           Cnt = Cnt + 1
                           ReDim Preserve BDateArray(Cnt)
                           ReDim Preserve BatchNoArray(Cnt)
                           ReDim Preserve PaycodeArray(Cnt)
                           ReDim Preserve HLTcodeArray(Cnt)
                           ReDim Preserve MemberNoArray(Cnt)
                           BDateArray(Cnt) = tmpBDate
                           BatchNoArray(Cnt) = tmpBatchNo
                           PaycodeArray(Cnt) = TMPPayCode
                           HLTcodeArray(Cnt) = tmpHLTCode
                           MemberNoArray(Cnt) = tmpMemberNo
                        End If
                    Next ii
            End If
End Select


If cmdMCOPrintPreview Then
    Call MCOListingVer
Else
    If cmdMCOPrint Then
        
        Select Case cboMCOCriteriaSel.ListIndex
            Case 0 'MCO Criteria Selection for Bordx Batch Print
                    
                    Call MCOListingVer
            
            Case 1 'MCO Criteria Selection for Bordx Range Print
                    
                    Call MCOListingVer
        
        End Select
    End If
End If
Screen.MousePointer = Default

End Sub

Private Sub MCOListingVer()
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

    Select Case cboMCOListingVer
        Case "MCO Listing"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            TMPPayCode = lvwMCO.SelectedItem
            sDate = Trim(txtMCOBordxDate)
                                 
            ss = "{VIEW_MCO_LISTING_All2.MBMBordxDate} = '" & sDate & "'"
            cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WithoutPremium_AllNep.rpt"
            cRptMCO.Destination = crptToWindow
            cRptMCO.WindowState = crptMaximized
            cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All2.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                       " {VIEW_MCO_LISTING_All2.PayCode} = '" & Mid(cboMCOPayor, 1, 2) & "'" & " And " & _
                                       " {VIEW_MCO_LISTING_All2.MBMBatchNo} = " & txtMCOBatchNo & ""
                                                    
            Call MCOListingBBPrintSeqPrev
                         
        Case "MCO Listing With Premium"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            TMPPayCode = lvwMCO.SelectedItem
            sDate = Trim(txtMCOBordxDate)
                                 
            ss = "{VIEW_MCO_LISTING_All2.MBMBordxDate} = '" & sDate & "'"
            cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_AllNep.rpt"
            cRptMCO.Destination = crptToWindow
            cRptMCO.WindowState = crptMaximized
            cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All2.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                       " {VIEW_MCO_LISTING_All2.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                       " {VIEW_MCO_LISTING_All2.MBMBatchNo} = " & tmpBatchNo & ""
                         
            Call MCOListingBBPrintSeqPrev
                         
    End Select

End Sub

Private Sub MCOListingVer_BRPrintPreV()
    
    Select Case cboMCOListingVer
        Case "MCO Listing"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            TMPPayCode = lvwMCO.SelectedItem
            DateFrom = Trim(txtMCOBordxDateFr)
            DateTo = Trim(txtMCOBordxDateTo)
                                        
            cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WithoutPremium_AllNep.rpt"
            cRptMCO.Destination = crptToWindow
            cRptMCO.WindowState = crptMaximized
            cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All2.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                       " {VIEW_MCO_LISTING_All2.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                       " {VIEW_MCO_LISTING_All2.PayCode} = '" & TMPPayCode & "'"
                                                           
            Call MCOListingBRPrintSeqPrev
                                    
                                        
        Case "MCO Listing With Premium"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            TMPPayCode = lvwMCO.SelectedItem
            DateFrom = Trim(txtMCOBordxDateFr)
            DateTo = Trim(txtMCOBordxDateTo)
                                        
            ss = "{VIEW_MCO_LISTING_All2.MBMBordxDate} = '" & sDate & "'"
            cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_AllNep.rpt"
            cRptMCO.Destination = crptToWindow
            cRptMCO.WindowState = crptMaximized
            cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All2.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                       " {VIEW_MCO_LISTING_All2.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                       " {VIEW_MCO_LISTING_All2.PayCode} = '" & TMPPayCode & "'"
            Call MCOListingBRPrintSeqPrev
    
                                
    End Select

End Sub

Private Sub MCOListingVer_BBPrint()

    Select Case cboMCOListingVer
        Case "MCO Listing"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            TMPPayCode = lvwMCO.SelectedItem
            sDate = Trim(txtMCOBordxDate)
                                        
            ss = "{VIEW_MCO_LISTING_All2.MBMBordxDate} = '" & sDate & "'"
            cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WithoutPremium_AllNep.rpt"
            cRptMCO.Destination = crptToPrinter
            cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All2.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                       " {VIEW_MCO_LISTING_All2.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                       " {VIEW_MCO_LISTING_All2.MBMBatchNo} = " & tmpBatchNo & ""
                                
            Call MCOListingBBPrintSeq
        
                                    
        Case "MCO Listing With Premium"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            TMPPayCode = lvwMCO.SelectedItem
            sDate = Trim(txtMCOBordxDate)
                                        
            ss = "{VIEW_MCO_LISTING_All2.MBMBordxDate} = '" & sDate & "'"
                                
            cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_AllNep.rpt"
            cRptMCO.Destination = crptToPrinter
            cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All2.MBMBordxDate} = date(" & Year(txtMCOBordxDate) & "," & Month(txtMCOBordxDate) & "," & Day(txtMCOBordxDate) & ") " & " And " & _
                                       " {VIEW_MCO_LISTING_All2.PayCode} = '" & TMPPayCode & "'" & " And " & _
                                       " {VIEW_MCO_LISTING_All2.MBMBatchNo} = " & tmpBatchNo & ""
                                    
            Call MCOListingBBPrintSeq
                                    
    End Select

End Sub
Private Sub MCOListingVer_BRPrint()
                                    
    Select Case cboMCOListingVer
        Case "MCO Listing"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            TMPPayCode = lvwMCO.SelectedItem
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            DateFrom = Trim(txtMCOBordxDateFr)
            DateTo = Trim(txtMCOBordxDateTo)
                                        
            cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_WithoutPremium_AllNep.rpt"
            cRptMCO.Destination = crptToPrinter
            cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All2.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                       " {VIEW_MCO_LISTING_All2.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                       " {VIEW_MCO_LISTING_All2.PayCode} = '" & TMPPayCode & "'"
                                                       
            Call MCOListingBRPrintSeq
        
                                
        Case "MCO Listing With Premium"
            TmpHLT = lvwMCO.SelectedItem.SubItems(MCO_HLTCODE)
            TMPPayCode = lvwMCO.SelectedItem
            tmpBatchNo = lvwMCO.SelectedItem.SubItems(MCO_BATCHNO)
            DateFrom = Trim(txtMCOBordxDateFr)
            DateTo = Trim(txtMCOBordxDateTo)
            
            ss = "{VIEW_MCO_LISTING_All2.MBMBordxDate} = '" & sDate & "'"
            cRptMCO.ReportFileName = crdPath & "\Report\MCOListing_Premium_AllNep.rpt"
            cRptMCO.Destination = crptToPrinter
            cRptMCO.SelectionFormula = " {VIEW_MCO_LISTING_All2.MBMBordxDate} >= date(" & Year(txtMCOBordxDateFr) & "," & Month(txtMCOBordxDateFr) & "," & Day(txtMCOBordxDateFr) & ") " & " And " & _
                                       " {VIEW_MCO_LISTING_All2.MBMBordxDate} <= date(" & Year(txtMCOBordxDateTo) & "," & Month(txtMCOBordxDateTo) & "," & Day(txtMCOBordxDateTo) & ") " & "AND " & _
                                       " {VIEW_MCO_LISTING_All2.PayCode} = '" & TMPPayCode & "'"
                                                      
            Call MCOListingBRPrintSeq
        
                                
    End Select

End Sub

Private Sub MCOListingBBPrintSeqPrev()

    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All2.MBMPolicyNo}"
    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_ALL2.MBMNumber}"
    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL2.MBMCCoverID}"
    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
    cRptMCO.ParameterFields(1) = "FromDate;" & sDate & ";true"
    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
    cRptMCO.PrintReport

End Sub

Private Sub MCOListingBRPrintSeqPrev()

    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All2.MBMPolicyNo}"
    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_ALL2.MBMNumber}"
    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL2.MBMCCoverID}"
    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
    cRptMCO.PrintReport

End Sub

Private Sub MCOListingBBPrintSeq()
                
    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All2.MBMPolicyNo}"
    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_ALL2.MBMNumber}"
    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL2.MBMCCoverID}"
    cRptMCO.ParameterFields(0) = "TFrom;" & "" & ";true"
    cRptMCO.ParameterFields(1) = "FromDate;" & sDate & ";true"
    cRptMCO.ParameterFields(2) = "ToDate;" & "Batch No :" & ";true"
    cRptMCO.ParameterFields(3) = "LtDate;" & txtMCOBatchNo & ";true"
    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
    
    If bPrintPass = False And CancelPrinting = False Then
        bPrintPass = PrintFunction
        If bPrintPass = False Then
            Screen.MousePointer = Default
            Exit Sub
        End If
    Else
        cRptMCO.PrintReport
    End If

End Sub

Private Sub MCOListingBRPrintSeq()
            
    cRptMCO.SortFields(0) = "+{VIEW_MCO_LISTING_All2.MBMPolicyNo}"
    cRptMCO.SortFields(1) = "+{VIEW_MCO_Listing_ALL2.MBMNumber}"
    cRptMCO.SortFields(2) = "+{VIEW_MCO_LISTING_ALL2.MBMCCoverID}"
    cRptMCO.ParameterFields(0) = "TFrom;" & "FROM" & ";true"
    cRptMCO.ParameterFields(1) = "FromDate;" & DateFrom & ";true"
    cRptMCO.ParameterFields(2) = "ToDate;" & "TO" & ";true"
    cRptMCO.ParameterFields(3) = "LtDate;" & DateTo & ";true"
    cRptMCO.ParameterFields(4) = "PayorName;" & cboMCOPayor & ";true"
    
    If bPrintPass = False And CancelPrinting = False Then
        bPrintPass = PrintFunction
        If bPrintPass = False Then
            Screen.MousePointer = Default
            Exit Sub
        End If
    Else
        cRptMCO.PrintReport
    End If
                    
            
End Sub

Sub ComboListing()
    Dim PAY As New ADODB.Recordset
    Dim sqlstr As String
   
    sqlstr = "select PAYCode , PAYCompanyName from PAY "
    
    PAY.Open sqlstr, medixneptuneconn, adOpenKeyset, adLockOptimistic
    
    Do While Not PAY.EOF
        cboMCOPayor.AddItem (PAY!PAYCode) & " - " & (PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    
    PAY.Close
    Set PAY = Nothing
End Sub
 
Private Sub Form_Load()
    Call ComboListing
End Sub

Private Function PrintFunction() As Boolean
On Error GoTo cancelPrint

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
    End Select
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
