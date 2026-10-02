VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMCOManualUpd 
   Caption         =   "MCO Invoice Update"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   495
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lstMBM 
      Height          =   5295
      Left            =   60
      TabIndex        =   19
      Top             =   1200
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9340
      View            =   3
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Payor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Bordx.Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Batch"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Invoice No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Inv. Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "MBMNo"
         Text            =   "Member No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Cover ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "MCO"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   60
      TabIndex        =   16
      Top             =   240
      Width           =   11775
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   4335
      End
      Begin MSMask.MaskEdBox txtBordxDate1 
         Height          =   315
         Left            =   6960
         TabIndex        =   3
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtBatchNo1 
         Height          =   285
         Left            =   6960
         MaxLength       =   3
         TabIndex        =   5
         Top             =   525
         Width           =   1335
      End
      Begin VB.TextBox TxtMBMNo2 
         Height          =   285
         Left            =   3600
         MaxLength       =   25
         TabIndex        =   2
         Top             =   525
         Width           =   1935
      End
      Begin VB.TextBox TxtMBMNo1 
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   1
         Top             =   525
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtBordxDate2 
         Height          =   315
         Left            =   8760
         TabIndex        =   4
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtBatchNo2 
         Height          =   285
         Left            =   8760
         MaxLength       =   3
         TabIndex        =   6
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   26
         Top             =   560
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   570
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Batch No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   23
         Top             =   555
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Member No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   555
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Bordx. Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Payor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame FrameUpdate 
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   6480
      Width           =   11775
      Begin VB.TextBox txtInvoiceNo 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtInvoiceDate 
         Height          =   285
         Left            =   4080
         TabIndex        =   8
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         Caption         =   "Invoice Date"
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Invoice No."
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   7080
      Width           =   11775
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   8160
         TabIndex        =   11
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   8880
         TabIndex        =   12
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "St&op"
         Height          =   315
         Left            =   7320
         TabIndex        =   10
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   315
         Left            =   9720
         TabIndex        =   13
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   315
         Left            =   10560
         TabIndex        =   14
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Height          =   315
         Left            =   6480
         TabIndex        =   9
         Top             =   200
         Width           =   855
      End
      Begin VB.Label lblMCO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   4920
         TabIndex        =   35
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label13 
         Caption         =   "Total MCO Amt"
         Height          =   255
         Left            =   3720
         TabIndex        =   34
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "of"
         Height          =   255
         Left            =   1560
         TabIndex        =   32
         Top             =   270
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   1080
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1920
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.Label Label33 
      Caption         =   "Bold - Searchable "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   7680
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   " MCO Invoice Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   220
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmMCOManualUpd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intExit As Integer
Dim StopChk As Boolean
Dim Chk0Mco As Boolean
Dim Rec0Cnt As Integer
Dim TotRecSel As Integer
Dim CurrRecProc As Integer
Dim TotalMCO As String

Private Sub ComboListing()
Dim Payor As New ADODB.Recordset

    Payor.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    CboPayor.Clear
    Do Until Payor.EOF
        CboPayor.AddItem Payor!PAYCode & " - " & Payor!PAYCompanyName
        Payor.MoveNext
    Loop
 
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

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    CboPayor.Text = ""
    lstMBM.listitems.Clear
    lblCurrent = 0
    lblTotal = 0
    lblMCO = 0
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
    
    If lstMBM.listitems.Count <> 0 Then
        Call ExportListViewtoExcel(lstMBM)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = Default
End Sub

Private Sub CmdSearch_Click()
Dim MBMSql As String
Dim CovSql As String

    CmdSearch.Enabled = False
    CmdPrint.Enabled = False
    CmdClear.Enabled = False
    CmdUpdate.Enabled = False
    
    lblCurrent = ""
    lblMCO = ""
    
    CurrRecProc = "0"
    TotalMCO = "0"
                        
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    MBMSql = "SELECT MBMNumber, MBMCOVID, MBMBordxDate, MBMBatchNo, MBMInvoiceNo " & _
            " from View_MBM_MBMOthers WHERE (MBMInvoiceNo IS NULL OR MBMInvoiceNo = '') "
    
    CovSql = "SELECT MBMCNumber AS MBMNumber, MBMCCoverID AS MBMCOVID, MBMCBordxDate AS MBMBordxDate, MBMCMCO as MBMMCo," & _
            "  MBMCBatchNo AS MBMBatchNo, MBMInvoiceNo FROM MBMCoveredPersons WHERE (MBMInvoiceNo IS NULL OR MBMInvoiceNo = '') "
    
    If Trim(CboPayor) <> "" Then
        MBMSql = MBMSql & " AND PAYCode='" & Mid(CboPayor, 1, 2) & "'"
        CovSql = CovSql & " AND SUBSTRING(MBMCNumber,1,2)='" & Mid(CboPayor, 1, 2) & "'"
    End If
    If Trim(TxtMBMNo1) <> "" Then
        MBMSql = MBMSql & " AND MBMNumber >='" & Trim(TxtMBMNo1) & "'"
        CovSql = CovSql & " AND MBMCNumber >='" & Trim(TxtMBMNo1) & "'"
    End If
    If Trim(TxtMBMNo2) <> "" Then
        MBMSql = MBMSql & " AND MBMNumber <='" & Trim(TxtMBMNo2) & "'"
        CovSql = CovSql & " AND MBMCNumber <='" & Trim(TxtMBMNo2) & "'"
    End If
    If IsDate(txtBordxDate1) Then
        MBMSql = MBMSql & " AND MBMBordxDate >='" & Format(txtBordxDate1, "mm/dd/yyyy") & "'"
        CovSql = CovSql & " AND MBMCBordxDate >='" & Format(txtBordxDate1, "mm/dd/yyyy") & "'"
    End If
    If IsDate(txtBordxDate2) Then
        MBMSql = MBMSql & " AND MBMBordxDate <='" & Format(txtBordxDate2, "mm/dd/yyyy") & "'"
        CovSql = CovSql & " AND MBMCBordxDate <='" & Format(txtBordxDate2, "mm/dd/yyyy") & "'"
    End If
    If Trim(txtBatchNo1) <> "" Then
        MBMSql = MBMSql & " AND MBMBatchNo >='" & Trim(txtBatchNo1) & "'"
        CovSql = CovSql & " AND MBMCBatchNo >='" & Trim(txtBatchNo1) & "'"
    End If
    If Trim(txtBatchNo2) <> "" Then
        MBMSql = MBMSql & " AND MBMBatchNo <='" & Trim(txtBatchNo2) & "'"
        CovSql = CovSql & " AND MBMCBatchNo <='" & Trim(txtBatchNo2) & "'"
    End If
    Rec0Cnt = 0
    TotRecSel = 0
    lstMBM.listitems.Clear
    Chk0Mco = False
    lblCurrent = 0
    lblTotal = 0
    lblMCO = 0
    Screen.MousePointer = vbHourglass
    Call MBMListing(MBMSql)
    Call MBMListing(CovSql)
    Screen.MousePointer = Default
    If Rec0Cnt = 0 Then
        MsgBox "No record found ! ", vbInformation, DialogBoxTitle
    Else
        If Chk0Mco = True Then
            MsgBox "Zero MCO Fees Found ! ", vbOKOnly, DialogBoxTitle
        End If
    End If
    
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    
    CmdSearch.Enabled = True
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    CmdUpdate.Enabled = True
    
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdSearch.Enabled = True
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    CmdUpdate.Enabled = True
End Sub

Private Sub CmdUpdate_Click()
Dim BB As Integer
Dim Saved As Boolean
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    If lstMBM.listitems.Count > 0 Then
        If Chk0Mco = True Then
            MsgBox "Zero MCO found ! Please refer to IT department ! ", vbOKOnly, DialogBoxTitle
        Else
            If Trim(txtInvoiceNo) = "" Or IsDate(txtInvoiceDate) = False Then
                MsgBox "Please Keyin Invoice Number and Invoice Date ! ", vbOKOnly, DialogBoxTitle
                txtInvoiceNo.SetFocus
            Else
                For BB = 1 To lstMBM.listitems.Count
                    If Trim(lstMBM.listitems(BB).SubItems(3)) <> "" Then
                        Saved = True
                        Exit For
                    End If
                Next BB
                If Saved Then
                    MsgBox "Data Has Been Saved ! ", vbCritical + vbOKOnly, DialogBoxTitle
                Else
                    Call SaveRecord
                End If
            End If
        End If
    End If
    
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
End Sub

Private Sub Form_Load()
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Call ComboListing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
End Sub

Private Sub txtBatchNo1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtBatchNo2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtBordxDate1_LostFocus()
    If IsDate(txtBordxDate1) = False Then
        If txtBordxDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDate1 = "__/__/____"
            txtBordxDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtBordxDate2_LostFocus()
    If IsDate(txtBordxDate2) = False Then
        If txtBordxDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtBordxDate2 = "__/__/____"
            txtBordxDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtInvoiceNo_LostFocus()
    txtInvoiceNo = ChkStr(txtInvoiceNo)
End Sub

Private Sub TxtMBMNo1_LostFocus()
    TxtMBMNo1 = ChkStr(TxtMBMNo1)
End Sub

Private Sub txtMBMNo2_LostFocus()
    TxtMBMNo2 = ChkStr(TxtMBMNo2)
End Sub

Private Sub MBMListing(Optional StrAppend As String)
Dim MBMRec As New ADODB.Recordset
Dim MBMOthRec As New ADODB.Recordset
Dim MBMOthSql As String
Dim lstRec As ListItem
Dim TempMco As String

    'TotalMCO = "0"

    MBMRec.Open StrAppend, medixmasterconn, adOpenKeyset, adLockOptimistic
    If MBMRec.recordCount Then
        'Do
        Rec0Cnt = Rec0Cnt + 1
        'TotRecSel = TotRecSel + MBMRec.recordCount
        'lblTotal = TotRecSel
        Do Until MBMRec.EOF
        'CurrRecProc = CurrRecProc + 1
        'lblCurrent = lblCurrent + CurrRecProc
            With MBMRec
                Set lstRec = lstMBM.listitems.Add(, , Mid(!MBMNumber, 1, 2))
                If Len(!MBMBordxDate) Then
                    lstRec.SubItems(1) = !MBMBordxDate
                End If
                If Len(!MBMBatchNo) Then
                    lstRec.SubItems(2) = !MBMBatchNo
                End If
                lstRec.SubItems(3) = ""
'                LstREc.SubItems(4) = !MBMInvoiceDate
                lstRec.SubItems(5) = !MBMNumber
                lstRec.SubItems(6) = !MBMCOVID
                If !MBMCOVID = "00" Then
                    MBMOthSql = "SELECT MBMMCOFee FROM MBMOthers WHERE MBMNumber='" & Trim(!MBMNumber) & "'"
                    MBMOthRec.Open MBMOthSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                    TempMco = IIf(IsNumeric(MBMOthRec!MBMMCOFee), MBMOthRec!MBMMCOFee, 0)
                    MBMOthRec.Close
                    Set MBMOthRec = Nothing
                Else
                    TempMco = IIf(IsNumeric(!MBMMCO), !MBMMCO, 0)
                End If
                lstRec.SubItems(7) = CDbl(TempMco)
                'TotalMCO = CDbl(TotalMCO) + CDbl(lstRec.SubItems(7))
                'lblMCO = Format(TotalMCO, "#,##0.00")
                
                If !MBMCOVID = "00" Then        ' Check for MCo Amt
                    If IsNumeric(TempMco) = False Then
                        Chk0Mco = True
                    ElseIf CDbl(TempMco) <= 0 Then
                        Chk0Mco = True
                    End If
                End If
                
            End With
            MBMRec.MoveNext
            
            CurrRecProc = CurrRecProc + 1
            lblCurrent = CurrRecProc
            
            TotalMCO = CDbl(TotalMCO) + CDbl(lstRec.SubItems(7))
            lblMCO = Format(TotalMCO, "#,##0.00")
            
            'Loop
            DoEvents
                If intExit = 1 Then
                   StopChk = False
                    Screen.MousePointer = Default

                   Exit Do
                End If
            Loop
        MBMRec.Close
        Set MBMRec = Nothing
    End If
    
End Sub

Private Sub SaveRecord()
Dim AA As Integer
Dim MBMRec As New ADODB.Recordset
Dim MBMSql As String
Dim MBMNo As String
Dim CoverID As String

    Screen.MousePointer = vbHourglass
    For AA = 1 To lstMBM.listitems.Count
        MBMNo = Trim(lstMBM.listitems(AA).SubItems(5))
        CoverID = Trim(lstMBM.listitems(AA).SubItems(6))
    
        If CoverID = "00" Then
            MBMSql = "SELECT MBMInvoiceNo, MBMInvoiceDate, MBMInvoiceBy " & _
                    " from MBMOthers WHERE MBMNumber='" & MBMNo & "'"
        Else
            MBMSql = "SELECT MBMInvoiceNo, MBMInvoiceDate, MBMInvoiceBy " & _
                    " FROM MBMCoveredPersons WHERE MBMCNumber='" & MBMNo & "' AND MBMCCoverID='" & CoverID & "'"
        End If
        
        MBMRec.Open MBMSql, medixmasterconn, adOpenKeyset, adLockOptimistic

        With MBMRec
            lstMBM.listitems(AA).SubItems(3) = Trim(txtInvoiceNo)
            lstMBM.listitems(AA).SubItems(4) = Trim(txtInvoiceDate)
            !MBMInvoiceNo = Trim(txtInvoiceNo)
            !MBMInvoiceDate = Trim(txtInvoiceDate)
            !MBMInvoiceBy = Logon
        End With
        MBMRec.Update
        MBMRec.Close
        Set MBMRec = Nothing
    Next AA
    MsgBox "Records Updated ........ ! ", vbInformation, DialogBoxTitle
    Screen.MousePointer = Default
End Sub
