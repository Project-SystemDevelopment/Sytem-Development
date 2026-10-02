VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPrintBordx 
   Caption         =   "Print Bordx"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   7200
      Width           =   11655
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   2760
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox lblCurrent 
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   975
      End
      Begin Crystal.CrystalReport crptSummary 
         Left            =   3000
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
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
         Height          =   495
         Left            =   7560
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdDetailPreview 
         Caption         =   "De&tail Preview"
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
         Height          =   495
         Left            =   6720
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSummaryPreview 
         Caption         =   "S&ummary Preview"
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
         Height          =   495
         Left            =   4080
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdPrintDetails 
         Caption         =   "Print &Detail"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdPrintSummary 
         Caption         =   "&Print Summary"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10800
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
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
         Height          =   495
         Left            =   10080
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8640
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9360
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin MSComDlg.CommonDialog dlgCommonDialog_PrintBordx 
         Left            =   1560
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
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
         Left            =   1200
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
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
         Left            =   1920
         TabIndex        =   20
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   900
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   11655
      Begin VB.ComboBox CboReportType 
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
         ItemData        =   "FrmPrintBordx.frx":0000
         Left            =   1320
         List            =   "FrmPrintBordx.frx":000A
         TabIndex        =   0
         Top             =   150
         Width           =   4575
      End
      Begin VB.TextBox TxtBordxNo 
         Height          =   285
         Left            =   4800
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   10320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox TxtBordxNoFr 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   440
         Width           =   2055
      End
      Begin VB.TextBox TxtBordxNoTo 
         Height          =   285
         Left            =   3840
         TabIndex        =   3
         Top             =   440
         Width           =   2055
      End
      Begin VB.ComboBox cboPrintOpt 
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
         ItemData        =   "FrmPrintBordx.frx":0035
         Left            =   7200
         List            =   "FrmPrintBordx.frx":003F
         TabIndex        =   5
         Top             =   180
         Width           =   1935
      End
      Begin VB.ComboBox CboBTBG 
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
         ItemData        =   "FrmPrintBordx.frx":0057
         Left            =   7200
         List            =   "FrmPrintBordx.frx":0061
         TabIndex        =   6
         Top             =   465
         Width           =   1935
      End
      Begin VB.ComboBox CboAmt 
         Height          =   315
         ItemData        =   "FrmPrintBordx.frx":0076
         Left            =   7200
         List            =   "FrmPrintBordx.frx":0083
         TabIndex        =   4
         Text            =   "All"
         Top             =   180
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label2 
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
         Index           =   1
         Left            =   9480
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   30
         Top             =   500
         Width           =   255
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   29
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "BTBG Status"
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
         Index           =   0
         Left            =   6240
         TabIndex        =   26
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Print Status"
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
         Left            =   6240
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Bordx No"
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
         TabIndex        =   18
         Top             =   500
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Amount"
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
         Left            =   6240
         TabIndex        =   17
         Top             =   195
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstPrintBordx 
      Height          =   6015
      Left            =   120
      TabIndex        =   24
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10610
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
         Text            =   "Plan Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Claim No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Patient Name"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Member No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Hospital"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Bordx Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Bill Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Approve Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Recovery Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Case Mgmt Fee"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Debit Credit Ref No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Bordx No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Print Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "HLT Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Payor"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Pln"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "BTBG"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Print Bordx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   220
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmPrintBordx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public intExit As Integer
Dim ScanCode As String
Dim Check As Boolean
Dim CancelPrinting As Boolean
Dim bPrintPass As Boolean
Dim TmpmgmtFeesStatus As Boolean
Dim tmpFloatStatus As Boolean
Dim tmpFloatEffDate As String
Dim TmpBordxDate As String
Dim TempMBMName As String
Dim TempGrpComp As String
Dim TempDiagDesc1 As String
Dim TempDiagDesc2 As String
Dim TempDiagDesc3 As String
Dim tmpBordxAmt As Double
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
Private Sub DetailPrintExcel()
Dim sql As String
Dim strsqlCLM As String
Dim strsqlClmInv As String
Dim LstCnt As Integer
Dim PrntCnt As Integer
Dim NPrntCnt As Integer
Dim PrintedDate As String
Dim tmpHltCode As String
Dim tmpAmtType As String
Dim tmpPayCode As String
Dim tmpPayName As String
Dim PayRec As New ADODB.Recordset
Dim strsqlPay As String
Dim tmpFormula As String
Dim PAYName As String
Dim PayAdd1 As String
Dim PayAdd2 As String
Dim PayAdd3 As String
Dim tmpPln As String
Dim tmpProdtName As String
Dim sqlbor As String
Dim sqlinsert As String
Dim BOR As New ADODB.Recordset
Dim TmpBordxDate As String
Dim tmpDOA As String
Dim tmpDOD As String
Dim TmpEffDate As String
Dim tmpExpDate As String
Dim tempPath As String
Dim PlutoBor As New ADODB.Recordset
Dim strsqlPlutoBor As String
Dim PLUTOStatus As Boolean
Dim tmpCLMWSRemarks As String
Dim tmpHospital As String
Dim tmpDept As String
Dim tmpMBMCovName As String
'New Excel for Print Out
Dim objExcel As excel.Application
Dim objworksheet As excel.WORKSHEET
Dim objWorkbook As excel.Workbook

    PLUTOStatus = False
    PrintedDate = Format(Date, "YYYY/MM/DD")
    bPrintPass = False
    CancelPrinting = False
    PrntCnt = 0
    NPrntCnt = 0
    For LstCnt = 1 To lstPrintBordx.ListItems.Count
        tmpHltCode = lstPrintBordx.ListItems(LstCnt).SubItems(13)
        tmpPayCode = lstPrintBordx.ListItems(LstCnt).SubItems(14)
        tmpPln = lstPrintBordx.ListItems(LstCnt).Text
        If lstPrintBordx.ListItems(LstCnt).SubItems(12) = "YES" Then
            PrntCnt = PrntCnt + 1
        Else
            NPrntCnt = NPrntCnt + 1
        End If
    Next LstCnt
    
    If (PrntCnt = 0 And NPrntCnt <> 0) And cboPrintOpt.ListIndex = 1 Then
        MsgBox "Please select 'Print New' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        If (PrntCnt <> 0 And NPrntCnt = 0) And cboPrintOpt.ListIndex = 0 Then
            MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    tmpProdtName = GetProdtName(tmpPln, tmpHltCode, tmpPayCode)
    Call GetPayorInfo(tmpPayCode, PAYName, PayAdd1, PayAdd2, PayAdd3)
    
    tmpPayName = tmpPayCode & " - " & PAYName
    
    
    strsqlPlutoBor = "Select BordxRefNo,BordxDate from TempBordxSummary where BordxRefNo >= '" & Trim(txtBordxNoFr) & "' And BordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
    PlutoBor.Open strsqlPlutoBor, medixmasterconnWSPluto, adOpenForwardOnly, adLockReadOnly
    
    Do While Not PlutoBor.EOF
        PLUTOStatus = True
        TmpBordxDate = PlutoBor!BordxDate
        PlutoBor.MoveNext
    Loop
    PlutoBor.Close
    Set PlutoBor = Nothing
    
        sqlbor = "Select * from VIEW_PrintClaimBordx_06   where BordxRefNo >= '" & Trim(txtBordxNoFr) & "' And BordxRefNo <= '" & Trim(txtBordxNoTo) & "' order by BordxRefNo "
        BOR.Open sqlbor, medixmasterconnWS, adOpenForwardOnly, adLockReadOnly
        
        If CboReportType.ListIndex = 0 Then
            Do While Not BOR.EOF
                Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
                objExcel.Visible = False
                objExcel.WindowState = xlMaximized
                Set objWorkbook = objExcel.Workbooks.Add(crdPath & "\Report\ClaimBordx\PrintBordx_Single.xlt")
                Set objworksheet = objWorkbook.Sheets(1)
                objExcel.DisplayAlerts = False
                Set objworksheet = objWorkbook.Worksheets.Item(1)
                TmpBordxDate = Format(BOR!BordxDate, "yyyy/mm/dd")
                tmpDOA = Format(BOR!CASDateAdmitted, "yyyy/mm/dd")
                tmpDOD = Format(BOR!CASDischargeDate, "yyyy/mm/dd")
                TmpEffDate = Format(BOR!MBMPayorEffDate, "yyyy/mm/dd")
                tmpExpDate = Format(BOR!MBMPayorEXPDate, "yyyy/mm/dd")
                TempMBMName = Replace(Trim(BOR!MBMName), "'", "''")
                TempGrpComp = Replace(Trim(BOR!GrpCompany), "'", "''")
                TempDiagDesc1 = Replace(Trim(BOR!ICDDesc1), "'", "''")
                If BOR!ICDDesc2 <> "" Then
                    TempDiagDesc2 = Replace(Trim(BOR!ICDDesc2), "'", "''")
                Else
                    TempDiagDesc2 = ""
                End If
                If BOR!ICDDesc3 <> "" Then
                    TempDiagDesc3 = Replace(Trim(BOR!ICDDesc3), "'", "''")
                Else
                    TempDiagDesc3 = ""
                End If
                If BOR!CLMWSRemarks <> "" Then
                    tmpCLMWSRemarks = Replace(Trim(BOR!CLMWSRemarks), "'", "''")
                Else
                    tmpCLMWSRemarks = ""
                End If
                If BOR!HOSName <> "" Then
                    tmpHospital = Replace(Trim(BOR!HOSName), "'", "''")
                Else
                    tmpHospital = ""
                End If
                If BOR!MBMDepartment <> "" Then
                    tmpDept = Replace(Trim(BOR!MBMDepartment), "'", "''")
                Else
                    tmpDept = ""
                End If
                If BOR!MBMCName <> "" Then
                    tmpMBMCovName = Replace(BOR!MBMCName, "'", "")
                Else
                    tmpMBMCovName = ""
                End If
                 objworksheet.Cells(11, "B") = tmpPayName
                 objworksheet.Cells(12, "B") = PayAdd1
                 objworksheet.Cells(13, "B") = PayAdd2
                 objworksheet.Cells(14, "B") = PayAdd3
                 objworksheet.Cells(12, "H") = BOR!BordxRefNo
                 objworksheet.Cells(13, "H") = TmpBordxDate
                 objworksheet.Cells(14, "H") = CStr(BOR!CasNumber) + "-" + CStr(BOR!claimversion)
                 If BOR!MBMPatientCovID = "00" Then
                    objworksheet.Cells(18, "C") = TempMBMName
                 Else
                    objworksheet.Cells(18, "C") = tmpMBMCovName
                 End If
                 objworksheet.Cells(19, "C") = BOR!MBMCRELCode   'Relatioship
                 objworksheet.Cells(20, "C") = PAYName
                 objworksheet.Cells(21, "C") = BOR!MBMNumber
                 objworksheet.Cells(22, "C") = BOR!MBMPolicyNo
                 objworksheet.Cells(23, "C") = CStr(Format(TmpEffDate, "DD/MM/YYYY")) + " To " + CStr(Format(tmpExpDate, "dd/MM/YYYY"))
                 objworksheet.Cells(24, "C") = tmpProdtName
                 objworksheet.Cells(25, "C") = BOR!PLNCode
                 objworksheet.Cells(26, "C") = tmpHospital
                 objworksheet.Cells(27, "C") = CStr(Format(tmpDOA, "DD/MM/YYYY"))
                 objworksheet.Cells(28, "C") = CStr(Format(tmpDOD, "DD/MM/YYYY"))
                 objworksheet.Cells(29, "C") = TempDiagDesc1
                 objworksheet.Cells(30, "C") = TempDiagDesc2
                 objworksheet.Cells(31, "C") = TempDiagDesc3
                 objworksheet.Cells(32, "C") = tmpCLMWSRemarks
                 objworksheet.Cells(26, "H") = IIf(IsNumeric(BOR!CLMTotalActual), BOR!CLMTotalActual, 0)
                 objworksheet.Cells(27, "H") = BOR!CLMTotalApproved
                 objworksheet.Cells(28, "H") = IIf(IsNumeric(BOR!CLMBordxAmt), BOR!CLMBordxAmt, 0)
                 If Len(Dir("C:\PrintBordx\", vbDirectory)) = 0 Then
                    MkDir "C:\PrintBordx\"
                 End If
                 Dim DirPath As String
                 DirPath = "C:\PrintBordx\" + BOR!BordxRefNo
                 If Len(Dir(DirPath, vbDirectory)) = 0 Then
                    MkDir DirPath
                 End If
                 DirPath = DirPath + "\" + BOR!BordxRefNo + ".xls"
                 objworksheet.SaveAs DirPath
                 objworksheet.PrintOut
                 Set objworksheet = Nothing
                 objExcel.Quit
                 bPrintPass = True
            BOR.MoveNext
            Loop
        Else
                Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
                objExcel.Visible = False
                objExcel.WindowState = xlMaximized
                Set objWorkbook = objExcel.Workbooks.Add(crdPath & "\Report\ClaimBordx\PrintBordx_Multiple.xlt")
                Set objworksheet = objWorkbook.Sheets(1)
                objExcel.DisplayAlerts = False
                Set objworksheet = objWorkbook.Worksheets.Item(1)
                Dim isFrst As Boolean
                Dim rcnt As Long
                Dim RowCnt As Long
                Dim TmpRefNum As String
                
                Dim HosAmt As Double
                Dim MBMAmt As Double
                Dim AppAmt As Double
                
                TmpRefNum = ""
                HosAmt = 0
                MBMAmt = 0
                AppAmt = 0
                RowCnt = 10
                rcnt = 1
                isFrst = True
            Do While Not BOR.EOF
                If isFrst Then
                    objworksheet.Cells(5, "E") = Format(DateTime.Now, "DD/MM/YYYY")
                    objworksheet.Cells(6, "E") = tmpPayName
                    objworksheet.Cells(7, "E") = BOR!BordxRefNo
                    TmpRefNum = BOR!BordxRefNo
                isFrst = False
                End If
                
                objworksheet.Cells(RowCnt, "A") = CStr(rcnt) & CStr(BOR!HLTCode)
                objworksheet.Cells(RowCnt, "B") = IIf(Len(BOR!PLNCode), BOR!PLNCode, "")
                objworksheet.Cells(RowCnt, "C") = IIf(Len(BOR!CasNumber), BOR!CasNumber, "")
                objworksheet.Cells(RowCnt, "D") = IIf(Len(BOR!claimversion), BOR!claimversion, "")
                objworksheet.Cells(RowCnt, "E") = IIf(Len(BOR!GrpCompany), BOR!GrpCompany, "")
                objworksheet.Cells(RowCnt, "F") = IIf(Len(BOR!MBMName), BOR!MBMName, "")
                objworksheet.Cells(RowCnt, "G") = IIf(Len(BOR!MBMCRELCode), BOR!MBMCRELCode, "")
                objworksheet.Cells(RowCnt, "H") = IIf(Len(BOR!MBMPolicyNo), BOR!MBMPolicyNo, "")
                objworksheet.Cells(RowCnt, "I") = IIf(Len(BOR!MBMNumber), BOR!MBMNumber, "")
                objworksheet.Cells(RowCnt, "J") = IIf(Len(BOR!CASDateAdmitted), Format(BOR!CASDateAdmitted, "DD/MM/YYYY"), "")
                objworksheet.Cells(RowCnt, "K") = IIf(Len(BOR!CASDischargeDate), Format(BOR!CASDischargeDate, "DD/MM/YYYY"), "")
                objworksheet.Cells(RowCnt, "L") = IIf(Len(BOR!ICDDesc1), BOR!ICDDesc1, "")
                objworksheet.Cells(RowCnt, "M") = IIf(Len(BOR!HOSName), BOR!HOSName, "")
                objworksheet.Cells(RowCnt, "N") = IIf(IsNumeric(BOR!CLMPayableByMedix2H), BOR!CLMPayableByMedix2H, 0)
                HosAmt = HosAmt + CDbl(IIf(IsNumeric(BOR!CLMPayableByMedix2H), BOR!CLMPayableByMedix2H, 0))
                objworksheet.Cells(RowCnt, "O") = IIf(IsNumeric(BOR!CLMTotalApproved), BOR!CLMTotalApproved, 0)
                AppAmt = AppAmt + CDbl(IIf(IsNumeric(BOR!CLMTotalApproved), BOR!CLMTotalApproved, 0))
                objworksheet.Cells(RowCnt, "P") = IIf(IsNumeric(BOR!CLMPayableByMedix2M), BOR!CLMPayableByMedix2M, 0)
                MBMAmt = MBMAmt + CDbl(IIf(IsNumeric(BOR!CLMPayableByMedix2M), BOR!CLMPayableByMedix2M, 0))
                objworksheet.Cells(RowCnt, "Q") = ""
                objworksheet.Cells(RowCnt, "R") = IIf(Len(BOR!CLMWSRemarks), BOR!CLMWSRemarks, "")
                objworksheet.Cells(RowCnt, "S") = ""
                
                RowCnt = RowCnt + 1
                rcnt = rcnt + 1
                BOR.MoveNext
            Loop
            
            Dim sPos As String
            Dim lPos As String
            
                RowCnt = RowCnt + 2
                objworksheet.Cells(RowCnt, "M") = "Sub Total"
                objworksheet.Cells(RowCnt, "N") = HosAmt
                objworksheet.Cells(RowCnt, "O") = AppAmt
                objworksheet.Cells(RowCnt, "P") = MBMAmt
                
                RowCnt = RowCnt + 2
                objworksheet.Cells(RowCnt, "M") = "TOTAL"
                objworksheet.Cells(RowCnt, "N") = HosAmt
                objworksheet.Cells(RowCnt, "O") = AppAmt
                objworksheet.Cells(RowCnt, "P") = MBMAmt
                
                RowCnt = RowCnt + 5
                sPos = "A" + CStr(RowCnt)
                lPos = "C" + CStr(RowCnt)
                objworksheet.Range(sPos + ":" + lPos).MergeCells = True
                objworksheet.Cells(RowCnt, "A") = "Approved by:"
                
                sPos = "N" + CStr(RowCnt)
                lPos = "O" + CStr(RowCnt)
                objworksheet.Range(sPos + ":" + lPos).MergeCells = True
                objworksheet.Cells(RowCnt, "N") = "Submitted by:"
                
                RowCnt = RowCnt + 5
                sPos = "A" + CStr(RowCnt)
                lPos = "C" + CStr(RowCnt)
                objworksheet.Range(sPos + ":" + lPos).MergeCells = True
                objworksheet.Cells(RowCnt, "A") = "Company Stamp:"
                
                sPos = "N" + CStr(RowCnt)
                lPos = "O" + CStr(RowCnt)
                objworksheet.Range(sPos + ":" + lPos).MergeCells = True
                objworksheet.Cells(RowCnt, "N") = "Company Stamp:"
                
                RowCnt = RowCnt + 1
                sPos = "A" + CStr(RowCnt)
                lPos = "C" + CStr(RowCnt)
                objworksheet.Range(sPos + ":" + lPos).MergeCells = True
                objworksheet.Cells(RowCnt, "A") = "Date:"
                
                If Len(Dir("C:\PrintBordx\", vbDirectory)) = 0 Then
                    MkDir "C:\PrintBordx\"
                 End If
                 'Dim DirPath As String
                 DirPath = ""
                 DirPath = "C:\PrintBordx\" + TmpRefNum
                 If Len(Dir(DirPath, vbDirectory)) = 0 Then
                    MkDir DirPath
                 End If
                 DirPath = DirPath + "\" + TmpRefNum + ".xls"
                 objworksheet.SaveAs DirPath
                 objworksheet.PrintOut
                 Set objworksheet = Nothing
                 objExcel.Quit
                 bPrintPass = True
        End If
        
    'End If
    
    
    
    'tmpFormula = "{VIEW_PrintClaimBordx_06.HLTCode} ='" & tmpHltCode & "'" & _
    '             " AND {VIEW_PrintClaimBordx_06.BordxRefNo} >= '" & Trim(TxtBordxNoFr) & "'" & _
    '            " AND {VIEW_PrintClaimBordx_06.BordxRefNo} <= '" & Trim(TxtBordxNoTo) & "'"
    'Select Case CboAmt.ListIndex
    '    Case 0
    '        tmpAmtType = "( BELOW RM2000 )"
    '        tmpFormula = tmpFormula & " AND {VIEW_PrintClaimBordx_06.CLMTotalApproved} <= 2000"
    '    Case 1
    '        tmpAmtType = "( ABOVE RM2000 )"
    '        tmpFormula = tmpFormula & " AND {VIEW_PrintClaimBordx_06.CLMTotalApproved} > 2000"
    '    Case 2
            tmpAmtType = ""
    'End Select
    'If CboReportType.ListIndex = 0 Then   ' single claim
    '    If tmpFloatStatus = True Then
    '        If CDate(TmpBordxDate) >= CDate(tmpFloatEffDate) Then
    '            crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_ClaimFloat.rpt"
    '        Else
    '            'If Trim(tmpHLTCode) = "N" Then
    '                crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_Claim.rpt"
    '                'CrystalActiveXReportViewer1.ReportSource = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_Claim.rpt"
    '            'Else
    '            '    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_Claim_Sihat.rpt"
    '            'End If
    '        End If
    '    Else
    '        If Mid(CboBTBG, 1, 1) = "N" Then
    '            If TmpmgmtFeesStatus = False Then
    '                If Trim(tmpHltCode) = "N" Then
    '                    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_Claim.rpt"
    '                Else
    '                    If CDbl(tmpBordxAmt) >= 2000 Then
    '                        crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_Claim_Sihat.rpt"
    '                    Else
    '                        crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_Claim_Sihat_Below.rpt"
    '                    End If
    '                End If
    '            Else
    '                crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_Claim_XP.rpt"
    '            End If
    '        Else
    '            crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_Claim_BTBG.rpt"
    '        End If
    '    End If
    '    crptSummary.ParameterFields(1) = "HLTCode;" & tmpHltCode & ";true"
    '    crptSummary.ParameterFields(2) = "PayAdd1;" & PayAdd1 & ";true"
    '    crptSummary.ParameterFields(3) = "PayAdd2;" & PayAdd2 & ";true"
    '    crptSummary.ParameterFields(4) = "PayAdd3;" & PayAdd3 & ";true"
    '    crptSummary.ParameterFields(5) = "BTBG;" & Mid(CboBTBG, 1, 1) & ";true"
    '    crptSummary.ParameterFields(6) = "ProdtType;" & tmpProdtName & ";true"
    'Else  'multiple claims
    '    If Trim(tmpHltCode) = "S" Then
    '        If Mid(CboBTBG, 1, 1) = "N" Then
    '            crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Detail_Sihat.rpt"
    '        ElseIf Mid(CboBTBG, 1, 1) = "Y" Then
    '            crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Detail_Sihat_BTB.rpt"
    '        End If
    '    Else 'non sihat
    '        If Mid(CboBTBG, 1, 1) = "N" Then
    '            crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Detail_NS.rpt"
    '        ElseIf Mid(CboBTBG, 1, 1) = "Y" Then
    '            crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Detail_NS_BTB.rpt"
    '        End If
    '    End If
    '    crptSummary.ParameterFields(1) = "AmtType;" & tmpAmtType & ";true"
    'End If
    'crptSummary.SelectionFormula = tmpFormula
    'crptSummary.SortFields(0) = "+{VIEW_PrintClaimBordx_06.BordxRefNo}"
    ''crptSummary.SortFields(1) = "+{VIEW_PrintClaimBordx_06.PLNCode}"
    'crptSummary.ParameterFields(0) = "txtPayor;" & tmpPayName & ";true"
   '
   ' If cmdDetailPreview Then
   '     crptSummary.Destination = crptToWindow
   '     crptSummary.WindowState = crptMaximized
   '     crptSummary.Action = 1
   ' End If
   ' If cmdPrintDetails Then
   '     crptSummary.Destination = crptToPrinter
   ' End If
   ' If cmdPrintDetails Then
   '     If bPrintPass = False And CancelPrinting = False Then
   '         bPrintPass = PrintFunction
   '         If bPrintPass = False Then
   '             Screen.MousePointer = Default
   '             Exit Sub
   '         End If
   '     Else
   '         crptSummary.PrintReport
   '     End If
  '
        If bPrintPass = True Then
            sql = " update Bordx set BordxPrintedDate = '" & PrintedDate & "', " & _
                  " BordxPrintedStatus = '1', BordxPrintedBy = '" & Logon & "' " & _
                  " Where BordxRefNo >= '" & Trim(txtBordxNoFr) & "' And BordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
            medixmasterconnWS.Execute sql
                    
            strsqlCLM = " update Claim set CLMPending = 'B' Where BordxRefNo >= '" & Trim(txtBordxNoFr) & "' And BordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
            medixmasterconnWS.Execute strsqlCLM
            
            strsqlClmInv = " update ClaimInvoice set CLMInvPending = 'B' " & _
                        " Where INVBordxRefNo >= '" & Trim(txtBordxNoFr) & "' And INVBordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
            medixmasterconnWS.Execute strsqlClmInv
            CboReportType.ListIndex = 0
        End If
    'End If
    Screen.MousePointer = Default

End Sub

Private Sub DetailPrint()
Dim sql As String
Dim strsqlCLM As String
Dim strsqlClmInv As String
Dim LstCnt As Integer
Dim PrntCnt As Integer
Dim NPrntCnt As Integer
Dim PrintedDate As String
Dim tmpHltCode As String
Dim tmpAmtType As String
Dim tmpPayCode As String
Dim tmpPayName As String
Dim PayRec As New ADODB.Recordset
Dim strsqlPay As String
Dim tmpFormula As String
Dim PAYName As String
Dim PayAdd1 As String
Dim PayAdd2 As String
Dim PayAdd3 As String
Dim tmpPln As String
Dim tmpProdtName As String
Dim sqlbor As String
Dim sqlinsert As String
Dim BOR As New ADODB.Recordset
Dim TmpBordxDate As String
Dim tmpDOA As String
Dim tmpDOD As String
Dim TmpEffDate As String
Dim tmpExpDate As String
Dim tempPath As String
Dim PlutoBor As New ADODB.Recordset
Dim strsqlPlutoBor As String
Dim PLUTOStatus As Boolean
Dim tmpCLMWSRemarks As String
Dim tmpHospital As String
Dim tmpDept As String
Dim tmpMBMCovName As String

    PLUTOStatus = False
    PrintedDate = Format(Date, "YYYY/MM/DD")
    bPrintPass = False
    CancelPrinting = False
    
    PrntCnt = 0
    NPrntCnt = 0
    
    For LstCnt = 1 To lstPrintBordx.ListItems.Count
        tmpHltCode = lstPrintBordx.ListItems(LstCnt).SubItems(13)
        tmpPayCode = lstPrintBordx.ListItems(LstCnt).SubItems(14)
        tmpPln = lstPrintBordx.ListItems(LstCnt).Text
        If lstPrintBordx.ListItems(LstCnt).SubItems(12) = "YES" Then
            PrntCnt = PrntCnt + 1
        Else
            NPrntCnt = NPrntCnt + 1
        End If
    Next LstCnt
    
    If (PrntCnt = 0 And NPrntCnt <> 0) And cboPrintOpt.ListIndex = 1 Then
        MsgBox "Please select 'Print New' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        If (PrntCnt <> 0 And NPrntCnt = 0) And cboPrintOpt.ListIndex = 0 Then
            MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    tmpProdtName = GetProdtName(tmpPln, tmpHltCode, tmpPayCode)
    Call GetPayorInfo(tmpPayCode, PAYName, PayAdd1, PayAdd2, PayAdd3)
    
    tmpPayName = tmpPayCode & " - " & PAYName
    
    
    strsqlPlutoBor = "Select BordxRefNo,BordxDate from TempBordxSummary where BordxRefNo >= '" & Trim(txtBordxNoFr) & "' And BordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
    PlutoBor.Open strsqlPlutoBor, medixmasterconnWSPluto, adOpenForwardOnly, adLockReadOnly
    
    Do While Not PlutoBor.EOF
        PLUTOStatus = True
        TmpBordxDate = PlutoBor!BordxDate
        PlutoBor.MoveNext
    Loop
    PlutoBor.Close
    Set PlutoBor = Nothing
    
    If PLUTOStatus = False Then
        sqlbor = "Select * from VIEW_PrintClaimBordx_06   where BordxRefNo >= '" & Trim(txtBordxNoFr) & "' And BordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
        BOR.Open sqlbor, medixmasterconnWS, adOpenForwardOnly, adLockReadOnly
    
        Do While Not BOR.EOF
            
            TmpBordxDate = Format(BOR!BordxDate, "yyyy/mm/dd")
            tmpDOA = Format(BOR!CASDateAdmitted, "yyyy/mm/dd")
            tmpDOD = Format(BOR!CASDischargeDate, "yyyy/mm/dd")
            TmpEffDate = Format(BOR!MBMPayorEffDate, "yyyy/mm/dd")
            tmpExpDate = Format(BOR!MBMPayorEXPDate, "yyyy/mm/dd")
            TempMBMName = Replace(Trim(BOR!MBMName), "'", "''")
            TempGrpComp = Replace(Trim(BOR!GrpCompany), "'", "''")
            TempDiagDesc1 = Replace(Trim(BOR!ICDDesc1), "'", "''")
            If BOR!ICDDesc2 <> "" Then
                TempDiagDesc2 = Replace(Trim(BOR!ICDDesc2), "'", "''")
            Else
                TempDiagDesc2 = ""
            End If
            If BOR!ICDDesc3 <> "" Then
                TempDiagDesc3 = Replace(Trim(BOR!ICDDesc3), "'", "''")
            Else
                TempDiagDesc3 = ""
            End If
            If BOR!CLMWSRemarks <> "" Then
                tmpCLMWSRemarks = Replace(Trim(BOR!CLMWSRemarks), "'", "''")
            Else
                tmpCLMWSRemarks = ""
            End If
            If BOR!HOSName <> "" Then
                tmpHospital = Replace(Trim(BOR!HOSName), "'", "''")
            Else
                tmpHospital = ""
            End If
            If BOR!MBMDepartment <> "" Then
                tmpDept = Replace(Trim(BOR!MBMDepartment), "'", "''")
            Else
                tmpDept = ""
            End If
            
            If BOR!MBMCName <> "" Then
                tmpMBMCovName = Replace(BOR!MBMCName, "'", "")
            Else
                tmpMBMCovName = ""
            End If
        
            
            sqlinsert = ""
            sqlinsert = "Insert into TempBordxSummary(CASNumber, ClaimVersion, MBMName, ClaimBordxStatus," & _
                        "BordxDate, CLMTotalApproved, CLMCaseMgmtFees, TotalMgmtFee,CLMPayableByMedix2H, CLMTotalActual," & _
                        " PAYCode,BordxRefNo, BordxPrintedStatus, CLMPayableByMedix2M, DebitCreditRefNo, " & _
                        "INVPayTo, HOS.HOSName,INSCode, PLNCode,HLTCode, MBMNumber,MBMBTBG,MBMPatientCovID, " & _
                        "GRPCompany,MBMDepartment,MBMPolicyNo,CASDischargeDate,CASDateAdmitted,CLMWSRemarks, " & _
                        "MBMCRELCode,MBMCName,MBMMemberType,MBMPayorEffDate,MBMPayorExpDate,CLMBordxAmt, " & _
                        "ICDDesc1,ICDDesc2,ICDDesc3,ICDCode1, ICDCode2,ICDCode3,CASHosName) " & _
                        "Values('" & BOR!CasNumber & "','" & BOR!claimversion & "', '" & TempMBMName & "',  '" & BOR!ClaimBordxStatus & "'," & _
                        "'" & TmpBordxDate & "', " & BOR!CLMTotalApproved & ", " & BOR!CLMCaseMgmtFees & ", " & BOR!TotalMgmtFee & ", " & IIf(IsNumeric(BOR!CLMPayableByMedix2H), BOR!CLMPayableByMedix2H, 0) & ", " & IIf(IsNumeric(BOR!CLMTotalActual), BOR!CLMTotalActual, 0) & "," & _
                        "'" & BOR!PAYCode & "', '" & BOR!BordxRefNo & "', '" & IIf(Trim(BOR!BordxPrintedStatus) = "TRUE", "1", "0") & "', " & BOR!CLMPayableByMedix2M & ", '" & BOR!DebitCreditRefNo & "', " & _
                        "'" & BOR!InvPayTo & "', '" & tmpHospital & "', '" & BOR!INSCode & "','" & BOR!PLNCode & "', '" & BOR!HLTCode & "', '" & BOR!MBMNumber & "','" & BOR!MBMBTBG & "', '" & BOR!MBMPatientCovID & "', " & _
                         "'" & TempGrpComp & "' , '" & tmpDept & "', '" & BOR!MBMPolicyNo & "','" & tmpDOD & "','" & tmpDOA & "','" & tmpCLMWSRemarks & "', " & _
                        "'" & BOR!MBMCRELCode & "','" & tmpMBMCovName & "','" & BOR!MBMMemberType & "','" & TmpEffDate & "','" & tmpExpDate & "'," & BOR!CLMBordxAmt & ", " & _
                        "'" & TempDiagDesc1 & "','" & TempDiagDesc2 & "' ,'" & TempDiagDesc3 & "', '" & BOR!ICDCode1 & "','" & BOR!ICDCode2 & "', '" & BOR!ICDCode3 & "', '" & BOR!CASHOSName & "')"
                        
            medixmasterconnWSPluto.Execute sqlinsert
            
                       'BordxMNRBSql = "Insert Into SUNPosting(TransRef, BordxMNRBStatus," & _
                        '            "PostStatus, PostType, PostDate, PostBy, " & _
                        '            "AccMonth, AccYear, BatchNo) " & _
                        '            "Values('" & TempReceiptNo & "', '1','1','" & PostingType & _
                        '            "','" & Format(Date, "yyyy/mm/dd") & _
                        '            "','" & Logon & "','" & Mid(cboAccPeriod, 1, 2) & _
                        '            "','" & Mid(cboAccPeriod, 4, 4) & "','" & BatchNo & "')"
                        'medixmasterconnWS.Execute BordxMNRBSql
    
        
        BOR.MoveNext
        Loop
    End If
    tmpFormula = "{VIEW_PrintClaimBordx_06.HLTCode} ='" & tmpHltCode & "'" & _
                 " AND {VIEW_PrintClaimBordx_06.BordxRefNo} >= '" & Trim(txtBordxNoFr) & "'" & _
                " AND {VIEW_PrintClaimBordx_06.BordxRefNo} <= '" & Trim(txtBordxNoTo) & "'"
    'Select Case CboAmt.ListIndex
    '    Case 0
    '        tmpAmtType = "( BELOW RM2000 )"
    '        tmpFormula = tmpFormula & " AND {VIEW_PrintClaimBordx_06.CLMTotalApproved} <= 2000"
    '    Case 1
    '        tmpAmtType = "( ABOVE RM2000 )"
    '        tmpFormula = tmpFormula & " AND {VIEW_PrintClaimBordx_06.CLMTotalApproved} > 2000"
    '    Case 2
            tmpAmtType = ""
    'End Select
    If CboReportType.ListIndex = 0 Then   ' single claim
        If tmpFloatStatus = True Then
            If CDate(TmpBordxDate) >= CDate(tmpFloatEffDate) Then
                crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_ClaimFloat.rpt"
            Else
                'If Trim(tmpHLTCode) = "N" Then
                    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_Claim.rpt"
                    'CrystalActiveXReportViewer1.ReportSource = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_Claim.rpt"
                'Else
                '    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_Claim_Sihat.rpt"
                'End If
            End If
        Else
            If Mid(cboBTBG, 1, 1) = "N" Then
                If TmpmgmtFeesStatus = False Then
                    If Trim(tmpHltCode) = "N" Then
                        crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_Claim.rpt"
                    Else
                        If CDbl(tmpBordxAmt) >= 2000 Then
                            crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_Claim_Sihat.rpt"
                        Else
                            crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_Claim_Sihat_Below.rpt"
                        End If
                    End If
                Else
                    crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_Claim_XP.rpt"
                End If
            Else
                crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Single_Claim_BTBG.rpt"
            End If
        End If
        crptSummary.ParameterFields(1) = "HLTCode;" & tmpHltCode & ";true"
        crptSummary.ParameterFields(2) = "PayAdd1;" & PayAdd1 & ";true"
        crptSummary.ParameterFields(3) = "PayAdd2;" & PayAdd2 & ";true"
        crptSummary.ParameterFields(4) = "PayAdd3;" & PayAdd3 & ";true"
        crptSummary.ParameterFields(5) = "BTBG;" & Mid(cboBTBG, 1, 1) & ";true"
        crptSummary.ParameterFields(6) = "ProdtType;" & tmpProdtName & ";true"
    Else  'multiple claims
        If Trim(tmpHltCode) = "S" Then
            If Mid(cboBTBG, 1, 1) = "N" Then
                crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Detail_Sihat.rpt"
            ElseIf Mid(cboBTBG, 1, 1) = "Y" Then
                crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Detail_Sihat_BTB.rpt"
            End If
        Else 'non sihat
            If Mid(cboBTBG, 1, 1) = "N" Then
                crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Detail_NS.rpt"
            ElseIf Mid(cboBTBG, 1, 1) = "Y" Then
                crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Detail_NS_BTB.rpt"
            End If
        End If
        crptSummary.ParameterFields(1) = "AmtType;" & tmpAmtType & ";true"
    End If
    crptSummary.SelectionFormula = tmpFormula
    crptSummary.SortFields(0) = "+{VIEW_PrintClaimBordx_06.BordxRefNo}"
    'crptSummary.SortFields(1) = "+{VIEW_PrintClaimBordx_06.PLNCode}"
    crptSummary.ParameterFields(0) = "txtPayor;" & tmpPayName & ";true"
    
    If cmdDetailPreview Then
        crptSummary.Destination = crptToWindow
        crptSummary.WindowState = crptMaximized
        crptSummary.Action = 1
    End If
    If cmdPrintDetails Then
        crptSummary.Destination = crptToPrinter
    End If
    If cmdPrintDetails Then
        If bPrintPass = False And CancelPrinting = False Then
            bPrintPass = PrintFunction
            If bPrintPass = False Then
                Screen.MousePointer = Default
                Exit Sub
            End If
        Else
            crptSummary.PrintReport
        End If
   
        If bPrintPass = True And CancelPrinting = False Then
            sql = " update Bordx set BordxPrintedDate = '" & PrintedDate & "', " & _
                  " BordxPrintedStatus = '1', BordxPrintedBy = '" & Logon & "' " & _
                  " Where BordxRefNo >= '" & Trim(txtBordxNoFr) & "' And BordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
            medixmasterconnWS.Execute sql
                    
            strsqlCLM = " update Claim set CLMPending = 'B' Where BordxRefNo >= '" & Trim(txtBordxNoFr) & "' And BordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
            medixmasterconnWS.Execute strsqlCLM
            
            strsqlClmInv = " update ClaimInvoice set CLMInvPending = 'B' " & _
                        " Where INVBordxRefNo >= '" & Trim(txtBordxNoFr) & "' And INVBordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
            medixmasterconnWS.Execute strsqlClmInv
            CboReportType.ListIndex = 0
        End If
    End If
    Screen.MousePointer = Default

End Sub

Private Function PrintFunction() As Boolean
On Error GoTo CancelPrint
        If crptSummary Is Nothing Then Exit Function
        
        With dlgCommonDialog_PrintBordx
            .DialogTitle = "Print Crystal Report "
            .CancelError = True
            .Flags = cdlPDNoSelection + cdlPDNoPageNums
            .ShowPrinter
            
            
            If Err <> MSComDlg.cdlCancel Then
                crptSummary.Action = 1
                'crptSummary.PrintReport
            End If
        End With
        PrintFunction = True
        Exit Function

CancelPrint:
        PrintFunction = False
        CancelPrinting = True
    
End Function

Private Sub SummaryPrint()
Dim sql As String
Dim strsqlCLM As String
Dim strsqlClmInv As String
Dim LstCnt As Integer
Dim PrntCnt As Integer
Dim NPrntCnt As Integer
Dim PrintedDate As String
Dim tmpHltCode As String
Dim tmpAmtType As String
Dim tmpPayCode As String
Dim tmpPayName As String
Dim PayRec As New ADODB.Recordset
Dim strsqlPay As String
Dim tmpFormula As String
Dim PAYName As String
Dim PayAdd1 As String
Dim PayAdd2 As String
Dim PayAdd3 As String
Dim sqlbor As String
Dim sqlinsert As String
Dim BOR As New ADODB.Recordset
Dim TmpBordxDate As String
Dim tmpDOA As String
Dim tmpDOD As String
Dim TmpEffDate As String
Dim tmpExpDate As String
Dim strsqlPlutoBor As String
Dim PlutoBor As New ADODB.Recordset
Dim PLUTOStatus As Boolean
    
    PLUTOStatus = False
    PrintedDate = Format(Date, "YYYY/MM/DD")
    bPrintPass = False
    CancelPrinting = False
    
    PrntCnt = 0
    NPrntCnt = 0
    
    For LstCnt = 1 To lstPrintBordx.ListItems.Count
        tmpHltCode = lstPrintBordx.ListItems(LstCnt).SubItems(13)
        tmpPayCode = lstPrintBordx.ListItems(LstCnt).SubItems(14)
        If lstPrintBordx.ListItems(LstCnt).SubItems(12) = "YES" Then
            PrntCnt = PrntCnt + 1
        Else
            NPrntCnt = NPrntCnt + 1
        End If
    Next LstCnt
    
    If (PrntCnt = 0 And NPrntCnt <> 0) And cboPrintOpt.ListIndex = 1 Then
        MsgBox "Please select 'Print New' option for this report", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        If (PrntCnt <> 0 And NPrntCnt = 0) And cboPrintOpt.ListIndex = 0 Then
            MsgBox "This report already printed, please select 'Reprint' ", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    Call GetPayorInfo(tmpPayCode, PAYName, PayAdd1, PayAdd2, PayAdd3)
    tmpPayName = tmpPayCode & " - " & PAYName
    
    strsqlPlutoBor = "Select BordxRefNo,BordxDate from TempBordxSummary where BordxRefNo >= '" & Trim(txtBordxNoFr) & "' And BordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
    PlutoBor.Open strsqlPlutoBor, medixmasterconnWSPluto, adOpenForwardOnly, adLockReadOnly
    
    Do While Not PlutoBor.EOF
        PLUTOStatus = True
        TmpBordxDate = PlutoBor!BordxDate
        PlutoBor.MoveNext
    Loop
    PlutoBor.Close
    Set PlutoBor = Nothing

    If PLUTOStatus = False Then
        sqlbor = "Select * from VIEW_PrintClaimBordx_06   where BordxRefNo >= '" & Trim(txtBordxNoFr) & "' And BordxRefNo <= '" & Trim(txtBordxNoTo) & "'"
        BOR.Open sqlbor, medixmasterconnWS, adOpenForwardOnly, adLockReadOnly
    
        Do While Not BOR.EOF
            
            TmpBordxDate = Format(BOR!BordxDate, "yyyy/mm/dd")
            tmpDOA = Format(BOR!CASDateAdmitted, "yyyy/mm/dd")
            tmpDOD = Format(BOR!CASDischargeDate, "yyyy/mm/dd")
            TmpEffDate = Format(BOR!MBMPayorEffDate, "yyyy/mm/dd")
            tmpExpDate = Format(BOR!MBMPayorEXPDate, "yyyy/mm/dd")
            TempMBMName = Replace(Trim(BOR!MBMName), "'", "''")
            TempGrpComp = Replace(Trim(BOR!GrpCompany), "'", "''")
            TempDiagDesc1 = Replace(Trim(BOR!ICDDesc1), "'", "''")
            TempDiagDesc2 = Replace(Trim(BOR!ICDDesc2), "'", "''")
            TempDiagDesc3 = Replace(Trim(BOR!ICDDesc3), "'", "''")
            
            sqlinsert = ""
            sqlinsert = "Insert into TempBordxSummary(CASNumber, ClaimVersion, MBMName, ClaimBordxStatus," & _
                        "BordxDate, CLMTotalApproved, CLMCaseMgmtFees, TotalMgmtFee,CLMPayableByMedix2H, CLMTotalActual," & _
                        " PAYCode,BordxRefNo, BordxPrintedStatus, CLMPayableByMedix2M, DebitCreditRefNo, " & _
                        "INVPayTo, HOS.HOSName,INSCode, PLNCode,HLTCode, MBMNumber,MBMBTBG,MBMPatientCovID, " & _
                        "GRPCompany,MBMDepartment,MBMPolicyNo,CASDischargeDate,CASDateAdmitted,CLMWSRemarks, " & _
                        "MBMCRELCode,MBMCName,MBMMemberType,MBMPayorEffDate,MBMPayorExpDate,CLMBordxAmt, " & _
                        "ICDDesc1,ICDDesc2,ICDDesc3,ICDCode1, ICDCode2,ICDCode3,CASHosName) " & _
                        "Values('" & BOR!CasNumber & "','" & BOR!claimversion & "', '" & TempMBMName & "',  '" & BOR!ClaimBordxStatus & "'," & _
                        "'" & TmpBordxDate & "', " & BOR!CLMTotalApproved & ", " & BOR!CLMCaseMgmtFees & ", " & BOR!TotalMgmtFee & ", " & BOR!CLMPayableByMedix2H & ", " & BOR!CLMTotalActual & "," & _
                        "'" & BOR!PAYCode & "', '" & BOR!BordxRefNo & "', '" & IIf(Trim(BOR!BordxPrintedStatus) = "TRUE", "1", "0") & "', " & BOR!CLMPayableByMedix2M & ", '" & BOR!DebitCreditRefNo & "', " & _
                        "'" & BOR!InvPayTo & "', '" & BOR!HOSName & "', '" & BOR!INSCode & "','" & BOR!PLNCode & "', '" & BOR!HLTCode & "', '" & BOR!MBMNumber & "','" & BOR!MBMBTBG & "', '" & BOR!MBMPatientCovID & "', " & _
                        "'" & TempGrpComp & "' , '" & BOR!MBMDepartment & "', '" & BOR!MBMPolicyNo & "','" & tmpDOD & "','" & tmpDOA & "','" & BOR!CLMWSRemarks & "', " & _
                        "'" & BOR!MBMCRELCode & "','" & BOR!MBMCName & "','" & BOR!MBMMemberType & "','" & TmpEffDate & "','" & tmpExpDate & "'," & BOR!CLMBordxAmt & ", " & _
                        "'" & TempDiagDesc1 & "','" & TempDiagDesc2 & "' ,'" & TempDiagDesc3 & "', '" & BOR!ICDCode1 & "','" & BOR!ICDCode2 & "', '" & BOR!ICDCode3 & "', '" & BOR!CASHOSName & "')"
            medixmasterconnWSPluto.Execute sqlinsert
            
                       'BordxMNRBSql = "Insert Into SUNPosting(TransRef, BordxMNRBStatus," & _
                        '            "PostStatus, PostType, PostDate, PostBy, " & _
                        '            "AccMonth, AccYear, BatchNo) " & _
                        '            "Values('" & TempReceiptNo & "', '1','1','" & PostingType & _
                        '            "','" & Format(Date, "yyyy/mm/dd") & _
                        '            "','" & Logon & "','" & Mid(cboAccPeriod, 1, 2) & _
                        '            "','" & Mid(cboAccPeriod, 4, 4) & "','" & BatchNo & "')"
                        'medixmasterconnWS.Execute BordxMNRBSql
    
        
        BOR.MoveNext
        Loop
    End If
    tmpFormula = "{VIEW_PrintClaimBordx_06.HLTCode} ='" & tmpHltCode & "'" & _
                 " AND {VIEW_PrintClaimBordx_06.BordxRefNo} >= '" & Trim(txtBordxNoTo) & "'" & _
                 " AND {VIEW_PrintClaimBordx_06.BordxRefNo} <= '" & Trim(txtBordxNoFr) & "'"
    'Select Case CboAmt.ListIndex
    '    Case 0
    '        tmpAmtType = "( BELOW RM2000 )"
    '    Case 1
    '        tmpAmtType = "( ABOVE RM2000 )"
        'Case 2
            tmpAmtType = ""
   ' End Select

    If Trim(tmpHltCode) = "S" Then
        If Mid(cboBTBG, 1, 1) = "N" Then
            crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Summ_Sihat.rpt"
        ElseIf Mid(cboBTBG, 1, 1) = "Y" Then
            crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Summ_Sihat_BTB.rpt"
        End If
    Else 'non sihat
        If Mid(cboBTBG, 1, 1) = "N" Then
            crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Summ_NS.rpt"
        ElseIf Mid(cboBTBG, 1, 1) = "Y" Then
            crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Summ_NS_BTB.rpt"
        End If
    End If
    crptSummary.SelectionFormula = tmpFormula
    'crptSummary.SortFields(0) = "+{VIEW_PrintClaimBordx_06.INSCode}"
    'crptSummary.SortFields(1) = "+{VIEW_PrintClaimBordx_06.PLNCode}"
    crptSummary.SortFields(0) = "+{VIEW_PrintClaimBordx_06.BordxRefNo}"
    
    crptSummary.ParameterFields(1) = "txtPayor;" & tmpPayName & ";true"
    crptSummary.ParameterFields(0) = "AmtType;" & tmpAmtType & ";true"
    crptSummary.ParameterFields(2) = "txtBordxRef;" & Trim(txtBordxNO) & ";true"
    If CmdSummaryPreview Then
        crptSummary.Destination = crptToWindow
        crptSummary.WindowState = crptMaximized
        crptSummary.Action = 1
    End If
    If CmdPrintSummary Then
        crptSummary.Destination = crptToPrinter
    End If
    If CmdPrintSummary Then
        If bPrintPass = False And CancelPrinting = False Then
            bPrintPass = PrintFunction
            If bPrintPass = False Then
                Screen.MousePointer = Default
                Exit Sub
            End If
        Else
            crptSummary.PrintReport
        End If
    End If
Screen.MousePointer = Default

End Sub

Private Sub cmdDetailPreview_Click()
Dim tempPath  As String
    If lstPrintBordx.ListItems.Count = 0 Then
        MsgBox "Please search records", vbInformation
    Else
        tempPath = "\\PLUTO\HISSaturn\Conn\mediconnectWSPluto.UDL"
        'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        medixmasterconnWSPluto.Open "File Name=" & tempPath
            Call DetailPrint
        'medixmasterconnWS.Close
        'Set medixmasterconnWS = Nothing
        medixmasterconnWSPluto.Close
        Set medixmasterconnWSPluto = Nothing
        crptSummary.Reset
    End If
End Sub

Private Sub cmdPrintDetails_Click()
Dim tempPath As String
    If lstPrintBordx.ListItems.Count = 0 Then
        MsgBox "Please search records", vbInformation
    Else

        tempPath = ""
        tempPath = "\\PLUTO\HISSaturn\Conn\mediconnectWSPluto.UDL"
        'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        medixmasterconnWSPluto.Open "File Name=" & tempPath
    
        Call DetailPrintExcel
        'Call DetailPrint
        medixmasterconnWSPluto.Close
        Set medixmasterconnWSPluto = Nothing
        crptSummary.Reset
    End If
End Sub

Private Sub CmdPrintSummary_Click()
Dim tempPath As String
    If lstPrintBordx.ListItems.Count = 0 Then
        MsgBox "Please search records", vbInformation
    Else
        tempPath = "\\PLUTO\HISSaturn\Conn\mediconnectWSPluto.UDL"
        'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        medixmasterconnWSPluto.Open "File Name=" & tempPath
    
        Call SummaryPrint
       ' medixmasterconnWS.Close
       ' Set medixmasterconnWS = Nothing
        medixmasterconnWSPluto.Close
        Set medixmasterconnWSPluto = Nothing
        crptSummary.Reset
    End If
End Sub

Private Sub CmdSummaryPreview_Click()
Dim tempPath As String
    If lstPrintBordx.ListItems.Count = 0 Then
        MsgBox "Please search records", vbInformation
    Else

        tempPath = "\\PLUTO\HISSaturn\Conn\mediconnectWSPluto.UDL"
       ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        medixmasterconnWSPluto.Open "File Name=" & tempPath
        Call SummaryPrint
      '  medixmasterconnWS.Close
      '  Set medixmasterconnWS = Nothing
        medixmasterconnWSPluto.Close
        Set medixmasterconnWSPluto = Nothing
        crptSummary.Reset
    End If
End Sub

Private Sub CmdPrintFile_Click()
    
    Screen.MousePointer = vbHourglass
        
    If lstPrintBordx.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstPrintBordx)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    'CboAmt.ListIndex = 0
    cboBTBG.ListIndex = 0
    cboPrintOpt.ListIndex = 0
    CmdPrintSummary.Enabled = False
    cmdPrintDetails.Enabled = False
    CboReportType.ListIndex = 0
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call combolist
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing
End Sub

'******************** Start Command Button ********************************
Private Sub cmdClear_Click()
    lstPrintBordx.ListItems.Clear
    txtBordxNO.Text = 0
    cboPrintOpt.ListIndex = 0
    txtBordxNO.Text = ""
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    CmdPrintSummary.Enabled = False
    cmdPrintDetails.Enabled = False
    CmdSummaryPreview.Enabled = False
    cmdDetailPreview.Enabled = False
    CmdPrintFile.Enabled = False
    CboReportType.ListIndex = 0
End Sub

Private Sub CmdSearch_Click()
    
    Screen.MousePointer = vbHourglass
    CmdSearch.Enabled = False
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
  '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    'If Trim(CboPayor) = "" Then
    '    MsgBox "Please enter Payor", vbOKOnly + vbCritical, DialogBoxTitle
    
    If Trim(txtBordxNoFr) = "" Or Trim(txtBordxNoTo) = "" Then
        MsgBox "Please enter Bordx Reference No", vbOKOnly + vbCritical, DialogBoxTitle
        'TxtBordxNo.SetFocus
    Else
        'ScanCode = TxtBordxNo
        If Mid(ScanCode, 1, 2) = "LO" Then
            txtBordxNO = ""
            Call SearchForm(ScanCode)
        Else
            CmdClear.Enabled = False
            Call SearchRecord
        End If
    End If
    'medixmasterconn.Close
   ' Set medixmasterconn = Nothing
   ' medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    CmdSearch.Enabled = True
    Screen.MousePointer = Default
    
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String

    strAppend = ""
    
    'If Len(CboPayor) Then
    '    StrAppend = StrAppend + " And PAYCode = '" & Mid(CboPayor, 1, 2) & "'"
    'End If
    
    If Len(Trim(txtBordxNoFr)) Then
        strAppend = strAppend + " And BordxRefNo >= '" & ChkStr(txtBordxNoFr) & "'"
    End If

    If Len(Trim(txtBordxNoTo)) Then
        strAppend = strAppend + " And BordxRefNo <= '" & ChkStr(txtBordxNoTo) & "'"
    End If

    'If CboAmt.ListIndex = 0 Then
    '    StrAppend = StrAppend + " AND TotalMgmtFee <= 2000"
    'End If

    'If CboAmt.ListIndex = 1 Then
    '    StrAppend = StrAppend + " AND TotalMgmtFee > 2000"
    'End If
                          
    If txtBordxNoFr = "" Or txtBordxNoTo = "" Then
        CmdPrintSummary.Enabled = False
        cmdPrintDetails.Enabled = False
        CmdSummaryPreview.Enabled = False
        cmdDetailPreview.Enabled = False
        CmdPrintFile.Enabled = False
    Else
        CmdPrintSummary.Enabled = True
        cmdPrintDetails.Enabled = True
        CmdSummaryPreview.Enabled = True
        cmdDetailPreview.Enabled = True
        CmdPrintFile.Enabled = True
    End If
    Call PrintBordxListing(strAppend)
End Function

Private Sub PrintBordxListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim TotalAmt As Variant
Dim cmd As New ADODB.Command
Dim Prm1 As ADODB.Parameter
Dim Prm2 As ADODB.Parameter
Dim PlanCheck As Boolean
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset

    total = 0
    Current = 0
    PlanCheck = False
    lstPrintBordx.ListItems.Clear
               
    Set cmd.ActiveConnection = medixmasterconnWS
    cmd.CommandText = "Print_Bordx"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, 1)
    cmd.Parameters.Append Prm2
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    
    'sql = " Select PLNCode, INSCode, CASNumber, ClaimVersion, " & _
          " PatientName, MBMNumber, HOSName, CLMPayableByMedix2H, CLMTotalActual," & _
          " CLMFinalApprove, CLMPayableByMedix2M, DebitCreditRefNo, BordxRefNo, " & _
          " ClaimBordxStatus, BordxPrintedStatus, HLTCode, PAYCode, MBMBTBG, " & _
          " INVPAYTo, CLMCaseMgmtFees, TotalMgmtFee From View_PrintBordx where 0=0 "
    
     If rst2.EOF = True Then
        MsgBox "No record found ", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrintSummary.Enabled = False
        cmdPrintDetails.Enabled = False
        CmdSummaryPreview.Enabled = False
        cmdDetailPreview.Enabled = False
        CmdPrintFile.Enabled = False
    Else
        
    Do
        Do Until rst2.EOF
        total = rst2.recordCount
        lblTotal = total
             With rst2
                Set Record = lstPrintBordx.ListItems.Add(, , rst2!PLNCode)
                If Len(rst2!CasNumber & "-" & rst2!claimversion) Then
                    Record.SubItems(1) = (rst2!CasNumber & "-" & rst2!claimversion)
                End If
                If Len(rst2!PatientName) Then
                    Record.SubItems(2) = rst2!PatientName
                End If
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(3) = rst2!MBMNumber
                End If
                If Len(rst2!HOSName) Then
                   Record.SubItems(4) = rst2!HOSName
                End If
                                
                If rst2!MBMBTBG = "Y" Then
                    If Len(rst2!CLMPayableByMedix2H) Then
                        Record.SubItems(5) = Format(rst2!CLMPayableByMedix2H, "#,##0.00")
                        tmpBordxAmt = Record.SubItems(5)
                    End If
                Else
                    If Len(rst2!CLMTotalApproved) Then
                        Record.SubItems(5) = Format(rst2!CLMTotalApproved, "#,##0.00")
                        tmpBordxAmt = Record.SubItems(5)
                    End If
                End If
                
                If Len(rst2!CLMTotalActual) Then
                    Record.SubItems(6) = Format(rst2!CLMTotalActual, "#,##0.00")
                End If
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(7) = Format(rst2!CLMTotalApproved, "#,##0.00")
                End If
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(8) = Format(rst2!CLMPayableByMedix2M, "#,##0.00")
                End If
                If Len(rst2!CLMCaseMgmtFees) Then
                    Record.SubItems(9) = Format(rst2!CLMCaseMgmtFees, "#,##0.00")
                End If
                If Len(rst2!DebitCreditRefNo) Then
                    Record.SubItems(10) = rst2!DebitCreditRefNo
                End If
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(11) = rst2!BordxRefNo
                End If
                  
                If rst2!BordxPrintedStatus = True Then
                    Record.SubItems(12) = "YES"
                    cboPrintOpt.ListIndex = 1
                ElseIf rst2!BordxPrintedStatus = False Then
                    cboPrintOpt.ListIndex = 0
                    Record.SubItems(12) = "NO"
                End If
                
                If Len(rst2!HLTCode) Then
                    Record.SubItems(13) = rst2!HLTCode
                End If
                If Len(rst2!PAYCode) Then
                    Record.SubItems(14) = rst2!PAYCode
                End If
                If Len(rst2!PLNCode) Then
                    Record.SubItems(15) = rst2!PLNCode
                End If
                If Trim(rst2!MBMBTBG) = "Y" Then
                    Record.SubItems(15) = rst2!MBMBTBG
                Else
                    Record.SubItems(15) = "N"
                End If
                
                TmpBordxDate = Format(rst2!BordxDate, "dd/mm/yyyy")
                If rst2!HLTCode = "N" Then
                'rst2.Close
                'Set rst2 = Nothing
                '    strsqlPLN = ""
                    strsqlPLN = " Select  PLNCode ,PAYCode,PLNMgmtFee,PLNFloatStatus,PLNFloatEffDate from  PLN where  PLNCode = '" & Trim(rst2!PLNCode) & "' And PAYCode= '" & Trim(rst2!PAYCode) & "' "
                    PLN.Open strsqlPLN, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If PLN.recordCount Then
                            PlanCheck = True
                            
                            If Trim(PLN!PLNMgmtFee) = "Y" Then
                                TmpmgmtFeesStatus = True
                            Else
                                TmpmgmtFeesStatus = False
                            End If
                            If Trim(PLN!PLNFloatStatus) = "Y" Then
                                tmpFloatStatus = True
                                tmpFloatEffDate = PLN!PLNFloatEffDate
                            Else
                                tmpFloatStatus = False
                            End If
                        End If
                        PLN.Close
                        Set PLN = Nothing
                Else
                    TmpmgmtFeesStatus = False
                    tmpFloatStatus = False
                End If
                
                Current = Current + 1
                lblCurrent = Current
                
                
                If rst2!MBMBTBG = "Y" And rst2!InvPayTo = "H" Then
                    TotalAmt = TotalAmt + rst2!CLMPayableByMedix2H
                ElseIf rst2!MBMBTBG = "Y" And rst2!InvPayTo = "M" Then
                    TotalAmt = TotalAmt + rst2!CLMPayableByMedix2M
                Else
                'If Len(rst2!MBMBTBG) Then
                    TotalAmt = TotalAmt + rst2!TotalMgmtFee
                End If
                
                'TotalAmt = TotalAmt + rst2!CLMTotalApproved
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
End Sub

Private Sub lstPrintBordx_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstPrintBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstPrintBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstPrintBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstPrintBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
    Case 10
        SortListView lstPrintBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstPrintBordx, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
    Case 12
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstPrintBordx, ColumnHeader.Index, ldtString, m_blnDirection(14)
    End Select
End Sub



Private Sub GetPayorInfo(tmpPayCode As String, PayorName As String, PayAdd1 As String, PayAdd2 As String, PayAdd3 As String)
Dim PayRec As New ADODB.Recordset
Dim strsqlPay As String
    
    If Trim(tmpPayCode) <> "" Then
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        strsqlPay = "Select PAYCompanyName, PAYAddress1, PAYAddress2, PAYAddress3 " & _
                    " from Pay where PAYCode='" & Trim(tmpPayCode) & "'"
        PayRec.Open strsqlPay, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do While Not PayRec.EOF
            PayorName = IIf(Len(PayRec!PAYCompanyName), Trim(PayRec!PAYCompanyName), "")
            PayAdd1 = IIf(Len(PayRec!PAYAddress1), Trim(PayRec!PAYAddress1), "")
            PayAdd2 = IIf(Len(PayRec!PAYAddress2), Trim(PayRec!PAYAddress2), "")
            PayAdd3 = IIf(Len(PayRec!PAYAddress3), Trim(PayRec!PAYAddress3), "")
            PayRec.MoveNext
        Loop
        PayRec.Close
        Set PayRec = Nothing
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
    End If
End Sub

Private Function GetProdtName(tmpPlnCode As String, tmpHltCode As String, tmpPayCode As String) As String
Dim ProdtRec As New ADODB.Recordset
Dim strsqlPay As String
Dim PLNRec As New ADODB.Recordset
Dim PLNSql As String
Dim tmpProdCat As String
    tmpProdCat = "SH"
    If Trim(tmpHltCode) <> "S" Then
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        PLNSql = "SELECT PLNProCat FROM PLN WHERE PLNCode ='" & Trim(tmpPlnCode) & "' And PayCode ='" & Trim(tmpPayCode) & "'"
        PLNRec.Open PLNSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If PLNRec.recordCount > 0 Then
            If Trim(PLNRec!PLNProCat) <> "" Then
                tmpProdCat = Trim(PLNRec!PLNProCat)
            End If
        End If
        PLNRec.Close
        Set PLNRec = Nothing
       ' medixmasterconn.Close
      '  Set medixmasterconn = Nothing
    End If
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    strsqlPay = "Select ProductCode, ProductName from ProductCategory where ProductCode='" & Trim(tmpProdCat) & "'"
    ProdtRec.Open strsqlPay, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
    Do While Not ProdtRec.EOF
        GetProdtName = IIf(Len(ProdtRec!ProductName), Trim(ProdtRec!ProductName), "")
        ProdtRec.MoveNext
    Loop
    ProdtRec.Close
    Set ProdtRec = Nothing
  '  medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
End Function


Private Sub combolist()
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
   
    Do Until PAY.EOF
        CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
  
End Sub

