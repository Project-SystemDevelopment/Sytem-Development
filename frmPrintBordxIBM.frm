VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintBordxIBM 
   Caption         =   "Print Bordx - IBM"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   11895
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   660
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
         ItemData        =   "frmPrintBordxIBM.frx":0000
         Left            =   7200
         List            =   "frmPrintBordxIBM.frx":000A
         TabIndex        =   22
         Top             =   750
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.TextBox TxtBordxNo 
         Height          =   285
         Left            =   4800
         TabIndex        =   21
         Top             =   1200
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   10320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   960
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox TxtBordxNoFr 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   195
         Width           =   2055
      End
      Begin VB.TextBox TxtBordxNoTo 
         Height          =   285
         Left            =   3480
         TabIndex        =   1
         Top             =   195
         Visible         =   0   'False
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
         ItemData        =   "frmPrintBordxIBM.frx":0035
         Left            =   13080
         List            =   "frmPrintBordxIBM.frx":003F
         TabIndex        =   19
         Top             =   780
         Visible         =   0   'False
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
         ItemData        =   "frmPrintBordxIBM.frx":0057
         Left            =   13080
         List            =   "frmPrintBordxIBM.frx":0061
         TabIndex        =   18
         Top             =   1065
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox CboAmt 
         Height          =   315
         ItemData        =   "frmPrintBordxIBM.frx":0076
         Left            =   8280
         List            =   "frmPrintBordxIBM.frx":0083
         TabIndex        =   17
         Text            =   "All"
         Top             =   780
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
         TabIndex        =   29
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Top             =   220
         Visible         =   0   'False
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
         Left            =   6120
         TabIndex        =   27
         Top             =   795
         Visible         =   0   'False
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
         Left            =   12120
         TabIndex        =   26
         Top             =   1110
         Visible         =   0   'False
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
         Left            =   12120
         TabIndex        =   25
         Top             =   840
         Visible         =   0   'False
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
         TabIndex        =   24
         Top             =   220
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
         Left            =   7320
         TabIndex        =   23
         Top             =   795
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   7200
      Width           =   11655
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   2760
         TabIndex        =   33
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   975
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
         TabIndex        =   11
         Top             =   240
         Width           =   1095
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
         Left            =   4200
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
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
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdPrintDetails 
         Caption         =   "Print &Bordx"
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
         Left            =   6600
         TabIndex        =   8
         Top             =   240
         Width           =   975
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
         Left            =   5040
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   2
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
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin Crystal.CrystalReport crptSummary 
         Left            =   3000
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   120
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstPrintBordx 
      Height          =   6255
      Left            =   120
      TabIndex        =   30
      Top             =   960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11033
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
      Caption         =   "Print Bordx - IBM"
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
      TabIndex        =   31
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmPrintBordxIBM"
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
Dim strsqlPAY As String
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
Dim CaseCnt As String
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.WORKSHEET

    PLUTOStatus = False
    PrintedDate = Format(Date, "YYYY/MM/DD")
    bPrintPass = False
    CancelPrinting = False
    
    PrntCnt = 0
    NPrntCnt = 0
    CaseCnt = 0
    
    For LstCnt = 1 To lstPrintBordx.ListItems.Count
        tmpHltCode = lstPrintBordx.ListItems(LstCnt).SubItems(13)
        tmpPayCode = lstPrintBordx.ListItems(LstCnt).SubItems(14)
        tmpPln = lstPrintBordx.ListItems(LstCnt).Text
    Next LstCnt
    
    Screen.MousePointer = vbHourglass
    tmpProdtName = GetProdtName(tmpPln, tmpHltCode, tmpPayCode)
    Call GetPayorInfo(tmpPayCode, PAYName, PayAdd1, PayAdd2, PayAdd3)
    
    tmpPayName = tmpPayCode & " - " & PAYName
    
    
        sqlbor = "Select * from VIEW_PrintClaimBordx_IBM   where BordxRefNo = '" & Trim(TxtBordxNoFr) & "'"
        BOR.Open sqlbor, medixmasterconnWS, adOpenForwardOnly, adLockReadOnly
    
        Do While Not BOR.EOF
            If objExcel Is Nothing Then
                Set objExcel = CreateObject("Excel.application")
                    objExcel.Interactive = False
                If objExcel.Visible = False Then
                    objExcel.Visible = True
                End If
                
                objExcel.WindowState = xlMaximized
                Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "Claim\IBMInpatientMedicalClaims.xlt")
                Set objWorksheet = objWorkbook.Sheets(2)
                objExcel.DisplayAlerts = False
                Set objWorksheet = objWorkbook.Worksheets.Item(2)
            End If
            
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
            If BOR!HosName <> "" Then
                tmpHospital = Replace(Trim(BOR!HosName), "'", "''")
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
                    
                With BOR
                    CaseCnt = CDbl(CaseCnt) + 1
                    
                        If CaseCnt = 1 Then
                            objWorksheet.Cells(1 + CDbl(CaseCnt), "C") = BOR!BordxRefNo
                            objWorksheet.Cells(2 + CDbl(CaseCnt), "C") = TmpBordxDate
                        End If
                        
                        objWorksheet.Cells(6 + CDbl(CaseCnt), "A") = BOR!MBMDepartment
                        objWorksheet.Cells(6 + CDbl(CaseCnt), "C") = BOR!CasNumber & "-" & BOR!claimversion
                        objWorksheet.Cells(6 + CDbl(CaseCnt), "D") = BOR!MBMEmployeeNo
                        objWorksheet.Cells(6 + CDbl(CaseCnt), "E") = TempMBMName
                        objWorksheet.Cells(6 + CDbl(CaseCnt), "F") = BOR!MBMNumber
                        objWorksheet.Cells(6 + CDbl(CaseCnt), "G") = BOR!MBMPatientCovID
                        If BOR!MBMPatientCovID = "00" Then
                            objWorksheet.Cells(6 + CDbl(CaseCnt), "H") = TempMBMName
                        Else
                            objWorksheet.Cells(6 + CDbl(CaseCnt), "H") = tmpMBMCovName
                        End If
                        objWorksheet.Cells(6 + CDbl(CaseCnt), "I") = BOR!MBMIcBcPp

                        If BOR!MBMPatientCovID = "00" Then
                            If Len(BOR!MBMIcBcPp) <= 4 Then
                                objWorksheet.Cells(6 + CDbl(CaseCnt), "J") = "-"
                            Else
                                objWorksheet.Cells(6 + CDbl(CaseCnt), "J") = BOR!MBMIcBcPp
                            End If
                        Else
                            If Len(BOR!MBMCICBCPPNo) <= 4 Then
                                objWorksheet.Cells(6 + CDbl(CaseCnt), "J") = "-"
                            Else
                                objWorksheet.Cells(6 + CDbl(CaseCnt), "J") = BOR!MBMCICBCPPNo
                            End If
                        End If
                        objWorksheet.Cells(6 + CDbl(CaseCnt), "K") = tmpHospital
                        objWorksheet.Cells(6 + CDbl(CaseCnt), "L") = BOR!ICDCode1
                        'If BOR!ICDCode1 = "K08" Then
                         '   objWorksheet.Cells(6 + CDbl(CaseCnt), "K") = "Dental Treatment"
                        'Else
                            objWorksheet.Cells(6 + CDbl(CaseCnt), "M") = TempDiagDesc1
                        'End If
                        If BOR!ICDCode1 = "K08" Then
                            objWorksheet.Cells(6 + CDbl(CaseCnt), "N") = "Dental"
                        Else
                            objWorksheet.Cells(6 + CDbl(CaseCnt), "N") = "Inpatient"
                        End If
                        objWorksheet.Cells(6 + CDbl(CaseCnt), "O") = BOR!CLMTotalActual
                        objWorksheet.Cells(6 + CDbl(CaseCnt), "P") = CDbl(BOR!CLMBordxAmt) + CDbl(BOR!CLMBTBGAmt)
                        objWorksheet.Cells(6 + CDbl(CaseCnt), "Q") = BOR!CLMBTBGAmt * -1
                        objWorksheet.Cells(6 + CDbl(CaseCnt), "R") = BOR!CLMBordxAmt
                        
                        If BOR!InvPayTo = "H" Then
                            objWorksheet.Cells(6 + CDbl(CaseCnt), "S") = "C"
                        Else
                            objWorksheet.Cells(6 + CDbl(CaseCnt), "S") = "R"
                        End If
                        
                        objWorksheet.Cells(6 + CDbl(CaseCnt), "T") = tmpDOA
                        objWorksheet.Cells(6 + CDbl(CaseCnt), "U") = tmpDOD
                        objWorksheet.Cells(6 + CDbl(CaseCnt), "V") = BOR!ClaimNatureCode
                        
                        
                End With
                    
                BOR.MoveNext
                Loop
                BOR.Close
                Set BOR = Nothing
            
        
        Set objWorksheet = objWorkbook.Sheets(1)
        objExcel.DisplayAlerts = False
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
        sqlbor = ""
        CaseCnt = 0
        sqlbor = "Select * from VIEW_PrintBordxSummary_IBM   where BordxRefNo = '" & Trim(TxtBordxNoFr) & "'"
        BOR.Open sqlbor, medixmasterconnWS, adOpenForwardOnly, adLockReadOnly
        Do While Not BOR.EOF
            CaseCnt = CDbl(CaseCnt) + 1
            If CaseCnt = 1 Then
                objWorksheet.Cells(10, "K") = BOR!BordxRefNo
                objWorksheet.Cells(12, "K") = TmpBordxDate
                objWorksheet.Cells(20, "B") = "Actual Inpatient medical cost incurred for the month of" & " " & Format(TmpBordxDate, "mmmm") & " " & Year(TmpBordxDate)
            End If
            
            If Trim(BOR!MBMDepartment) = "REGULAR" Then
            
            
                objWorksheet.Cells(24, "I") = BOR!TotalCases
                'objWorksheet.Cells(24, "K") = BOR!CLMTotalActual
                
                objWorksheet.Cells(24, "K") = CDbl(BOR!CLMBordxAmt) + CDbl(BOR!CLMBTBGAmt)
                objWorksheet.Cells(24, "L") = (BOR!CLMBTBGAmt) * -1
                objWorksheet.Cells(24, "M") = BOR!CLMBordxAmt
            ElseIf Trim(BOR!MBMDepartment) = "FTH" Then
                objWorksheet.Cells(26, "I") = BOR!TotalCases
                objWorksheet.Cells(26, "K") = CDbl(BOR!CLMBordxAmt) + CDbl(BOR!CLMBTBGAmt)
                objWorksheet.Cells(26, "L") = (BOR!CLMBTBGAmt) * -1
                objWorksheet.Cells(26, "M") = BOR!CLMBordxAmt
            ElseIf Trim(BOR!MBMDepartment) = "RETIREE" Then
                objWorksheet.Cells(28, "I") = BOR!TotalCases
                objWorksheet.Cells(28, "K") = CDbl(BOR!CLMBordxAmt) + CDbl(BOR!CLMBTBGAmt)
                objWorksheet.Cells(28, "L") = (BOR!CLMBTBGAmt) * -1
                objWorksheet.Cells(28, "M") = BOR!CLMBordxAmt
            Else
                objWorksheet.Cells(30, "I") = BOR!TotalCases
                objWorksheet.Cells(30, "K") = CDbl(BOR!CLMBordxAmt) + CDbl(BOR!CLMBTBGAmt)
                objWorksheet.Cells(30, "L") = (BOR!CLMBTBGAmt) * -1
                objWorksheet.Cells(30, "M") = BOR!CLMBordxAmt
            End If
        
            CaseCnt = CDbl(CaseCnt) + 1
        BOR.MoveNext
        Loop
        BOR.Close
        Set BOR = Nothing
    'End If
    objExcel.DisplayAlerts = True
    objExcel.Interactive = True
    Set objWorksheet = Nothing
    Set objExcel = Nothing
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
            
                crptSummary.PrintReport
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
Dim strsqlPAY As String
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
    
    strsqlPlutoBor = "Select BordxRefNo,BordxDate from TempBordxSummary where BordxRefNo >= '" & Trim(TxtBordxNoFr) & "' And BordxRefNo <= '" & Trim(TxtBordxNoTo) & "'"
    PlutoBor.Open strsqlPlutoBor, medixmasterconnWSPluto, adOpenForwardOnly, adLockReadOnly
    
    Do While Not PlutoBor.EOF
        PLUTOStatus = True
        TmpBordxDate = PlutoBor!BordxDate
        PlutoBor.MoveNext
    Loop
    PlutoBor.Close
    Set PlutoBor = Nothing

    If PLUTOStatus = False Then
        sqlbor = "Select * from VIEW_PrintClaimBordx_06   where BordxRefNo >= '" & Trim(TxtBordxNoFr) & "' And BordxRefNo <= '" & Trim(TxtBordxNoTo) & "'"
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
                        "'" & BOR!InvPayTo & "', '" & BOR!HosName & "', '" & BOR!INSCode & "','" & BOR!PLNCode & "', '" & BOR!HLTCode & "', '" & BOR!MBMNumber & "','" & BOR!MBMBTBG & "', '" & BOR!MBMPatientCovID & "', " & _
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
                 " AND {VIEW_PrintClaimBordx_06.BordxRefNo} >= '" & Trim(TxtBordxNoTo) & "'" & _
                 " AND {VIEW_PrintClaimBordx_06.BordxRefNo} <= '" & Trim(TxtBordxNoFr) & "'"
    'Select Case CboAmt.ListIndex
    '    Case 0
    '        tmpAmtType = "( BELOW RM2000 )"
    '    Case 1
    '        tmpAmtType = "( ABOVE RM2000 )"
        'Case 2
            tmpAmtType = ""
   ' End Select

    If Trim(tmpHltCode) = "S" Then
        If Mid(CboBTBG, 1, 1) = "N" Then
            crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Summ_Sihat.rpt"
        ElseIf Mid(CboBTBG, 1, 1) = "Y" Then
            crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Summ_Sihat_BTB.rpt"
        End If
    Else 'non sihat
        If Mid(CboBTBG, 1, 1) = "N" Then
            crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Summ_NS.rpt"
        ElseIf Mid(CboBTBG, 1, 1) = "Y" Then
            crptSummary.ReportFileName = crdPath & "\Report\ClaimBordx\ClaimBordx_Summ_NS_BTB.rpt"
        End If
    End If
    crptSummary.SelectionFormula = tmpFormula
    'crptSummary.SortFields(0) = "+{VIEW_PrintClaimBordx_06.INSCode}"
    'crptSummary.SortFields(1) = "+{VIEW_PrintClaimBordx_06.PLNCode}"
    crptSummary.SortFields(0) = "+{VIEW_PrintClaimBordx_06.BordxRefNo}"
    
    crptSummary.ParameterFields(1) = "txtPayor;" & tmpPayName & ";true"
    crptSummary.ParameterFields(0) = "AmtType;" & tmpAmtType & ";true"
    crptSummary.ParameterFields(2) = "txtBordxRef;" & Trim(TxtBordxNo) & ";true"
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
      '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        medixmasterconnWSPluto.Open "File Name=" & tempPath
            Call DetailPrint
      '  medixmasterconnWS.Close
       ' Set medixmasterconnWS = Nothing
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

       ' tempPath = ""
       ' tempPath = "\\PLUTO\HISSaturn\Conn\mediconnectWSPluto.UDL"
       ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        'medixmasterconnWSPluto.Open "File Name=" & tempPath
    
        Call DetailPrint
       ' medixmasterconnWS.Close
        'Set medixmasterconnWS = Nothing
        'medixmasterconnWSPluto.Close
        'Set medixmasterconnWSPluto = Nothing
        'crptSummary.Reset
    End If
End Sub

Private Sub CmdPrintSummary_Click()
Dim tempPath As String
    If lstPrintBordx.ListItems.Count = 0 Then
        MsgBox "Please search records", vbInformation
    Else
        tempPath = "\\PLUTO\HISSaturn\Conn\mediconnectWSPluto.UDL"
       ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
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
        'medixmasterconnWS.Close
       ' Set medixmasterconnWS = Nothing
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
    CboBTBG.ListIndex = 0
    cboPrintOpt.ListIndex = 0
    CmdPrintSummary.Enabled = False
    cmdPrintDetails.Enabled = False
    CboReportType.ListIndex = 0
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call combolist
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
End Sub

'******************** Start Command Button ********************************
Private Sub CmdClear_Click()
    lstPrintBordx.ListItems.Clear
    TxtBordxNo.Text = 0
    cboPrintOpt.ListIndex = 0
    TxtBordxNo.Text = ""
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
    
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
 '   medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    'If Trim(CboPayor) = "" Then
    '    MsgBox "Please enter Payor", vbOKOnly + vbCritical, DialogBoxTitle
    
    If Trim(TxtBordxNoFr) = "" Then
        MsgBox "Please enter Bordx Reference No", vbOKOnly + vbCritical, DialogBoxTitle
        'TxtBordxNo.SetFocus
    Else
        'ScanCode = TxtBordxNo
        If Mid(ScanCode, 1, 2) = "LO" Then
            TxtBordxNo = ""
            Call SearchForm(ScanCode)
        Else
            CmdClear.Enabled = False
            Call SearchRecord
        End If
    End If
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    CmdSearch.Enabled = True
    Screen.MousePointer = Default
    
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Function SearchRecord()
Dim StrSql As String
Dim strAppend As String
Dim sql As String

    strAppend = ""
    
    'If Len(CboPayor) Then
    '    StrAppend = StrAppend + " And PAYCode = '" & Mid(CboPayor, 1, 2) & "'"
    'End If
    
    If Len(Trim(TxtBordxNoFr)) Then
        strAppend = strAppend + " And BordxRefNo = '" & ChkStr(TxtBordxNoFr) & "'"
    End If

    'If Len(Trim(TxtBordxNoTo)) Then
    '    strAppend = strAppend + " And BordxRefNo <= '" & ChkStr(TxtBordxNoTo) & "'"
    'End If'

    'If CboAmt.ListIndex = 0 Then
    '    StrAppend = StrAppend + " AND TotalMgmtFee <= 2000"
    'End If

    'If CboAmt.ListIndex = 1 Then
    '    StrAppend = StrAppend + " AND TotalMgmtFee > 2000"
    'End If
                          
    If TxtBordxNoFr = "" Then
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
                If Len(rst2!HosName) Then
                   Record.SubItems(4) = rst2!HosName
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
Dim strsqlPAY As String
    
    If Trim(tmpPayCode) <> "" Then
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        strsqlPAY = "Select PAYCompanyName, PAYAddress1, PAYAddress2, PAYAddress3 " & _
                    " from Pay where PAYCode='" & Trim(tmpPayCode) & "'"
        PayRec.Open strsqlPAY, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        Do While Not PayRec.EOF
            PayorName = IIf(Len(PayRec!PAYCompanyName), Trim(PayRec!PAYCompanyName), "")
            PayAdd1 = IIf(Len(PayRec!PAYAddress1), Trim(PayRec!PAYAddress1), "")
            PayAdd2 = IIf(Len(PayRec!PAYAddress2), Trim(PayRec!PAYAddress2), "")
            PayAdd3 = IIf(Len(PayRec!PAYAddress3), Trim(PayRec!PAYAddress3), "")
            PayRec.MoveNext
        Loop
        PayRec.Close
        Set PayRec = Nothing
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
    End If
End Sub

Private Function GetProdtName(tmpPlnCode As String, tmpHltCode As String, tmpPayCode As String) As String
Dim ProdtRec As New ADODB.Recordset
Dim strsqlPAY As String
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
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
    End If
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    strsqlPAY = "Select ProductCode, ProductName from ProductCategory where ProductCode='" & Trim(tmpProdCat) & "'"
    ProdtRec.Open strsqlPAY, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
    Do While Not ProdtRec.EOF
        GetProdtName = IIf(Len(ProdtRec!ProductName), Trim(ProdtRec!ProductName), "")
        ProdtRec.MoveNext
    Loop
    ProdtRec.Close
    Set ProdtRec = Nothing
    'medixmasterconnMaint.Close
    'Set medixmasterconnMaint = Nothing
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



