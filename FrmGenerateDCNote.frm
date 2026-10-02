VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmGenerateDCNote 
   Caption         =   "Generate Debit/Credit Note Listing"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   7320
      Width           =   11655
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "&Exp To Excel"
         Height          =   375
         Left            =   8880
         TabIndex        =   35
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10680
         TabIndex        =   12
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9960
         TabIndex        =   11
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdReprint 
         Caption         =   "&Reprocess"
         Height          =   375
         Left            =   7920
         TabIndex        =   10
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Process"
         Height          =   375
         Left            =   7200
         TabIndex        =   9
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   6120
         TabIndex        =   8
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   5400
         TabIndex        =   7
         Top             =   180
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblTotalAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3720
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1920
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblCurrent 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label22 
         Caption         =   "of"
         Height          =   255
         Left            =   840
         TabIndex        =   28
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   2880
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1260
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   11655
      Begin VB.Frame Frame3 
         Caption         =   "Print Status"
         Height          =   615
         Left            =   5880
         TabIndex        =   34
         Top             =   180
         Width           =   2295
         Begin VB.OptionButton OptReprint 
            Caption         =   "Reprint"
            Height          =   255
            Left            =   1200
            TabIndex        =   3
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptPrint 
            Caption         =   "Print"
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.OptionButton OptCredit 
         Caption         =   "Credit Note"
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         Top             =   200
         Width           =   1335
      End
      Begin VB.OptionButton OptDebit 
         Caption         =   "Debit Note"
         Height          =   255
         Left            =   1440
         TabIndex        =   0
         Top             =   200
         Width           =   1335
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtAmtFr 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   1455
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtAmtTo 
         Height          =   285
         Left            =   3840
         TabIndex        =   16
         Top             =   1455
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtRefFr 
         Height          =   285
         Left            =   6840
         MaxLength       =   15
         TabIndex        =   13
         Top             =   1380
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtRefTo 
         Height          =   285
         Left            =   9240
         MaxLength       =   15
         TabIndex        =   14
         Top             =   1380
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSMask.MaskEdBox txtDate1 
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDate2 
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "DN / CN"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   585
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Amount"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1515
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         Height          =   255
         Left            =   3360
         TabIndex        =   22
         Top             =   1515
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Ref No From"
         Height          =   255
         Left            =   5640
         TabIndex        =   21
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   20
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Date From"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   885
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   3360
         TabIndex        =   18
         Top             =   885
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView lstDebitCreditList 
      Height          =   5775
      Left            =   120
      TabIndex        =   31
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10186
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
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Payor"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ref No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Patient Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Revising"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Previous Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Difference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "DR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "CR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Hospital"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "Plan Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Hosp. Type"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Object.Width           =   2540
      EndProperty
   End
   Begin Crystal.CrystalReport crptSummary 
      Left            =   480
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog_PrintBordx 
      Left            =   1200
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Generate Debit/Credit Note Listing"
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
      TabIndex        =   36
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmGenerateDCNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const PRINTDCNOTE_PAYOR = 0
Const PRINTDCNOTE_DCNO = 1
Dim DCNoArray() As String
Dim DRCRAmt As String
Public DNCNCriteria As String
Public intExit As Integer
Private m_blnDirection(1 To 10) As Boolean

Private Sub CmdPrintFile_Click()
    Call ExportToExcelDrCrListing(lstDebitCreditList)
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

Private Sub cmdReprint_Click()
    
    Screen.MousePointer = vbHourglass
    cmdReprint.Enabled = False
    If lstDebitCreditList.listitems.count <> 0 Then
        medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        'RptDebitCredit.show
        If OptDebit.value = True Then
               crptSummary.ReportFileName = crdPath & "\Report\DebitNote1.rpt"
               crptSummary.Destination = crptToWindow
               crptSummary.WindowState = crptMaximized
               crptSummary.SelectionFormula = "{VIEW_DebitListing1.PAYCODE} = '" & Trim(Mid(CboPayor, 1, 2)) & "'"
               crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
               crptSummary.Action = 1
            
        ElseIf OptCredit.value = True Then
               crptSummary.ReportFileName = crdPath & "\Report\CreditNote1.rpt"
               crptSummary.Destination = crptToWindow
               crptSummary.WindowState = crptMaximized
               crptSummary.SelectionFormula = "{VIEW_CreditListing1.PAYCODE} = '" & Trim(Mid(CboPayor, 1, 2)) & "'"
               crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
               crptSummary.Action = 1
        End If
        medixmasterconnWS.Close
        Set medixmasterconnWS = Nothing
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    cmdReprint.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    DNCNCriteria = ""
    OptPrint.value = True
    Call ComboListing
End Sub

'******************* Start Command Button ************************************

Private Sub CmdClear_Click()
    
    lstDebitCreditList.listitems.Clear
    CboPayor.Clear
    txtDate1.Text = "__/__/____"
    txtDate2.Text = "__/__/____"
    txtRefFr.Text = ""
    txtRefTo.Text = ""
    txtAmtFr.Text = ""
    txtAmtTo.Text = ""
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    Call ComboListing
End Sub

Private Sub cmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub CmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub CmdSearch_Click()
    DRCRAmt = 0
    CmdSearch.Enabled = False
    If CboPayor = "" Then
        MsgBox "Please Enter Payor Code.", vbCritical
        CboPayor.SetFocus
    Else
        CmdClear.Enabled = False
        Screen.MousePointer = vbHourglass
            medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                Call SearchRecord
            medixmasterconnWS.Close
            Set medixmasterconnWS = Nothing
        Screen.MousePointer = Default
    End If
    CmdSearch.Enabled = True
End Sub

Private Sub CmdPrint_Click()
Dim cnt As Integer
Dim ii As Integer
Dim tmpDCNo As String
Dim tmpPayor As String
Static Arraycnt As Integer
Dim rsUpdate As New ADODB.Recordset
Dim strsql As String
Dim Sql As String


Screen.MousePointer = vbHourglass
    CmdPrint.Enabled = False
    If lstDebitCreditList.listitems.count <> 0 Then
        
        medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                        
        If OptDebit.value = True Then
               crptSummary.ReportFileName = crdPath & "\Report\DebitNote.rpt"
               crptSummary.Destination = crptToWindow
               crptSummary.WindowState = crptMaximized
               crptSummary.SelectionFormula = "{VIEW_DebitListing.PAYCODE} = '" & Trim(Mid(CboPayor, 1, 2)) & "'"
               crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
               crptSummary.Action = 1
               Call PrintDebitNote
            
        ElseIf OptCredit.value = True Then
               crptSummary.ReportFileName = crdPath & "\Report\CreditNote.rpt"
               crptSummary.Destination = crptToWindow
               crptSummary.WindowState = crptMaximized
               crptSummary.SelectionFormula = "{VIEW_CreditListing.PAYCODE} = '" & Trim(Mid(CboPayor, 1, 2)) & "'"
               crptSummary.ParameterFields(0) = "txtPayor;" & CboPayor & ";true"
               crptSummary.Action = 1
               PrintCreditNote
        End If
        medixmasterconnWS.Close
        Set medixmasterconnWS = Nothing
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    CmdPrint.Enabled = True
    Screen.MousePointer = vbDefault
    

End Sub
'******************* End Command Button ************************************


'******************* Start SearchRecord ************************************
Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim Sql As String

    
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    
    strAppend = ""
        
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 4, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If (txtDate1) <> "__/__/____" And IsDate(txtDate1) = True _
        And (txtDate2) <> "__/__/____" And IsDate(txtDate2) = True Then
        Sql = ""
        Sql = " AND DebitCreditDate >= '" & Format(Trim(txtDate1), "mm/dd/yyyy") & _
              "' And DebitCreditDate <= '" & Format(Trim(txtDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + Sql
    End If
    
    If Len(Trim(txtDate1)) > 0 And IsDate(txtDate1) = True Then
        strAppend = strAppend + " And DebitCreditDate >= '" & Format(txtDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDate2)) > 0 And IsDate(txtDate2) = True Then
        strAppend = strAppend + " And DebitCreditDate <= '" & Format(txtDate2, "mm/dd/yyyy") & "'"
    End If
               
    If OptDebit.value = True Then
        If OptPrint.value = True Then
            DNCNCriteria = strAppend
            Call DNListing(strAppend)
        Else
            DNCNCriteria = strAppend
            Call DNListing1(strAppend)
        End If
    Else
        If OptPrint.value = True Then
            DNCNCriteria = strAppend
            Call CRListing(strAppend)
        Else
            DNCNCriteria = strAppend
            Call CRListing1(strAppend)
        End If
    End If
    
End Function
'******************* End SearchRecord ************************************

'******************* Start DNListing ************************************

Private Sub DNListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
'Dim DRCRAmt As String
    Total = 0
    Current = 0
    
    lstDebitCreditList.listitems.Clear
            
    Sql = " Select DebitCreditRefNo, DebitCreditDate, " & _
          " PLNCode, MBMNumber, PatientName, " & _
          " HOSName, PAYCode, DbCrListPrintstatus, " & _
          " CLMPayableByMedix2M, DebitCreditRevised, DebitCreditPreviousAmt " & _
          " From VIEW_DebitListing where 0 = 0"
        'ClaimType,
     
    If Len(Trim(strAppend)) Then
        Sql = Sql + strAppend
    End If
    'If OptPrint.value = True Then
       'Sql = Sql + " AND (DbCrListPrintstatus is  null or DbCrListPrintstatus = 0)"
    'End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Do
            Do Until rst2.EOF
            Total = rst2.recordCount
            lblTotal = Total
            
                With rst2
                    Set Record = lstDebitCreditList.listitems.Add(, , rst2!PAYCode)
                    If Len(rst2!DebitCreditRefNo) Then
                        Record.SubItems(1) = rst2!DebitCreditRefNo
                    End If
                    If Len(rst2!DebitCreditDate) Then
                       Record.SubItems(2) = Format(rst2!DebitCreditDate, "dd/mm/yyyy")
                    End If
                    If Len(rst2!PatientName) Then
                        Record.SubItems(3) = rst2!PatientName
                    End If
                    If Len(rst2!CLMPayableByMedix2M) Then
                        Record.SubItems(4) = Format(rst2!CLMPayableByMedix2M, "#,##0.00;(#,##0.00)")
                    End If
                    If Len(rst2!debitCreditRevised) Then
                        Record.SubItems(5) = rst2!debitCreditRevised
                    End If
                    If Len(rst2!debitCreditPreviousAmt) Then
                        Record.SubItems(6) = Format(rst2!debitCreditPreviousAmt, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If IsNumeric(rst2!debitCreditPreviousAmt) = False Then
                        rst2!debitCreditPreviousAmt = 0
                    End If
                    
                    If IsNumeric(rst2!CLMPayableByMedix2M) = False Then
                        rst2!CLMPayableByMedix2M = 0
                    End If
                    
                    If rst2!debitCreditPreviousAmt >= 0 And rst2!CLMPayableByMedix2M >= 0 Then
                        DRCRAmt = CDbl(rst2!CLMPayableByMedix2M) - CDbl(rst2!debitCreditPreviousAmt)
                        Record.SubItems(7) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                    ElseIf rst2!debitCreditPreviousAmt < 0 And rst2!CLMPayableByMedix2M >= 0 Then
                        DRCRAmt = CDbl(rst2!CLMPayableByMedix2M) + CDbl(rst2!debitCreditPreviousAmt)
                        Record.SubItems(7) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                    ElseIf rst2!debitCreditPreviousAmt < 0 And rst2!CLMPayableByMedix2M < 0 Then
                        DRCRAmt = CDbl(rst2!CLMPayableByMedix2M) - CDbl(rst2!debitCreditPreviousAmt)
                        Record.SubItems(7) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                    Else
                        DRCRAmt = rst2!CLMPayableByMedix2M
                        Record.SubItems(7) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If DRCRAmt < 0 Then
                        Record.SubItems(8) = Format(DRCRAmt, "#,##0.00;#,##0.00")
                    ElseIf DRCRAmt >= 0 Then
                        Record.SubItems(9) = Format(DRCRAmt, "(#,##0.00)")
                    End If
                'If DRCRAmt < 0 Then
                '    Record.SubItems(8) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                'ElseIf DRCRAmt >= 0 Then
                '    Record.SubItems(9) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                'End If

                    If Len(rst2!HOSName) Then
                        Record.SubItems(10) = rst2!HOSName
                    End If
                    'If Len(rst2!PAYCode) Then
                        'Record.SubItems(7) = rst2!PAYCode
                    'End If
                    If Len(rst2!PLNCode) Then
                        Record.SubItems(11) = rst2!PLNCode
                    End If
                    'If Len(rst2!ClaimType) Then
                        'Record.SubItems(9) = rst2!ClaimType
                    'End If
                    Current = Current + 1
                    lblCurrent = Current
                    If Len(rst2!CLMPayableByMedix2M) Then
                        TotalAmt = TotalAmt + rst2!CLMPayableByMedix2M
                    End If
                    LblTotalAmt = Format(TotalAmt, "#,##0.00;(#,##0.00)")
                    
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

Private Sub DNListing1(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
    
    Total = 0
    Current = 0
    
    lstDebitCreditList.listitems.Clear
            
    Sql = " Select DebitCreditRefNo, DebitCreditDate, " & _
          " PLNCode, MBMNumber, PatientName, " & _
          " HOSName, PAYCode, DbCrListPrintstatus, " & _
          " CLMPayableByMedix2M, DebitCreditRevised, DebitCreditPreviousAmt " & _
          " From VIEW_DebitListing1 where 0 = 0 "
     'ClaimType,
     
    If Len(Trim(strAppend)) Then
        Sql = Sql + strAppend
    End If
    'If OptPrint.value = True Then
       'Sql = Sql + " AND (DbCrListPrintstatus is  null or DbCrListPrintstatus = 0)"
    'End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Do
            Do Until rst2.EOF
            Total = rst2.recordCount
            lblTotal = Total
            
                With rst2
                    Set Record = lstDebitCreditList.listitems.Add(, , rst2!PAYCode)
                    If Len(rst2!DebitCreditRefNo) Then
                        Record.SubItems(1) = rst2!DebitCreditRefNo
                    End If
                    If Len(rst2!DebitCreditDate) Then
                       Record.SubItems(2) = Format(rst2!DebitCreditDate, "dd/mm/yyyy")
                    End If
                    If Len(rst2!PatientName) Then
                        Record.SubItems(3) = rst2!PatientName
                    End If
                    If Len(rst2!CLMPayableByMedix2M) Then
                        Record.SubItems(4) = Format(rst2!CLMPayableByMedix2M, "#,##0.00;(#,##0.00)")
                    End If
                    If Len(rst2!debitCreditRevised) Then
                        Record.SubItems(5) = rst2!debitCreditRevised
                    End If
                    If Len(rst2!debitCreditPreviousAmt) Then
                        Record.SubItems(6) = Format(rst2!debitCreditPreviousAmt, "#,##0.00;(#,##0.00)")
                    End If
                    If rst2!debitCreditPreviousAmt >= 0 And rst2!CLMPayableByMedix2M >= 0 Then
                        DRCRAmt = CDbl(rst2!CLMPayableByMedix2M) - CDbl(rst2!debitCreditPreviousAmt)
                        Record.SubItems(7) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                    ElseIf rst2!debitCreditPreviousAmt < 0 And rst2!CLMPayableByMedix2M >= 0 Then
                        DRCRAmt = CDbl(rst2!CLMPayableByMedix2M) + CDbl(rst2!debitCreditPreviousAmt)
                        Record.SubItems(7) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                    ElseIf rst2!debitCreditPreviousAmt < 0 And rst2!CLMPayableByMedix2M < 0 Then
                        DRCRAmt = CDbl(rst2!CLMPayableByMedix2M) - CDbl(rst2!debitCreditPreviousAmt)
                        Record.SubItems(7) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                    Else
                        DRCRAmt = rst2!CLMPayableByMedix2M
                        Record.SubItems(7) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If DRCRAmt >= 0 Then
                        Record.SubItems(8) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                    ElseIf DRCRAmt < 0 Then
                        Record.SubItems(9) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                    End If
                    If Len(rst2!HOSName) Then
                        Record.SubItems(10) = rst2!HOSName
                    End If
                    'If Len(rst2!PAYCode) Then
                        'Record.SubItems(7) = rst2!PAYCode
                    'End If
                    If Len(rst2!PLNCode) Then
                        Record.SubItems(11) = rst2!PLNCode
                    End If
                    'If Len(rst2!ClaimType) Then
                        'Record.SubItems(9) = rst2!ClaimType
                    'End If
                    Current = Current + 1
                    lblCurrent = Current
                    If Len(rst2!CLMPayableByMedix2M) Then
                        TotalAmt = TotalAmt + rst2!CLMPayableByMedix2M
                    End If
                    LblTotalAmt = Format(TotalAmt, "#,##0.00;(#,##0.00)")
                    
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
'******************* End DNListing ************************************

'******************* Start CRListing ************************************

Private Sub CRListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant

    
    Total = 0
    Current = 0
    
    lstDebitCreditList.listitems.Clear
            
    Sql = " Select DebitCreditRefNo, DebitCreditDate, " & _
          " PLNCode, MBMNumber, PatientName, " & _
          " HOSName, PAYCode, DbCrListPrintstatus, " & _
          " DbCrListPrintstatus, CLMPayableByMedix2M, DebitCreditRevised, DebitCreditPreviousAmt " & _
          " From VIEW_CreditListing where 0 = 0 "
          
     
    If Len(Trim(strAppend)) Then
        Sql = Sql + strAppend
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
         
    Do
        Do Until rst2.EOF
        Total = rst2.recordCount
        lblTotal = Total
        
            With rst2
                Set Record = lstDebitCreditList.listitems.Add(, , rst2!PAYCode)
                If Len(rst2!DebitCreditRefNo) Then
                    Record.SubItems(1) = rst2!DebitCreditRefNo
                End If
                If Len(rst2!DebitCreditDate) Then
                   Record.SubItems(2) = Format(rst2!DebitCreditDate, "dd/mm/yyyy")
                End If
                'If rst2!mbmpatientcovid = "0" Then
                    If Len(rst2!PatientName) Then
                        Record.SubItems(3) = rst2!PatientName
                    End If
                'Else
                    'If Len(rst2!MBMCName) Then
                        'Record.SubItems(2) = rst2!MBMCName
                    'End If
                'End If
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(4) = Format(rst2!CLMPayableByMedix2M, "#,##0.00;(#,##0.00)")
                End If
                If Len(rst2!debitCreditRevised) Then
                    Record.SubItems(5) = rst2!debitCreditRevised
                End If
                If Len(rst2!debitCreditPreviousAmt) Then
                    Record.SubItems(6) = Format(rst2!debitCreditPreviousAmt, "#,##0.00;(#,##0.00)")
                End If
                
                If rst2!debitCreditPreviousAmt >= 0 And rst2!CLMPayableByMedix2M >= 0 Then
                    DRCRAmt = CDbl(rst2!CLMPayableByMedix2M) - CDbl(rst2!debitCreditPreviousAmt)
                    Record.SubItems(7) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                ElseIf rst2!debitCreditPreviousAmt < 0 And rst2!CLMPayableByMedix2M >= 0 Then
                    DRCRAmt = CDbl(rst2!CLMPayableByMedix2M) + CDbl(rst2!debitCreditPreviousAmt)
                    Record.SubItems(7) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                ElseIf rst2!debitCreditPreviousAmt < 0 And rst2!CLMPayableByMedix2M < 0 Then
                    DRCRAmt = CDbl(rst2!CLMPayableByMedix2M) - CDbl(rst2!debitCreditPreviousAmt)
                    Record.SubItems(7) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                Else
                    DRCRAmt = rst2!CLMPayableByMedix2M
                    Record.SubItems(7) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                End If
                
                If DRCRAmt < 0 Then
                    Record.SubItems(8) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                ElseIf DRCRAmt >= 0 Then
                    Record.SubItems(9) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(10) = rst2!HOSName
                End If
                'If Len(rst2!PAYCode) Then
                    'Record.SubItems(7) = rst2!PAYCode
                'End If
                If Len(rst2!PLNCode) Then
                    Record.SubItems(11) = rst2!PLNCode
                End If
                'If Len(rst2!ClaimType) Then
                    'Record.SubItems(9) = rst2!ClaimType
                'End If
                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!CLMPayableByMedix2M) Then
                    TotalAmt = TotalAmt + rst2!CLMPayableByMedix2M
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00;(#,##0.00)")
                
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

Private Sub CRListing1(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant

    Total = 0
    Current = 0
    
    lstDebitCreditList.listitems.Clear
            
    Sql = " Select DebitCreditRefNo, DebitCreditDate, " & _
          " PLNCode, MBMNumber, PatientName, " & _
          " HOSName, PAYCode, DbCrListPrintstatus, " & _
          " DbCrListPrintstatus, CLMPayableByMedix2M, DebitCreditRevised, DebitCreditPreviousAmt " & _
          " From VIEW_CreditListing1 where 0 = 0"
            
        'ClaimType,
            
    If Len(Trim(strAppend)) Then
        Sql = Sql + strAppend
    End If
    'If OptPrint.value = True Then
       'Sql = Sql + " AND (DbCrListPrintstatus is  null or DbCrListPrintstatus = 0)"
    'End If
       
    'End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
         
    Do
        Do Until rst2.EOF
        Total = rst2.recordCount
        lblTotal = Total
        
            With rst2
                Set Record = lstDebitCreditList.listitems.Add(, , rst2!PAYCode)
                If Len(rst2!DebitCreditRefNo) Then
                    Record.SubItems(1) = rst2!DebitCreditRefNo
                End If
                If Len(rst2!DebitCreditDate) Then
                   Record.SubItems(2) = Format(rst2!DebitCreditDate, "dd/mm/yyyy")
                End If
                'If rst2!mbmpatientcovid = "0" Then
                    If Len(rst2!PatientName) Then
                        Record.SubItems(3) = rst2!PatientName
                    End If
                'Else
                    'If Len(rst2!MBMCName) Then
                        'Record.SubItems(2) = rst2!MBMCName
                    'End If
                'End If
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(4) = Format(rst2!CLMPayableByMedix2M, "#,##0.00;(#,##0.00)")
                End If
                If Len(rst2!debitCreditRevised) Then
                    Record.SubItems(5) = rst2!debitCreditRevised
                End If
                If Len(rst2!debitCreditPreviousAmt) Then
                    Record.SubItems(6) = Format(rst2!debitCreditPreviousAmt, "#,##0.00;(#,##0.00)")
                End If
                If rst2!debitCreditPreviousAmt >= 0 And rst2!CLMPayableByMedix2M >= 0 Then
                    DRCRAmt = CDbl(rst2!CLMPayableByMedix2M) - CDbl(rst2!debitCreditPreviousAmt)
                    Record.SubItems(7) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                ElseIf rst2!debitCreditPreviousAmt < 0 And rst2!CLMPayableByMedix2M >= 0 Then
                    DRCRAmt = CDbl(rst2!CLMPayableByMedix2M) + CDbl(rst2!debitCreditPreviousAmt)
                    Record.SubItems(7) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                ElseIf rst2!debitCreditPreviousAmt < 0 And rst2!CLMPayableByMedix2M < 0 Then
                    DRCRAmt = CDbl(rst2!CLMPayableByMedix2M) - CDbl(rst2!debitCreditPreviousAmt)
                    Record.SubItems(7) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                Else
                    DRCRAmt = rst2!CLMPayableByMedix2M
                    Record.SubItems(7) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                End If
                
                If DRCRAmt < 0 Then
                    Record.SubItems(8) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                ElseIf DRCRAmt >= 0 Then
                    Record.SubItems(9) = Format(DRCRAmt, "#,##0.00;(#,##0.00)")
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(10) = rst2!HOSName
                End If
                'If Len(rst2!PAYCode) Then
                    'Record.SubItems(7) = rst2!PAYCode
                'End If
                If Len(rst2!PLNCode) Then
                    Record.SubItems(11) = rst2!PLNCode
                End If
                'If Len(rst2!ClaimType) Then
                    'Record.SubItems(9) = rst2!ClaimType
                'End If
                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!CLMPayableByMedix2M) Then
                    TotalAmt = TotalAmt + rst2!CLMPayableByMedix2M
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00;(#,##0.00)")
                
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
'******************* End CRListing ************************************



'******************* Start ComboListing ************************************

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim Sql As String
    
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
 
End Sub
'******************* End ComboListing ************************************

'**********************Start Sorted ListView *******************************
Private Sub lstDebitCreditList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    
    Case 1
        SortListView lstDebitCreditList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstDebitCreditList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstDebitCreditList, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    Case 4
        SortListView lstDebitCreditList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstDebitCreditList, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstDebitCreditList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstDebitCreditList, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstDebitCreditList, ColumnHeader.Index, ldtString, m_blnDirection(8)
    'Case 9
        'SortListView lstDebitCreditList, ColumnHeader.Index, ldtString, m_blnDirection(9)
    'Case 10
        'SortListView lstDebitCreditList, ColumnHeader.Index, ldtString, m_blnDirection(10)
    End Select
End Sub
'**********************End Sorted ListView *******************************

'******************* Start ComboBox Change/keyPress ************************************

Private Sub CboPayor_Click()
If CboPayor.Text = "" Then
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    lstDebitCreditList.listitems.Clear
End If
End Sub

Private Sub CboPayor_Change()
If CboPayor.Text = "" Then
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
End If
    lstDebitCreditList.listitems.Clear
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
End Sub

'Private Sub CboPayor_KeyPress(KeyAscii As Integer)
'D 'im NewText As String
'Di 'm ValidCount As Integer
'Dim ValidValue As String
'Dim i As Integer
 '   If KeyAscii >= 32 And KeyAscii <> 127 Then
  '  NewText = LCase(Left(CboPayor.Text, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor.Text, CboPayor.SelStart + CboPayor.SelLength + 1))
 '
    'ValidCount = 0
  ''  ValidValue = ""
 
    'For i = 0 To CboPayor.ListCount - 1
 
     '   If NewText = LCase(Left(CboPayor.list(i), Len(NewText))) Then
      '          ValidCount = ValidCount + 1
       '         ValidValue = CboPayor.list(i)
'        End If
 '       Next
  '      If ValidCount <= 1 Then KeyAscii = 0
   '         If ValidCount = 1 Then
 '               CboPayor.Text = ValidValue
    '            CboPayor.SelStart = 0
     '           CboPayor.SelLength = Len(ValidValue)
      '      End If
'        End If
'End Sub
'******************* End ComboBox Change/keyPress ************************************

Private Sub OptPrint_Click()
    If OptPrint.value = True Then
        CmdPrint.Enabled = True
        cmdReprint.Enabled = False
    End If
End Sub

Private Sub OptReprint_Click()
    If OptReprint.value = True Then
        cmdReprint.Enabled = True
        CmdPrint.Enabled = False
    End If
End Sub

Private Function PrintDebitNote()
Dim strsql As String
Dim PrintDNCN As New ADODB.Recordset


    If CboPayor <> "" Then
        
        
        If lstDebitCreditList.listitems.count = 0 Then
            MsgBox "Please select record.", vbCritical
        Else
            strsql = "Select DebitCreditRefNo, DbCrListPrintstatus, DebitCreditRevised, DebitCreditPreviousAmt From Claim " & _
                     "where substring(DebitCreditRefNo ,1 ,1) = 'D' " & _
                     "AND (DbCrListPrintstatus is null or DbCrListPrintstatus = 0) " & _
                     "AND substring(CASNumber, 1 ,2) = '" & Trim(Mid(CboPayor, 1, 2)) & "'"
       
            PrintDNCN.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
            If PrintDNCN.recordCount Then
                Do Until PrintDNCN.EOF
                    With PrintDNCN
                        !DbCrListPrintstatus = True
                        !debitCreditRevised = False
                        !debitCreditPreviousAmt = 0
                    End With
                    PrintDNCN.Update
                    PrintDNCN.MoveNext
                Loop
            End If
            PrintDNCN.Close
            Set PrintDNCN = Nothing
        End If

           'lstDebitCreditList.listitems.clear
           'CboPayor.ListIndex = 0
           'txtDate1.Text = "__/__/____"
           'txtDate2.Text = "__/__/____"
           'txtRefFr.Text = ""
           'txtRefTo.Text = ""
           'txtAmtFr.Text = ""
           'txtAmtTo.Text = ""
           'lblCurrent = 0
           'lblTotal = 0
           'LblTotalAmt = 0
    End If
End Function

Private Function PrintCreditNote()
Dim strsql As String
Dim PrintDNCN As New ADODB.Recordset


    If CboPayor <> "" Then
           
        If lstDebitCreditList.listitems.count = 0 Then
            MsgBox "Please select record.", vbCritical
        Else
            strsql = "Select DebitCreditRefNo, DbCrListPrintstatus, DebitCreditRevised, DebitCreditPreviousAmt From Claim " & _
                     "where substring(DebitCreditRefNo ,1 ,1) = 'C' " & _
                     "AND DbCrListPrintstatus is null " & _
                     "AND substring(CASNumber, 1 ,2) = '" & Trim(Mid(CboPayor, 1, 2)) & "'"
                                             
            PrintDNCN.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
            
            Do Until PrintDNCN.EOF
                With PrintDNCN
                    !DbCrListPrintstatus = True
                    !debitCreditRevised = False
                    !debitCreditPreviousAmt = 0
                End With
                PrintDNCN.Update
                PrintDNCN.MoveNext
            Loop
            PrintDNCN.Close
            Set PrintDNCN = Nothing
        End If
        'lstDebitCreditList.listitems.clear
        'CboPayor.ListIndex = 0
        'txtDate1.Text = "__/__/____"
        'txtDate2.Text = "__/__/____"
        'txtRefFr.Text = ""
        'txtRefTo.Text = ""
        'txtAmtFr.Text = ""
        'txtAmtTo.Text = ""
        'lblCurrent = 0
        'lblTotal = 0
        'LblTotalAmt = 0
    End If
End Function


Private Sub txtDate1_LostFocus()
    If IsDate(txtDate1) Then
    Else
        If txtDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtDate2_LostFocus()
    If IsDate(txtDate2) Then
    Else
        If txtDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDate2.SetFocus
        End If
    End If
End Sub
