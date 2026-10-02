VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInvoiceHOSAC 
   Caption         =   "Invoice Hospital Accounting"
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
   Begin MSComctlLib.ListView lstInvoiceAccList 
      Height          =   5940
      Left            =   120
      TabIndex        =   22
      Top             =   1440
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10478
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hospital"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Inv No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Inv Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Inv Reg Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Worksheet No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Bordx No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Member No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Member Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Patient Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Pay2HOS"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Previous Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "DR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "CR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Date Seq"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Bat No"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   120
      TabIndex        =   23
      Top             =   240
      Width           =   11655
      Begin VB.ComboBox cboHospital 
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   210
         Width           =   3840
      End
      Begin VB.ComboBox cboPrintStatus 
         Height          =   315
         ItemData        =   "FrmInvoiceHOSAC.frx":0000
         Left            =   1560
         List            =   "FrmInvoiceHOSAC.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   450
         Width           =   1695
      End
      Begin VB.TextBox txtBatchNo1 
         Height          =   285
         Left            =   10080
         TabIndex        =   8
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtBatchNo2 
         Height          =   285
         Left            =   10920
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin MSMask.MaskEdBox TxtINVDateFr 
         Height          =   285
         Left            =   6720
         TabIndex        =   2
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtINVDateTo 
         Height          =   285
         Left            =   8160
         TabIndex        =   3
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtREGDateFR 
         Height          =   285
         Left            =   6720
         TabIndex        =   4
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtREGDateTo 
         Height          =   285
         Left            =   8160
         TabIndex        =   5
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtBatchDateFr 
         Height          =   285
         Left            =   6720
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtBatchDateTo 
         Height          =   285
         Left            =   8160
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Inv Date Fr"
         Height          =   255
         Left            =   5520
         TabIndex        =   34
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   7920
         TabIndex        =   33
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   10680
         TabIndex        =   32
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   7920
         TabIndex        =   31
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "Batch Seq Fr"
         Height          =   255
         Left            =   5520
         TabIndex        =   30
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Batch No"
         Height          =   255
         Left            =   9360
         TabIndex        =   29
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   250
         Width           =   975
      End
      Begin VB.Label Label45 
         Caption         =   "Print Status"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label41 
         Caption         =   "To"
         Height          =   255
         Left            =   7920
         TabIndex        =   25
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label40 
         Caption         =   "Inv Reg Date Fr"
         Height          =   255
         Left            =   5520
         TabIndex        =   24
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   7440
      Width           =   11655
      Begin VB.TextBox lblSQENo 
         Height          =   285
         Left            =   4800
         TabIndex        =   39
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox LblTotalAmt 
         Height          =   285
         Left            =   1920
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview"
         Height          =   330
         Left            =   8520
         TabIndex        =   28
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   330
         Left            =   7920
         TabIndex        =   13
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   330
         Left            =   7320
         TabIndex        =   12
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   330
         Left            =   6000
         TabIndex        =   10
         Top             =   200
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   330
         Left            =   7680
         TabIndex        =   18
         Top             =   200
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9360
         TabIndex        =   14
         Top             =   200
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10920
         TabIndex        =   16
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10320
         TabIndex        =   15
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   6720
         TabIndex        =   11
         Top             =   200
         Width           =   615
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   7440
         TabIndex        =   17
         Top             =   200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Seq No"
         Height          =   255
         Left            =   4080
         TabIndex        =   35
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Records"
         Height          =   255
         Left            =   1200
         TabIndex        =   21
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Total Amt"
         Height          =   255
         Left            =   3240
         TabIndex        =   20
         Top             =   195
         Width           =   735
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Invoice Hospital Accounting"
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
Attribute VB_Name = "FrmInvoiceHOSAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim CaseNoArray() As String
Public searchCriteria As String
Dim CASNo As String
Dim CASVer As String
Dim PAYTo As String
Dim INVNo As String
Dim strAppend As String
Private Hospital As String
Private InvoiceNo As String
'Sort Direction for each column in the ListView
Private m_blnDirection(1 To 13) As Boolean

Private Sub cmdPreview_Click()
    InvoiceNo = ""
    
    If lstInvoiceAccList.ListItems.Count = 0 Then
        MsgBox "Please select a record(s) before proceeding!", vbOKOnly + vbCritical

    ElseIf cboPrintStatus = "" Then
        MsgBox "Please select Print Status", vbOKOnly + vbCritical
    Else

        cnt = 0
        recordFound = False

        Screen.MousePointer = vbHourglass
            Call ExportToExcelInvoice(lstInvoiceAccList)
        Screen.MousePointer = Default
    End If

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


Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim X As Boolean
End Sub

Private Sub Form_Load()
    intExit = 0
    lblCurrent = 0
    'lblTotal = 0
    LblTotalAmt = 0
    Call ComboListing
    txtPath = DocPath & "HOSInvoice"
    txtLocalFileLocation = "C:\"
    X = CreatePath("c:\cardPrinting")
    
    If X = False Then
        MsgBox "Error Creating Folder c:\CardPrinting !" & vbNewLine & _
                " Please create that folder manually to store the output of leaflet in sofe copy!"
    End If

End Sub

'************************Start Command Button**************************
Private Sub cmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lstInvoiceAccList.ListItems.Clear
    lblCurrent = 0
    'lblTotal = 0
    LblTotalAmt = 0
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdPrintFile.Enabled = True
End Sub

Public Sub cmdSearch_Click()
    CmdClear.Enabled = False
    CmdExit.Enabled = False
    CmdPrintFile.Enabled = False
    CmdSearch.Enabled = False
    cmdPreview.Enabled = False
    Screen.MousePointer = vbHourglass
        medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        
        If cboPrintStatus = "" Then
            MsgBox "Please select Print Status", vbOKOnly + vbCritical
            cboPrintStatus.SetFocus
        Else
            Arraycnt = 0
            lstInvoiceAccList.Height = 6900
            lstInvoiceAccList.Left = 120
            lstInvoiceAccList.Top = 480
            lstInvoiceAccList.Width = 11655
            Frame1.Visible = False
        
            Call SearchRecord
        End If
        medixmasterconnWS.Close
        Set medixmasterconnWS = Nothing
    Screen.MousePointer = Default
    CmdSearch.Enabled = True
    cmdPreview.Enabled = True

End Sub

Private Sub CmdBack_Click()
    lstInvoiceAccList.Height = 6000
    lstInvoiceAccList.Left = 120
    lstInvoiceAccList.Top = 1500
    lstInvoiceAccList.Width = 11655
    Frame1.Visible = True
End Sub

Private Sub CmdPrintFile_Click()

    InvoiceNo = ""
    medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    
    If lstInvoiceAccList.ListItems.Count = 0 Then
        MsgBox "Please select a record(s) before proceeding!", vbOKOnly + vbCritical
    ElseIf cboPrintStatus = "" Then
        MsgBox "Please select Print Status", vbOKOnly + vbCritical
    Else

        Update = MsgBox("Do you want to print the record?", vbYesNo + vbExclamation)
            If Update = vbYes Then
                cnt = 0
                recordFound = False
        
                Screen.MousePointer = vbHourglass
                
                If Mid(cboPrintStatus, 1, 1) = "N" Then
                    InvoiceNo = GenerateINVHSeqNo(Date)
                    lblSQENo = InvoiceNo
                End If
        
                If Mid(cboPrintStatus, 1, 1) = "N" Then
                    Call UpdatePrintRecord
                End If

                Call ExportListViewtoExcel(lstInvoiceAccList)
                
                Screen.MousePointer = Default
            End If
    End If
    medixmasterconnWS.Close
    Set medixmasterconnWS = Nothing
End Sub

Private Sub CmdView_Click()
    lstInvoiceAccList.Height = 7000
    lstInvoiceAccList.Left = 120
    lstInvoiceAccList.Top = 480
    lstInvoiceAccList.Width = 11655
    Frame1.Visible = False
End Sub
'************************End Command Button**************************


'************************Start sorted ListView**************************

Private Sub lstInvoiceAccList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    Case 4
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
    Case 12
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    Case 13
        SortListView lstInvoiceAccList, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
    End Select
End Sub
'************************End sorted ListView**************************


'************************Start ComboListing**************************

Private Sub ComboListing()
Dim HOS As New ADODB.Recordset
Dim cmd As New ADODB.Command
    
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHOS"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    HOS.CursorLocation = adUseClient
    HOS.CursorType = adOpenForwardOnly
    HOS.CacheSize = 100
    Set HOS = cmd.Execute

    Do Until HOS.EOF
        With HOS
            cboHospital.AddItem Trim(!HosName)
            HOS.MoveNext
        End With
    Loop
    HOS.Close
    Set HOS = Nothing
    
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    
    With cboPrintStatus
        .AddItem "N - New"
        .AddItem "R - Reprint"
        .Text = "N - New"
    End With
     
End Sub
'************************Start ComboListing**************************

'************************Start Search Record**************************

Private Function SearchRecord()
Dim strsql As String
Dim sql As String
    
    strAppend = ""
    Hospital = ""
    
    If Len(Trim(TxtINVDateFr)) And IsDate(TxtINVDateFr) Then
        strAppend = strAppend + " And INVDate >= '" & Format(TxtINVDateFr, "mm/dd/YYYY") & "'"
    End If
    
    If Len(Trim(TxtINVDateTo)) And IsDate(TxtINVDateTo) Then
        strAppend = strAppend + " And INVDate <= '" & Format(TxtINVDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtREGDateFR)) And IsDate(TxtREGDateFR) Then
        strAppend = strAppend + " And INVEnteredDate >= '" & Format(TxtREGDateFR, "mm/dd/YYYY") & "'"
    End If
    
    If Len(Trim(TxtREGDateTo)) And IsDate(TxtREGDateTo) Then
        strAppend = strAppend + " And INVEnteredDate <= '" & Format(TxtREGDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(cboHospital) Then
        Hospital = Replace(cboHospital, "-", " ")
        strAppend = strAppend + " And HOSName = '" & Hospital & "'"
    End If
    
    If Mid(cboPrintStatus, 1, 1) = "N" Then
       sql = ""
       sql = " AND INVPrintStatus = '0' "
       strAppend = strAppend + sql
    End If
    
    If Mid(cboPrintStatus, 1, 1) = "R" Then
       sql = ""
       sql = " AND INVPrintStatus = '1' "
       strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtBatchDateFr)) > 0 And IsDate(txtBatchDateFr) = True Then
        strAppend = strAppend + " And INVBatchDate >= '" & txtBatchDateFr & "'"
    End If
    
    If Len(Trim(txtBatchDateTo)) > 0 And IsDate(txtBatchDateTo) = True Then
        strAppend = strAppend + " And INVBatchDate <= '" & txtBatchDateTo & "'"
    End If
    
    If Len(Trim(txtBatchNo1)) Then
        strAppend = strAppend + " AND INVBatchNo >= '" & Trim(txtBatchNo1) & "'"
    End If
    
    If Len(Trim(txtBatchNo2)) Then
        strAppend = strAppend + " AND INVBatchNo <= '" & Trim(txtBatchNo2) & "'"
    End If
        
    searchCriteria = strAppend
    Call InvoiceAccountListing(strAppend)

End Function
'************************End Search Record**************************


'************************Start List Record**************************
Private Sub InvoiceAccountListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim HOS As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim strsqlHOS As String
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
Dim TotalInvNo As String
Dim DRCRAmt As String
    
    Total = 0
    Current = 0
    TotalAmt = 0
    DRCRAmt = 0
    LblTotalAmt = ""
    lstInvoiceAccList.ListItems.Clear
       
    sql = " Select CASNumber, INVWSVersion, MBMNumber, " & _
          " INVNo, INVDate, MBMName, PatientName, HOSName, " & _
          " INVPayee, PatientName, DebitCreditRefNo, INVBatchDate, INVBatchNo," & _
          " CLMPayableByMedix2H, INVPreviousAmt, BordxRefNo, INVEnteredDate " & _
          " From VIEW_HOSInvoice Where INVPayTo = 'H'  "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If

    rst2.Open sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        CmdPrintFile.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
        'Total = rst2.recordCount
        'lblTotal = Total
        
            With rst2
                                
                If Len(!INVPayee) Then
                    Set Record = lstInvoiceAccList.ListItems.Add(, , !INVPayee)
                Else
                    Set Record = lstInvoiceAccList.ListItems.Add(, , "Empty")
                End If
                
                If Len(rst2!INVNo) Then
                    Record.SubItems(1) = Trim(rst2!INVNo)
                End If
                
                If Len(rst2!InvDate) Then
                    Record.SubItems(2) = Format(rst2!InvDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!INVEnteredDate) Then
                    Record.SubItems(3) = rst2!INVEnteredDate
                End If
                
                If Len(rst2!CASNumber & "-" & rst2!INVWSVersion) Then
                    Record.SubItems(4) = rst2!CASNumber & "-" & rst2!INVWSVersion
                End If
                
                'If Len(rst2!DebitCreditRefNo) Then
                '    Record.SubItems(5) = Trim(rst2!DebitCreditRefNo)
                'End If
                
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(5) = Trim(rst2!BordxRefNo)
                End If
                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(6) = Trim(rst2!MBMNumber)
                End If
                
                If Len(rst2!MBMName) Then
                    Record.SubItems(7) = Trim(rst2!MBMName)
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(8) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!CLMPayableByMedix2H) Then
                    Record.SubItems(9) = Format(rst2!CLMPayableByMedix2H, "#,##0.00;(#,##0.00)")
                End If
                
                If Len(rst2!INVPreviousAmt) Then
                    Record.SubItems(10) = Format(rst2!INVPreviousAmt, "#,##0.00;(#,##0.00)")
                End If
                                
                Current = Current + 1
                lblCurrent = Current
                If Len(rst2!CLMPayableByMedix2H) Then
                    TotalAmt = TotalAmt + rst2!CLMPayableByMedix2H
                End If
                LblTotalAmt = Format(TotalAmt, "#,##0.00")
                
                If rst2!INVPreviousAmt >= 0 And rst2!CLMPayableByMedix2H >= 0 Then
                    DRCRAmt = CDbl(rst2!CLMPayableByMedix2H) - CDbl(rst2!INVPreviousAmt)
                ElseIf rst2!INVPreviousAmt < 0 And rst2!CLMPayableByMedix2H >= 0 Then
                    DRCRAmt = CDbl(rst2!CLMPayableByMedix2H) + CDbl(rst2!INVPreviousAmt)
                ElseIf rst2!INVPreviousAmt < 0 And rst2!CLMPayableByMedix2H < 0 Then
                    DRCRAmt = CDbl(rst2!CLMPayableByMedix2H) - CDbl(rst2!INVPreviousAmt)
                Else
                    If Len(rst2!CLMPayableByMedix2H) Then
                        DRCRAmt = rst2!CLMPayableByMedix2H
                    Else
                        DRCRAmt = 0
                    End If
                End If
                
                If DRCRAmt < 0 Then
                    Record.SubItems(11) = Format(DRCRAmt, "#,##0.00;#,##0.00")
                ElseIf DRCRAmt > 0 Then
                    Record.SubItems(12) = Format(DRCRAmt, "(#,##0.00)")
                ElseIf DRCRAmt = 0 Then
                    Record.SubItems(12) = Format(DRCRAmt, "#,##0.00;#,##0.00")
                End If
                
                If Len(rst2!INVBatchDate) Then
                    Record.SubItems(13) = rst2!INVBatchDate
                End If
                
                If Len(rst2!INVBatchNo) Then
                    Record.SubItems(14) = rst2!INVBatchNo
                End If
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   check = False
                   Exit Do
                End If
            
        Loop
    Loop Until check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdPrintFile.Enabled = True
    'cmdPrint.Enabled = True
    'CmdUpdate.Enabled = True
End Sub
'************************End List Record**************************

'************************Start ComboBox Change/KeyPress**************************

Private Sub cboPrintStatus_Change()
If cboPrintStatus.Text = "" Then
    If InStr(cboPrintStatus.Text, "null") Then
       cboPrintStatus.Enabled = False
    End If
End If
End Sub


Private Sub cboPrintStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer


If cboPrintStatus.Text = "" Then
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPrintStatus.Text, cboPrintStatus.SelStart) + Chr(KeyAscii) + Mid(cboPrintStatus.Text, cboPrintStatus.SelStart + cboPrintStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPrintStatus.ListCount - 1
 
        If NewText = LCase(Left(cboPrintStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPrintStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPrintStatus.Text = ValidValue
                cboPrintStatus.SelStart = 0
                cboPrintStatus.SelLength = Len(ValidValue)
            End If
        End If
End If
End Sub

Private Sub CboHospital_Change()
    If InStr(cboHospital.Text, "null") Then
       cboHospital.Enabled = False
    End If
End Sub

Private Sub CboHospital_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHospital.Text, cboHospital.SelStart) + Chr(KeyAscii) + Mid(cboHospital.Text, cboHospital.SelStart + cboHospital.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboHospital.ListCount - 1
 
        If NewText = LCase(Left(cboHospital.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHospital.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboHospital.Text = ValidValue
                cboHospital.SelStart = 0
                cboHospital.SelLength = Len(ValidValue)
            End If
        End If
End Sub
'************************End ComboBox Change/KeyPress**************************


'**********************Start Update Record*******************************
Private Sub UpdatePrintRecord()
Dim strsqlCLM As String
Dim CLM As New ADODB.Recordset
Dim Update
Dim strsql As String
Dim AccPeriod As New ADODB.Recordset
             
        
        
        strsqlCLM = " Select CASNumber, INVWSVersion, CLMPayableByMedix2H " & _
                    " From VIEW_HOSInvoice where  INVPayTo = 'H'"
        strsqlCLM = strsqlCLM + strAppend
        
        CLM.Open strsqlCLM, medixmasterconnWS, adOpenKeyset, adLockOptimistic
      'CLM.Close
      'Set CLM = Nothing
      'AccPeriod.Close
      'Set AccPeriod = Nothing
    Do Until CLM.EOF
    
        strsql = " Select CASNumber, INVWSVersion, InvPrintStatus, InvPreviousAmt, " & _
                 " InvBatchDate, InvBatchNo, InvPrintLogon, LastAccPeriodDate, INVPayTo " & _
                 " From ClaimInvoice Where CASNumber = '" & Trim(CLM!CASNumber) & "' And INVWSVersion = '" & Trim(CLM!INVWSVersion) & "'"
        
        AccPeriod.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        'AccPeriod.Close
        Do Until AccPeriod.EOF
            With AccPeriod
                !INVBatchDate = Mid(InvoiceNo, 1, 9)
                !INVBatchNo = Trim(Mid(InvoiceNo, 11, 1))
                !InvPrintLogon = Logon
                !InvPrintStatus = "1"
                If Len(CLM!CLMPayableByMedix2H) Then
                    !INVPreviousAmt = Format(CLM!CLMPayableByMedix2H, "#,##0.00;(#,##0.00)")
                Else
                    !INVPreviousAmt = Format(0, "#,##0.00")
                End If
            AccPeriod.Update
            AccPeriod.MoveNext
            End With
        Loop
        AccPeriod.Close
        Set AccPeriod = Nothing
        
    CLM.MoveNext
    Loop
    CLM.Close
    Set CLM = Nothing
End Sub
'**********************End Update Record*******************************

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

Private Sub HOSAccASCII()
Dim strsql As String
Dim AccPeriod As New ADODB.Recordset
Dim TotalSupplementary As Integer
Dim strsplOthers As String
Dim MBMOthers As New ADODB.Recordset
Dim MBMGroupCompany As String
Dim MBMCRELCode As String
Dim check As Boolean
Dim MBMDateJoin As String
        
    cnt = 0
        
    strsql = " Select CASNumber, INVWSVersion, MBMNumber, INSCode, " & _
             " INVNo, CASDischargeDate, CaseRegDate, PLNCode, " & _
             " CASDateAdmitted, INVPayTo, INVPayee, PatientName, AccPeriodMonth, AccPeriodYear " & _
             " MBMName, InvPrintStatus, INVPreviousAmt " & _
             " From VIEW_InvoiceAccount where ClaimStatus = 'C' And Substring(CASNumber,1,2) = '" & Mid(CboPayor, 1, 2) & "'"
               
    AccPeriod.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    Do Until AccPeriod.EOF
        For ii = 1 To AccPeriod.recordCount

         DataString = DataString & l("M", 1)
         DataString = DataString & l(AccPeriod!CASNumber & "-" & AccPeriod!INVWSVersion, 20)
         DataString = DataString & l(AccPeriod!MBMNumber, 25)
         DataString = DataString & l(AccPeriod!MBMName, 60)
         DataString = DataString & l(AccPeriod!PatientName, 60)
         DataString = DataString & l(AccPeriod!INSCode, 1)
         DataString = DataString & l(AccPeriod!PLNCode, 4)
         DataString = DataString & l(AccPeriod!INVNo, 10)
         'DataString = DataString & l(AccPeriod!INVAmount, 10)
         'DataString = DataString & l(AccPeriod!CLMTotalApproved, 10)
        
         AccPeriod.MoveNext
         
         'Debug.Print DataString
         Print #1, DataString
         Print #2, DataString

        Next ii
    Loop

End Sub


'Private Sub UpdateRecord()
'Dim Update
'Dim strsql As String
'Dim AccPeriod As New ADODB.Recordset
         
    'strsql = " Select CASNumber, INVWSVersion, AccPeriodStatus, " & _
    " AccPeriodMonth, AccPeriodYear, LastAccPeriodUser, LastAccPeriodDate " & _
    " From ClaimInvoice where CASNumber = '" & Trim(CASNo) & "' " & _
    " AND INVWSVersion = '" & Trim(CASVer) & "' AND INVNo = '" & Trim(INVNo) & "'"
    'AccPeriod.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
      
    'With AccPeriod
        '!LastAccPeriodUser = Logon
        '!LastAccPeriodDate = Date
        '!AccPeriodMonth = Trim(CboPeriodMth)
        '!AccPeriodYear = Trim(CboPeriodYr)
    'End With
    'AccPeriod.Update
    'AccPeriod.Close
    'Set AccPeriod = Nothing
'End Sub


'Private Sub cmdUpdate_Click()
'Dim ii As Integer
'Dim cnt As Integer
'Dim recordFound As Boolean
'Dim delimiter1, delimiter2, delimiter3 As Integer
    
        
        'If lstInvoiceAccList.listitems.count = 0 Then
            'MsgBox "Please select a record(s) before proceeding!", vbOKOnly + vbCritical
        'ElseIf UBound(CaseNoArray) = 0 Then
            'MsgBox "Please select at least one record to be updated! ", vbCritical
        'ElseIf CboPeriodMth = "" Then
            'MsgBox "Month cannot be Empty.", vbCritical
        'ElseIf CboPeriodYr = "" Then
            'MsgBox "Year cannot be empty!", vbCritical
        'Else
    
        'Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            'If Update = vbYes Then
                'cnt = 0
           
                'recordFound = False
                'If lstInvoiceAccList.listitems.count <> 0 Then
                    'For ii = 0 To UBound(CaseNoArray)
                        'If Trim(CaseNoArray(ii)) <> "" Then
                                                                                
                            'delimiter1 = InStr(1, CaseNoArray(ii), "~")
                            'CASNo = Trim(Mid(CaseNoArray(ii), 1, IIf(delimiter1 = 0, 10, delimiter1 - 1)))
                            'delimiter2 = InStr(1, CaseNoArray(ii), "-")
                            'CASVer = Trim(Mid(CaseNoArray(ii), delimiter1 + 1, delimiter2 - delimiter1 - 1))
                            'INVNo = Trim(Mid(CaseNoArray(ii), delimiter2 + 1)) ', delimiter3 - 1))
        
                            'Call UpdateRecord
                            'recordFound = True
                        
                        'End If
                    'Next ii
                    'Erase CaseNoArray
                    'Arraycnt = 0
                    'ReDim CaseNoArray(0)
                    'MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    'Call ClearForm(Me)
                    'Call InvoiceAccountListing(searchCriteria)
                'End If
            'End If
            
        'End If
'End Sub

'Private Sub CmdPrint_Click()
'Dim ii As Integer
'Dim cnt As Integer
'Dim recordFound As Boolean
'Dim delimiter1, delimiter2, delimiter3 As Integer
'Dim SystemDate As Date

        'If txtInvNo1 = "" Then
            'MsgBox "Please select Invoice No", vbOKOnly + vbCritical
            'CboPayor.SetFocus
        'ElseIf cboPrintStatus = "" Then
            'MsgBox "Please select Print Status", vbOKOnly + vbCritical
            'cboPrintStatus.SetFocus
        'ElseIf lstInvoiceAccList.listitems.count = 0 Then
            'MsgBox "Please select a record(s) before proceeding!", vbOKOnly + vbCritical
        'Else
    
        'Update = MsgBox("Do you want to print the record?", vbYesNo + vbExclamation)
            'If Update = vbYes Then
                'cnt = 0
                
                'recordFound = False
                
                    'If Mid(cboPrintStatus, 1, 1) = "N" Then
                        'InvoiceNo = GenerateINVMSeqNo(Mid(CboPayor, 1, 2), Date)
                    'Else
                        'InvoiceNo = Mid(CboPayor, 1, 2) & Format(Date, "ddmmyyyy")
                    'End If
                        'Open txtPath & "\" & InvoiceNo For Output As #1
                        'Open txtLocalFileLocation & InvoiceNo For Output As #2
                            
                            'Call HOSAccASCII
                            
                            'If Mid(cboPrintStatus, 1, 1) = "N" Then
                                'Call UpdatePrintRecord
                            'End If
                            
                            'recordFound = True
                        
                      'End If
                    
                    'Close #1
                    'Close #2
                    'MsgBox "Record Printed...", vbOKOnly + vbInformation, DialogBoxTitle
                    'Call ClearForm(Me)
                    'Call InvoiceAccountListing(searchCriteria)
                'End If
            'End If
'End Sub


