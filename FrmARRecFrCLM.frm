VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmARRecFrCLM 
   Caption         =   "Date Send to Audit Review"
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
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   6960
      Width           =   11655
      Begin VB.TextBox lblCurrent 
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print to &File"
         Height          =   375
         Left            =   9000
         TabIndex        =   10
         Top             =   220
         Width           =   1095
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   7440
         TabIndex        =   8
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   10080
         TabIndex        =   11
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10680
         TabIndex        =   12
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   6720
         TabIndex        =   7
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton CmdToAC 
         Caption         =   "&Update"
         Height          =   375
         Left            =   8160
         TabIndex        =   9
         Top             =   220
         Width           =   855
      End
      Begin VB.CommandButton CmdCheck 
         Caption         =   "&Check All"
         Height          =   375
         Left            =   5880
         TabIndex        =   6
         Top             =   220
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   11655
      Begin VB.TextBox TxtWksNo2 
         Height          =   285
         Left            =   3240
         MaxLength       =   20
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox TxtWksNo1 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   5040
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtDateSent2 
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Top             =   620
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDateSent1 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   620
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboTokenStatus 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   18
         Top             =   670
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Date Sent"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   670
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   405
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Case No"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   405
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Token Status"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1005
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstClaimDelivery 
      Height          =   5295
      Left            =   120
      TabIndex        =   21
      Top             =   1680
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Case Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Date Send to A/R"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Date Returned to CLM"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Passed By"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Date Send to Audit Review"
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
      TabIndex        =   22
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmARRecFrCLM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Dim searchCriteria As String
Dim listloop1 As Integer
Dim listloop2 As Integer
Dim Check As Boolean
Dim Check2 As Boolean
Dim Current2 As String
Dim TotalAmt2 As String
Dim CASNo As String
Private m_blnDirection(1 To 6) As Boolean

Private Sub ChkScan_Click()
    If ChkScan.Value = 1 Then
        lstClaimDelivery.ListItems.Clear
        Label14.Visible = False
        TxtWksNo2.Visible = False
        TxtWksNo1.SetFocus
    Else
        lstClaimDelivery.ListItems.Clear
        Label14.Visible = True
        TxtWksNo2.Visible = True
        TxtWksNo1.SetFocus
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

Private Sub Form_Load()
    intExit = 0
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    Check = False
    Check2 = False
    Current2 = 0
    TotalAmt2 = 0
    Label14.Visible = False
    TxtWksNo2.Visible = False
    ReDim ClaimNoArray(0)
    Call ComboListing
    ItemChecked = False
End Sub
'**********************Start Command Button *******************************
Private Sub cmdPrint_Click()
    TxtWksNo1.SetFocus
    Screen.MousePointer = vbHourglass
        'Call ExportToExcelWksToAC(lstClaimDelivery)
    If lstClaimDelivery.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstClaimDelivery)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdCheck_Click()
    SetCheckAllItems True
End Sub

Private Sub CmdClear_Click()
    CboTokenStatus.Clear
    Call ClearForm(Me)
    lstClaimDelivery.ListItems.Clear
    Current2 = 0
    TotalAmt2 = 0
    TxtWksNo1.SetFocus
    Check = False
    Check2 = False
    lblCurrent = 0
    lblTotal = 0
    LblTotalAmt = 0
    ReDim ClaimNoArray(0)
    Call ComboListing
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
    ItemChecked = False
    LblTotalAmt = ""
    lblTotal = ""
    lblCurrent = ""
    ReDim ClaimNoArray(0)
    CmdSearch.Enabled = False
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If ChkScan.Value = 1 Then
        If TxtWksNo1 = "" Then
            MsgBox "Please Enter Case No", vbOKOnly + vbCritical, DialogBoxTitle
            TxtWksNo1.SetFocus
            CmdClear.Enabled = True
            CmdExit.Enabled = True
            cmdPrint.Enabled = True
            CmdToAC.Enabled = True
        Else
            Screen.MousePointer = vbHourglass
            listrecord = lstClaimDelivery.ListItems.Count
            If listrecord <> 0 Then
                For listloop1 = 1 To listrecord
                    Call CaseNoCheck
                Next listloop1
            End If
                If Check = True Then
                    MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                ElseIf Check = False Then
                    CmdClear.Enabled = True
                    CmdExit.Enabled = True
                    cmdPrint.Enabled = True
                    CmdToAC.Enabled = True
                    Call SearchRecord
                End If
                Check = False
                CmdClear.Enabled = True
                CmdExit.Enabled = True
                cmdPrint.Enabled = True
                CmdToAC.Enabled = True
            Screen.MousePointer = Default
            TxtWksNo1 = ""
            TxtWksNo1.SetFocus
        End If
        TxtWksNo1.SetFocus
    Else
            Screen.MousePointer = vbHourglass
            CmdClear.Enabled = False
            CmdExit.Enabled = False
            cmdPrint.Enabled = False
            CmdToAC.Enabled = False
            Call SearchRecord
            Screen.MousePointer = Default
        CmdClear.Enabled = True
        CmdExit.Enabled = True
        cmdPrint.Enabled = True
        CmdToAC.Enabled = True
    End If
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    CmdSearch.Enabled = True
End Sub

Private Function SearchRecord()
On Error Resume Next
Dim strsql As String
Dim strAppend As String
Dim sql As String
    strAppend = ""
    If ChkScan.Value = 1 Then
        If Len(Trim(TxtWksNo1)) Then
            strAppend = strAppend + " AND CASNumber = '" & Trim(TxtWksNo1) & "'"
        End If
    Else
        If Len(Trim(TxtWksNo1)) Then
            strAppend = strAppend + " AND CASNumber >= '" & Trim(TxtWksNo1) & "'"
        End If
        
        If Len(Trim(TxtWksNo2)) Then
            strAppend = strAppend + " AND CASNumber <= '" & Trim(TxtWksNo2) & "'"
        End If
    End If
    
    If IsDate(txtDateSent1) = True Then
        strAppend = strAppend + " And CASDRC >= '" & Format(txtDateSent1, "mm/dd/yyyy") & "'"
    End If
      
    If IsDate(txtDateSent2) = True Then
        strAppend = strAppend + " And CASDRC <= '" & Format(txtDateSent2, "mm/dd/yyyy") & "'"
    End If
    
    If Mid(CboTokenStatus, 1, 1) = "0" Then
        strAppend = strAppend + " And (CASRecStatus = '0' OR CASRecStatus = '1' OR CASRecStatus IS NULL)"
    End If
    
    If Mid(CboTokenStatus, 1, 1) = "1" Then
        strAppend = strAppend + " And CASRecStatus = '1'"
    End If
    
    If Mid(CboTokenStatus, 1, 1) = "2" Then
        strAppend = strAppend + " And CASRecStatus = '0'"
    End If
    
    searchCriteria = strAppend
    
    If ChkScan.Value = 0 Then
        Call WksDeliveryListing(strAppend)
    Else
        Call WksDeliveryListing2(strAppend)
    End If
End Function

Private Sub CmdStop_Click()
    intExit = 1
    CmdClear.Enabled = True
    CmdToAC.Enabled = True
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
    Check = False
    Check2 = False
    TxtWksNo1.SetFocus
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdToAC_Click()
On Error Resume Next
Dim ClmNo As String
Dim ClmVer As String
Dim i As Integer
Dim i2 As Integer
Dim ItemChecked As Boolean
Dim Update
    TxtWksNo1.SetFocus
    ItemChecked = False
    CmdToAC.Enabled = False
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        If lstClaimDelivery.ListItems.Count = 0 Then
            MsgBox "Please select a record before proceeding!", vbOKOnly + vbCritical
        Else
            Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            
            If Update = vbYes Then
                For i = 1 To lstClaimDelivery.ListItems.Count
                    If lstClaimDelivery.ListItems(i).Checked = True Then
                        ItemChecked = True
                    Exit For
                    End If
                Next i
                
                If ItemChecked = True Then
                    
                    For i = 1 To lstClaimDelivery.ListItems.Count
                       If lstClaimDelivery.ListItems(i).Checked = True Then
                          
                          ClmNo = Trim(lstClaimDelivery.ListItems(i).ListSubItems(1).Text)
                          Call UpdateSentCLM(ClmNo)
                          
                       End If
                    Next i
                    
                    MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
                    
                    lstClaimDelivery.ListItems.Clear
                
                    If ChkScan.Value = 0 Then
                        Call WksDeliveryListing(searchCriteria)
                    End If
                End If
            End If
        End If
        medixmasterconn.Close
        Set medixmasterconn = Nothing
        medixmasterconnWS.Close
        Set medixmasterconnWS = Nothing
    CmdToAC.Enabled = True
End Sub
'**********************End Command Button *******************************

'**********************Start BordxDelivery Listing  *******************************
Private Sub WksDeliveryListing(Optional strAppend As String)
On Error Resume Next
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlCount As String
Dim ClaimAdmin As Boolean
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
    Total = 0
    Current = 0
    lstClaimDelivery.ListItems.Clear
    sql = " Select CASNumber, CASDRC, CASDTC, CASRecBy, CASRecStatus " & _
          " From CASOthers where 0 = 0 "
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdToAC.Enabled = True
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
                
            With rst2
                Set Record = lstClaimDelivery.ListItems.Add(, , "")
                
                If Len(!CAsNumber) Then
                    Record.SubItems(1) = !CAsNumber
                End If
                If Len(!CASDRC) Then
                    Record.SubItems(2) = Format(!CASDRC, "dd/mm/yyyy")
                End If
                If Len(!CASDTC) Then
                    Record.SubItems(3) = Format(!CASDTC, "dd/mm/yyyy")
                End If
                If Len(!CASRecBy) Then
                    Record.SubItems(4) = Trim(!CASRecBy)
                End If
                If Len(!CASRecStatus) Then
                    Record.SubItems(5) = !CASRecStatus
                End If
                Current = Current + 1
                lblCurrent = Current
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
    CmdToAC.Enabled = True
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
End Sub

Private Sub WksDeliveryListing2(Optional strAppend As String)
On Error Resume Next
Dim rst2 As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sqlCount As String
Dim ClaimAdmin As Boolean
Dim Total As String
Dim Current As String
Dim TotalAmt As Variant
    Current = 0
    sql = " Select CASNumber, CASDRC, CASDTC, CASRecBy, CASRecStatus " & _
          " From CASOthers where 0 = 0 "
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    rst2.Open sql, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdClear.Enabled = True
        CmdToAC.Enabled = True
        cmdPrint.Enabled = True
        CmdExit.Enabled = True
    Else
         
    Do
        Do Until rst2.EOF
                
            With rst2
                Set Record = lstClaimDelivery.ListItems.Add(, , "")
                
                If Len(!CAsNumber) Then
                    Record.SubItems(1) = !CAsNumber
                End If
                If Len(!CASDRC) Then
                    Record.SubItems(2) = Format(!CASDRC, "dd/mm/yyyy")
                End If
                If Len(!CASDTC) Then
                    Record.SubItems(3) = Format(!CASDTC, "dd/mm/yyyy")
                End If
                If Len(!CASRecBy) Then
                    Record.SubItems(4) = Trim(!CASRecBy)
                End If
                If Len(!CASRecStatus) Then
                    Record.SubItems(5) = !CASRecStatus
                End If
                                
                Current = Current + 1
                lblCurrent = Current
            
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
    CmdToAC.Enabled = True
    cmdPrint.Enabled = True
    CmdExit.Enabled = True
End Sub
'**********************End BordxDelivery Listing *******************************

'**********************Start Update all Record functions*******************************
Private Sub UpdateSentCLM(ClmNo As String)
On Error Resume Next
Dim strsql As String
Dim StrSql2 As String
Dim ToAC As New ADODB.Recordset
Dim CheckInv As New ADODB.Recordset
Dim TokenDate As String
Dim InvoiceNo As String
Dim InvoiceDate As String
Dim InvoiceAmt As Double
    TokenDate = ""
    InvoiceNo = ""
    InvoiceDate = ""
    InvoiceAmt = 0
    StrSql2 = " Select CASNumber, INVWSversion, InvNo, InvDate, INVAmount " & _
              " From ClaimInvoice Where CASNumber = '" & Trim(ClmNo) & "' And INVWSversion = '1'"
    CheckInv.Open StrSql2, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
    
    If CheckInv.EOF = False Then
        With CheckInv
            InvoiceNo = Trim(!INVNo)
            InvoiceDate = Format(!InvDate, "YYYY/MM/DD")
            InvoiceAmt = Format(!INVAmount, "#,##0.00")
        End With
    End If
    CheckInv.Close
    Set CheckInv = Nothing
    TokenDate = Format(Date, "YYYY/MM/DD")
    strsql = " update CASOthers set " & _
             " CASDRC = '" & TokenDate & "', " & _
             " CASRecStatus = '1', " & _
             " CASRecBy = '" & Logon & "', " & _
             " CASINVNo = '" & InvoiceNo & "', " & _
             " InvDate = '" & InvoiceDate & "', " & _
             " CASInvAmt = " & InvoiceAmt & " " & _
             " Where CASNumber = '" & Trim(ClmNo) & "'"
    medixmasterconn.Execute strsql
End Sub
'**********************End Update all Record functions*******************************

'**********************Start Sorted ListView *******************************
Private Sub lstClaimDelivery_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstClaimDelivery, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstClaimDelivery, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstClaimDelivery, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    Case 4
        SortListView lstClaimDelivery, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
    Case 5
        SortListView lstClaimDelivery, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstClaimDelivery, ColumnHeader.Index, ldtString, m_blnDirection(6)
    End Select
End Sub
'**********************End Sorted ListView *******************************

'**********************Start Combo Listing*******************************
Private Sub ComboListing()
    With CboTokenStatus
        .AddItem "0 - Both"
        .AddItem "1 - Passed"
        .AddItem "2 - Not Yet Passed"
        .Text = "0 - Both"
    End With
End Sub
'**********************End Combo Listing*******************************

Private Sub txtDateSent1_LostFocus()
    If IsDate(txtDateSent1) Then
    Else
        If txtDateSent1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateSent1.SetFocus
        End If
    End If
End Sub

Private Sub txtDateSent2_LostFocus()
    If IsDate(txtDateSent2) Then
    Else
        If txtDateSent2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateSent2.SetFocus
        End If
    End If
End Sub

Private Sub SetCheckAllItems(bState As Boolean)
On Error Resume Next
Dim LV As LVITEM
Dim lvCount As Long
Dim lvIndex As Long
Dim lvState As Long
Dim r As Long
    lvState = IIf(bState, &H2000, &H1000)
    lvCount = lstClaimDelivery.ListItems.Count - 1
    Do
        With LV
        .mask = LVIF_STATE
        .state = lvState
        .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstClaimDelivery.hwnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
End Sub

Public Sub SetCheck(hwnd As Long, lItemIndex As Long, bState As Boolean)
Dim LV As LVITEM
Dim r As Long
    With LV
    .mask = LVIF_STATE
    .state = IIf(bState, &H2000, &H1000)
    .stateMask = LVIS_STATEIMAGEMASK
    End With
    r = SendMessageAny(hwnd, LVM_SETITEMSTATE, lItemIndex, LV)
End Sub

Private Function CaseNoCheck()
    CASNo = Trim(lstClaimDelivery.ListItems(listloop1).SubItems(1))
    If Trim(ChkStr(TxtWksNo1)) = ChkStr(CASNo) Then
        Check = True
    End If
End Function


