VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBordxRevise 
   Caption         =   "Bordereaux Revise"
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
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   6720
      Width           =   11655
      Begin VB.TextBox lblTotalAmt 
         Height          =   285
         Left            =   9120
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Approved Amount"
         Height          =   255
         Left            =   7680
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Items"
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   7320
      Width           =   11655
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9720
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10680
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdReject 
         Caption         =   "&Remove "
         Height          =   375
         Left            =   7680
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   8640
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   11655
      Begin VB.CommandButton CmdGo 
         Caption         =   "&Go"
         Height          =   280
         Left            =   2760
         TabIndex        =   1
         Top             =   240
         Width           =   765
      End
      Begin VB.TextBox TxtBordx 
         Height          =   285
         Left            =   960
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblSearhCriteria 
         Height          =   375
         Left            =   9480
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblBordxDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8040
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   7080
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Bordx Printed Date"
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblBordxPrintedDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5280
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Bordx No "
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lstWS 
      Height          =   5775
      Left            =   120
      TabIndex        =   9
      Top             =   960
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
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Claim No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Claim Reg Date"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Mem No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Patient"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "DOA"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "DOD"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Approved Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Case Status"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Bordereaux Revise"
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
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmBordxRevise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bExit As Boolean
'Public SearchCriteria As String
Dim WSNoArray() As String
Dim ItemChecked As Boolean
Private m_blnDirection(1 To 8) As Boolean
Dim BordxAmt As Double
Dim ScanCode As String
Dim TotalApproved As Double


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

Private Sub CmdClear_Click()
    TxtBordx.Text = ""
    lblBordxPrintedDate.Caption = ""
    lblBordxDate.Caption = ""
    lblTotal = ""
    lblTotalAmt = ""
    lstWS.ListItems.Clear
    ReDim WSNoArray(0)
    CmdReject.Enabled = True
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
            
    If lstWS.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstWS)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    
    Screen.MousePointer = vbDefault
    
    'If lstWS.listitems.count <> 0 Then
        'RptBordxRevise.show
    'Else
        'MsgBox "Report cannot be printed without any record.", vbInformation
    'End If
    
End Sub

Private Sub Form_Load()
    ReDim WSNoArray(0)
    bExit = False
End Sub

Private Sub CmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Dim RecordExit
'    If bExit = False Then
'        RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'        If RecordExit = vbYes Then
            Unload Me
'        Else
'            Cancel = True
'        End If
'    End If
End Sub


'***************************

Private Sub CmdGo_Click()
    ScanCode = TxtBordx
    If Mid(ScanCode, 1, 2) = "LO" Then
        TxtBordx = ""
        Call SearchForm(ScanCode)
    Else
        Screen.MousePointer = vbHourglass
        ReDim WSNoArray(0)
        If Trim(TxtBordx) = "" Then
            MsgBox "Please select a Bordx No ", vbOKOnly + vbCritical, DialogBoxTitle
            TxtBordx.SetFocus
        Else
            CmdGo.Enabled = False
            CmdReject.Enabled = False
           ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                Call SearchRecord
           ' medixmasterconnWS.Close
           ' Set medixmasterconnWS = Nothing
            CmdGo.Enabled = True
            CmdReject.Enabled = True
        End If
        Screen.MousePointer = Default
    End If
End Sub


Private Function SearchRecord()
Dim StrAppend As String
    
    StrAppend = ""
    If Len(Trim(TxtBordx)) Then
        StrAppend = StrAppend + " And BordxRefNo = '" & Trim(TxtBordx) & "'"
    End If
    
    lblSearhCriteria = StrAppend
    Call BordxReviseListing(StrAppend)

End Function
'****************** End Search Record *****************************

'****************** Start BordxReviseListing Listing *****************************

Private Sub BordxReviseListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim ClaimAdmin As Boolean
Dim total As String
Dim TotalAmt As Double
Dim BordxStatus1 As String
Dim BordxStatus2 As String
    
    total = 0
    BordxStatus1 = ""
    BordxStatus2 = ""
    
    lstWS.ListItems.Clear
    
    '===Run Sp=========================================================
    Sql = "Exec BordxRevise '" & Trim(TxtBordx) & "'"
    
    Set rst2 = New ADODB.Recordset
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    If Not rst2.EOF Then
    '=============================================================
    
    'Sql = "Select * From VIEW_BordxRevise where 0=0 "
    
    'If Len(Trim(StrAppend)) Then
        'Sql = Sql + StrAppend
    'End If
       
    'rst2.Open Sql, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
   
    'If rst2.EOF = True Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = False
    'Else
        CmdPrint.Enabled = True
        
        If Len(rst2!BordxPrintedDate) Then
            lblBordxPrintedDate = rst2!BordxPrintedDate
        End If
        lblBordxDate = rst2!BordxDate
        total = rst2.recordCount
        BordxAmt = rst2!BordxClaimAmt
        BordxStatus1 = rst2!BordxPayToStatus
        BordxStatus2 = rst2!BordxMNRBToStatus
        lblTotal = total
        Do Until rst2.EOF
            With rst2
                Set Record = lstWS.ListItems.Add(, , !CasNumber & " - " & !ClaimVersion)
                If Len(!ClaimOpendate) Then
                    Record.SubItems(1) = !ClaimOpendate
                End If
                If Len(!MBMNumber) Then
                   Record.SubItems(2) = !MBMNumber
                End If
                If Len(!PatientName) Then
                    Record.SubItems(3) = !PatientName
                End If
                If Len(!CASDateAdmitted) Then
                   Record.SubItems(4) = !CASDateAdmitted
                End If
                If Len(!CASDischargeDate) Then
                    Record.SubItems(5) = !CASDischargeDate
                End If
                
                If Len(!CLMTotalApproved) Then
                    Record.SubItems(6) = Format(!CLMTotalApproved, "#,##0.00")
                End If
                If Len(!CasStatus) Then
                    Record.SubItems(7) = !CasStatus
                End If
                If Len(rst2!CLMTotalApproved) Then
                    TotalAmt = TotalAmt + rst2!CLMTotalApproved
                End If
            End With
            rst2.MoveNext
        Loop
        lblTotalAmt = Format(TotalAmt, "#,##0.00;(#,#.00)")
        
        If (BordxStatus1 = "True" Or BordxStatus2 = "True") Then
            MsgBox "This Bordx has been sent out.", vbOKOnly + vbCritical, DialogBoxTitle
            CmdReject.Enabled = False
        Else
            CmdReject.Enabled = True
        End If
        
    Else
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = False
    End If
    rst2.Close
    Set rst2 = Nothing
End Sub


Private Sub lstWS_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'**********************Start Check ListView Record*******************************
Dim i, newbound  As Integer
Dim tempArray() As String
Dim Found As Boolean
Dim temp
Static Arraycnt As Integer
    If Item.Checked = True Then
            ReDim Preserve WSNoArray(Arraycnt)
            WSNoArray(Arraycnt) = Item.Text
            Arraycnt = Arraycnt + 1
            ItemChecked = True
    Else
        ' Remove from the array
        newbound = 0
    
        If UBound(WSNoArray) <> 0 Then
            For i = 0 To UBound(WSNoArray)
                If Trim(WSNoArray(i)) <> "" Then
                    If (Item.SubItems(2)) <> WSNoArray(i) Then
                        ReDim Preserve tempArray(newbound)
                        tempArray(newbound) = WSNoArray(i)
                        newbound = newbound + 1
                    End If
                End If
            Next
            Erase WSNoArray
            If UBound(tempArray) = 0 Then
                Arraycnt = 1
            Else
                Arraycnt = UBound(tempArray) + 1
            End If
            ReDim WSNoArray(UBound(tempArray))
            For i = 0 To UBound(tempArray)
                WSNoArray(i) = tempArray(i)
            Next
            Erase tempArray
        Else
            Erase WSNoArray
            ReDim WSNoArray(0)
            Arraycnt = 0
            ItemChecked = False
        End If
    End If
'**********************End Check ListView Record*******************************
End Sub

'**********************Start Sorted ListView *******************************
Private Sub lstWS_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    
    Case 1
        SortListView lstWS, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstWS, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
    Case 3
        SortListView lstWS, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstWS, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstWS, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
    Case 6
        SortListView lstWS, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lstWS, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstWS, ColumnHeader.Index, ldtString, m_blnDirection(8)
    End Select
End Sub
'**********************End Sorted ListView *******************************

Private Sub CmdReject_Click()
Dim ii As Integer
Dim Cnt As Integer
Dim Update
Dim recordFound As Boolean
Static Arraycnt As Integer
    
    CmdGo.Enabled = False
    CmdReject.Enabled = False
    If ItemChecked = True Then
       ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        Update = MsgBox("You are going to remove a worksheet from current Bordereaux. " & vbNewLine & _
                "Are you sure?", vbYesNo + vbExclamation)
        
        If Update = vbYes Then
            Cnt = 0
            recordFound = False
            If lstWS.ListItems.Count <> 0 Then
                For ii = 0 To UBound(WSNoArray)
                    If Trim(WSNoArray(ii)) <> "" Then
                       RemoveFromBordx WSNoArray(ii)
                       recordFound = True
                    End If
                Next ii
                If recordFound = True Then
                    MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
                End If
                Erase WSNoArray
                Arraycnt = 0
                ReDim WSNoArray(0)
                Call ClearForm(Me)
                lstWS.ListItems.Clear
                lblBordxPrintedDate = ""
                lblBordxDate = ""
            End If
        End If
        'medixmasterconnWS.Close
        'Set medixmasterconnWS = Nothing
    End If
    CmdGo.Enabled = True
    CmdReject.Enabled = True
End Sub

'**********************Start Update all Record functions*******************************

Private Sub RemoveFromBordx(tWSNo As String)
Dim StrSql As String
Dim strReduceBordx As String
Dim Claim As New ADODB.Recordset
Dim BordxArchive As New ADODB.Recordset
Dim strsqlArchive As String
Dim pos1 As Integer
Dim CasNumber, VersionNo As String
Dim strsqlBordx As String
Dim strsqlClmInv As String
Dim Bordx As New ADODB.Recordset
Dim ClmInv As New ADODB.Recordset

  pos1 = InStr(1, tWSNo, "-")
  CasNumber = Trim(Mid(tWSNo, 1, pos1 - 1))
  VersionNo = Trim(Mid(tWSNo, pos1 + 1, Len(tWSNo) - pos1))
  StrSql = " Select CASNUmber, CLaimVersion, ClaimBordxStatus, " & _
           " BordxRefNo, CLMTotalApproved, CLMBordxWSStatus, " & _
           " CLMBordxWSAscii From Claim where CASNumber ='" & Trim(CasNumber) & "' " & _
           " AND ClaimVersion = " & VersionNo
           
  Claim.Open StrSql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
  
  TotalApproved = Claim!CLMTotalApproved
  strReduceBordx = " update bordx set BordxNoClaims = BordxNoClaims - 1 " & _
                   " where BordxRefNo = '" & Claim!BordxRefNo & "'"
                    
  medixmasterconnWS.Execute strReduceBordx
  
  strsqlBordx = "Select BordxClaimAmt, BordxDate from Bordx where BordxRefNo = '" & Trim(TxtBordx) & "'"
  Bordx.Open strsqlBordx, medixmasterconnWS, adOpenKeyset, adLockOptimistic

  strsqlClmInv = " Select CASNumber, INVWSVersion, INVBordxRefNo, INVBordxed from ClaimInvoice where INVBordxRefNo = '" & Trim(TxtBordx) & "'" & _
                 " AND CASNumber = '" & Trim(CasNumber) & "' AND INVWSVersion = '" & Trim(VersionNo) & "'"
  ClmInv.Open strsqlClmInv, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    BordxArchive.Open "BordxWSArchive", medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    BordxArchive.AddNew
     With BordxArchive
        !CasNumber = Trim(CasNumber)
        !ClaimVersion = Trim(VersionNo)
        !BordxRefNo = Trim(TxtBordx)
        !BordxDate = Bordx!BordxDate
        !LastUpdateDate = Format(Date, "dd/mm/yyyy")
        !LastUpdateUser = Logon
        .Update
    End With
    BordxArchive.Close
    Set BordxArchive = Nothing
    
    If Bordx.recordCount > 0 Then
        Bordx!BordxClaimAmt = CDbl(Bordx!BordxClaimAmt) - TotalApproved
        Bordx.Update
        Bordx.Close
        Set Bordx = Nothing
    End If
    
    If ClmInv.recordCount > 0 Then
        Do Until ClmInv.EOF
            With ClmInv
                !INVBordxRefNo = ""
                !INVBordxed = "0"
            End With
            ClmInv.Update
            ClmInv.MoveNext
       Loop
       ClmInv.Close
       Set ClmInv = Nothing
    End If
    
    If Claim.recordCount > 0 Then
       With Claim
          !ClaimBordxStatus = "0"
          !BordxRefNo = ""
          If !CLMBordxWSASCII = "Y" Then
            !CLMBordxWSASCII = "N"
            !CLMBordxWSStatus = "D"
          End If
       End With
       Claim.Update
       Claim.Close
       Set Claim = Nothing
    End If
End Sub

Private Sub TxtBordx_GotFocus()
    CmdGo.Default = True
End Sub

Private Sub TxtBordx_LostFocus()
    TxtBordx = ChkStr(TxtBordx)
    CmdGo.Default = False
End Sub
