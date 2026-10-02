VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchCase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search For Member"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   315
      Left            =   5760
      TabIndex        =   14
      Top             =   5520
      Width           =   960
   End
   Begin MSComctlLib.ListView lstMBMInquiryList 
      Height          =   4200
      Left            =   120
      TabIndex        =   19
      Top             =   1080
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7408
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Member No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cov. ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Member Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IC No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Pol Number"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Relationship"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   960
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   8535
      Begin VB.ComboBox CboMbmType 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   540
         Width           =   1530
      End
      Begin VB.ComboBox cboPlanCode 
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
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   4770
      End
      Begin VB.ComboBox cboINSCode 
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
         Left            =   960
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   540
         Width           =   1890
      End
      Begin VB.ComboBox CboPayorCode 
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
         Left            =   960
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   4770
      End
      Begin VB.Label Label3 
         Caption         =   "Mem Type"
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label75 
         Caption         =   "Plan Code"
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
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label73 
         Caption         =   "Payor "
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
         Left            =   120
         TabIndex        =   22
         Top             =   210
         Width           =   705
      End
      Begin VB.Label Label74 
         Caption         =   "Ins Type"
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
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   5280
      Width           =   8535
      Begin VB.TextBox lblTotal 
         Height          =   285
         Left            =   1440
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   27
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
         Height          =   315
         Left            =   6600
         TabIndex        =   15
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
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
         Left            =   3720
         TabIndex        =   12
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   315
         Left            =   4680
         TabIndex        =   13
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdExit 
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
         Height          =   315
         Left            =   7560
         TabIndex        =   16
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Records"
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "of"
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame11 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8475
      Begin VB.OptionButton OptPOLSubNo 
         Caption         =   "Policy Sub Num"
         Height          =   240
         Left            =   6660
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton OptEmployeeNo 
         Caption         =   "Employee Num"
         Height          =   240
         Left            =   6660
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton txtMBMPolicyNo1 
         Caption         =   "Policy Num"
         Height          =   240
         Left            =   5520
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optMBMName1 
         Caption         =   "Mem Name"
         Height          =   240
         Left            =   4320
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton txtIcBcPc1 
         Caption         =   "IC Num"
         Height          =   240
         Left            =   4320
         TabIndex        =   5
         Top             =   600
         Width           =   1065
      End
      Begin VB.OptionButton txtMBMNo1 
         Caption         =   "Mem Num"
         Height          =   240
         Left            =   5520
         TabIndex        =   6
         Top             =   600
         Width           =   1500
      End
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         MaxLength       =   60
         TabIndex        =   1
         Top             =   300
         Width           =   3465
      End
      Begin VB.Label Label16 
         Caption         =   "Search "
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
         Left            =   120
         TabIndex        =   17
         Top             =   285
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmSearchCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Public Source As Integer
Public intCaseAdmission  As Integer
Public intCaseAdjustment As Integer
Private m_blnDirection(1 To 15) As Boolean

Private Sub cmdPrint_Click()
    If lstMBMInquiryList.ListItems.Count = 0 Then
        MsgBox "Empty record found.", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    Else
        Screen.MousePointer = vbHourglass
        Call ExportListViewtoExcel(lstMBMInquiryList)
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

Private Sub CmdClear_Click()
txtSearch.Text = ""
optMBMName1.Value = True
CboPayorCode.Text = ""
cboPlanCode.Text = ""
cboINSCode.Text = ""
CboMbmType.ListIndex = 0
lstMBMInquiryList.ListItems.Clear
lblCurrent = ""
lblTotal = ""
End Sub

Private Sub Form_Load()
    intExit = 0
    intCaseAdmission = 1
    intCaseAdjustment = 2
    Call CenterForm(Me)
    'Call LoadInsuredType
    'Call LoadPayor
    'Call ComboListing
    'Call LoadPlanType
End Sub

'Private Sub ComboListing()
    'CboMbmType.AddItem "Both"
    'CboMbmType.AddItem "P" & " - " & "Principal"
    'CboMbmType.AddItem "S" & " - " & "Supplementary"
    'CboMbmType.Text = "Both"

    'cboStatus.AddItem "A" & " - " & "Active"
    'cboStatus.AddItem "C" & " - " & "Cancel"
    'cboStatus.AddItem "D" & " - " & "Delete"
    'cboStatus.AddItem "S" & " - " & "Suspend"
'End Sub

'Private Sub LoadPayor()
    'Dim PAY As New ADODB.Recordset
    'sqlPAY = "select PAYCode , PAYCompanyName , HLTCode from PAY "
    
    'CboPayorCode.Clear
    'PAY.Open sqlPAY, medixmasterconn, adOpenKeyset, adLockOptimistic
    'Do Until PAY.EOF
        'CboPayorCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName) & "*" & PAY!HLTCode
        'PAY.MoveNext
    'Loop
    'PAY.Close
    'Set PAY = Nothing
'End Sub

'Private Sub LoadPlanType(Optional HLTCode As String)

    'Dim strsql As String
    'Dim PLN As New ADODB.Recordset
    'Dim planstring As String
    
    
    'planstring = ""   ' Initialize string.
    'cboPlanCode.Clear
   
    'strsql = "Select PLNCode ,PLNDescription from PLN where 0=0  "

    'If InStr(1, Trim(CboPayorCode), "*") > 0 And InStr(1, Trim(CboPayorCode), "-") > 0 Then
        'If Trim(Mid(ChkStr(CboPayorCode), InStr(1, CboPayorCode, "*") + 1, Len(Trim(CboPayorCode)))) = "S" Then
            'strsql = strsql & " and Hltcode = 'S'"
        'Else
            'strsql = strsql & " and Hltcode = 'N' and PAYcode = '" & Trim(Mid(ChkStr(CboPayorCode), 1, InStr(1, CboPayorCode, "-") - 1)) & "'"
        'End If
    'End If
    
    'PLN.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    'If PLN.recordCount <> 0 Then
        'Do Until PLN.EOF
            'planstring = ChkStr(PLN!PLNCode)
            'cboPlanCode.AddItem planstring & " - " & PLN!PLNDescription
            'PLN.MoveNext
        'Loop
    'End If
    'PLN.Close
    'Set PLN = Nothing
'End Sub

'Private Sub LoadInsuredType()
'On Error Resume Next
    'Dim strsql As String
    'Dim HLTCode As String
    'Dim INS As New ADODB.Recordset
    
    'cboINSCode.Clear
    'strsql = "Select INSCode, INSDescription from INS "
    'If HLTCode = "N" Then
        'strsql = strsql & "Where INSCode = 'H' or INSCode = 'G'"
    'End If
    'INS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'Do Until INS.EOF
        'cboINSCode.AddItem ChkStr(INS!INSCode) & " - " & ChkStr(INS!INSDescription)
        'INS.MoveNext
    'Loop
    'INS.Close
    'Set INS = Nothing
'End Sub


' #################################
'Private Sub CboPayorCode_LostFocus()
    'Call LoadPlanType
'End Sub


Private Sub MBMInquiryList(Optional str As String)
    On Error GoTo ErrListing
    counter = 0
    Dim MBM As New ADODB.Recordset
    Dim sqlstr, strCriteria As String
    Dim sqlstrMBM, sqlstrSupp As String
    Dim MEMList As ListItem
    Dim TotalPrincipalCount As Integer
    Dim MBMSupplementary As New ADODB.Recordset
    Dim CASMBMNumber As String
    Dim CASNumber As String
    
    Screen.MousePointer = vbHourglass
    CmdExit.Enabled = False
    strCriteria = ""
    lstMBMInquiryList.ListItems.Clear
     '--------------------------------------------------------------------------------------------------------------------------------------------
      If Len(Trim(CboPayorCode)) Then
         If InStr(1, CboPayorCode, "-") <> 0 Then
           strCriteria = strCriteria & " AND PAYCode = '" & Trim(Mid(CboPayorCode, 1, InStr(1, CboPayorCode, "-") - 1)) & "' "
         Else
             strCriteria = strCriteria & " AND PAYCode = '" & Trim(CboPayorCode) & "' "
         End If
      End If
      
      
      If Len(Trim(cboPlanCode)) Then
          strCriteria = strCriteria & " And PLNCode ='" & Trim(Mid(cboPlanCode, 1, InStr(1, cboPlanCode, "-") - 1)) & "' "
      End If
      
    '--------------------------------------------------------------------------------------------------------------------------------------------
      If Len(Trim(cboINSCode)) Then
           strCriteria = strCriteria & " AND INSCode = '" & Trim(Mid(cboINSCode, 1, IIf(InStr(1, cboINSCode, "-") = 0, 1, InStr(1, cboINSCode, "-") - 1))) & "' "
      End If

     
     '--------------------------------------------------------------------------------------------------------------------------------------------
'      If Len(Trim(cboStatus)) Then
 '         strCriteria = strCriteria & "AND MBMStatus = '" & Trim(Mid(cboStatus, 1, InStr(1, cboStatus, "-") - 1)) & "' "
  '    End If
     
      If Len(Trim(txtSearch)) Then
         If TxtMBMNo1.Value = True Then
             strCriteria = strCriteria & " AND MBMNumber = '" & RemStr(txtSearch) & "'"
         ElseIf txtMBMPolicyNo1.Value = True Then
             strCriteria = strCriteria & " AND MBMPolicyNo = '" & RemStr(txtSearch) & "'"
         ElseIf optMBMName1.Value = True Then
             'strCriteria = strCriteria & " AND MBMName LIKE '%" & RemStr(txtSearch) & "%' "
             strCriteria = strCriteria & " AND MBMName = '" & RemStr(txtSearch) & "' "
         ElseIf txtIcBcPc1.Value = True Then
             strCriteria = strCriteria & " AND MBMIcBcPp =  '" & RemStr(txtSearch) & "'"
         ElseIf OptEmployeeNo.Value = True Then
             strCriteria = strCriteria & " AND MBMEmployeeNo =  '" & RemStr(txtSearch) & "' "
         ElseIf OptPOLSubNo.Value = True Then
             strCriteria = strCriteria & " AND MBMPOLSubNo1 like  '%" & RemStr(txtSearch) & "%' "
         End If
      End If
     
    sqlstr = "Select MBMNumber, MBMPolicyNo, MBMName, MBMIcBcPp, " & _
         " MBMMemberType, MBMStatus, " & _
         " RElCode, MBMCCoverID "
    
    ' Form MBM table
    sqlstrMBM = sqlstr
    ' Form MBMCoverpersons  table
    sqlstrSupp = sqlstr
     
    If Mid(ChkStr(CboMbmType), 1, 1) = "P" Then
       ' sqlstr = sqlstr & " , MBMAmendEdtType, MBMAEndtBordxDate as MBMAEndtBrodxDate, MBMAEndtBatchNo "
        
        sqlstr = sqlstr & " From View_MBMSearch Where MBMStatus <> 'D' "
        
        '======== testing =================================================
        'sqlstr = sqlstr & " From View_MBM_Search  Where MBMStatus <> 'D' "
        '======== testing =================================================
        
    ElseIf Mid(ChkStr(CboMbmType), 1, 1) = "S" Then
        sqlstr = sqlstr & " From View_MBM_Supplementary where MBMStatus <> 'D' "
    Else
        'sqlstrMBM = sqlstrMBM & " , MBMAmendEdtType, MBMAEndtBordxDate as MBMAEndtBrodxDate , MBMAEndtBatchNo "
        
        sqlstrMBM = sqlstrMBM & " From View_MBMSearch Where MBMStatus <> 'D' "
        
        '======== testing =================================================
        'sqlstrMBM = sqlstrMBM & " From View_MBM_Search Where MBMStatus <> 'D' "
        '======== testing =================================================
        
        sqlstrSupp = sqlstrSupp & " ,RElCode, MBMCCoverID From View_MBM_Supplementary Where MBMStatus <> 'D' "
    End If
    
      If Len(Trim(tempsqlstrBrodx)) And Len(Trim(tempsqlstrEndt)) Then
        strCriteria = strCriteria & tempsqlstrBrodx & " OR " & Mid(strCriteria, 5, Len(strCriteria)) & tempsqlstrEndt
        If InStr(1, CboMbmType, "-") <> 0 Then
            sqlstr = sqlstr & strCriteria
        Else
            ' select both principal and supp.
            sqlstrMBM = sqlstrMBM & strCriteria
            sqlstrSupp = sqlstrSupp & strCriteria
        End If
      ElseIf Len(Trim(tempsqlstrBrodx)) = 0 And Len(Trim(tempsqlstrEndt)) Then
        If InStr(1, CboMbmType, "-") <> 0 Then
            ' BOTH C AND S
            sqlstr = sqlstr & strCriteria & tempsqlstrEndt
        Else
            sqlstrMBM = sqlstrMBM & strCriteria & tempsqlstrEndt
            sqlstrSupp = sqlstrSupp & strCriteria & tempsqlstrEndt
        End If
      ElseIf Len(Trim(tempsqlstrBrodx)) And Len(Trim(tempsqlstrEndt)) = 0 Then
        If InStr(1, CboMbmType, "-") <> 0 Then
            sqlstr = sqlstr & strCriteria & tempsqlstrBrodx
        Else
            sqlstrMBM = sqlstrMBM & strCriteria & tempsqlstrBrodx
            sqlstrSupp = sqlstrSupp & strCriteria & tempsqlstrBrodx
        End If
      Else
        If InStr(1, CboMbmType, "-") <> 0 Then
            sqlstr = sqlstr & strCriteria
        Else
            sqlstrMBM = sqlstrMBM & strCriteria
            sqlstrSupp = sqlstrSupp & strCriteria
        End If
      End If
      
     ' END Query ------------------------------------------
        MemExist = False
        If InStr(1, CboMbmType, "-") <> 0 Then
           MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
           If MBM.recordCount <> 0 Then
              MemExist = True
           End If
        Else
           MBM.Open sqlstrMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
           MBMSupplementary.Open sqlstrSupp, medixmasterconn, adOpenKeyset, adLockOptimistic
            If MBM.recordCount <> 0 Or MBMSupplementary.recordCount <> 0 Then
                MemExist = True
            Else
                MemExist = False
            End If
        End If


        If MemExist = True Then  'record Found
         lblTotal = MBM.recordCount
         If InStr(1, CboMbmType, "-") = 0 Then
            lblTotal = CLng(lblTotal) + CLng(MBMSupplementary.recordCount)
         End If
         Do
             Do Until MBM.EOF
                 Set MEMList = lstMBMInquiryList.ListItems.Add(, , MBM!MBMNumber)
                 With MEMList
                     If MBM!MBMStatus = "S" Then
                         strName = Trim(MBM!MBMName) & " (Suspended)"
                     ElseIf MBM!MBMStatus = "C" Then
                         strName = Trim(MBM!MBMName) & " (Canceled)"
                     ElseIf MBM!MBMStatus = "D" Then
                         strName = Trim(MBM!MBMName) & " (Deleted)"
                     Else
                        strName = Trim(MBM!MBMName)
                     End If
                     If Trim(Mid(CboMbmType, 1, 1)) = "P" Then
                         .SubItems(1) = "00"
                     ElseIf Trim(Mid(CboMbmType, 1, 1)) = "Q" Or Trim(Mid(CboMbmType, 1, 1)) = "C" Then
                         .SubItems(1) = MBM!MBMCCoverID
                     Else
                         '.SubItems(1) = IIf(Len(MBM!MBMCCoverID) = "0", "0", MBM!MBMCCoverID) & 0
                         .SubItems(1) = IIf(Len(MBM!MBMCCoverID) <> "00", MBM!MBMCCoverID, "00")
                     End If
                     .SubItems(2) = strName
                     If Len(Trim(MBM!MBMIcBcPp)) Then
                        .SubItems(3) = Trim(MBM!MBMIcBcPp)
                    End If
                     If Trim(MBM!MBMPolicyNo) <> "" Then
                         .SubItems(4) = Trim(MBM!MBMPolicyNo)
                     End If
                     
                     '.SubItems(5) = Trim(MBM!RELCODE)
                     .SubItems(5) = "P"
                     
                     
                 '    If Len(Trim(MBM!MBMLastEndorseStatus)) Then
                  '       .SubItems(6) = Trim(MBM!MBMLastEndorseStatus)
                   '  End If
               
                 End With
                 If Trim(MBM!MBMMemberType) <> "" Then
                        If ChkStr(MBM!MBMMemberType) = "P" Then
                           TotalPrincipalCount = TotalPrincipalCount + 1
                           lblCurrent = TotalPrincipalCount
                        ElseIf ChkStr(MBM!MBMMemberType) = "Q" Or ChkStr(MBM!MBMMemberType) = "C" Then
                           totalSuppCount = totalSuppCount + 1
                           lblCurrent = totalSuppCount
                        Else
                          lblCurrent = lblCurrent + 1
                        End If
                End If
                 MBM.MoveNext
                 DoEvents
                 
                 If intExit = 1 Then
                     Check = False
                     Exit Do
                 End If
               Loop
        Loop Until Check = False
        Do
               If InStr(1, CboMbmType, "-") = 0 Then
                    Do Until MBMSupplementary.EOF
                        
                        Set MEMList = lstMBMInquiryList.ListItems.Add(, , MBMSupplementary!MBMNumber)
                        With MEMList
                            If MBMSupplementary!MBMStatus = "S" Then
                                strName = Trim(MBMSupplementary!MBMName) & " (Suspended)"
                            ElseIf MBMSupplementary!MBMStatus = "C" Then
                                strName = Trim(MBMSupplementary!MBMName) & " (Canceled)"
                            ElseIf MBMSupplementary!MBMStatus = "D" Then
                                strName = Trim(MBMSupplementary!MBMName) & " (Deleted)"
                            Else
                               strName = Trim(MBMSupplementary!MBMName)
                            End If
                            If Mid(ChkStr(CboMbmType), 1, 1) = "P" Then
                                .SubItems(1) = "00"
                            Else
                                .SubItems(1) = MBMSupplementary!MBMCCoverID
                            End If
                            .SubItems(2) = strName
                            If Len(Trim(MBMSupplementary!MBMIcBcPp)) Then
                               .SubItems(3) = Trim(MBMSupplementary!MBMIcBcPp)
                           End If
                            If Trim(MBMSupplementary!MBMPolicyNo) <> "" Then
                                .SubItems(4) = Trim(MBMSupplementary!MBMPolicyNo)
                            End If
                            
                            .SubItems(5) = Trim(MBMSupplementary!RELCODE)
                    '
                     '       If Len(Trim(MBMSupplementary!MBMLastEndorseStatus)) Then
                      '          .SubItems(6) = Trim(MBMSupplementary!MBMLastEndorseStatus)
                       '     End If
                           End With
                           totalSuppCount = totalSuppCount + 1
                           lblCurrent = CInt(TotalPrincipalCount) + CInt(totalSuppCount)
                           MBMSupplementary.MoveNext
                           DoEvents
                            
                           If intExit = 1 Then
                               Check = False
                               Exit Do
                           End If
                    
                    
                    Loop
                End If
         Loop Until Check = False
     Else
        MsgBox "No record found!", vbCritical
     End If
    CmdExit.Enabled = True
    Screen.MousePointer = vbDefault
    intExit = 0
    MBM.Close
    Set MBM = Nothing
    CmdClear.Enabled = True
    Exit Sub
ErrListing:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
    CmdExit.Enabled = True
End Sub

'**********************Start Sorted ListView *******************************
Private Sub lstMBMInquiryList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtDatetime, m_blnDirection(6)
    Case 7
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtDatetime, m_blnDirection(7)
    Case 8
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
    Case 11
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
    Case 12
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    Case 13
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
    Case 14
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(15)
    End Select
End Sub
'**********************End Sorted ListView *******************************


Private Sub lstMBMInquiryList_DblClick()
'    Dim MBMCov As New ADODB.Recordset
    Dim sqlstr As String
    Dim ctl As Form
    Dim sqlstr1 As String
    Dim sqlstr2 As String
    Dim sqlstr3 As String
    Dim sqlstr4 As String
    Dim sqlstr5 As String
    Dim sqlstr6 As String
    Dim sqlstr7 As String
    Dim sqlstr8 As String
    Dim ICD As String
    Dim HosCode As String
    Dim HOSName As String
    Dim ICDCode As String
    Dim ICDDesc As String
 '   Dim HospitalMaster As ListItem
  '  Dim InitialDiag As ListItem
   ' Dim FinalDiag As ListItem
    'Dim HOS As New ADODB.Recordset
'    D'im CAS As New ADODB.Recordset
 '   Dim CASHOS As New ADODB.Recordset
  '  Dim diafirst As New ADODB.Recordset
   ' Dim DIAFinal As New ADODB.Recordset
    'Dim AUT As New ADODB.Recordset
'    Dim rst2 As New ADODB.Recordset
 '   Dim ICDCod As New ADODB.Recordset
    
If lstMBMInquiryList.ListItems.Count <> 0 Then
    
    If Len(Trim(lstMBMInquiryList.SelectedItem.SubItems(1))) Then
        If Source = intCaseAdmission Then
           Set ctl = frmCaseAdmission
           'Call ctl.LoadPlanType(lstMBMInquiryList.SelectedItem.SubItems(5))
           ctl.txtMBMNumber = Trim(lstMBMInquiryList.SelectedItem.Text)
           ctl.cmdShow_Click
        Else
           Set ctl = frmCaseAdjustment2016
           ctl.txtMBMNumber = Trim(lstMBMInquiryList.SelectedItem.Text)
           ctl.cmdShow_Click
        End If
 '       sqlstr = "select MBMCNumber, MBMCCoverID, MBMCName, " & _
  '          " MBMCRELCode,  MBMCStatus , MBMCAnnualLimit " & _
   '         " from  MBMCoveredPersons where MBMCNumber = '" & Trim(lstMBMInquiryList.SelectedItem.Text) & "'"
    '    MBMCov.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
     '   With ctl
      '
       '         .txtMBMName = lstMBMInquiryList.SelectedItem.SubItems(1)
        '        .txtMBMPolNo = lstMBMInquiryList.SelectedItem.SubItems(2)
         '       .txtMBMIcBcPP = lstMBMInquiryList.SelectedItem.SubItems(3)
          '      .txtDOB.Mask = ""
           '     .txtDOB = lstMBMInquiryList.SelectedItem.SubItems(6)
            '    If Source = intCaseAdmission Then
'                    .txtINSCode = lstMBMInquiryList.SelectedItem.SubItems(4)
 '               End If
  '
   '             .txtPLNCode = lstMBMInquiryList.SelectedItem.SubItems(7)
          '      .txtAdd1 = lstMBMInquiryList.SelectedItem.SubItems(8)
    '            .txtAdd2 = lstMBMInquiryList.SelectedItem.SubItems(9)
     '           .txtAdd3 = lstMBMInquiryList.SelectedItem.SubItems(10)
      '          .txtPostCode = lstMBMInquiryList.SelectedItem.SubItems(11)
       '         .txtCity = lstMBMInquiryList.SelectedItem.SubItems(12)
        '        .txtState = lstMBMInquiryList.SelectedItem.SubItems(13)
         '       .txtTakeOver = lstMBMInquiryList.SelectedItem.SubItems(14)
'                .txtPayCode = lstMBMInquiryList.SelectedItem.SubItems(15)
 '               If Source = intCaseAdmission Then
  '                  .txtAgeCode = lstMBMInquiryList.SelectedItem.SubItems(16)
   '                 '---------Add by Jason Loh -----------------------------------------
    '                .txtCancellationDate = lstMBMInquiryList.SelectedItem.SubItems(21)
     '               .txtLastAdjustment = lstMBMInquiryList.SelectedItem.SubItems(22)
      '              '-------------------------------------------------------------------
       '         End If
            '    .txtMBMEffDate.Mask = ""
        '        .txtMBMEffDate = lstMBMInquiryList.SelectedItem.SubItems(17)
         '       .txtMBMExpiryDate.Mask = ""
          '      .txtMBMExpiryDate = lstMBMInquiryList.SelectedItem.SubItems(18)
           '     '---------Add by Jason Loh ---------------------------------------------
             '   If Source = intCaseAdmission Then
              '      If DateDiff("ww", .txtMBMExpiryDate, DateAdd("d", 14, Date)) >= 0 Then
               '         MsgBox "Members Expired (+14 days grace period..!) ", vbOKOnly + vbCritical, DialogBoxTitle
                '    End If
'                End If
                '-----------------------------------------------------------------------
 '               .txtAnnualLimit = GetAnnualLimit(lstMBMInquiryList.SelectedItem.SubItems(4), .txtPLNCode, "", .txtMBMEffDate)
  '              .txtAvailableAnnualLimit = lstMBMInquiryList.SelectedItem.SubItems(19)
   '             .txtPolStatus = lstMBMInquiryList.SelectedItem.SubItems(20)
    '            '---------Add by Jason Loh ---------------------------------------------
     '           If Source = intCaseAdmission Then
      '              If .txtPolStatus = "C" Then
       '                 MsgBox "Members is already CANCELLED..! ", vbOKOnly + vbCritical, DialogBoxTitle
            '        ElseIf .txtPolStatus = "S" Then
        '''                MsgBox "Members is already SUSPENDED..! ", vbOKOnly + vbCritical, DialogBoxTitle
           '         ElseIf .txtPolStatus = "R" Then
             '           MsgBox "Members is already REJECTED..! ", vbOKOnly + vbCritical, DialogBoxTitle
              '      ElseIf .txtPolStatus = "D" Then
               '         MsgBox "Members is already DELETED..! ", vbOKOnly + vbCritical, DialogBoxTitle
                '    End If
'                End If
 '               '-----------------------------------------------------------------------
  '          End With
   '
    '         If Source = intCaseAdmission Then
     '           ctl.cboPatientList.clear
      '''          ctl.cboPatientList.AddItem "Self -" & ctl.txtMBMName
             '   If MBMCov.recordCount Then
            '        Do While Not MBMCov.EOF
         '                ctl.cboPatientList.AddItem MBMCov!MBMCCoverID & "- " & MBMCov!MBMCName
          '               MBMCov.MoveNext
           '         Loop
              '  End If
            
'             ElseIf Source = intCaseAdjustment Then
   '             Set ctl = frmCaseAdjustment
 '               frmCaseAdjustment.lstHosList.listitems.clear
  '                  sqlstr1 = "select * from  CAS where MBMNumber = '" & Trim(lstMBMInquiryList.SelectedItem.Text) & "'"
    '                CAS.Open sqlstr1, medixmasterconn, adOpenKeyset, adLockOptimistic
                                        
     '                   If CAS.recordCount Then
      '                      ctl.txtCaseNo = CAS!CASNumber
       '                     ctl.txtPatientCovID = CAS!MBMPatientCovID
        '''                    ctl.txtRegDate.Mask = ""
           '                 ctl.txtRegDate = CAS!CaseRegDate
                '            ctl.txtDateAdmitted.Mask = ""
            '                ctl.txtDateAdmitted = CAS!CASDateAdmitted
                            'ctl.txtCASRemark = CAS!CASRemark
             '               'ctl.txtCASAdmissionReason = CAS!CASAdmissionReason
              '              ctl.txtSecondReserveAmount = CAS!CASReserveAmount
               '             'ctl.txtCASRepudiation = CAS!CASRepReason
                            'ctl.txtSecondReserveAmount = CAS!CASSecondRAmount
                 '           ctl.txtDateOfLoss1.Mask = ""
                  '          ctl.txtDateOfLoss1 = CAS!CaseDateOfLoss
                    'rst2.Open "HOS", medixmasterconn, adOpenKeyset, adLockOptimistic
                    '        Do Until rst2.EOF
                    '        Set HospitalMaster = ctl.lstHosList.ListItems.Add(, , rst2!HOSCode)
                    '            'ctl.lstHosList.SubItems(1) = rst2!HosName
                    '            rst2.MoveNext
                    '        Loop

                    '    End If
                    
'                    sqlstr2 = "select * from CAS_HOS where CaseCode = '" & Trim(frmCaseAdjustment.txtCaseNo) & "'"
 '                   CASHOS.Open sqlstr2, medixmasterconn, adOpenKeyset, adLockOptimistic
  '
   '                 'If Len(CASHOS!HOSCode) Then
    '                'HOSCode = CASHOS!HOSCode
     '               'End If
      '              sqlstr7 = "select * from HOS where HOSCode = '" & Trim(HosCode) & "'"
       '             HOS.Open sqlstr7, medixmasterconn, adOpenKeyset, adLockOptimistic
        '
                    'HOSName = HOS!HOSName
         '
'         '               If CASHOS.recordCount Then
 '         '                  Do Until CASHOS.EOF
  '         '''                 Set HospitalMaster = ctl.lstHosList.listitems.Add(, , CASHOS!HosCode)
   '           '                  HospitalMaster.SubItems(1) = Trim(HOSName)
    '                            CASHOS.MoveNext
     '          '             Loop
      '          '        End If
                  '      ctl.txtDateDischarge.Mask = ""
                   ''     'ctl.txtDateDischarge = CASHOS!DateDischarge

 '                   S 'qlstr3 = "select * from DIAFirst where CasNumber = '" & Trim(frmCaseAdjustment.txtCaseNo) & "'"
'                    di 'afirst.Open sqlstr3, medixmasterconn, adOpenKeyset, adLockOptimistic
                        
                        'ICDCode = diafirst!ICDCode
                    
                    'sqlstr8 = "select * from ICD where ICDCode = '" & Trim(ICDCode) & "'"
                    'ICDCod.Open sqlstr8, medixmasterconn, adOpenKeyset, adLockOptimistic
                        
                        'ICDDesc = ICDCod!ICDDescription
                        
'                        If diafirst.recordCount Then
 '                           Do Until diafirst.EOF
  '                              Set InitialDiag = ctl.lstInitialDiag.listitems.Add(, , diafirst!ICDCode)
   '                                 'InitialDiag.SubItems(1) = Trim(ICDDesc)
    '                                InitialDiag.SubItems(2) = diafirst!DIADiagnoseBy
     '                               'InitialDiag.SubItems(3) = DIAFirst!DIADiagDate
      '                              diafirst.MoveNext
       '                     Loop
                            'ctl.txtDiagnosis2 = DIAFirst!ICDCode
        '                    'ctl.txtDiagBy2 = DIAFirst!DIADiagnoseBy
         '                   'ctl.txtDiagDate2.Mask = ""
          '                  'ctl.txtDiagDate2 = DIAFirst!DIADiagDate
           '             End If
            '        'ICDCod.Close
             '
              '      sqlstr4 = "select * from DIAFinal where CasNumber = '" & Trim(frmCaseAdjustment.txtCaseNo) & "'"
               '     DIAFinal.Open sqlstr4, medixmasterconn, adOpenKeyset, adLockOptimistic
                        
                '        If DIAFinal.recordCount Then
                 '           Do Until DIAFinal.EOF
                  '          Set FinalDiag = ctl.lstFinalDiag.listitems.Add(, , DIAFinal!ICDCode)
                   '                 'FinalDiag.SubItems(2) = DIAFirst!
                    '                FinalDiag.SubItems(2) = DIAFinal!DIADiagnoseBy
                     '               'FinalDiag.SubItems(3) = DIAFinal!DIADiagDate
                      '              DIAFinal.MoveNext
                       ''     Loop
'
 '                           'ctl.txtDiagnosis = DIAFinal!ICDCode
                            'ctl.txtDiagBy = DIAFinal!DIADiagnoseBy
  '                          'ctl.txtDiagDate = DIAFinal!DIADiagDate
   '                     End If
    '
     '               sqlstr5 = "select * from AUT where CasNumber = '" & Trim(frmCaseAdjustment.txtCaseNo) & "'"
      '              AUT.Open sqlstr5, medixmasterconn, adOpenKeyset, adLockOptimistic
       '
        '                If AUT.recordCount Then
                            'ctl.txtFirstValueDate = AUT!AUTFirstValueDate
         '                   'ctl.txtFirstInitiator = AUT!AUTFirstInitiator
          '                  ctl.txtSecondValueDate = AUT!AUTSecondValueDate
           '                 ctl.txtSecondInitiator = Trim(AUT!AUTSecondInitiator)
            '                ctl.txtSecondAuthorisationCode = Trim(AUT!AUTSecondAuthoriseCode)
             '           End If
                '    CAS.Close
              '      CASHOS.Close
               '     Set CAS = Nothing
'                    Set CASHOS = Nothing
 '                   diafirst.Close
  '                  DIAFinal.Close
   '                 AUT.Close
          '          Set diafirst = Nothing
    '                Set DIAFinal = Nothing
     '               Set AUT = Nothing
      '            End If
       '        End If
        '    'End With
         '
        End If
'    MBMCov.Close
 '   Set MBMCov = Nothing
    Unload Me
End If
End Sub

' ########################################################
Private Sub cmdSearch_Click()
    CmdStop.Enabled = True
    CmdSearch.Enabled = False
    lblTotal = 0
    lblCurrent = 0
    If Len(txtSearch) >= 3 Then
        CmdClear.Enabled = False
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call MBMInquiryList
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
    Else
        MsgBox "Search String must be at least 3 characters!", vbInformation
        txtSearch.SetFocus
        CmdStop.Enabled = False
    End If
    CmdSearch.Enabled = True
End Sub

'Private Sub cmdListAll_Click()
    'Call MBMInquiryList("All")
'End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdStop_Click()
    ' stop
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub
' #####################################################



Private Sub cboPlanCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboPlanCode.Text, cboPlanCode.SelStart) + Chr(KeyAscii) + Mid(cboPlanCode.Text, cboPlanCode.SelStart + cboPlanCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPlanCode.ListCount - 1
            
            If NewText = LCase(Left(cboPlanCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPlanCode.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   cboPlanCode.Text = ValidValue
                   cboPlanCode.SelStart = 0
                   cboPlanCode.SelLength = Len(ValidValue)
                 End If
            End If
End Sub

Private Sub cboINSCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboINSCode.Text, cboINSCode.SelStart) + Chr(KeyAscii) + Mid(cboINSCode.Text, cboINSCode.SelStart + cboINSCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboINSCode.ListCount - 1
            
            If NewText = LCase(Left(cboINSCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboINSCode.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   cboINSCode.Text = ValidValue
                   cboINSCode.SelStart = 0
                   cboINSCode.SelLength = Len(ValidValue)
                 End If
            End If
End Sub


Private Sub cboPayorCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(CboPayorCode.Text, CboPayorCode.SelStart) + Chr(KeyAscii) + Mid(CboPayorCode.Text, CboPayorCode.SelStart + CboPayorCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To CboPayorCode.ListCount - 1
            
            If NewText = LCase(Left(CboPayorCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = CboPayorCode.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   CboPayorCode.Text = ValidValue
                   CboPayorCode.SelStart = 0
                   CboPayorCode.SelLength = Len(ValidValue)
                 End If
            End If
End Sub

'Private Sub cbostatus_KeyPress(KeyAscii As Integer)
'On Error Resume Next
    'Dim NewText As String
    'Dim ValidCount As Integer
    'Dim ValidValue As String
    'Dim i As Integer
    '    If KeyAscii >= 32 And KeyAscii <> 127 Then
   '     NewText = LCase(Left(cboStatus.Text, cboStatus.SelStart) + Chr(KeyAscii) + Mid(cboStatus.Text, cboStatus.SelStart + cboStatus.SelLength + 1))
  '
 '       ValidCount = 0
'        ValidValue = ""
        
      '  For i = 0 To cboStatus.ListCount - 1
     '
    '        If NewText = LCase(Left(cboStatus.list(i), Len(NewText))) Then
'                    ValidCount = ValidCount + 1
   '                 ValidValue = cboStatus.list(i)
  '          End If
 '           Next
'            If ValidCount <= 1 Then KeyAscii = 0
       ''
      '           If ValidCount = 1 Then
     '              cboStatus.Text = ValidValue
    '               cboStatus.SelStart = 0
   '                cboStatus.SelLength = Len(ValidValue)
  '               End If
 '           End If
'End Sub


