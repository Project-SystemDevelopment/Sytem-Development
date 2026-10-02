VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchMBM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search For Member"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   315
      Left            =   8280
      TabIndex        =   4
      Top             =   6480
      Width           =   720
   End
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   4800
      TabIndex        =   1
      Top             =   0
      Width           =   5775
      Begin VB.OptionButton txtIcBcPc1 
         Caption         =   "IC Num"
         Height          =   240
         Left            =   1560
         TabIndex        =   18
         Top             =   120
         Width           =   825
      End
      Begin VB.OptionButton txtMBMNo1 
         Caption         =   "Mem Num"
         Height          =   240
         Left            =   4800
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.OptionButton txtMBMPolicyNo1 
         Caption         =   "Policy Num"
         Height          =   240
         Left            =   1560
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton txtMBMName6 
         Caption         =   "Mem Name"
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton OptEmployee 
         Caption         =   "Employee Num"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1425
      End
      Begin VB.OptionButton OptPolSubNo 
         Caption         =   "Policy Sub Num"
         Height          =   240
         Left            =   2640
         TabIndex        =   16
         Top             =   120
         Width           =   1545
      End
   End
   Begin VB.Frame Frame10 
      Height          =   690
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox cboEXPStatus 
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
         Left            =   7080
         TabIndex        =   8
         Top             =   900
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.ComboBox cboStatus 
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
         Left            =   7080
         TabIndex        =   9
         Top             =   570
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.ComboBox CboDataStatus 
         Height          =   315
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1785
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
         TabIndex        =   0
         Top             =   240
         Width           =   3825
      End
      Begin VB.Label Label11 
         Caption         =   "Exp. Status"
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
         Left            =   5880
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label72 
         Caption         =   "Policy Status"
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
         Left            =   5880
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label2 
         Caption         =   "Data Status"
         Height          =   255
         Left            =   5880
         TabIndex        =   13
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label67 
         Caption         =   "Search"
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
         TabIndex        =   12
         Top             =   240
         Width           =   810
      End
   End
   Begin MSComctlLib.ListView lstMBMInquiryList 
      Height          =   5565
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   9816
      SortKey         =   7
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Member No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Covered ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Member Name"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IC No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Policy No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Relation"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Eff Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Exp Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Bordx Date"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "PLNCode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Payor"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Group Company"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   0
      TabIndex        =   22
      Top             =   6240
      Width           =   10575
      Begin VB.TextBox totalSuppCount 
         Height          =   285
         Left            =   5040
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TotalPrincipalCount 
         Height          =   285
         Left            =   3480
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   315
         Left            =   7560
         TabIndex        =   3
         Top             =   240
         Width           =   720
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   9720
         TabIndex        =   6
         Top             =   240
         Width           =   720
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
         Left            =   6840
         TabIndex        =   2
         Top             =   240
         Width           =   720
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
         Left            =   9000
         TabIndex        =   5
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1320
         TabIndex        =   25
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Principal"
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Supp"
         Height          =   255
         Left            =   4440
         TabIndex        =   23
         Top             =   270
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmSearchMBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
' To determine what is the source
Public Source As Integer
Public intMBMRegistration As Integer
Public intMBMAdjustment As Integer
Public intMBMInquiry As Integer
Public intCASAdmission As Integer
Dim intTempMBMReg As Integer
Dim intTempMBMAdj As Integer
Dim intMBMExpired As Integer
Dim intMBMPanelHosp As Integer
Dim Current As String
Dim TotalS As String
Dim TotalP As String
Dim tmpMBMPolEffDate As String
Dim tmpMBMPolExpDate As String
Private m_blnDirection(1 To 9) As Boolean

Private Sub cmdPrint_Click()
    If lstMBMInquiryList.ListItems.Count = 0 Then
        MsgBox "No record found.", vbOKOnly + vbCritical, DialogBoxTitle
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
txtMBMName6.Value = True
lstMBMInquiryList.ListItems.Clear
lblCurrent = ""
TotalPrincipalCount = ""
totalSuppCount = ""
TotalP = 0
TotalS = 0
Current = 0

End Sub

Private Sub Form_Load()
    intExit = 0
    intMBMRegistration = 1
    intMBMAdjustment = 2
    intMBMInquiry = 3
    intCASAdmission = 4
    intTempMBMReg = 5
    intTempMBMAdj = 6
    intMBMExpired = 7
    intMBMPanelHosp = 8
    Current = 0
    TotalP = 0
    TotalS = 0
    total = 0
    Call CenterForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intExit = 1
End Sub

Private Sub MBMInquiryList()
Dim rst2 As New Adodb.Recordset
Dim Record As ListItem
Dim sqlstr As String
Dim strCriteria As String
Dim cmd As New Adodb.Command
Dim Prm1 As New Adodb.Parameter
    
Dim cmd2 As New Adodb.Command
Dim Prm2 As New Adodb.Parameter
    
    strCriteria = ""
'    sqlstr = "SELECT MBMNumber, MBMPolicyNo, MBMIcBcPp,MBMName,MBMEmployeeNo, " & _
'             " MBMPOLSubNo1, MBMCOVID,MBMPayorEffDate , MBMPayorEXPDate "
    If Len(Trim(txtSearch)) Then
       If TxtMBMNo1.Value = True Then
           strCriteria = strCriteria & " AND MBMNumber =  '" & RemStr(txtSearch) & "'"
       ElseIf txtMBMPolicyNo1.Value = True Then
           'strCriteria = strCriteria & " AND MBMPolicyNo = '" & RemStr(txtSearch) & "'"
           strCriteria = strCriteria & " AND MBMPolicyNo = '" & txtSearch & "'"
       ElseIf txtMBMName6.Value = True Then
       'LIKE '%" & RemStr(txtSearch) & "%' "
            'strCriteria = strCriteria & " AND MBMName = '" & RemStr(txtSearch) & "'"
            strCriteria = strCriteria & " AND  MBMName LIKE '" & RemStr(txtSearch) & "%'"
            'strCriteria = strCriteria & " AND left(MBMName) = '" & txtSearch & "'"
            
       ElseIf txtIcBcPc1.Value = True Then
           If Len(txtSearch) >= 12 Then
                Dim LPosition As Integer
                Dim StrIc1 As String
                Dim StrIc2 As String
                LPosition = InStr(txtSearch, "-")
                If LPosition >= 1 Then
                    StrIc1 = txtSearch
                    StrIc2 = Replace(txtSearch, "-", "")
                Else
                    StrIc1 = txtSearch
                    StrIc2 = Mid(txtSearch, 1, 6) & "-" & Mid(txtSearch, 7, 2) & "-" & Mid(txtSearch, 9, 4)
                End If
                
                strCriteria = strCriteria & " AND ( MBMIcBcPp = '" & RemStr(StrIc1) & "'  or  MBMIcBcPp = '" & RemStr(StrIc2) & "')"
           Else
                strCriteria = strCriteria & " AND MBMIcBcPp = '" & RemStr(txtSearch) & "'"
           End If
           
           'strCriteria = strCriteria & " AND Replace( MBMIcBcPp,'-','') = '" & RemStr(Replace(txtSearch, "-", "")) & "'"
       ElseIf OptEmployee.Value = True Then
           strCriteria = strCriteria & " AND MBMEmployeeNo =  '" & RemStr(txtSearch) & "'"
       ElseIf OptPOLSubNo.Value = True Then
           strCriteria = strCriteria & " AND MBMPOLSubNo1  = '" & RemStr(txtSearch) & "'"
       End If
    
  
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Search_MBM2013"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strCriteria)
    cmd.Parameters.Append Prm1
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    Do
        Do Until rst2.EOF
                    
            With rst2
                Set Record = lstMBMInquiryList.ListItems.Add(, , !MBMNumber)
                Record.SubItems(1) = IIf(Len(!MBMCOVID), !MBMCOVID, "")
                Record.SubItems(2) = IIf(Len(!MBMName), Trim(!MBMName), "")
                Record.SubItems(3) = IIf(Len(!MBMIcBcPp), Trim(!MBMIcBcPp), "")
                Record.SubItems(4) = IIf(Len(!MBMPolicyNo), Trim(!MBMPolicyNo), "")
                If Len(!MBMPayorEffDate) Then
                    Record.SubItems(6) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    tmpMBMPolEffDate = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                End If
                If Len(!MBMPayorEXPDate) Then
                   Record.SubItems(7) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                   tmpMBMPolExpDate = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                End If
                                                
                'Record.SubItems(9) = IIf(Len(!PLNCode), Trim(!PLNCode), "")
                'Record.SubItems(10) = IIf(Len(!PAYCode), Trim(!PAYCode), "")
                Record.SubItems(11) = IIf(Len(!MBMGroupCompany), Trim(!MBMGroupCompany), "")
                Record.SubItems(12) = IIf(Len(!MBMStatus), Trim(!MBMStatus), "")
                TotalP = TotalP + 1
                TotalPrincipalCount = TotalP
                
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
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    
    If txtMBMName6.Value = True Then
    
        strCriteria = ""
        strCriteria = strCriteria & " AND MBMName LIKE '%" & RemStr(txtSearch) & "'"
        
        Set cmd2.ActiveConnection = medixmasterconn
        cmd2.CommandText = "Search_MBM2013"
        cmd2.CommandType = adCmdStoredProc
        cmd2.CommandTimeout = 300
        Set Prm2 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strCriteria)
        cmd2.Parameters.Append Prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        Do
            Do Until rst2.EOF
                        
                With rst2
                    'If lstMBMInquiryList.ListItems.Count = 0 Then
                        Set Record = lstMBMInquiryList.ListItems.Add(, , !MBMNumber)
                    'Else
                    '    Record.Text = !MBMNumber
                    'End If
                    Record.SubItems(1) = IIf(Len(!MBMCOVID), !MBMCOVID, "")
                    Record.SubItems(2) = IIf(Len(!MBMName), Trim(!MBMName), "")
                    Record.SubItems(3) = IIf(Len(!MBMIcBcPp), Trim(!MBMIcBcPp), "")
                    Record.SubItems(4) = IIf(Len(!MBMPolicyNo), Trim(!MBMPolicyNo), "")
                    If Len(!MBMPayorEffDate) Then
                        Record.SubItems(6) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                        tmpMBMPolEffDate = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                    If Len(!MBMPayorEXPDate) Then
                       Record.SubItems(7) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                       tmpMBMPolExpDate = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                                                    
                    'Record.SubItems(9) = IIf(Len(!PLNCode), Trim(!PLNCode), "")
                    'Record.SubItems(10) = IIf(Len(!PAYCode), Trim(!PAYCode), "")
                    Record.SubItems(11) = IIf(Len(!MBMGroupCompany), Trim(!MBMGroupCompany), "")
                    Record.SubItems(12) = IIf(Len(!MBMStatus), Trim(!MBMStatus), "")
                    TotalP = TotalP + 1
                    TotalPrincipalCount = TotalP
                    
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
            intExit = 0
            rst2.Close
            Set rst2 = Nothing
            
        End If
    End If
End Sub

Private Sub SuppInquiryList()
Dim rst2 As New Adodb.Recordset
Dim rst3 As New Adodb.Recordset
Dim Record As ListItem
Dim sqlstr As String
Dim sqlstr2 As String
Dim strCriteria As String
Dim cmd1 As New Adodb.Command
Dim Prm1 As New Adodb.Parameter
Dim cmd2 As New Adodb.Command
Dim Prm2 As New Adodb.Parameter
    
    
Dim cmd3 As New Adodb.Command
Dim Prm3 As New Adodb.Parameter
    
'    strCriteria = "SELECT MBMCNumber, MBMCName, MBMCCoverID, MBMCIcBcPpNo,MBMCEffDate, MBMCExpDate, MBMCPLNCode FROM MBMCoveredPersons WHERE 0=0 "
'    If Len(Trim(txtSearch)) Then
'         If txtMBMNo1.Value = True Then
'             strCriteria = strCriteria & " AND MBMCNumber =  '" & RemStr(txtSearch) & "'"
'         ElseIf txtMBMName6.Value = True Then
'             strCriteria = strCriteria & " AND MBMCName = '" & RemStr(txtSearch) & "' "
'         ElseIf txtIcBcPc1.Value = True Then
'             strCriteria = strCriteria & " AND MBMCIcBcPpNo = '" & RemStr(txtSearch) & "' "
'         End If
'    End If
'    rst2.Open strCriteria, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
    
     If Len(Trim(txtSearch)) Then
       If TxtMBMNo1.Value = True Then
           strCriteria = strCriteria & " AND MBMCNumber =  '" & RemStr(txtSearch) & "'"
       ElseIf txtMBMName6.Value = True Then
           strCriteria = strCriteria & " AND MBMCName LIKE '%" & RemStr(txtSearch) & "' "
       ElseIf txtIcBcPc1.Value = True Then
           'strCriteria = strCriteria & " AND RTRIM(LTRIM(REPLACE(MBMCIcBcPpNo,'-',''))) = RTRIM(LTRIM(REPLACE('" & RemStr(txtSearch) & "','-','')))"
           strCriteria = strCriteria & " AND MBMCIcBcPpNo = '" & RemStr(txtSearch) & "'"
       End If
    
  
         Set cmd1.ActiveConnection = medixmasterconn
         cmd1.CommandText = "Search_MBMCov2013"
         cmd1.CommandType = adCmdStoredProc
         cmd1.CommandTimeout = 300
         Set Prm2 = cmd1.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strCriteria)
         cmd1.Parameters.Append Prm2
         rst2.CursorLocation = adUseClient
         rst2.CursorType = adOpenForwardOnly
         rst2.CacheSize = 100
         Set rst2 = cmd1.Execute
        
         Do
             Do Until rst2.EOF
                 With rst2
                     Set Record = lstMBMInquiryList.ListItems.Add(, , !MBMCNumber)
                     Record.SubItems(1) = IIf(Len(!MBMCCoverID), !MBMCCoverID, "")
                     Record.SubItems(2) = IIf(Len(!MBMCName), Trim(!MBMCName), "")
                     Record.SubItems(3) = IIf(Len(!MBMCICBCPPNo), Trim(!MBMCICBCPPNo), "")
                     'sqlstr2 = " Select MBMPayorEffDate, MBMPayorExpDate " & _
                     '          " FROM MBM Where MBMNumber = '" & Trim(!MBMCNumber) & "'"
         
                     'rst3.Open sqlstr2, medixmasterconn, adOpenKeyset, adLockOptimistic
                     'If rst3.recordCount <> 0 Then
                         If Len(!MBMCEffDate) Then
                             Record.SubItems(6) = Format(!MBMCEffDate, "dd/mm/yyyy")
                         End If
                         If Len(!MBMCExpDate) Then
                             Record.SubItems(7) = Format(!MBMCExpDate, "dd/mm/yyyy")
                         End If
                         'Record.SubItems(9) = IIf(Len(!MBMCPLNCode), Trim(!MBMCPLNCode), "")
                         'Record.SubItems(10) = IIf(Len(!PAYCode), Trim(!PAYCode), "")
                         Record.SubItems(10) = Mid(!MBMCNumber, 1, 2)
                         Record.SubItems(12) = IIf(Len(!MBMStatus), Trim(!MBMStatus), "")
                     'End If
                     
                     'rst3.Close
                     'Set rst3 = Nothing
                     
                     TotalS = TotalS + 1
                     totalSuppCount = TotalS
                     
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
         'End If
         intExit = 0
         
         'If lstMBMInquiryList.ListItems.Count = 0 Then
         '    MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
         'End If
         
         rst2.Close
         Set rst2 = Nothing
    
    
    If txtMBMName6.Value = True Then
            strCriteria = ""
           strCriteria = strCriteria & " AND MBMCName LIKE '" & RemStr(txtSearch) & "%' "
    
             Set cmd1.ActiveConnection = medixmasterconn
             cmd3.CommandText = "Search_MBMCov2013"
             cmd3.CommandType = adCmdStoredProc
             cmd3.CommandTimeout = 300
             Set Prm3 = cmd3.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strCriteria)
             cmd3.Parameters.Append Prm3
             rst2.CursorLocation = adUseClient
             rst2.CursorType = adOpenForwardOnly
             rst2.CacheSize = 100
             Set rst2 = cmd1.Execute
            
             Do
                 Do Until rst2.EOF
                     With rst2
                         Set Record = lstMBMInquiryList.ListItems.Add(, , !MBMCNumber)
                         Record.SubItems(1) = IIf(Len(!MBMCCoverID), !MBMCCoverID, "")
                         Record.SubItems(2) = IIf(Len(!MBMCName), Trim(!MBMCName), "")
                         Record.SubItems(3) = IIf(Len(!MBMCICBCPPNo), Trim(!MBMCICBCPPNo), "")
                         'sqlstr2 = " Select MBMPayorEffDate, MBMPayorExpDate " & _
                         '          " FROM MBM Where MBMNumber = '" & Trim(!MBMCNumber) & "'"
             
                         'rst3.Open sqlstr2, medixmasterconn, adOpenKeyset, adLockOptimistic
                         'If rst3.recordCount <> 0 Then
                             If Len(!MBMCEffDate) Then
                                 Record.SubItems(6) = Format(!MBMCEffDate, "dd/mm/yyyy")
                             End If
                             If Len(!MBMCExpDate) Then
                                 Record.SubItems(7) = Format(!MBMCExpDate, "dd/mm/yyyy")
                             End If
                             'Record.SubItems(9) = IIf(Len(!MBMCPLNCode), Trim(!MBMCPLNCode), "")
                             'Record.SubItems(10) = IIf(Len(!PAYCode), Trim(!PAYCode), "")
                             Record.SubItems(10) = Mid(!MBMCNumber, 1, 2)
                             Record.SubItems(12) = IIf(Len(!MBMStatus), Trim(!MBMStatus), "")
                         'End If
                         
                         'rst3.Close
                         'Set rst3 = Nothing
                         
                         TotalS = TotalS + 1
                         totalSuppCount = TotalS
                         
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
             'End If
             intExit = 0
             
             'If lstMBMInquiryList.ListItems.Count = 0 Then
             '    MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
             'End If
             
             rst2.Close
             Set rst2 = Nothing
        End If
    End If
End Sub

Private Sub CmdExit_Click()
    intExit = 1
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    
    Screen.MousePointer = vbHourglass
    CmdSearch.Enabled = False

    lstMBMInquiryList.ListItems.Clear
    lblCurrent = ""
    TotalPrincipalCount = ""
    totalSuppCount = ""
    Current = 0
    TotalP = 0
    TotalS = 0
    
    CmdStop.Enabled = True
    If Len(txtSearch) >= 4 Then
        CmdClear.Enabled = False
        CmdExit.Enabled = False
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        lstMBMInquiryList.Enabled = False
        Call MBMInquiryList
        If txtMBMName6.Value = True Or TxtMBMNo1.Value = True Or txtIcBcPc1.Value = True Then
            Call SuppInquiryList
        End If
        lstMBMInquiryList.Enabled = True
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing

        If lstMBMInquiryList.ListItems.Count = 0 Then
            MsgBox "Please verify with Membership Department!", vbOKOnly + vbCritical, DialogBoxTitle
        End If
    Else
        MsgBox "Search String must be at least 4 characters!", vbInformation
        txtSearch.SetFocus
        CmdStop.Enabled = False
    End If
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdStop.Enabled = True
    CmdSearch.Enabled = True
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdStop_Click()
    ' stop
    intExit = 1
    CmdClear.Enabled = True
    CmdExit.Enabled = True
End Sub

'**********************Start Sorted ListView *******************************
Private Sub lstMBMInquiryList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
    Case 3
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    End Select
End Sub
'**********************End Sorted ListView *******************************

Private Sub lstMBMInquiryList_DblClick()
On Error GoTo ErrorShow
    
    If lstMBMInquiryList.ListItems.Count Then
        If Source = intMBMRegistration Then
            frmNewMBMReg05.txtMBMPrevMEMNo = lstMBMInquiryList.SelectedItem.Text
            frmNewMBMReg05.cmdShow_Click
        ElseIf Source = intMBMInquiry Then
            frmNewMBMEnquiry.txtMBMNumber = lstMBMInquiryList.SelectedItem.Text
            frmNewMBMEnquiry.cmdShow_Click
        ElseIf Source = intMBMAdjustment Then
           frmNewMBMAdjustment.txtMBMNumber = lstMBMInquiryList.SelectedItem.Text
           frmNewMBMAdjustment.cmdShow_Click
        ElseIf Source = intCASAdmission Then
           frmCaseAdmission.txtMBMNumber = lstMBMInquiryList.SelectedItem.Text
           frmCaseAdmission.cmdShow_Click
        ElseIf Source = intTempMBMReg Then
            FrmTempMBMReg.txtMBMPrevMEMNo = lstMBMInquiryList.SelectedItem.Text
            FrmTempMBMReg.cmdShow_Click
        ElseIf Source = intTempMBMAdj Then
            FrmTempMBMAdj.txtMBMNumber = lstMBMInquiryList.SelectedItem.Text
            FrmTempMBMAdj.cmdShow_Click
        ElseIf Source = intMBMExpired Then
            frmMBMExpiredPolicy.txtMBMNumber = lstMBMInquiryList.SelectedItem.Text
            frmMBMExpiredPolicy.cmdShow_Click
        ElseIf Source = intMBMPanelHosp Then
            frmPanelHospitals.txtPlanNo = lstMBMInquiryList.SelectedItem.SubItems(9)
            frmPanelHospitals.cboPLNPayor = lstMBMInquiryList.SelectedItem.SubItems(10)
            frmPanelHospitals.cmdSearch_Click
            
        End If
        Unload Me
   End If
   
    Exit Sub
ErrorShow:

   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
   
End Sub


