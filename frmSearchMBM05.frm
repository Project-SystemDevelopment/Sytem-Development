VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmSearchMBM05 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search for MBM 05 "
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   5880
      Width           =   9615
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   255
         Left            =   8520
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   255
         Left            =   7440
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame10 
      Height          =   690
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox CboDataStatus 
         Height          =   315
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
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
         TabIndex        =   5
         Top             =   570
         Visible         =   0   'False
         Width           =   1785
      End
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
         TabIndex        =   4
         Top             =   900
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label txtSearch 
         BackColor       =   &H00C0E0FF&
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
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   3255
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
         TabIndex        =   10
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label2 
         Caption         =   "Data Status"
         Height          =   255
         Left            =   5880
         TabIndex        =   9
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
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
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   1020
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
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   1020
      End
   End
   Begin MSComctlLib.ListView lstMBMInquiryList 
      Height          =   5115
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9022
      View            =   3
      LabelEdit       =   1
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Member No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Member Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IC No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Policy No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Eff Date"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Exp Date"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Group Comp"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Bordx Date"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "RegDate"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Reg By"
         Object.Width           =   1411
      EndProperty
   End
End
Attribute VB_Name = "frmSearchMBM05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Source As Integer
Public SearchRecCnt As Integer
Public tmpMBMNo As String
Private m_blnDirection(1 To 11) As Boolean

Private Sub MBMInquiryList()
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim strCriteria As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
    SearchRecCnt = 0
    lstMBMInquiryList.Enabled = False
    strCriteria = RemStr(txtSearch) & "%"
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "Search_MBM_05"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set Prm1 = cmd.CreateParameter("StrSql", adVarChar, adParamInput, 20, strCriteria)
    cmd.Parameters.Append Prm1
    rst2.CursorLocation = adUseClient
    rst2.CursorType = adOpenForwardOnly
    rst2.CacheSize = 100
    Set rst2 = cmd.Execute
    Do While Not rst2.EOF
        
        SearchRecCnt = SearchRecCnt + 1
        With rst2
            Set Record = lstMBMInquiryList.ListItems.Add(, , !MBMNumber)
            If SearchRecCnt = 1 Then
                tmpMBMNo = !MBMNumber
            End If
            Record.SubItems(1) = IIf(Len(!MBMCOVID), !MBMCOVID, "")
            Record.SubItems(2) = IIf(Len(!MBMName), Trim(!MBMName), "")
            Record.SubItems(3) = IIf(Len(!MBMIcBcPp), Trim(!MBMIcBcPp), "")
            Record.SubItems(4) = IIf(Len(!MBMPolicyNo), Trim(!MBMPolicyNo), "")
            Record.SubItems(5) = IIf(Len(!MBMStatus), Trim(!MBMStatus), "")
            If Len(!MBMPayorEffDate) Then
                Record.SubItems(6) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
            End If
            If Len(!MBMPayorEXPDate) Then
               Record.SubItems(7) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
            End If
            If Len(!MBMGroupCompany) Then
               Record.SubItems(8) = IIf(Len(!MBMGroupCompany), !MBMGroupCompany, "")
            End If
            
            If Len(!MBMBordxDate) Then
               Record.SubItems(9) = Format(!MBMBordxDate, "dd/mm/yyyy")
            End If
            If Len(!MBMValueDate) Then
               Record.SubItems(10) = Format(!MBMValueDate, "dd/mm/yyyy")
            End If
            Record.SubItems(11) = IIf(Len(!MBMRegUser), Trim(!MBMRegUser), "")
        End With
        rst2.MoveNext
    Loop
    rst2.Close
    Set rst2 = Nothing
'    medixmasterconn.Close
 '   Set medixmasterconn = Nothing
    lstMBMInquiryList.Enabled = True
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

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

Private Sub Form_Load()
    If Source = 1 Then
        txtSearch = frmNewMBMReg05.txtMBMPrevMEMNo
    ElseIf Source = 2 Then
       txtSearch = frmNewMBMAdjustment.txtMBMNumber
    ElseIf Source = 3 Then
        txtSearch = frmNewMBMEnquiry.txtMBMNumber
    ElseIf Source = 4 Then
        txtSearch = frmCaseAdmission.txtMBMNumber
    ElseIf Source = 5 Then
        txtSearch = frmRECNonDisclosure.txtMBMNumber
    ElseIf Source = 7 Then
        'txtSearch = frmCaseRecording.txtMBMNumber
        
    End If
    If Len(Trim(txtSearch)) > 0 Then
        Call MBMInquiryList
    End If
End Sub

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
    
    If lstMBMInquiryList.ListItems.Count Then
        If Source = 1 Then
            'frmNewMBMReg05.txtMBMPrevMEMNo = lstMBMInquiryList.SelectedItem.Text
            'frmNewMBMReg05.cmdShow_Click
        ElseIf Source = 2 Then
           frmNewMBMAdjustment.txtMBMNumber = lstMBMInquiryList.SelectedItem.Text
           frmNewMBMAdjustment.cmdShow_Click
        ElseIf Source = 3 Then
            frmNewMBMEnquiry.txtMBMNumber = lstMBMInquiryList.SelectedItem.Text
            frmNewMBMEnquiry.cmdShow_Click
        ElseIf Source = 4 Then
           frmCaseAdmission.txtMBMNumber = frmSearchMBM05.lstMBMInquiryList.SelectedItem.Text
           frmCaseAdmission.cmdShow_Click
        ElseIf Source = 5 Then
           frmRECNonDisclosure.txtMBMNumber = frmSearchMBM05.lstMBMInquiryList.SelectedItem.Text
           frmRECNonDisclosure.cmdGo_Click
        ElseIf Source = 6 Then
           FrmTempMBMAdj.txtMBMNumber = frmSearchMBM05.lstMBMInquiryList.SelectedItem.Text
           FrmTempMBMAdj.cmdShow_Click
        ElseIf Source = 7 Then
          ' frmCaseRecording.txtMBMNumber = frmSearchMBM05.lstMBMInquiryList.SelectedItem.Text
          ' frmCaseRecording.cmdShow_Click
        
        End If
        Unload Me
   End If
   
End Sub

