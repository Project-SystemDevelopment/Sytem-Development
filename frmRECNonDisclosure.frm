VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRECNonDisclosure 
   Caption         =   "Record Of Non-Disclosure"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   10590
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   0
      TabIndex        =   14
      Top             =   4560
      Width           =   10575
      Begin MSMask.MaskEdBox txtPreDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboPreCategories 
         Height          =   315
         Left            =   3240
         TabIndex        =   4
         Top             =   450
         Width           =   1815
      End
      Begin MSMask.MaskEdBox txtUndDate 
         Height          =   285
         Left            =   6600
         TabIndex        =   6
         Top             =   465
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboUndCategories 
         Height          =   315
         Left            =   8760
         TabIndex        =   7
         Top             =   430
         Width           =   1695
      End
      Begin VB.TextBox txtUndInfBy 
         Height          =   285
         Left            =   6600
         MaxLength       =   30
         TabIndex        =   8
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtPreInfBy 
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   5
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "Pre-Existing Illness"
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
         TabIndex        =   24
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   495
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Categories"
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Informed by"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Underlying Illness"
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
         TabIndex        =   28
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Date"
         Height          =   255
         Left            =   5640
         TabIndex        =   27
         Top             =   495
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Categories"
         Height          =   255
         Left            =   7920
         TabIndex        =   26
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Informed by"
         Height          =   255
         Left            =   5640
         TabIndex        =   25
         Top             =   720
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lstMBMListing 
      Height          =   3255
      Left            =   0
      TabIndex        =   13
      Top             =   960
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "MBM Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cover ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "MBM Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Policy No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Eff Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Exp Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Pre-Existing Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Pre-Existing Cat"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Pre-Existing Inf"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Und Illness Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Und Illness Cat"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Und Illness Inf"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   240
      Width           =   10575
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
         Height          =   285
         Left            =   4200
         TabIndex        =   34
         Top             =   240
         Width           =   810
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&Pre"
         Height          =   285
         Left            =   3360
         TabIndex        =   33
         Top             =   240
         Width           =   810
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   285
         Left            =   5040
         TabIndex        =   12
         Top             =   240
         Width           =   810
      End
      Begin VB.CommandButton cmdGO 
         Caption         =   "Go"
         Default         =   -1  'True
         Height          =   285
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   810
      End
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Mem No"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   0
      TabIndex        =   18
      Top             =   5640
      Width           =   10575
      Begin VB.TextBox tempMBMCovID 
         Height          =   285
         Left            =   1680
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox tempMBMNo 
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   7560
         TabIndex        =   9
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   9840
         TabIndex        =   21
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   9240
         TabIndex        =   20
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdEXPtoExcel 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   8280
         TabIndex        =   19
         Top             =   180
         Width           =   975
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "of"
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Label lblMBMName 
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
      Left            =   1320
      TabIndex        =   30
      Top             =   4320
      Width           =   6615
   End
   Begin VB.Label Label12 
      Caption         =   "Member Name : "
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Record of Non-Disclosure"
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
      TabIndex        =   1
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmRECNonDisclosure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADOParameterMBM As ADODB.Parameter
Dim ADOCommandMBM As New ADODB.Command
Dim ADOrstMBM As New ADODB.Recordset

Dim ADOParameterMBMCov As ADODB.Parameter
Dim ADOCommandMBMCov As New ADODB.Command
Dim ADOrstMBMCov As New ADODB.Recordset
Private m_blnDirection(1 To 10) As Boolean


Private Sub ComboListing()
    cboPreCategories.Clear
    cboUndCategories.Clear
    
    cboPreCategories.AddItem "0 - Illness"
    cboPreCategories.AddItem "1 - Injury"
    cboPreCategories.AddItem "2 - Hospitalisation"
    
    cboUndCategories.AddItem "0 - Diabetes"
    cboUndCategories.AddItem "1 - Hypertension"
End Sub

Private Sub cboPreCategories_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPreCategories.Text, cboPreCategories.SelStart) + Chr(KeyAscii) + Mid(cboPreCategories.Text, cboPreCategories.SelStart + cboPreCategories.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPreCategories.ListCount - 1
 
        If NewText = LCase(Left(cboPreCategories.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPreCategories.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPreCategories.Text = ValidValue
                cboPreCategories.SelStart = 0
                cboPreCategories.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboUndCategories_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboUndCategories.Text, cboUndCategories.SelStart) + Chr(KeyAscii) + Mid(cboUndCategories.Text, cboUndCategories.SelStart + cboUndCategories.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboUndCategories.ListCount - 1
 
        If NewText = LCase(Left(cboUndCategories.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboUndCategories.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboUndCategories.Text = ValidValue
                cboUndCategories.SelStart = 0
                cboUndCategories.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub CmdClear_Click()
    Call ClearEntries
    txtMBMNumber = ""
    lstMBMListing.listitems.Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Public Sub cmdGO_Click()
Dim MBMNumber As String
Dim tmpREcCnt As Integer
    
    MBMNumber = txtMBMNumber
    Call ClearEntries
    lstMBMListing.listitems.Clear
    txtMBMNumber = MBMNumber
    Screen.MousePointer = vbHourglass
    If txtMBMNumber = "" Then
        MsgBox "Please Enter Membership No", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMNumber.SetFocus
    Else
        If Len(txtMBMNumber) = 8 Or Len(txtMBMNumber) = 9 Then
            tmpREcCnt = 1
            frmSearchMBM05.Source = 5
            frmSearchMBM05.Show
            tmpREcCnt = frmSearchMBM05.SearchRecCnt
            If frmSearchMBM05.SearchRecCnt < 1 Then
                Unload frmSearchMBM05
            ElseIf frmSearchMBM05.SearchRecCnt = 1 Then
                txtMBMNumber = frmSearchMBM05.tmpMBMNo
                Unload frmSearchMBM05
            End If

        End If
        If tmpREcCnt < 2 Then
            medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
                frmNewMBMOthers.Source = 2
                Call FindMBM
            medixmasterconn.Close
            Set medixmasterconn = Nothing
        End If
    End If
    
    Screen.MousePointer = Default
End Sub

Private Sub CmdNext_Click()
    Dim MBM As New ADODB.Recordset
    Dim sql As String
    Dim MBMNumber As String
    Dim status As Boolean
        
        status = False
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        sql = " select MBMNumber from MBM where MBMnumber > '" & RemStr(txtMBMNumber) & "' " & _
            " Order by MBMnumber "
        MBM.MaxRecords = 1
        MBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBM.recordCount Then
            MBMNumber = MBM!MBMNumber
            status = True
        End If
        MBM.Close
        Set MBM = Nothing
        medixmasterconn.Close
        Set medixmasterconn = Nothing
        
        If status = True Then
            txtMBMNumber = MBMNumber
            cmdGO_Click
        Else
            MsgBox "Last Record Reached! " & vbNewLine & "Not a valid Membership Number Format! ", vbCritical + vbOKOnly
        End If
End Sub

Private Sub cmdPrev_Click()
    Dim MBM As New ADODB.Recordset
    Dim sql As String
    Dim MBMNumber As String
    Dim status As Boolean
        
        status = False
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        sql = " select MBMNumber from MBM where MBMnumber < '" & RemStr(txtMBMNumber) & "' " & _
            " Order by MBMnumber DESC "
        MBM.MaxRecords = 1
        MBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBM.recordCount Then
            MBMNumber = MBM!MBMNumber
            status = True
        End If
        MBM.Close
        Set MBM = Nothing
        medixmasterconn.Close
        Set medixmasterconn = Nothing
        
        If status = True Then
            txtMBMNumber = MBMNumber
            cmdGO_Click
        Else
            MsgBox "First Record Reached! " & vbNewLine & "Not a valid Membership Number Format! ", vbCritical + vbOKOnly
        End If
End Sub

Private Sub CmdUpdate_Click()
Dim RecordSaved
Dim bSucess As Boolean
    If Trim(tempMBMNo) = "" Or Trim(tempMBMCovID) = "" Then
        MsgBox "Please Select a record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
        On Error GoTo ErrorSave
            Screen.MousePointer = vbHourglass
                medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
                medixmasterconn.beginTrans
                bSucess = True
                
                If tempMBMCovID = "00" Then
                    Call SavePrincipal
                Else
                    Call SaveSupplementary
                End If
                
                medixmasterconn.CommitTrans
                bSucess = False
                medixmasterconn.Close
                Set medixmasterconn = Nothing
                MsgBox "Record Updated Successfully", vbOKOnly + vbInformation, DialogBoxTitle
                Call cmdGO_Click
            Screen.MousePointer = vbDefault
            
        End If
    End If

    
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If bSucess = True Then
        medixmasterconn.RollbackTrans
    End If
    
    medixmasterconn.Close
    Set medixmasterconn = Nothing

End Sub

Private Sub Form_Load()
    Call ComboListing
End Sub
Private Function FindMBM() As Integer
Dim strCount As String
Dim sqlstr2 As String
Dim sql As String

    
    sql = "Exec SearchMBMCovPersonsProc '" & Trim(txtMBMNumber) & "'"
    Set ADOrstMBM = New ADODB.Recordset
    ADOrstMBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If Not ADOrstMBM.EOF Then
        Call showMBM
        If Trim(ADOrstMBM!INSCode) = "F" Or Trim(ADOrstMBM!INSCode) = "H" Then
            If Len(Trim(ADOrstMBM!MBMCCoverID)) And Trim(ADOrstMBM!MBMCCoverID) <> "00" Then
                Call showMBMCov
            End If
        End If
    Else
        MsgBox "No record found!", vbCritical
    End If
    ADOrstMBM.Close
    Set ADOrstMBM = Nothing
    
End Function

Private Function showMBM() As Integer
Dim MemberListItem As ListItem

    'Do While Not ADOrstMBM.EOF
        With ADOrstMBM
            Set MemberListItem = lstMBMListing.listitems.Add(, , IIf(Trim(!MBMNumber) <> "", !MBMNumber, ""))
                MemberListItem.SubItems(1) = "00"
                MemberListItem.SubItems(2) = IIf(Len(Trim(!MBMCName)), Trim(!MBMCName), "")
                MemberListItem.SubItems(3) = IIf(Len(Trim(!MBMPolicyNo)), Trim(!MBMPolicyNo), "")
                MemberListItem.SubItems(4) = !MBMPayorEffDate
                MemberListItem.SubItems(5) = !MBMPayorEXPDate
                MemberListItem.SubItems(6) = IIf(Len(!MBMPreIllDate), !MBMPreIllDate, "")
                MemberListItem.SubItems(7) = IIf(Len(!MBMPreIllCAT), !MBMPreIllCAT, "")
                MemberListItem.SubItems(8) = IIf(Len(!MBMPreIllInfBy), !MBMPreIllInfBy, "")
                MemberListItem.SubItems(9) = IIf(Len(!MBMUndIllDate), !MBMUndIllDate, "")
                MemberListItem.SubItems(10) = IIf(Len(!MBMUndIllCAT), !MBMUndIllCAT, "")
                MemberListItem.SubItems(11) = IIf(Len(!MBMUndIllInfBy), !MBMUndIllInfBy, "")
        End With
        'ADOrstMBM.MoveNext
    'Loop
End Function

Private Sub showMBMCov()
Dim MemberConveredPerson As ListItem
    
    Do While Not ADOrstMBM.EOF
        With ADOrstMBM
            Set MemberConveredPerson = lstMBMListing.listitems.Add(, , IIf(Trim(!MBMCNumber) <> "", !MBMCNumber, ""))
                MemberConveredPerson.SubItems(1) = IIf(Len(Trim(!MBMCCoverID)), Trim(!MBMCCoverID), "")
                MemberConveredPerson.SubItems(2) = IIf(Len(Trim(!MBMCName)), Trim(!MBMCName), "")
                MemberConveredPerson.SubItems(3) = IIf(Len(Trim(!MBMPolicyNo)), Trim(!MBMPolicyNo), "")
                MemberConveredPerson.SubItems(4) = IIf(Len(!MBMCEffDate), !MBMCEffDate, !MBMPayorEffDate)
                MemberConveredPerson.SubItems(5) = IIf(Len(!MBMCExpDate), !MBMCExpDate, !MBMPayorEXPDate)
                MemberConveredPerson.SubItems(6) = IIf(Len(!MBMCPreIllDate), !MBMCPreIllDate, "")
                MemberConveredPerson.SubItems(7) = IIf(Len(!MBMCPreIllCAT), !MBMCPreIllCAT, "")
                MemberConveredPerson.SubItems(8) = IIf(Len(!MBMCPreInfBy), !MBMCPreInfBy, "")
                MemberConveredPerson.SubItems(9) = IIf(Len(!MBMCUndIllDate), !MBMCUndIllDate, "")
                MemberConveredPerson.SubItems(10) = IIf(Len(!MBMCUndIllCAT), !MBMCUndIllCAT, "")
                MemberConveredPerson.SubItems(11) = IIf(Len(!MBMCUndInfBy), !MBMCUndInfBy, "")
                
        End With
        ADOrstMBM.MoveNext
    Loop
End Sub

Private Sub ClearEntries()
    On Error Resume Next
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Name <> "txtMBMNumber" Then
               ctl.Enabled = True
               ctl.Text = ""
            End If
        ElseIf TypeOf ctl Is ComboBox Then
                ctl.Enabled = True
                ctl.Text = ""
        ElseIf TypeOf ctl Is CommandButton Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.Enabled = True
            ctl.mask = ""
            ctl.Text = ""
            ctl.mask = "##/##/####"
        End If
    Next
    lblMBMName = ""
End Sub


Private Sub lstMBMListing_Click()
    If lstMBMListing.listitems.count > 0 Then
        Call ClearEntries
        lblMBMName = ""
        tempMBMNo = lstMBMListing.SelectedItem.Text
        tempMBMCovID = lstMBMListing.SelectedItem.SubItems(1)
        lblMBMName = lstMBMListing.SelectedItem.SubItems(2)
        txtPreDate = IIf(Len(lstMBMListing.SelectedItem.SubItems(6)), lstMBMListing.SelectedItem.SubItems(6), "__/__/____")
        If Len(lstMBMListing.SelectedItem.SubItems(7)) Then
            If Trim(lstMBMListing.SelectedItem.SubItems(7)) = 0 Then
                cboPreCategories.ListIndex = 0
            ElseIf Trim(lstMBMListing.SelectedItem.SubItems(7)) = 1 Then
                cboPreCategories.ListIndex = 1
            ElseIf Trim(lstMBMListing.SelectedItem.SubItems(7)) = 2 Then
                cboPreCategories.ListIndex = 2
            End If
        End If
        txtPreInfBy = lstMBMListing.SelectedItem.SubItems(8)
        txtUndDate = IIf(Len(lstMBMListing.SelectedItem.SubItems(9)), lstMBMListing.SelectedItem.SubItems(9), "__/__/____")
        If Len(lstMBMListing.SelectedItem.SubItems(10)) Then
            If Trim(lstMBMListing.SelectedItem.SubItems(10)) = 0 Then
                cboUndCategories.ListIndex = 0
            ElseIf Trim(lstMBMListing.SelectedItem.SubItems(10)) = 1 Then
                cboUndCategories.ListIndex = 1
            End If
        End If
        
        txtUndInfBy = lstMBMListing.SelectedItem.SubItems(11)
    End If
End Sub

Private Sub lstMBMListing_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstMBMListing, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstMBMListing, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
    Case 3
        SortListView lstMBMListing, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstMBMListing, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstMBMListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
    Case 6
        SortListView lstMBMListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lstMBMListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
    Case 8
        SortListView lstMBMListing, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstMBMListing, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstMBMListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
    End Select

End Sub

Private Sub lstMBMListing_KeyDown(KeyCode As Integer, Shift As Integer)
    If lstMBMListing.listitems.count > 0 Then
        Call ClearEntries
        lblMBMName = ""
        tempMBMNo = lstMBMListing.SelectedItem.Text
        tempMBMCovID = lstMBMListing.SelectedItem.SubItems(1)
        lblMBMName = lstMBMListing.SelectedItem.SubItems(2)
        txtPreDate = IIf(Len(lstMBMListing.SelectedItem.SubItems(6)), lstMBMListing.SelectedItem.SubItems(6), "__/__/____")
        If Len(lstMBMListing.SelectedItem.SubItems(7)) Then
            If Trim(lstMBMListing.SelectedItem.SubItems(7)) = 0 Then
                cboPreCategories.ListIndex = 0
            ElseIf Trim(lstMBMListing.SelectedItem.SubItems(7)) = 1 Then
                cboPreCategories.ListIndex = 1
            ElseIf Trim(lstMBMListing.SelectedItem.SubItems(7)) = 2 Then
                cboPreCategories.ListIndex = 2
            End If
        End If
        txtPreInfBy = lstMBMListing.SelectedItem.SubItems(8)
        txtUndDate = IIf(Len(lstMBMListing.SelectedItem.SubItems(9)), lstMBMListing.SelectedItem.SubItems(9), "__/__/____")
        If Len(lstMBMListing.SelectedItem.SubItems(10)) Then
            If Trim(lstMBMListing.SelectedItem.SubItems(10)) = 0 Then
                cboUndCategories.ListIndex = 0
            ElseIf Trim(lstMBMListing.SelectedItem.SubItems(10)) = 1 Then
                cboUndCategories.ListIndex = 1
            End If
        End If
        
        txtUndInfBy = lstMBMListing.SelectedItem.SubItems(11)
    End If
End Sub

Private Sub lstMBMListing_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstMBMListing.listitems.count > 0 Then
        Call ClearEntries
        lblMBMName = ""
        tempMBMNo = lstMBMListing.SelectedItem.Text
        tempMBMCovID = lstMBMListing.SelectedItem.SubItems(1)
        lblMBMName = lstMBMListing.SelectedItem.SubItems(2)
        txtPreDate = IIf(Len(lstMBMListing.SelectedItem.SubItems(6)), lstMBMListing.SelectedItem.SubItems(6), "__/__/____")
        If Len(lstMBMListing.SelectedItem.SubItems(7)) Then
            If Trim(lstMBMListing.SelectedItem.SubItems(7)) = 0 Then
                cboPreCategories.ListIndex = 0
            ElseIf Trim(lstMBMListing.SelectedItem.SubItems(7)) = 1 Then
                cboPreCategories.ListIndex = 1
            ElseIf Trim(lstMBMListing.SelectedItem.SubItems(7)) = 2 Then
                cboPreCategories.ListIndex = 2
            End If
        End If
        txtPreInfBy = lstMBMListing.SelectedItem.SubItems(8)
        txtUndDate = IIf(Len(lstMBMListing.SelectedItem.SubItems(9)), lstMBMListing.SelectedItem.SubItems(9), "__/__/____")
        If Len(lstMBMListing.SelectedItem.SubItems(10)) Then
            If Trim(lstMBMListing.SelectedItem.SubItems(10)) = 0 Then
                cboUndCategories.ListIndex = 0
            ElseIf Trim(lstMBMListing.SelectedItem.SubItems(10)) = 1 Then
                cboUndCategories.ListIndex = 1
            End If
        End If
        
        txtUndInfBy = lstMBMListing.SelectedItem.SubItems(11)
    End If
End Sub

Private Sub SavePrincipal()
Dim RecordSaved
Dim strsqlTwo As String
Dim MBMTwo As New ADODB.Recordset
Dim strsqlMBMCovTwo As String
Dim MBMCovTwo As New ADODB.Recordset

    
        
    strsqlTwo = " SELECT MBMPreIllDate,MBMPreIllCAT,MBMPreIllInfBy, " & _
            " MBMUndIllDate, MBMUndIllCAT , MBMUndIllInfBy "
    strsqlTwo = strsqlTwo & " From MBMTwo Where MBMNumber ='" & Trim(txtMBMNumber) & "'"
    
    MBMTwo.Open strsqlTwo, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        With MBMTwo
            !MBMPreIllDate = IIf(IsDate(txtPreDate), txtPreDate, "")
            !MBMPreIllCAT = Mid(cboPreCategories, 1, 1)
            !MBMPreIllInfBy = ChkStr(txtPreInfBy)
            !MBMUndIllDate = IIf(IsDate(txtUndDate), txtUndDate, "")
            !MBMUndIllCAT = Mid(cboUndCategories, 1, 1)
            !MBMUndIllInfBy = ChkStr(txtUndInfBy)
        End With
    
    MBMTwo.Update
              
    MBMTwo.Close
    Set MBMTwo = Nothing
End Sub

Private Sub SaveSupplementary()
Dim RecordSaved
Dim strsqlMBMCovTwo As String
Dim MBMCovTwo As New ADODB.Recordset

    
    
    strsqlMBMCovTwo = " SELECT MBMCPreIllDate,MBMCPreIllCAT,MBMCPreInfBy, MBMCUndIllDate, MBMCUndIllCAT , MBMCUndInfBy "
    strsqlMBMCovTwo = strsqlMBMCovTwo & " From MBMCoveredPersonsTwo Where MBMCNumber ='" & Trim(txtMBMNumber) & "' And MBMCCoverID = '" & Trim(tempMBMCovID) & "'"
    
    MBMCovTwo.Open strsqlMBMCovTwo, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If MBMCovTwo.recordCount > 0 Then
        With MBMCovTwo
            !MBMCPreIllDate = IIf(IsDate(txtPreDate), txtPreDate, "")
            !MBMCPreIllCAT = Mid(cboPreCategories, 1, 1)
            !MBMCPreInfBy = ChkStr(txtPreInfBy)
            !MBMCUndIllDate = IIf(IsDate(txtUndDate), txtUndDate, "")
            !MBMCUndIllCAT = Mid(cboUndCategories, 1, 1)
            !MBMCUndInfBy = ChkStr(txtUndInfBy)
        End With
    
        MBMCovTwo.Update
    End If
    MBMCovTwo.Close
    Set MBMCovTwo = Nothing
End Sub

Private Sub txtPreDate_LostFocus()
    If IsDate(txtPreDate) Then
    Else
        If txtPreDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPreDate.SetFocus
        End If
    End If
End Sub


Private Sub txtUndDate_LostFocus()
   If IsDate(txtUndDate) Then
    Else
        If txtUndDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtUndDate.SetFocus
        End If
    End If
End Sub
