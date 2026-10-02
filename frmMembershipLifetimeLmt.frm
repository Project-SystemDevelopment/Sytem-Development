VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMembershipLifetimeLmt 
   Caption         =   "Lifetime Limit Upgrade"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   11655
      Begin VB.TextBox txtCoverID 
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
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   1
         Top             =   200
         Width           =   495
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Go"
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
         Left            =   3300
         TabIndex        =   2
         Top             =   180
         Width           =   600
      End
      Begin VB.TextBox txtMBMNo 
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
         Left            =   1080
         TabIndex        =   0
         Top             =   200
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Lifetime Lmt"
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
         Index           =   3
         Left            =   8760
         TabIndex        =   18
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblBalLT 
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
         Left            =   9720
         TabIndex        =   17
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblLTStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3960
         TabIndex        =   16
         Top             =   210
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Member No"
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
         Left            =   180
         TabIndex        =   5
         Top             =   210
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6375
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   11655
      Begin MSComctlLib.ListView lstMBMListing 
         Height          =   1695
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   2990
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Member Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cover ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "MBM Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "IC No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Policy No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Pol Eff Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Pol Exp Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Pol Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Data Status"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lstCaseListing 
         Height          =   1815
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   3201
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MBM Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Case Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Patient ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Patient Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Admission Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Discharge Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Claimability"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstFSYRCase 
         Height          =   1815
         Left            =   120
         TabIndex        =   14
         Top             =   4440
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   3201
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MBM Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Case Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Patient ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Patient Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Admission Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Discharge Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Claimability"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "1st and 2nd Year Case Listing"
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
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   4200
         Width           =   2220
      End
      Begin VB.Label Label2 
         Caption         =   "Membership Listing"
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
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   2220
      End
      Begin VB.Label Label2 
         Caption         =   "Case Listing"
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
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   2220
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   7200
      Width           =   11655
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
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
         Left            =   10560
         TabIndex        =   8
         Top             =   200
         Width           =   960
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
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
         Left            =   9600
         TabIndex        =   3
         Top             =   200
         Width           =   960
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Lifetime Limit Upgrade"
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
      TabIndex        =   6
      Top             =   0
      Width           =   13935
   End
End
Attribute VB_Name = "frmMembershipLifetimeLmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpTopupNumber As String
Dim tmpMBMNumber1 As String
Dim tmpMBMNumber2 As String
Dim tmpMBMNumber3 As String
Dim tmpMBMBalLmt1 As Double
Dim tmpMBMBalLmt2 As Double
Dim tmpMBMBalLmt3 As Double
Dim CASStatus1 As Boolean
Dim CASStatus2 As Boolean
Dim LTStatus As Boolean
Dim cnt As Integer
Dim Claimscnt As Integer
Dim Claimscnt2 As Integer
Dim tmpPlnCode As String
Dim tmpPayCode As String
Dim tmpINSCode As String

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdGo_Click()
    
    If Trim(txtMBMNo) = "" Then
        MsgBox "Please select a Member Number ! ", vbCritical
        txtMBMNo.SetFocus
    ElseIf txtCoverID = "" Then
        MsgBox "Please select a Cover ID ! ", vbCritical
        txtMBMNo.SetFocus

    Else
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        Screen.MousePointer = vbHourglass
        Call MBMInquiryList
        Screen.MousePointer = vbDefault
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
       ' medixmasterconnMaint.Close
      '  Set medixmasterconnMaint = Nothing
    End If
End Sub

Private Sub cmdUpdate_Click()
On Err GoTo ErrorSave
Dim strsql As String
Dim CompTrans As Boolean
Dim strsqlMBMLT As String
Dim strsqlMBMOthers As String
Dim strsqlMBMCov As String
Dim strsqlCovUpd As String
Dim MBMSupp As New ADODB.Recordset
Dim tmptotalLT As Double
Dim ann As New ADODB.Recordset
Dim StrSqlAnn As String

    If Trim(cnt) < 3 Then
        MsgBox "This Member is not qualified for Lifetime Limit! ", vbCritical
    'ElseIf CASStatus1 = True Then
    '    MsgBox "This Member has claims history! ", vbCritical
    ElseIf LTStatus = True Then
         MsgBox "Lifetime Limit has been allocated for this member! ", vbCritical
    Else
        Screen.MousePointer = vbHourglass
        
        tmptotalLT = 0
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        medixmasterconn.beginTrans
        StrSqlAnn = " Select AnnLimit From AnnualLimit Where PLNCode = '" & Trim(tmpPlnCode) & "' And PayCode = '" & Trim(tmpPayCode) & "'And INSCode = '" & Trim(tmpINSCode) & "'"
        ann.Open StrSqlAnn, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
            If ann.EOF = False Then
                tmptotalLT = ann!AnnLimit * 3
            End If
            
        ann.Close
        Set ann = Nothing
        CompTrans = False
            strsql = " update MBMOthers set " & _
                     " MBMGuaRenewal = '3' " & _
                     " Where MBMNumber = '" & Trim(txtMBMNo) & "'"
            medixmasterconn.Execute strsql
            
            strsqlMBMLT = ""
            If Trim(txtCoverID) = "00" Then
                strsqlMBMLT = "Insert Into MBMLifeTime(MBMNumber, MBMCovID,MBMLifeTimeLmt) " & _
                            "Values('" & tmpTopupNumber & "', '00'," & tmptotalLT & ")"
                medixmasterconn.Execute strsqlMBMLT
            Else
                strsqlMBMCov = ""
            'If Trim(tmpINSCode) = "F" Or Trim(tmpINSCode) = "H" Then
                strsqlMBMCov = "Select MBMCNumber, MBMCCoverID from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNo) & "' aND MBMCCoverID = '" & Trim(txtCoverID) & "'"
                MBMSupp.Open strsqlMBMCov, medixmasterconn, adOpenKeyset, adLockOptimistic
                Do While Not MBMSupp.EOF
                    strsqlMBMLT = ""
                    strsqlMBMLT = "Insert Into MBMLifeTime(MBMNumber, MBMCovID,MBMLifeTimeLmt) " & _
                    "Values('" & tmpTopupNumber & "', '" & MBMSupp!MBMCCoverID & "'," & tmptotalLT & ")"
                    medixmasterconn.Execute strsqlMBMLT
                    MBMSupp.MoveNext
                Loop
            End If
            
            'medixmasterconn.Execute strsqlMBMLT
            
        medixmasterconn.commitTrans
        CompTrans = True
        
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        MsgBox "Record is Updated! "
        lblLTStatus = "Lifetime Limit has been allocated for this member"
        lblBalLT = tmptotalLT
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
ErrorSave:
        Screen.MousePointer = vbDefault
        If CompTrans = False Then
            medixmasterconn.RollbackTrans
        End If
        'medixmasterconnWS.Close
        'Set medixmasterconnWS = Nothing
        If CompTrans = False Then
            MsgBox Err.Description
        End If
    

End Sub

Private Sub CmdNext_Click()

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

Private Sub MBMInquiryList()
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim CASRecord As ListItem
Dim sqlstr As String
Dim strCriteria As String
Dim cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim rstMBMRef As New ADODB.Recordset
Dim sqlMBMRef As String
Dim strsqlCAS As String
Dim CAS As New ADODB.Recordset
Dim strsqlMBMLT As String
Dim MBMLT As New ADODB.Recordset
Dim tmpMBMNumber As String
    lblBalLT = 0
    cnt = 0
    tmpTopupNumber = ""
    LTStatus = False
    lstMBMListing.ListItems.Clear
    lstCaseListing.ListItems.Clear
    lstFSYRCase.ListItems.Clear
    CASStatus1 = False
    CASStatus2 = False
    Claimscnt = 0
    Claimscnt2 = 0
    tmpPlnCode = ""
    tmpPayCode = ""
    tmpINSCode = ""
    sqlMBMRef = " Select MBMNumber, MBMTopupNumber From MBMCrossReference Where MBMNumber = '" & Trim(txtMBMNo) & "'"
    rstMBMRef.Open sqlMBMRef, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
        Do While Not rstMBMRef.EOF
            tmpTopupNumber = IIf(Len(rstMBMRef!MBMTopupNumber), rstMBMRef!MBMTopupNumber, "")
            rstMBMRef.MoveNext
        Loop
    rstMBMRef.Close
    Set rstMBMRef = Nothing
    
    If Len(tmpTopupNumber) Then
        strsqlMBMLT = " Select MBMNumber,MBMLifeTimeLmt From MBMLifeTime Where MBMNumber = '" & Trim(tmpTopupNumber) & "' And MBMCovID = '" & txtCoverID & "'"
        MBMLT.Open strsqlMBMLT, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
            
                Do While Not MBMLT.EOF
                    If LTStatus = False Then
                        LTStatus = True
                        lblBalLT = MBMLT!MBMLifeTimeLmt
                    End If
                MBMLT.MoveNext

                Loop
            
        MBMLT.Close
        Set MBMLT = Nothing
        
        If LTStatus = True Then
            lblLTStatus = "Lifetime Limit has been allocated for this member"
        Else
            lblLTStatus = "Lifetime Limit has not been allocated for this member"
        End If
    End If
    If tmpTopupNumber <> "" Then
        'If Trim(txtCoverID) = "00" Then
            strCriteria = ""
            strCriteria = strCriteria & " AND MBMTopupNumber =  '" & RemStr(tmpTopupNumber) & "' order by MBMPayorEffDate"
        
            Set cmd.ActiveConnection = medixmasterconn
            cmd.CommandText = "Search_MBM"
            cmd.CommandType = adCmdStoredProc
            cmd.CommandTimeout = 300
            Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strCriteria)
            cmd.Parameters.Append Prm1
            rst2.CursorLocation = adUseClient
            rst2.CursorType = adOpenForwardOnly
            rst2.CacheSize = 100
            Set rst2 = cmd.Execute
        'Else
        '    strCriteria = ""
        '    strCriteria = strCriteria & " AND MBMTopupNumber =  '" & RemStr(tmpTopupNumber) & "' AND MBMCCoverID =  '" & RemStr(txtCoverID) & "' order by MBMPayorEffDate"
       '
       '     Set cmd.ActiveConnection = medixmasterconn
        '    cmd.CommandText = "[Search_MBM_Supp]"
        '    cmd.CommandType = adCmdStoredProc
       '     cmd.CommandTimeout = 300
       '     Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strCriteria)
        '    cmd.Parameters.Append Prm1
       '     rst2.CursorLocation = adUseClient
       '     rst2.CursorType = adOpenForwardOnly
       '     rst2.CacheSize = 100
       '     Set rst2 = cmd.Execute
       ' End If
    
        Do Until rst2.EOF
            With rst2
                cnt = cnt + 1
                'If Cnt = 1 Then
                '    tmpMBMNumber1 = !MBMNumber
                '    tmpMBMBalLmt1 = !MBMAvailableLimit
                'End If
                
                'If Cnt = 2 Then
                '    tmpMBMNumber2 = !MBMNumber
                '    tmpMBMBalLmt2 = !MBMAvailableLimit
                'End If
                
                'If Cnt = 3 Then
                '    tmpMBMNumber3 = !MBMNumber
                '    tmpMBMBalLmt3 = !MBMAvailableLimit
                'End If
                Set Record = lstMBMListing.ListItems.Add(, , !MBMNumber)
                Record.SubItems(1) = IIf(Len(!MBMCOVID), !MBMCOVID, "")
                Record.SubItems(2) = IIf(Len(!MBMName), Trim(!MBMName), "")
                Record.SubItems(3) = IIf(Len(!MBMIcBcPp), Trim(!MBMIcBcPp), "")
                Record.SubItems(4) = IIf(Len(!MBMPolicyNo), Trim(!MBMPolicyNo), "")
                If Len(!MBMPayorEffDate) Then
                    Record.SubItems(5) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                End If
                If Len(!MBMPayorEXPDate) Then
                    Record.SubItems(6) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                End If
                Record.SubItems(7) = IIf(Len(!MBMStatus), Trim(!MBMStatus), "")
                Record.SubItems(8) = IIf(Len(!MBMDataStatus), Trim(!MBMDataStatus), "")
                tmpPlnCode = !PLNCode
                tmpPayCode = !PAYCode
                tmpINSCode = !INSCode
                
                strsqlCAS = ""
                If UCase(Trim(txtMBMNo)) <> UCase(Trim(!MBMNumber)) Then
                    Claimscnt2 = CDbl(Claimscnt2) + 1
                    strsqlCAS = " Select MBMNumber, CasNumber, MBMPatientCovID, MBMTopupNumber, " & _
                    "CASDischargeDate,CASDateAdmitted,MBMName,MBMCName,CASClaimability " & _
                    "From View_Case_Search Where MBMNumber = '" & Trim(!MBMNumber) & "' And MBMPatientCovID = '" & txtCoverID & "'"
                    CAS.Open strsqlCAS, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                    
                    If CAS.EOF = False Then
                        Claimscnt = CDbl(Claimscnt) + 1
                        CASStatus1 = True
                        
                        Do While Not CAS.EOF
                            With CAS
                                Set CASRecord = lstFSYRCase.ListItems.Add(, , !MBMNumber)
                                CASRecord.SubItems(1) = IIf(Len(!CasNumber), !CasNumber, "")
                                CASRecord.SubItems(2) = IIf(Len(!MBMPatientCovID), !MBMPatientCovID, "")
                                If !MBMPatientCovID = "00" Then
                                    CASRecord.SubItems(3) = IIf(Len(!MBMName), Trim(!MBMName), "")
                                Else
                                    CASRecord.SubItems(3) = IIf(Len(!MBMCName), Trim(!MBMCName), "")
                                End If
                                CASRecord.SubItems(4) = IIf(Len(!CASDateAdmitted), Format(!CASDateAdmitted, "dd/mm/yyyy"), "")
                                CASRecord.SubItems(5) = IIf(Len(!CASDischargeDate), Format(!CASDischargeDate, "dd/mm/yyyy"), "")
                                If Len(!CASClaimability) Then
                                    CASRecord.SubItems(6) = !CASClaimability
                                End If
                            End With
                        CAS.MoveNext
                        Loop
                    Else
                        'If CDbl(Claimscnt) = 1 Then
                        '    Claimscnt2 = Claimscnt2 + 1
                        'End If
                        If CDbl(Claimscnt) > 0 Then
                            Claimscnt = CDbl(Claimscnt) - 1
                            'CASStatus1 = True
                            'Claimscnt = Claimscnt + 1
                        Else
                            Claimscnt = 0
                        '    CASStatus1 = False
                        End If
                    End If
                    CAS.Close
                    Set CAS = Nothing
                Else
                
                      strsqlCAS = ""
                      strsqlCAS = " Select MBMNumber, CasNumber, MBMPatientCovID, MBMTopupNumber, " & _
                      "CASDischargeDate,CASDateAdmitted,MBMName,MBMCName,CASClaimability " & _
                      "From View_Case_Search Where MBMNumber = '" & Trim(txtMBMNo) & "'And MBMPatientCovID = '" & txtCoverID & "'"
                      CAS.Open strsqlCAS, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
                    
                      Do While Not CAS.EOF
                          With CAS
                              CASStatus1 = True
                              Set CASRecord = lstCaseListing.ListItems.Add(, , !MBMNumber)
                              CASRecord.SubItems(1) = IIf(Len(!CasNumber), !CasNumber, "")
                              CASRecord.SubItems(2) = IIf(Len(!MBMPatientCovID), !MBMPatientCovID, "")
                              If !MBMPatientCovID = "00" Then
                                  CASRecord.SubItems(3) = IIf(Len(!MBMName), Trim(!MBMName), "")
                              Else
                                  CASRecord.SubItems(3) = IIf(Len(!MBMCName), Trim(!MBMCName), "")
                              End If
                              CASRecord.SubItems(4) = IIf(Len(!CASDateAdmitted), Format(!CASDateAdmitted, "dd/mm/yyyy"), "")
                              CASRecord.SubItems(5) = IIf(Len(!CASDischargeDate), Format(!CASDischargeDate, "dd/mm/yyyy"), "")
                              If Len(!CASClaimability) Then
                                  CASRecord.SubItems(6) = !CASClaimability
                              End If
                          End With
                      CAS.MoveNext
                      Loop
                      CAS.Close
                      Set CAS = Nothing
                
                End If
            End With
        rst2.MoveNext
        Loop
        rst2.Close
        Set rst2 = Nothing
        
        If CDbl(Claimscnt2) + 1 = cnt And Claimscnt = 0 Then
            CASStatus1 = False
        Else
            CASStatus1 = True
        End If
        'If tmpMBMNumber1 <> "" Then
        'End If
        
        'If tmpMBMNumber2 <> "" Then
        '    strsqlCAS = ""
        '    strsqlCAS = " Select MBMNumber, CasNumber, MBMPatientCovID, MBMTopupNumber, " & _
        '    "CASDischargeDate,CASDateAdmitted,MBMName,MBMCName,CASClaimability " & _
        '    "From View_Case_Search Where MBMNumber = '" & Trim(tmpMBMNumber2) & "'"
        '    CAS.Open strsqlCAS, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
           
        '    Do While Not CAS.EOF
        '        With CAS
        '            CASStatus2 = True
        '            Set CASRecord = lstFSYRCase.ListItems.Add(, , !MBMNumber)
        '            CASRecord.SubItems(1) = IIf(Len(!CasNumber), !CasNumber, "")
        '            CASRecord.SubItems(2) = IIf(Len(!MBMPatientCovID), !MBMPatientCovID, "")
        '            If !MBMPatientCovID = "00" Then
        '                CASRecord.SubItems(3) = IIf(Len(!MBMName), Trim(!MBMName), "")
        '            Else
        '                CASRecord.SubItems(3) = IIf(Len(!MBMCName), Trim(!MBMCName), "")
        '            End If
        '            CASRecord.SubItems(4) = IIf(Len(!CASDateAdmitted), Format(!CASDateAdmitted, "dd/mm/yyyy"), "")
        '            CASRecord.SubItems(5) = IIf(Len(!CASDischargeDate), Format(!CASDischargeDate, "dd/mm/yyyy"), "")
        '            If Len(!CASClaimability) Then
        '                CASRecord.SubItems(6) = !CASClaimability
        '            End If
        '        End With
        '    CAS.MoveNext
        '    Loop
        '    CAS.Close
        '    Set CAS = Nothing
        'End If
        
    End If
    
    If tmpTopupNumber = "" Then
        MsgBox "No Record Found ! ", vbCritical
    End If
End Sub







