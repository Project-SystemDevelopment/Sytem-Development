VERSION 5.00
Begin VB.Form FrmITMaintainMBMDel 
   Caption         =   "Membership Maintenance - Deletion (IT Dept Only)"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      Width           =   6495
      Begin VB.TextBox txtTotCas 
         Height          =   285
         Left            =   3600
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtCISCas 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   255
         Left            =   4800
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtHISCas 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.Frame FrameCmd 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4080
         TabIndex        =   24
         Top             =   120
         Width           =   735
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   255
         Left            =   5520
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Clinical Case(s)"
         Height          =   255
         Left            =   1800
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "HIS Case(s)"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame FrameInfo 
      Enabled         =   0   'False
      Height          =   2175
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   6495
      Begin VB.TextBox txtPolStatus 
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   4920
         TabIndex        =   33
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtBatchNo 
         Height          =   285
         Left            =   4920
         TabIndex        =   22
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtBordxDate 
         Height          =   285
         Left            =   1320
         TabIndex        =   21
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtExpDate 
         Height          =   285
         Left            =   4920
         TabIndex        =   20
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtEffDate 
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtDOB 
         Height          =   285
         Left            =   4920
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TxtICNo 
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox TxtPolicyNo 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox TxtMBMName 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label13 
         Caption         =   "Policy Status"
         Height          =   255
         Left            =   3840
         TabIndex        =   32
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Date of Birth"
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Batch No"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Exp Date"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Eff Date"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Policy No"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "IC No"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Member Name"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      Begin VB.TextBox txtRenew 
         Height          =   285
         Left            =   5760
         TabIndex        =   34
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "&Go"
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtMBMNo 
         Height          =   285
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label14 
         Caption         =   "Latest MBMNo"
         Height          =   255
         Left            =   4560
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Membership No"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Membership Maintenance - Deletion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "FrmITMaintainMBMDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrPassword As String
Dim RenewKey As String
Dim HisFound As Boolean

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    txtMBMNo.Enabled = True
    txtMBMNo.SetFocus
End Sub

Private Sub cmdDelete_Click()
On Err GoTo ErrorSave
Dim startTrans As Boolean
Dim ASKMsg
Dim MBMSql As String
Dim SuppSql As String
Dim DelSql As String
Dim MBMRec As New ADODB.Recordset

    MBMSql = "WHERE MBMNumber='" & Trim(txtMBMNo) & "'"
    SuppSql = "WHERE MBMCNumber='" & Trim(txtMBMNo) & "'"
    
    ASKMsg = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
    If ASKMsg = vbYes Then
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        Screen.MousePointer = vbHourglass
' ----- MBMCrossReference table
        medixmasterconnMaint.beginTrans
        medixmasterconn.beginTrans
        DelSql = "Delete From MBMCrossReference " & MBMSql
        MBMRec.Open DelSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
' ----- MBMLifeTime table
        DelSql = "Delete From MBMLifeTime " & MBMSql
        MBMRec.Open DelSql, medixmasterconn, adOpenKeyset, adLockOptimistic
' ----- MBMBnfLmt
        DelSql = "Delete From MBMBnfLmt " & MBMSql
        MBMRec.Open DelSql, medixmasterconn, adOpenKeyset, adLockOptimistic
' ----- MBM table
        DelSql = ""
        DelSql = "Delete From MBM " & MBMSql
        Call DeleteRec(DelSql)
' ----- MBMTwo table
        DelSql = ""
        DelSql = "Delete From MBMTwo " & MBMSql
        Call DeleteRec(DelSql)
' ----- MBMOthers table
        DelSql = ""
        DelSql = "Delete From MBMOthers " & MBMSql
        Call DeleteRec(DelSql)
' ----- MBMOthersTwo table
        DelSql = ""
        DelSql = "Delete From MBMOthersTWo " & MBMSql
        Call DeleteRec(DelSql)
' ----- MBMCLNOthers table
        DelSql = ""
        DelSql = "Delete From MBMCLNOthers " & MBMSql
        Call DeleteRec(DelSql)
' ----- MBMCoveredPersons table
        DelSql = ""
        DelSql = "Delete From MBMCoveredPersons " & SuppSql
        Call DeleteRec(DelSql)
' ----- MBMCoveredPersonsTwo table
        DelSql = ""
        DelSql = "Delete From MBMCoveredPersonsTwo " & SuppSql
        Call DeleteRec(DelSql)
' ----- MBMCoveredPersonsCLN
        DelSql = ""
        DelSql = "Delete From MBMCoveredPersonsCLN " & SuppSql
        Call DeleteRec(DelSql)
        Call Resequence
        MsgBox "Record Deleted....", vbOKOnly + vbCritical, DialogBoxTitle
        
        Call ClearForm(Me)
        FrameCmd.Enabled = False
        txtMBMNo.Enabled = True
        txtMBMNo.SetFocus
        medixmasterconnMaint.CommitTrans
        medixmasterconn.CommitTrans
        startTrans = False
        medixmasterconn.Close
        Set medixmasterconn = Nothing
        medixmasterconnMaint.Close
        Set medixmasterconnMaint = Nothing

        Screen.MousePointer = vbDefault
        Exit Sub
    

ErrorSave:
        medixmasterconn.Close
        Set medixmasterconn = Nothing
        medixmasterconnMaint.Close
        Set medixmasterconnMaint = Nothing
        Screen.MousePointer = vbDefault
        If startTrans = True Then
            medixmasterconnMaint.RollbackTrans
            medixmasterconn.RollbackTrans
        End If
        
        MsgBox Err.Description
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdGo_Click()

    CmdGo.Enabled = False
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    Call InfoDet
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
    CmdGo.Enabled = True
    
    If Trim(TxtMBMName) <> "" Then
        txtMBMNo.Enabled = False
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

Private Sub txtMBMNo_GotFocus()
    CmdGo.Default = True
End Sub

Private Sub txtMBMNo_LostFocus()
    txtMBMNo = ChkStr(txtMBMNo)
    CmdGo.Default = False
End Sub

Private Sub DeleteRec(strsql As String)
Dim MBMRec As New ADODB.Recordset

    MBMRec.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
End Sub

Private Sub InfoDet()
Dim MBMXrefDB As New ADODB.Recordset
Dim XMBMSql As String
Dim RewSql As String
Dim RewREc As New ADODB.Recordset
Dim tmpPayCode, tmpMBMSeq As String
Dim tmpSeqID As String
Dim RewDel As Boolean
Dim TotalCAS As Integer
Dim FindStr As String
Dim StrPos As Integer
Dim tmpIns As String
Dim tmpMBMNo As String
    FrameCmd.Enabled = False
    TotalCAS = 0
    If Trim(txtMBMNo) = "" Then
        MsgBox "Please Keyin Membership No. ! ", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMNo.SetFocus
    Else
        RenewKey = ""
        XMBMSql = "SELECT MBMNAME, MBMPolicyNo, INSCode, MBMICBcPp, MBMPayorEffDate, MBMPayorExpDate, MBMDOB, " & _
               " MBMBordxDate, MBMBatchNO, MBMStatus FROM MBMCrossReference WHERE MBMNumber='" & Trim(txtMBMNo) & "'"
        MBMXrefDB.Open XMBMSql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
        If MBMXrefDB.EOF = True Then
            MsgBox "No Record Found ! ", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMNo.SetFocus
        Else
            Do While Not MBMXrefDB.EOF
                FrameCmd.Enabled = True
                TxtMBMName = Trim(MBMXrefDB!MBMName)
                TxtPolicyNo = IIf(Len(Trim(MBMXrefDB!MBMPolicyNo)), Trim(MBMXrefDB!MBMPolicyNo), "")
                TxtICNo = Trim(MBMXrefDB!MBMIcBcPp)
                txtEffDate = IIf(Len(MBMXrefDB!MBMPayorEffDate), MBMXrefDB!MBMPayorEffDate, "")
                txtExpDate = IIf(Len(MBMXrefDB!MBMPayorEXPDate), MBMXrefDB!MBMPayorEXPDate, "")
                txtDOB = MBMXrefDB!MBMDOB
                txtBordxDate = MBMXrefDB!MBMBordxDate
                txtBatchNo = IIf(Len(MBMXrefDB!MBMBatchNo), Trim(MBMXrefDB!MBMBatchNo), "")
                txtPolStatus = IIf(Len(MBMXrefDB!MBMStatus), Trim(MBMXrefDB!MBMStatus), "")
                tmpIns = IIf(Len(MBMXrefDB!INSCode), Trim(MBMXrefDB!INSCode), "")
                RenewKey = Mid(txtMBMNo, 1, 2) & "*" & tmpIns & "*" & Format(MBMXrefDB!MBMDOB, "yyyymmdd") & "*" & Trim(MBMXrefDB!MBMName)
                MBMXrefDB.MoveNext
            Loop
        End If
        MBMXrefDB.Close
        Set MBMXrefDB = Nothing
        Call FindCase(TotalCAS)
        HisFound = False
        RewDel = True
        StrPos = 0
        FindStr = "*"
        StrPos = InStr(1, txtMBMNo, FindStr)
        
        If StrPos <> 0 Then
            RewDel = False
            tmpPayCode = Mid(txtMBMNo, 1, 2)
            tmpMBMNo = Mid(txtMBMNo, 1, StrPos - 1)
            tmpMBMSeq = Right(tmpMBMNo, 6)
            tmpSeqID = Right(txtMBMNo, 2)
            If tmpSeqID = "01" Then
                Call FindHistory
            End If
            RewSql = "SELECT RewSeqNo, RewSEQID From MBMSEQRenew05 where substring(RewMBMSEQKey,1,2)='" & tmpPayCode & _
                     "' AND RewSEQNo = '" & tmpMBMSeq & "' AND RewSEQID='" & tmpSeqID & "'"
            RewREc.Open RewSql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
            Do While Not RewREc.EOF
                RewDel = True
                RewREc.MoveNext
            Loop
        End If
        txtRenew = RewDel
        If HisFound = True Then
            MsgBox "Please check on Member Renewal History ! ", vbCritical
        End If
        If TotalCAS > 0 Or txtPolStatus <> "D" Or RewDel = False Or HisFound = True Then
            FrameCmd.Enabled = False
        End If
    End If
End Sub

Private Sub Resequence()
Dim RewSql As String
Dim RewREc As New ADODB.Recordset
Dim tmpPayCode, tmpMBMSeq As String
Dim tmpSeqID, NewSeqID, CurrSeqID As String
Dim Seq05Rec As New ADODB.Recordset
Dim SeqSQl As String
Dim CurrMBMSeq As String
Dim CurrChar, NewChar As String
Dim CurrSeqNo As String
Dim NewMBMSeq As String
Dim NotRenew As Boolean
Dim tmpINSCode As String
    If Len(txtMBMNo) = 11 Or Len(txtMBMNo) = 12 Then
        tmpPayCode = Mid(txtMBMNo, 1, 2)
        If Len(txtMBMNo) = 11 Then
            tmpINSCode = ""
            tmpMBMSeq = Mid(txtMBMNo, 3, 6)
            tmpSeqID = Mid(txtMBMNo, 10, 2)
            RewSql = "SELECT RewSEQID  From MBMSEQRenew05 where substring(RewMBMSEQKey,1,2)='" & tmpPayCode & _
                 "' AND RewSEQNo = '" & tmpMBMSeq & "' AND RewSEQID='" & tmpSeqID & "'"
        Else
            tmpINSCode = Mid(txtMBMNo, 3, 1)
            tmpMBMSeq = Mid(txtMBMNo, 4, 6)
            tmpSeqID = Mid(txtMBMNo, 11, 2)
            RewSql = "SELECT RewSEQID  From MBMSEQRenew05 where substring(RewMBMSEQKey,1,2)='" & tmpPayCode & _
                 "' AND INSCode = '" & tmpINSCode & _
                 "' AND RewSEQNo = '" & tmpMBMSeq & "' AND RewSEQID='" & tmpSeqID & "'"
        End If
        
        RewREc.Open RewSql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        NotRenew = False
        Do While Not RewREc.EOF
            CurrSeqID = RewREc!RewSEQID
            NewSeqID = Format(RewREc!RewSEQID - 1, "00")
            If NewSeqID = "00" Then
                If HisFound = False Then
                    RewREc.Delete
                    NotRenew = True
                End If
                
            Else
                RewREc!RewSEQID = Format(NewSeqID, "00")
                RewREc.Update
            End If
            RewREc.MoveNext
        Loop
        RewREc.Close
        Set RewREc = Nothing
        If NotRenew = True Then
            SeqSQl = "Select PAYCode, SEQNumber From MBMSEQ05 WHERE PAYCode ='" & tmpPayCode & _
                     "' AND SEQNumber='" & tmpMBMSeq & "'"
            Seq05Rec.Open SeqSQl, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            Do While Not Seq05Rec.EOF
                CurrMBMSeq = Format(Seq05Rec!SEQNumber, "00000")
                If tmpMBMSeq = "A00001" Then
                    Seq05Rec.Delete
                Else
                    CurrChar = Mid(CurrMBMSeq, 1, 1)
                    CurrSeqNo = Mid(CurrMBMSeq, 2, 5)
                    NewMBMSeq = Format(CurrSeqNo - 1, "00000")
                    If NewMBMSeq = "00000" Then
                        CurrChar = Asc(CurrChar) - 1
                        CurrChar = Chr(CurrChar)
                        NewMBMSeq = "99999"
                    End If
                    NewMBMSeq = CurrChar & Format(NewMBMSeq, "00000")
                    Seq05Rec!SEQNumber = NewMBMSeq
                    Seq05Rec.Update
                End If
                Seq05Rec.MoveNext
            Loop
            Seq05Rec.Close
            Set Seq05Rec = Nothing
        End If
    End If
End Sub

Private Sub FindCase(tmpTotCas As Integer)
Dim CAsREc As New ADODB.Recordset
Dim CISCas As New ADODB.Recordset
Dim cmd1 As New ADODB.Command
Dim cmd2 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
    
    txtHISCas = 0
    txtCISCas = 0
    tmpTotCas = 0
    Set cmd1.ActiveConnection = medixmasterconnMaint
    cmd1.CommandText = "CntHisCasebyMBMNo"
    cmd1.CommandType = adCmdStoredProc
    cmd1.CommandTimeout = 300
    Set Prm1 = cmd1.CreateParameter("tmpMBMNo", adVarChar, adParamInput, 20, Trim(txtMBMNo))
    cmd1.Parameters.Append Prm1
    CAsREc.CursorLocation = adUseClient
    CAsREc.CursorType = adOpenForwardOnly
    CAsREc.CacheSize = 100
    Set CAsREc = cmd1.Execute
    Do While Not CAsREc.EOF
        txtHISCas = CAsREc!HISCAseCnt
        tmpTotCas = tmpTotCas + CAsREc!HISCAseCnt
        CAsREc.MoveNext
    Loop
    CAsREc.Close
    Set CAsREc = Nothing
    
    MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"

    Set cmd2.ActiveConnection = MediConnCISDB
    cmd2.CommandText = "CntCisCasebyMBMNo"
    cmd2.CommandType = adCmdStoredProc
    cmd2.CommandTimeout = 300
    Set Prm2 = cmd2.CreateParameter("tmpMBMNo", adVarChar, adParamInput, 20, Trim(txtMBMNo))
    cmd2.Parameters.Append Prm2
    CISCas.CursorLocation = adUseClient
    CISCas.CursorType = adOpenForwardOnly
    CISCas.CacheSize = 100
    Set CISCas = cmd2.Execute
    Do While Not CISCas.EOF
        txtCISCas = CISCas!CISCAseCnt
        tmpTotCas = tmpTotCas + CISCas!CISCAseCnt
        CISCas.MoveNext
    Loop
    CISCas.Close
    Set CISCas = Nothing
    MediConnCISDB.Close
    Set MediConnCISDB = Nothing
    txtTotCas = tmpTotCas
End Sub

Private Sub FindHistory()
Dim HisRec As New ADODB.Recordset
Dim HisSql As String
    HisFound = False
    HisSql = "SELECT SeqNo from MBMSEQRenewHistory Where MBMKey='" & RenewKey & "'"
    HisRec.Open HisSql, medixmasterconnMaint, adOpenForwardOnly, adLockReadOnly
    Do While Not HisRec.EOF
        HisFound = True
        HisRec.MoveNext
        Exit Do
    Loop
    HisRec.Close
    Set HisRec = Nothing
End Sub



