VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmClaimNotification 
   Caption         =   "Claim Notification"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4140
   ScaleWidth      =   7710
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   7545
      Begin VB.TextBox txtCASNumber 
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
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   0
         Top             =   195
         Width           =   2175
      End
      Begin VB.CommandButton cmdShow 
         Appearance      =   0  'Flat
         Caption         =   "&Go"
         Height          =   285
         Left            =   3600
         TabIndex        =   1
         Top             =   195
         Width           =   645
      End
      Begin VB.Label Label19 
         Caption         =   "Case Number"
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
         Top             =   225
         Width           =   1005
      End
   End
   Begin VB.Frame Frame4 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   7545
      Begin VB.TextBox txtMBMNo 
         Height          =   285
         Left            =   1620
         TabIndex        =   21
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox txtPolicyNo 
         Height          =   285
         Left            =   4455
         TabIndex        =   20
         Top             =   200
         Width           =   2955
      End
      Begin VB.TextBox txtPrincipleName 
         Height          =   285
         Left            =   4455
         TabIndex        =   18
         Top             =   450
         Width           =   2955
      End
      Begin MSMask.MaskEdBox txtAdmissionDate 
         Height          =   285
         Left            =   1620
         TabIndex        =   23
         Top             =   480
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDischargeDate 
         Height          =   285
         Left            =   1620
         TabIndex        =   22
         Top             =   720
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPatientName 
         Height          =   285
         Left            =   4455
         TabIndex        =   19
         Top             =   710
         Width           =   2955
      End
      Begin VB.Label Label1 
         Caption         =   "Membership No."
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
         TabIndex        =   29
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Policy No."
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
         Left            =   3375
         TabIndex        =   28
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Principle Name"
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
         Left            =   3360
         TabIndex        =   27
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label4 
         Caption         =   "Patient Name"
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
         Left            =   3360
         TabIndex        =   26
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label5 
         Caption         =   "Date of Admission"
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
         Left            =   135
         TabIndex        =   25
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label6 
         Caption         =   "Date of Discharge"
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
         Left            =   135
         TabIndex        =   24
         Top             =   720
         Width           =   1395
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Actions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   7545
      Begin VB.TextBox txtInv 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3315
         MaxLength       =   50
         TabIndex        =   6
         Top             =   855
         Width           =   4095
      End
      Begin VB.TextBox txtDoc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2235
         MaxLength       =   50
         TabIndex        =   4
         Top             =   525
         Width           =   5175
      End
      Begin VB.CheckBox chkUndProc 
         Caption         =   "Under Process"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   1440
      End
      Begin VB.CheckBox chkDoc 
         Caption         =   "Awaiting Document From"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   90
         TabIndex        =   3
         Top             =   525
         Width           =   2160
      End
      Begin VB.TextBox txtOther1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1035
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1200
         Width           =   6375
      End
      Begin VB.CheckBox chkInv 
         Caption         =   "Awaiting Reply to our investigation from"
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
         Left            =   90
         TabIndex        =   5
         Top             =   840
         Width           =   3240
      End
      Begin VB.CheckBox chkOther1 
         Caption         =   "Other 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   90
         TabIndex        =   7
         Top             =   1200
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   7545
      Begin VB.TextBox txtAdd3 
         Height          =   285
         Left            =   2880
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAdd2 
         Height          =   285
         Left            =   2280
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAdd1 
         Height          =   285
         Left            =   1680
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtPayMainContact 
         Height          =   285
         Left            =   840
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtPayName 
         Height          =   285
         Left            =   480
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtHOSName 
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdExit 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   6840
         TabIndex        =   14
         Top             =   240
         Width           =   645
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
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
         Left            =   6190
         TabIndex        =   13
         Top             =   240
         Width           =   645
      End
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         Caption         =   "&Save"
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
         Left            =   5550
         TabIndex        =   12
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Claim Notification"
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
      TabIndex        =   15
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmClaimNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Letter As String
Dim objword As Word.Application
Dim objdoc As Word.Document

Private Sub chkDoc_Click()
    If chkDoc.Value = 0 Then
        txtDoc.Enabled = False
    Else
        txtDoc.Enabled = True
        txtDoc.SetFocus
    End If
End Sub

Private Sub chkDoc_GotFocus()
    If chkDoc.Value = 0 Then
        txtDoc.Enabled = False
    Else
        txtDoc.Enabled = True
    End If
End Sub

Private Sub chkInv_Click()
    If chkInv.Value = 0 Then
        txtInv.Enabled = False
    Else
        txtInv.Enabled = True
        txtInv.SetFocus
    End If
End Sub

Private Sub chkInv_GotFocus()
    If chkInv.Value = 0 Then
        txtInv.Enabled = False
    Else
        txtInv.Enabled = True
    End If
End Sub

Private Sub chkOther1_Click()
    If chkOther1.Value = 0 Then
        txtOther1.Enabled = False
    Else
        txtOther1.Enabled = True
        txtOther1.SetFocus
    End If
End Sub

Private Sub chkOther1_GotFocus()
    If chkOther1.Value = 0 Then
        txtOther1.Enabled = False
    Else
        txtOther1.Enabled = True
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Call funcword
End Sub

Private Sub CmdSave_Click()
    medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        Call SaveRecord
    medixmasterconnWS.Close
    Set medixmasterconnWS = Nothing
End Sub

Private Sub cmdShow_Click()
Dim CLM As New ADODB.Recordset
Dim strsqlCLM As String

    cmdSave.Enabled = True
    If txtCASNumber = "" Then
        MsgBox "Please key in the Case No.!", vbCritical
        txtCASNumber.SetFocus
    Else
        medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        
        Call CleanForm
        Call ShowMBMDetails
        Call ShowCLMNotification
        chkUndProc.SetFocus
        medixmasterconnWS.Close
        Set medixmasterconnWS = Nothing
    End If
End Sub

Private Sub ShowCLMNotification()
Dim strsql As String
Dim rst As New ADODB.Recordset
    
    strsql = "Select CasNumber,CLMUndProcess,CLMAwaitingDoc,CLMAwaitingDocRmk, " & _
            "CLMAwaitingReply,CLMAwaitingReplyRmk,CLMOthers,CLMOthersRmk,CLMRegUSR,CLMRegDate " & _
            "From CLMNotification where CASNumber = '" & Trim(txtCASNumber) & "'"
    rst.Open strsql, medixmasterconnWS, adOpenForwardOnly, adLockOptimistic
    
    If Not rst.EOF Then
        If rst!CLMUndProcess = "1" Then
            chkUndProc.Value = 1
        End If
        If rst!CLMAwaitingDoc = "1" Then
            chkDoc.Value = 1
            txtDoc = rst!CLMAwaitingDocRmk
        End If
        If rst!CLMAwaitingReply = "1" Then
            chkInv.Value = 1
            txtInv = rst!CLMAwaitingReplyRmk
        End If
        If rst!CLMOthers = "1" Then
            chkOther1.Value = 1
            txtOther1 = rst!CLMOthersRmk
        End If
    End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub SaveRecord()
Dim strsql As String
Dim rst As New ADODB.Recordset
    If txtCASNumber = "" Then
        MsgBox "Please key in the Case No.!", vbCritical
        txtCASNumber.SetFocus
    ElseIf chkUndProc.Value = 0 And chkDoc.Value = 0 And chkInv.Value = 0 And chkOther1.Value = 0 Then
        MsgBox "Please Select action to be taken", vbOKOnly
    ElseIf chkDoc.Value = "1" And txtDoc = "" Then
        MsgBox "Please Enter Remarks", vbOKOnly
    ElseIf chkInv.Value = "1" And txtInv = "" Then
        MsgBox "Please Enter Remarks", vbOKOnly
    ElseIf chkOther1.Value = "1" And txtOther1 = "" Then
        MsgBox "Please Enter Remarks", vbOKOnly
    Else
        strsql = "Select CasNumber,CLMUndProcess,CLMAwaitingDoc,CLMAwaitingDocRmk, " & _
                "CLMAwaitingReply,CLMAwaitingReplyRmk,CLMOthers,CLMOthersRmk,CLMRegUSR,CLMRegDate " & _
                "From CLMNotification where CASNumber = '" & Trim(txtCASNumber) & "'"
        rst.Open strsql, medixmasterconnWS, adOpenForwardOnly, adLockOptimistic
        
        If rst.EOF Then
            rst.AddNew
        End If
        
         rst!CasNumber = txtCASNumber
        If chkUndProc.Value = 1 Then
           rst!CLMUndProcess = "1"
        Else
            rst!CLMUndProcess = "0"
        End If
        If chkDoc.Value = 1 Then
            rst!CLMAwaitingDoc = "1"
            rst!CLMAwaitingDocRmk = txtDoc
        Else
            rst!CLMAwaitingDoc = "0"
            rst!CLMAwaitingDocRmk = ""
        End If
        If chkInv.Value = 1 Then
            rst!CLMAwaitingReply = "1"
            rst!CLMAwaitingReplyRmk = txtInv
        Else
            rst!CLMAwaitingReply = "0"
            rst!CLMAwaitingReplyRmk = ""
        End If
        If chkOther1.Value = 1 Then
            rst!CLMOthers = "1"
            rst!CLMOthersRmk = txtOther1
        Else
            rst!CLMOthers = "0"
            rst!CLMOthersRmk = ""
        End If
        rst.Update
        rst.Close
        Set rst = Nothing
        MsgBox "Record Saved Successfully", vbOKOnly
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
            Case vbKeyReturn
                cmdShow_Click
        End Select
    End If
End Sub

Private Sub ShowMBMDetails()
Dim MBM As New ADODB.Recordset
Dim strsql As String
    
    strsql = "Exec CLM_CLMReminder '" & Trim(txtCASNumber) & "',1"
    MBM.Open strsql, medixmasterconnWS, adOpenForwardOnly, adLockOptimistic
    
    If Not MBM.EOF Then
        With MBM
            txtMBMNo = !MBMNumber
            txtPolicyNo = !MBMPolicyNo
            txtPrincipleName = !MBMName
            If !MBMPatientCovID = "00" Then
                txtPatientName = !MBMName
            Else
                txtPatientName = !MBMCName
            End If
            txtAdmissionDate = Format(!CASDateAdmitted, "dd/mm/yyyy")
            If !CASDischargeDate <> "" Then
                txtDischargeDate = Format(!CASDischargeDate, "dd/mm/yyyy")
            Else
                txtDischargeDate = "__/__/____"
            End If
            If !MBMAddress1 <> "" Then
                txtAdd1 = !MBMAddress1
            End If
            If !MBMAddress2 <> "" Then
                txtAdd2 = !MBMAddress2
            End If
            'If !MBMAddress3 <> "" Then
                txtAdd3 = !MBMAddress3 & "" & !CTYCode & "" & !MBMPostCode & "" & !STACode
            'End If
            txtHOSName = IIf(Len(!HOSName), !HOSName, "")
            txtPayMainContact = !PAYMAINContact
            txtPayName = !PAYCompanyName
        End With
    Else
       MsgBox "No Record Found", vbOKOnly
    End If
    MBM.Close
    Set MBM = Nothing

End Sub

Private Sub funcword()
On Error GoTo openerror

    Set objword = Nothing
    If objword Is Nothing Then
        Set objword = New Word.Application
        objword.Visible = True
        Set objdoc = objword.Documents.Add(LocCardPath & "Claim\" & "CLMNotification.dot\")
        '"Template\LG\" & Trim(LGNumber) & ".doc"
        '(LocCardPath & "\Case\" & "\BTBG.DOT\")
    ElseIf Not objdoc Is Nothing Then
        If MsgBox("Previous document " & objdoc & " not closed. Press OK to save and close it now or cancel to abort", _
            vbExclamation + vbOKCancel, DialogBoxTitle) = vbOK Then
        ElseIf vbCancel Then
            Exit Sub
        End If
        Set objdoc = Nothing
        Set objdoc = objword.Documents.Add(LocCardPath & "Claim\" & "CCLMNotification.dot\")
    End If
    If bookmark("CLMNotification") = True Then
        Me.Show
    End If
Exit Sub

openerror:
    Me.Show
    If Err.Number = "5356" Then
        MsgBox Err.Description & "Please close the previous document first using Word or save it to other name", vbExclamation + vbOKOnly, DialogBoxTitle
    ElseIf Err.Number = "462" Then
        Set objword = New Word.Application
        objword.Visible = True
    ElseIf Err.Number = "4198" Then
        Err.Clear
    Else
        MsgBox Err.Description & Err.Number, vbExclamation + vbOKOnly, DialogBoxTitle
    End If
End Sub

Private Function bookmark(Optional sql As String) As Integer
Dim MBM As New ADODB.Recordset
Dim CLM As New ADODB.Recordset
Dim USR As New ADODB.Recordset
Dim strsqlMBM As String
Dim strsqlCLM As String
Dim strsqlUSR As String
Dim strsql As String

    bookmark = False
    
    If txtCASNumber <> "" Then
        objdoc.Bookmarks("CaseNo").Range.Text = StrConv(txtCASNumber, vbUpperCase)
    End If
                       
    objdoc.Bookmarks("MBMName").Range.Text = StrConv(txtPrincipleName, vbUpperCase)
    
    
    objdoc.Bookmarks("INSName").Range.Text = StrConv(txtPrincipleName, vbUpperCase)
    
    If txtAdd1 <> "" Then
        objdoc.Bookmarks("Address1").Range.Text = StrConv(txtAdd1, vbUpperCase)
    End If
                
    If txtAdd2 <> "" Then
        objdoc.Bookmarks("Address2").Range.Text = StrConv(txtAdd2, vbUpperCase)
    End If
    
    If txtAdd3 <> "" Then
        objdoc.Bookmarks("Address3").Range.Text = StrConv(txtAdd3, vbUpperCase)
    End If
            
    If txtMBMNo <> "" Then
        objdoc.Bookmarks("MBMNo").Range.Text = txtMBMNo
    End If
            
    If txtPolicyNo <> "" Then
        objdoc.Bookmarks("PolicyNo").Range.Text = txtPolicyNo
    End If
            
    If txtPatientName <> "" Then
        objdoc.Bookmarks("PatientName").Range.Text = StrConv(txtPatientName, vbUpperCase)
    End If
            
    If txtAdmissionDate <> "" Then
        objdoc.Bookmarks("AdmissionDate").Range.Text = txtAdmissionDate
    End If
        
    If txtDischargeDate <> "" Then
        objdoc.Bookmarks("DischargeDate").Range.Text = txtDischargeDate
    Else
        objdoc.Bookmarks("DischargeDate").Range.Text = "-"
    End If
        
    objdoc.Bookmarks("Hospital").Range.Text = txtHOSName
        
    objdoc.Bookmarks("Date").Range.Text = Date
        
    objdoc.Bookmarks("Payor").Range.Text = txtPayName
        
    If chkUndProc.Value = 1 Then
        objdoc.Bookmarks("UndProcess").Range.Text = "X"
    End If
        
    If chkDoc.Value = 1 Then
        objdoc.Bookmarks("AwaitingDoc").Range.Text = "X"
        objdoc.Bookmarks("AwaitingDocRMK").Range.Text = txtDoc
    End If
        
    If chkInv.Value = 1 Then
        objdoc.Bookmarks("AwaitingRly").Range.Text = "X"
        objdoc.Bookmarks("AwaitingRlyRMK").Range.Text = txtInv
    End If
        
    If chkOther1.Value = 1 Then
        objdoc.Bookmarks("OTHER1").Range.Text = "X"
        objdoc.Bookmarks("TextOther1").Range.Text = txtOther1
    End If
                                        
    bookmark = True
End Function

Private Sub CleanForm()
    txtMBMNo = ""
    txtPolicyNo = ""
    txtPrincipleName = ""
    txtPatientName = ""
    txtAdmissionDate = "__/__/____"
    txtDischargeDate = "__/__/____"
    txtOther1 = ""
    txtDoc = ""
    txtInv = ""
    txtAdd1 = ""
    txtAdd2 = ""
    txtAdd3 = ""
    txtPayMainContact = ""
    txtPayName = ""
    txtHOSName = ""
    chkUndProc.Value = 0
    chkDoc.Value = 0
    chkInv.Value = 0
    chkOther1.Value = 0
End Sub

Private Sub txtCASNumber_LostFocus()
    txtCASNumber = ChkStr(txtCASNumber)
    cmdSave.Enabled = False
End Sub
