VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEmailNotification 
   Caption         =   "Email Notification"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7290
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   6480
      Width           =   11655
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print to Excel"
         Height          =   375
         Left            =   9240
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   2640
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   7440
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   4440
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   375
         Left            =   10440
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSend 
         Caption         =   "Email"
         Height          =   375
         Left            =   8280
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11925
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Email Notification"
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
         Height          =   315
         Left            =   240
         TabIndex        =   21
         Top             =   0
         Width           =   11445
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11655
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   6000
         TabIndex        =   6
         Top             =   180
         Width           =   3495
      End
      Begin VB.CheckBox ChkDFMNRB 
         Caption         =   "*"
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
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtRecMNRB1 
         Height          =   285
         Left            =   7800
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox TxtRecMNRB2 
         Height          =   285
         Left            =   9840
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSMask.MaskEdBox TxtDFMNRB2 
         Height          =   285
         Left            =   3480
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDFMNRB1 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPaymentDateFr 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   165
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPaymentDateTo 
         Height          =   285
         Left            =   3360
         TabIndex        =   10
         Top             =   165
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtPVNoTo 
         Height          =   285
         Left            =   3360
         MaxLength       =   15
         TabIndex        =   1
         Top             =   420
         Width           =   1680
      End
      Begin VB.TextBox TxtPVNoFr 
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   2
         Top             =   420
         Width           =   1440
      End
      Begin VB.Label Label66 
         Caption         =   "PV No"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   18
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Status"
         Height          =   255
         Left            =   5280
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Date Fr MNRB"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Payment Date Fr"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   200
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   195
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Rec. No (MNRB)fr."
         Height          =   255
         Left            =   6360
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   255
         Left            =   9480
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView lstEmailListing 
      Height          =   5415
      Left            =   120
      TabIndex        =   22
      Top             =   1120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9551
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
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Employee Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IC Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Company Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Patient Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Bill Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Paid Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Visit Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Provider Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Claim Number"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Plan Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Payment Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Account Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Date Sumitted"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "MBMNumber"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Email"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmEmailNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_blnDirection(1 To 17) As Boolean

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboStatus.Text, cboStatus.SelStart) + Chr(KeyAscii) + Mid(cboStatus.Text, cboStatus.SelStart + cboStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboStatus.ListCount - 1
 
        If NewText = LCase(Left(cboStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboStatus.Text = ValidValue
                cboStatus.SelStart = 0
                cboStatus.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        Call ExportToExcelAnalysis1(lstEmailListing)
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSearch_Click()
Dim strAppend As String
    lstEmailListing.ListItems.Clear
    Screen.MousePointer = vbHourglass

        If txtPaymentDateFr = "__/__/____" Then
            MsgBox "Please Enter PV Date From!", vbCritical
            txtPaymentDateFr.SetFocus
        ElseIf txtPaymentDateTo = "__/__/____" Then
            MsgBox "Please Enter PV Date To!", vbCritical
            txtPaymentDateTo.SetFocus
        ElseIf Trim(cboStatus) = "" Then
            MsgBox "Please Enter Status!", vbCritical
            cboStatus.SetFocus
        Else
        
            If Len(Trim(txtPaymentDateFr)) > 0 And IsDate(txtPaymentDateFr) = True Then
                strAppend = strAppend + " And PVDate >= '" & Format(txtPaymentDateFr, "yyyy/mm/dd") & "'"
            End If
        
            If Len(Trim(txtPaymentDateTo)) > 0 And IsDate(txtPaymentDateTo) = True Then
                strAppend = strAppend + " And PVDate <= '" & Format(txtPaymentDateTo, "yyyy/mm/dd") & "'"
            End If
    
            If Trim(TxtPVNoFr) <> "" Then
                strAppend = strAppend + " And PVNo >= '" & Trim(TxtPVNoFr) & "'"
            End If
        
            If Trim(TxtPVNoTo) <> "" Then
                strAppend = strAppend + " And PVNo <= '" & TxtPVNoTo & "'"
            End If
            
            If Mid(cboStatus, 1, 1) = "R" Then
                strAppend = strAppend + " And PVEmailStatus = 'Y'"
            Else
                strAppend = strAppend + " And ( PVEmailStatus  = '' or PVEmailStatus is null)"
            End If
            
            Call PaymentListing(strAppend)
        End If
        
    Screen.MousePointer = vbDefault
        
End Sub

Private Sub CmdSend_Click()
On Err GoTo ErrorSave
Dim oOApp As Outlook.Application
Dim oOMail As Outlook.MailItem
Dim i As Integer
Dim ItemChecked As Boolean
Dim Update
Dim strsql As String
Dim rst As New ADODB.Recordset
Dim tmpvisit As String
Dim tmpCLMTotalActual As Double
Dim tmpCLMBordxAmt As Double
Dim rstMBM As New ADODB.Recordset
Dim strsqlMBM As String
Dim tmpPatientName As String
Dim strsqlupdate As String
Dim emailupdate As New ADODB.Recordset

    i = 0
    Screen.MousePointer = vbHourglass
    'medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

        ItemChecked = False
        For i = 1 To lstEmailListing.ListItems.Count
            If lstEmailListing.ListItems(i).Checked = True Then
                ItemChecked = True
                Exit For
            End If
        Next i
        
        If ItemChecked = False Then
            MsgBox "Please select the record", vbOKOnly + vbCritical
        
        Else
            Update = MsgBox("Do you want to start sending email(s) ?", vbYesNo + vbExclamation)
            If Update = vbYes Then
            
                For i = 1 To lstEmailListing.ListItems.Count
                    If lstEmailListing.ListItems(i).Checked = True Then
                        tmpvisit = ""
                        tmpCLMTotalActual = 0
                        tmpCLMBordxAmt = 0
                        
                        Set oOApp = CreateObject("Outlook.Application")
                        Set oOMail = oOApp.CreateItem(olMailItem)
                        
                        strsql = ""
                        strsql = " Select MBMName,CASDateAdmitted,CLMTotalActual,CLMBordxAmt,MBMPatientCovID,MBMNumber,CasNumber,ClaimVersion from View_Email_Details Where PVNo = '" & Trim(lstEmailListing.ListItems(i).SubItems(9)) & "'"
                        rst.Open strsql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
                        
                            Do While Not rst.EOF
                            
                                tmpCLMTotalActual = CDbl(tmpCLMTotalActual) + CDbl(rst!CLMTotalActual)
                                
                                tmpCLMBordxAmt = CDbl(tmpCLMBordxAmt) + CDbl(rst!CLMBordxAmt)
                               ' rst.Close
                               ' Set rst = Nothing
                                If rst!MBMPatientCovID <> "00" Then
                                    strsqlMBM = ""
                                    strsqlMBM = " Select MBMCName FROM MBMCoveredpersons Where MBMCNumber = '" & Trim(rst!MBMNumber) & "' And MBMCCoverID = '" & rst!MBMPatientCovID & "'"
                                    rstMBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
                                 
                                    If rstMBM.EOF = False Then
                                        tmpPatientName = rstMBM!MBMCName
                                    End If
                                    
                                    rstMBM.Close
                                    Set rstMBM = Nothing
                                Else
                                    tmpPatientName = rst!MBMName
                                End If
                                
                                
                                If tmpvisit <> "" Then
                                    tmpvisit = tmpvisit & vbNewLine & rst!CasNumber & "   " & rst!CASDateAdmitted & "               " & Format(rst!CLMTotalActual, "###0.00") & "            " & Format(rst!CLMBordxAmt, "###0.00") & "   " & tmpPatientName
                                Else
                                    tmpvisit = "Claim Ref.   " & " " & "     DOA" & " " & "            Incurred Amt" & " " & "    Paid Amt" & "   " & "Claimant" & vbNewLine
                                    tmpvisit = tmpvisit & rst!CasNumber & "-" & rst!claimversion & "   " & rst!CASDateAdmitted & "             " & Format(rst!CLMTotalActual, "###0.00") & "          " & Format(rst!CLMBordxAmt, "###0.00") & "   " & tmpPatientName
                                End If
                            
                            rst.MoveNext
                            Loop
                            
                                tmpvisit = tmpvisit & vbNewLine
                                tmpvisit = tmpvisit & "                                                    " & Format(tmpCLMTotalActual, "###0.00") & "          " & Format(tmpCLMBordxAmt, "###0.00")
                        rst.Close
                        Set rst = Nothing
                        
                        'oOMail.Display vbModal
                        With oOMail
                            .To = lstEmailListing.ListItems(i).SubItems(16)
                            '.BCC = "sashi@medix.com.my;bryan@medix.com.my"
                            .Subject = lstEmailListing.ListItems(i).SubItems(1) & " (" & lstEmailListing.ListItems(i).SubItems(15) & ") " & "- " & lstEmailListing.ListItems(i).SubItems(9)
                            .Body = "Dear Sir/Madam," & vbNewLine & vbNewLine & _
                                    "IBM EB PAYMENT ADVICE" & vbNewLine & vbNewLine & _
                                    "We are pleased to advise that payment in respect of your claim as mentioned below has been banked into your MBB Account. Please allow reasonable time for the cheque to be cleared." & vbNewLine & vbNewLine & _
                                    "PV Number               :       " & lstEmailListing.ListItems(i).SubItems(9) & vbNewLine & _
                                    "Date Submitted         :       " & lstEmailListing.ListItems(i).SubItems(13) & vbNewLine & _
                                    "Incurred Amount      :       " & "RM " & Format(tmpCLMTotalActual, "###0.00") & vbNewLine & _
                                    "Paid Amount             :       " & "RM " & Format(tmpCLMBordxAmt, "###0.00") & vbNewLine & _
                                    "Cheque Issuance Date    :       " & lstEmailListing.ListItems(i).SubItems(11) & vbNewLine & _
                                    "Bank Account No.     :       " & lstEmailListing.ListItems(i).SubItems(12) & vbNewLine & vbNewLine & _
                                    "DETAILS OF PAYMENT" & vbNewLine & _
                                    tmpvisit & vbNewLine & vbNewLine & _
                                    "For your information, the difference(if any) between the paid amount and the amount incurred is due to MEMBER'S SHARE / CO-PAYMENT." & vbNewLine & vbNewLine & _
                                                         "Please contact our following staff or email to ibm@medix.com.my if you require any clarification:- " & vbNewLine & _
                                    "           Mr. Lim         -  03-7884 1924" & vbNewLine & _
                                    "           Mr. Sashi       -  03-7884 1905" & vbNewLine & _
                                    "           General Line    -  03-7803 2003" & vbNewLine & vbNewLine & _
                                    "Thank you." & vbNewLine & _
                                    "Regards, " & vbNewLine & _
                                    "MediExpress (Malaysia) Sdn Bhd"
                                    
                                    
                                    
                                    '"Employee Name       :       " & lstEmailListing.ListItems(i).SubItems(1) & vbNewLine & _
                                    '"IC Number                :       " & lstEmailListing.ListItems(i).SubItems(2) & vbNewLine & _
                                    '"Company Name        :       " & lstEmailListing.ListItems(i).SubItems(3) & vbNewLine & _
                                    '"Claimant Name        :       " & lstEmailListing.ListItems(i).SubItems(4) & vbNewLine & _

    '"Provider Name      :       " &' _
                                    'lstEmailListing.ListItems(i).SubItems(8)
    
                            '& vbNewLine & _
                            '        "IC Number      :   " & lstEmailListing.ListItems(i).SubItems(2)
                            '.Attachments.Add "\\server\drive\folder\filename", olByValue, 1
                        .Send
                        'oOApp.Quit
                        End With
                            strsqlupdate = ""
                            strsqlupdate = " update Payment2MBM set  PVEmailStatus = 'Y' " & _
                                     " Where PVNo = '" & Trim(lstEmailListing.ListItems(i).SubItems(9)) & "'"
                            medixmasterconnPayment.Execute strsqlupdate
                  End If
                Next i
           End If
        End If
  '  medixmasterconnPayment.Close
  '  Set medixmasterconnPayment = Nothing
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing

    Screen.MousePointer = vbDefault
    MsgBox "Emails have been sent out successfully!", vbOKOnly + vbCritical
    Exit Sub
ErrorSave:
        Screen.MousePointer = vbDefault
        'If CompTrans = False Then
        '    medixmasterconnPayment.RollbackTrans
        '    medixmasterconnPayment.Close
        '    Set medixmasterconnPayment = Nothing
        '    medixmasterconn.Close
        '    Set medixmasterconn = Nothing
        '    medixmasterconnMaint.Close
        '    Set medixmasterconnMaint = Nothing
        'End If
        MsgBox Err.Description

End Sub


Private Sub PaymentListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim CasNumber As String
Dim WSVersion As String
Dim Record As ListItem
Dim Check As Boolean
Dim RecSql As String
Dim ReceiptRec As New ADODB.Recordset
Dim ContrSql As String
Dim ContraREc As New ADODB.Recordset
Dim PayCurrRec As New ADODB.Recordset
Dim strsql As String
Dim Pay2MBMAmt As Double
Dim Paid2MBMAmt As Double
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim RecFound As Boolean
Dim tmpReceiptNo As String
Dim tmpReceiptAmt As Double
Dim sqlstr As String
Dim rstMBM As New ADODB.Recordset
Dim tmpPatientName As String
Dim tmpPaymentdate As String


    RecFound = False
  '  medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    sqlstr = ""
    sqlstr = "Select * from View_Email_notification Where 0 = 0"
            
        If Len(Trim(strAppend)) Then
            sqlstr = sqlstr + strAppend
        End If

    rst2.Open sqlstr, medixmasterconnPayment, adOpenKeyset, adLockOptimistic

    Do
        Do While Not rst2.EOF
        RecFound = True
        'txtPVStatus = ""
        'If Len(rst2!PVPayMBMStatus) Then
        '    txtPVStatus = rst2!PVPayMBMStatus
        'End If
        '    CboPayto.Clear
            With rst2
            
                If Trim(!MBMPatientCovID) <> "00" Then
                    RecSql = ""
                    RecSql = " Select MBMCName FROM MBMCoveredpersons Where MBMCNumber = '" & Trim(!MBMNumber) & "' And MBMCCoverID = '" & !MBMPatientCovID & "'"
                    rstMBM.Open RecSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                 
                    If rstMBM.EOF = False Then
                        tmpPatientName = rstMBM!MBMCName
                    End If
                    
                    rstMBM.Close
                    Set rstMBM = Nothing
                Else
                    tmpPatientName = !MBMName
                End If
                'tmpMBMName = IIf(Len(!MBMName), !MBMName, "")
                Set Record = lstEmailListing.ListItems.Add(, , "")
                If Len(!MBMName) Then
                    Record.SubItems(1) = !MBMName
                End If
                Record.SubItems(2) = !MBMIcBcPp
                Record.SubItems(3) = !MBMGroupCompany
                Record.SubItems(4) = tmpPatientName
                Record.SubItems(5) = Format(!CLMTotalActual, "###0.00")
                Record.SubItems(6) = Format(!CLMBordxAmt, "###0.00")
                Record.SubItems(7) = ""
                Record.SubItems(8) = ""
                Record.SubItems(9) = !PVNo
                Record.SubItems(11) = !PVDate
                Record.SubItems(12) = !MBMBankACNo
                Record.SubItems(13) = !CLMLastDocDate
                Record.SubItems(14) = IIf(Len(!PVEmailStatus), !PVEmailStatus, "")
                Record.SubItems(15) = !MBMNumber
                Record.SubItems(16) = IIf(Len(!MBMEmail), !MBMEmail, "")
                
                tmpPaymentdate = CDate(!PVDate) + 3
                
                
            End With
            rst2.MoveNext
            
            DoEvents
            'If intExit = 1 Then
            '   Check = False
            '   Exit Do
            'End If
        Loop
    Loop Until Check = False
    rst2.Close
    Set rst2 = Nothing
    
    If RecFound = False Then
       MsgBox "No record found !", vbOKOnly + vbCritical, DialogBoxTitle
    End If
 '   medixmasterconnPayment.Close
 '   Set medixmasterconnPayment = Nothing
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
End Sub

Private Sub Form_Load()
    cboStatus.AddItem "N - New"
    cboStatus.AddItem "R - Resend"
End Sub

Private Sub lstEmailListing_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstEmailListing, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstEmailListing, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstEmailListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstEmailListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(12)
    Case 13
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstEmailListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(14)
    Case 15
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(15)
    Case 16
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(16)
    Case 17
        SortListView lstEmailListing, ColumnHeader.Index, ldtString, m_blnDirection(17)
    End Select

End Sub

Private Sub lstEmailListing_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ProcessCheckedItem(Item)
End Sub

Private Sub txtPaymentDateFr_LostFocus()
    If txtPaymentDateFr <> "__/__/____" Then
        If IsDate(txtPaymentDateFr) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtPaymentDateFr.SetFocus
        End If
    End If
End Sub

Private Sub txtPaymentDateTo_LostFocus()
    If txtPaymentDateTo <> "__/__/____" Then
        If IsDate(txtPaymentDateTo) = False Then
            MsgBox "Invalid date format!", vbCritical
            txtPaymentDateTo.SetFocus
        End If
    End If
End Sub

Public Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
Dim RecordSaved

    If Item.Checked = True Then
        If Item.SubItems(16) = "" Then
            MsgBox "Please check member's email!", vbCritical
            Item.Checked = False
        ElseIf Item.SubItems(14) = "Y" Then
        
            RecordSaved = MsgBox("Email has been sent to the member! Do you want to send again?", vbYesNo + vbExclamation, DialogBoxTitle)
                If RecordSaved = vbNo Then
                    Item.Checked = False
                End If
        End If
    End If
End Function



