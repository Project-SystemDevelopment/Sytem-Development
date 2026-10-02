VERSION 5.00
Begin VB.Form frmGenerateCNNote 
   Caption         =   "Generate Credit/Debit Note"
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2280
   ScaleWidth      =   6405
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   6255
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
         Height          =   285
         Left            =   3600
         TabIndex        =   18
         Top             =   200
         Width           =   735
      End
      Begin VB.TextBox txtCLMNo 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   200
         Width           =   1455
      End
      Begin VB.TextBox txtVerNo 
         Height          =   285
         Left            =   3000
         TabIndex        =   1
         Top             =   200
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Claim Number"
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
         TabIndex        =   15
         Top             =   225
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "-"
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
         Left            =   2880
         TabIndex        =   14
         Top             =   225
         Width           =   135
      End
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   6255
      Begin VB.TextBox txtActAmt 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   8
         Top             =   200
         Width           =   1695
      End
      Begin VB.TextBox txtAppAmt 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   7
         Top             =   450
         Width           =   1695
      End
      Begin VB.TextBox txtPay2Hos 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   6
         Top             =   225
         Width           =   1695
      End
      Begin VB.TextBox txtPay2MBM 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   4440
         MaxLength       =   40
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Payable to MBM"
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
         Index           =   13
         Left            =   3240
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Payable to Hosp"
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
         Index           =   15
         Left            =   3240
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Approved Amt"
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
         Index           =   17
         Left            =   120
         TabIndex        =   10
         Top             =   450
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Actual Amt"
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
         Index           =   21
         Left            =   120
         TabIndex        =   9
         Top             =   195
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   1650
      Width           =   6255
      Begin VB.TextBox txtCRDNNote 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   19
         Top             =   200
         Width           =   1695
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   5520
         TabIndex        =   3
         Top             =   200
         Width           =   615
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
         Height          =   330
         Left            =   4800
         TabIndex        =   2
         Top             =   200
         Width           =   735
      End
      Begin VB.Label lblCNDRNote 
         Caption         =   "Credit Note :"
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
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Generate Credit/Debit Note"
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
      TabIndex        =   17
      Top             =   0
      Width           =   6525
   End
End
Attribute VB_Name = "frmGenerateCNNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim sql As String
Dim rst As New ADODB.Recordset
    
    
    If Trim(txtCLMNo) = "" Then
        MsgBox "Please enter claim number", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(txtVerNo) = "" Then
        MsgBox "Please enter version number", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Call ClearItems
        'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        sql = ""
        sql = "Select CASNumber,ClaimVersion,CLMPayableByMedix2H,CLMPayableByMedix2M,CLMTotalActual,CLMBordxAmt,DebitCreditRefNo from Claim where CASNumber = '" & Trim(txtCLMNo) & "' And ClaimVersion = '" & Trim(txtVerNo) & "'"
        rst.Open sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        If rst.recordCount Then
            txtActAmt = IIf(Len(rst!CLMTotalActual), rst!CLMTotalActual, 0)
            txtAppAmt = IIf(Len(rst!CLMBordxAmt), rst!CLMBordxAmt, 0)
            txtPay2Hos = IIf(Len(rst!CLMPayableByMedix2H), rst!CLMPayableByMedix2H, 0)
            txtPay2MBM = IIf(Len(rst!CLMPayableByMedix2M), rst!CLMPayableByMedix2M, 0)
            txtCRDNNote = IIf(Len(rst!DebitCreditRefNo), rst!DebitCreditRefNo, "")
            If Trim(txtCRDNNote) <> "" Then
                If Mid(txtCRDNNote, 1, 1) = "C" Then
                    lblCNDRNote.Caption = "Credit Note : "
                ElseIf Mid(txtCRDNNote, 1, 1) = "C" Then
                    lblCNDRNote.Caption = "Debit Note : "
                End If
            Else
                lblCNDRNote.Caption = "Credit Note : "
            End If
            
            If Trim(txtCRDNNote) <> "" Then
                cmdUpdate.Enabled = False
            ElseIf txtPay2MBM <> 0 And Trim(txtCRDNNote) = "" Then
                cmdUpdate.Enabled = True
            End If
        Else
            MsgBox "Record Not Found", vbOKOnly + vbCritical, DialogBoxTitle
        End If
       ' medixmasterconnWS.Close
       ' Set medixmasterconnWS = Nothing
    End If
End Sub

Private Sub ClearItems()
    txtActAmt = ""
    txtAppAmt = ""
    txtPay2Hos = ""
    txtPay2MBM = ""
    txtCRDNNote = ""
    lblCNDRNote.Caption = "Credit Note : "
End Sub

Private Sub cmdUpdate_Click()
Dim tmpcrdn As String
Dim tmpPayor As String
Dim RecordSaved

    RecordSaved = MsgBox("Do you want to save the record?", vbYesNo + vbExclamation)
    If RecordSaved = vbYes Then

        tmpcrdn = ""
        tmpPayor = Mid(txtCLMNo, 1, 2)
       ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
            If CDbl(txtPay2MBM) < 0 Then
                tmpcrdn = GenerateDebitNo(tmpPayor)
                Call UpdateCreditDebit(tmpcrdn)
                lblCNDRNote.Caption = "Debit Note : "
                txtCRDNNote = tmpcrdn
            ElseIf CDbl(txtPay2MBM) > 0 Then
                tmpcrdn = GenerateCreditNo(tmpPayor)
                Call UpdateCreditDebit(tmpcrdn)
                lblCNDRNote.Caption = "Credit Note : "
                txtCRDNNote = tmpcrdn
            End If
       ' medixmasterconnWS.Close
       ' Set medixmasterconnWS = Nothing
    End If
End Sub

Private Sub UpdateCreditDebit(str As String)
    Dim Claim As New ADODB.Recordset
    Dim strsql As String
    
    
    strsql = "Select CasNUmber , DebitCreditRefNo , debitCreditDate, CLMPayablebyMedix2m  from Claim " & _
        " where CasNUmber ='" & txtCLMNo & "' and ClaimVersion =" & txtVerNo
    Claim.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    If Claim.recordCount > 0 Then
        Claim!DebitCreditRefNo = Trim(ChkStr(str))
        Claim!DebitCreditDate = Date
    Claim.update
    End If
    
    Claim.Close
    Set Claim = Nothing
End Sub

