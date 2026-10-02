VERSION 5.00
Begin VB.Form FrmMBMAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member Account"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3780
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtOSAmt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtCLMNo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   3735
      Begin VB.CommandButton CmdOK 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Default         =   -1  'True
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox txtRecAmt 
      Height          =   285
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Claim No"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Amount O/S"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Receipt/Payment"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "FrmMBMAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Source As Integer

Private Sub CmdClear_Click()
    txtRecAmt.Text = ""
End Sub

Private Sub CmdOK_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
Dim strsqlUSR As String
Dim USR As New ADODB.Recordset

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If txtRecAmt = "" Then
        MsgBox "Please key in Amount!", vbCritical + vbOKOnly
        
        ElseIf CDbl(txtRecAmt) > CDbl(txtOSAmt) Then
            MsgBox "Recovery Amt exceeded Outstanding Amt", vbCritical + vbOKOnly
        Else
        
        If Source = 1 Then
            
            strsqlUSR = "Select USRCode, AuthorisedStatus From USR where USRCode = '" & Trim(Logon) & "'"
            USR.Open strsqlUSR, medixmasterconnMaint, adOpenKeyset, adLockPessimistic
        
            If Trim(USR!AuthorisedStatus) = "N" Then
            
                FrmReceiptMember.lstReceiptMBM.SelectedItem.SubItems(7) = txtRecAmt
                If Trim(txtRecAmt) = Trim(txtOSAmt) Then
                    FrmReceiptMember.lstReceiptMBM.SelectedItem.SubItems(2) = "Full"
                Else
                    FrmReceiptMember.lstReceiptMBM.SelectedItem.SubItems(2) = "Partial"
                End If
                
                FrmReceiptMember.lstReceiptMBM.SelectedItem.Checked = True
                Call FrmReceiptMember.ProcessCheckedItem(FrmReceiptMember.lstReceiptMBM.SelectedItem, True)
            
            Else
            
                FrmReceiptMember.lstReceiptMBM.SelectedItem.SubItems(7) = txtRecAmt
                If Trim(txtRecAmt) = Trim(txtOSAmt) Then
                    FrmReceiptMember.lstReceiptMBM.SelectedItem.SubItems(2) = "Full"
                Else
                    FrmReceiptMember.lstReceiptMBM.SelectedItem.SubItems(2) = "Partial"
                End If
                
                FrmReceiptMember.lstReceiptMBM.SelectedItem.Checked = True
                Call FrmReceiptMember.ProcessCheckedItem(FrmReceiptMember.lstReceiptMBM.SelectedItem, True)
                    
            End If
            USR.Close
            Set USR = Nothing
               
        ElseIf Source = 2 Then
            
            FrmPayMember.lstPaymentMBM.SelectedItem.SubItems(8) = txtRecAmt
            If Trim(txtRecAmt) = Trim(txtOSAmt) Then
                FrmPayMember.lstPaymentMBM.SelectedItem.SubItems(2) = "Full"
            Else
                FrmPayMember.lstPaymentMBM.SelectedItem.SubItems(2) = "Partial"
            End If
            
            FrmPayMember.lstPaymentMBM.SelectedItem.Checked = True
            Call FrmPayMember.ProcessCheckedItem(FrmPayMember.lstPaymentMBM.SelectedItem, True)
            
        ElseIf Source = 3 Then
          
            FrmPaymentHospital.lstPaymentHos.SelectedItem.SubItems(8) = txtRecAmt
            If Trim(txtRecAmt) = Trim(txtOSAmt) Then
                FrmPaymentHospital.lstPaymentHos.SelectedItem.SubItems(1) = "Full"
            Else
                FrmPaymentHospital.lstPaymentHos.SelectedItem.SubItems(1) = "Partial"
            End If
            
            FrmPaymentHospital.lstPaymentHos.SelectedItem.Checked = True
            Call FrmPaymentHospital.ProcessCheckedItem(FrmPaymentHospital.lstPaymentHos.SelectedItem, True)
        End If
        
        Unload Me
    End If
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub


'*************************** Start Diable Non Integer Value
Private Sub TxtRecAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

