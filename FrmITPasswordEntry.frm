VERSION 5.00
Begin VB.Form FrmITPasswordEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Entry"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   0
         X2              =   2880
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmITPasswordEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    txtPassword = ChkStr(txtPassword)
    If Trim(txtPassword) = "ITMBMDEL456" Then
        Unload Me
        If FormLoadedByName("FrmITMaintainMBMDel") = True Then
            FrmITMaintainMBMDel.SetFocus
        Else
            FrmITMaintainMBMDel.Show
        End If
    ElseIf Trim(txtPassword) = "ITCAS123" Then
        Unload Me
        If FormLoadedByName("frmITMaintainCLMDel") = True Then
            frmITMaintainCLMDel.SetFocus
        Else
            frmITMaintainCLMDel.Show
        End If
    ElseIf Trim(txtPassword) = "ITLGDEL123" Then
        Unload Me
        If FormLoadedByName("FrmLGCancellation") = True Then
            FrmLGCancellation.SetFocus
        Else
            FrmLGCancellation.Show
        End If
    ElseIf Trim(txtPassword) = "ITDEV123" Then
        Unload Me
        If FormLoadedByName("frmITDevelopment") = True Then
            frmITDevelopment.SetFocus
        Else
            frmITDevelopment.Show
        End If
        
    Else
        MsgBox "Access Denied", vbOKOnly + vbCritical, DialogBoxTitle
        txtPassword.SetFocus
    End If

End Sub

Private Sub Form_Load()
    Me.Caption = DialogBoxTitle

End Sub

Private Sub txtPassword_GotFocus()
    cmdOk.Default = True
End Sub

Private Sub txtPassword_LostFocus()
    cmdOk.Default = False
End Sub
