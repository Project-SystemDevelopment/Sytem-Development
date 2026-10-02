VERSION 5.00
Begin VB.Form frmBeginScan 
   BackColor       =   &H8000000C&
   Caption         =   "frmBeginScan"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGo 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   8400
      Width           =   75
   End
   Begin VB.TextBox txtScanCode 
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   8400
      Width           =   150
   End
End
Attribute VB_Name = "frmBeginScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdGo_Click()
Dim ScanCode As String
'On Error GoTo errOpen
    ScanCode = ""
    ScanCode = txtScanCode
    txtScanCode = ""
    Call SearchForm(ScanCode)
End Sub
    
Private Sub Form_GotFocus()
    txtScanCode.SetFocus
End Sub

