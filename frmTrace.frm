VERSION 5.00
Begin VB.Form frmTrace 
   Caption         =   "Trace Window"
   ClientHeight    =   4140
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   7065
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFF00&
      Height          =   735
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   720
      Width           =   4815
   End
End
Attribute VB_Name = "frmTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()

    Text1.Move 0, 0, ScaleWidth, ScaleHeight

End Sub


Private Sub Form_Unload(Cancel As Integer)

    frmDemo.traceEvents = False

End Sub


