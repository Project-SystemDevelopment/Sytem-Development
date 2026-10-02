VERSION 5.00
Begin VB.Form frmExpired 
   BackColor       =   &H80000005&
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   6120
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Left            =   120
      Top             =   1680
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HealthCare Information System"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   465
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   " This Software has been expired! Please contact your Administrator!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   3735
   End
End
Attribute VB_Name = "frmExpired"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Label1.Tag = "1"
    Timer2.Interval = 600
End Sub

Private Sub Timer2_Timer()
Dim JP2 As Control
     
     For Each JP2 In Me
        If TypeOf JP2 Is Label And JP2.Tag = "1" Then
           JP2.Visible = Not JP2.Visible
        End If
     Next

End Sub
