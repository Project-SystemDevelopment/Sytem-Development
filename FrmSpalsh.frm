VERSION 5.00
Begin VB.Form FrmSplash 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4275
   ClientLeft      =   5145
   ClientTop       =   -645
   ClientWidth     =   7665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4245
   ScaleMode       =   0  'User
   ScaleWidth      =   7380
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1530
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3060
      Width           =   4650
   End
   Begin VB.Timer timSplash 
      Interval        =   2000
      Left            =   6930
      Top             =   2835
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5490
      TabIndex        =   4
      Top             =   3915
      Width           =   1095
   End
   Begin VB.Label VersionNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6480
      TabIndex        =   3
      Top             =   3915
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   1993.033
      X2              =   7018.943
      Y1              =   1832.053
      Y2              =   1832.053
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Height          =   4290
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1455
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
      Left            =   2070
      TabIndex        =   1
      Top             =   1845
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   1980
      Picture         =   "FrmSpalsh.frx":0000
      Top             =   675
      Width           =   1620
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------+
' Project Name  : MedixHIS                                      '
' Department    : IT Department                                 '
' IT Manager    :                                               '
' Developer     : Lee Voon Swee                                 '
' Developer     : Jason Loh Wei Kee                             '
' Description   : HealthCare Information System                 '
'---------------------------------------------------------------+
Option Explicit

'Private Sub Form_Load()
 '   VersionNo = Version
'End Sub

Private Sub timSplash_Timer()

GetKeyboardState kbArray
kbArray.kbByte(VK_CAPITAL) = 1
SetKeyboardState kbArray

Unload Me
mdiMain.Show
'frmLogin.Show
'frmAGEBand.Show
End Sub

