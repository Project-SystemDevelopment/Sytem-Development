VERSION 5.00
Begin VB.Form frmITMaintainCLMDel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Worksheet deletion (only for IT dept)"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   8970
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   8895
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   255
         Left            =   7560
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   255
         Left            =   6720
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   255
         Left            =   5880
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   8895
      Begin VB.TextBox txtClmRegBy 
         Height          =   285
         Left            =   5520
         TabIndex        =   25
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox txtAppAmt 
         Height          =   285
         Left            =   7800
         TabIndex        =   19
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox txtPatientName 
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Top             =   180
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   7800
         TabIndex        =   21
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox txtMBMNo 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   420
         Width           =   1695
      End
      Begin VB.TextBox tctCovIID 
         Height          =   285
         Left            =   3720
         TabIndex        =   12
         Top             =   420
         Width           =   495
      End
      Begin VB.TextBox txtPolStatus 
         Height          =   285
         Left            =   3720
         TabIndex        =   18
         Top             =   680
         Width           =   495
      End
      Begin VB.TextBox txtMBMAVLmt 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   680
         Width           =   1695
      End
      Begin VB.TextBox txtBordxNo 
         Height          =   285
         Left            =   5520
         TabIndex        =   23
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Clm Open By"
         Height          =   255
         Left            =   4440
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   4440
         TabIndex        =   24
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Clm Bal Lmt"
         Height          =   255
         Left            =   6840
         TabIndex        =   22
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Approve Amt"
         Height          =   255
         Left            =   6840
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Pol Stat."
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   800
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   4320
         X2              =   4320
         Y1              =   120
         Y2              =   1080
      End
      Begin VB.Label Label7 
         Caption         =   "Available Lmt"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Patient Name"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Cover ID"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Member No"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   8895
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   7680
         TabIndex        =   29
         Top             =   200
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   5880
         TabIndex        =   27
         Top             =   200
         Width           =   855
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "&Go"
         Height          =   255
         Left            =   4080
         TabIndex        =   6
         Top             =   220
         Width           =   615
      End
      Begin VB.TextBox txtVersion 
         Height          =   285
         Left            =   3600
         TabIndex        =   5
         Top             =   200
         Width           =   375
      End
      Begin VB.TextBox txtCasNo 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   200
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Claimability"
         Height          =   255
         Left            =   6840
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Case Status"
         Height          =   255
         Left            =   4920
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Version"
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Case Number"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Worksheet Maintenance - Deletion"
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
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "frmITMaintainCLMDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub CmdExit_Click()
    Unload Me
End Sub
