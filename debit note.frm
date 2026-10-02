VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2160
      TabIndex        =   18
      Top             =   5400
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   840
      Width           =   1575
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1820
      TabIndex        =   15
      Top             =   4680
      Width           =   2895
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1800
      TabIndex        =   14
      Top             =   3960
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1810
      TabIndex        =   13
      Top             =   2760
      Width           =   2895
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1800
      TabIndex        =   12
      Top             =   3360
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1840
      TabIndex        =   11
      Top             =   2090
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1860
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8805
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Debit Note"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3285
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Admission Date"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Remarks"
      Height          =   255
      Left            =   340
      TabIndex        =   9
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Illness"
      Height          =   255
      Left            =   320
      TabIndex        =   8
      Top             =   4040
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Invoice No:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Hospital"
      Height          =   255
      Left            =   280
      TabIndex        =   6
      Top             =   3400
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Case No"
      Height          =   255
      Left            =   260
      TabIndex        =   5
      Top             =   2120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   255
      Left            =   300
      TabIndex        =   4
      Top             =   1480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Membership No:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'dim MBM as New adodb.
End Sub
