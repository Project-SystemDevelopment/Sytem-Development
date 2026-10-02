VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDailyMBMCASByIns 
   Caption         =   "MBM'ship & Claims Daily Statistic Report (By Ins Type)"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "&Print To File"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid InsFgrid 
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   6
      Cols            =   8
      BackColorFixed  =   -2147483638
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   7455
   End
   Begin VB.Label Label1 
      Caption         =   "MBM'ship && Claims Daily Statistic Report (By INS Type)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmDailyMBMCASByIns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
    frmDateSelection.show
End Sub

Private Sub InsInitial()
Dim AA As Integer
    
    InsFgrid.Row = 0
    InsFgrid.Col = 0
    InsFgrid = "Insured Type"
    InsFgrid.Col = 1
    InsFgrid = "Principal"
    InsFgrid.Col = 2
    InsFgrid = "Supplementary"
    InsFgrid.Col = 3
    InsFgrid = "Total"
    InsFgrid.Col = 4
    InsFgrid = "Premium"
    InsFgrid.Col = 5
    InsFgrid = "Cases"
    InsFgrid.Col = 6
    InsFgrid = "Approved"
    InsFgrid.Col = 7
    InsFgrid = "Prof/Loss"
    
    InsFgrid.Col = 0
    InsFgrid.Row = 1
    InsFgrid = "I - Individual"
    InsFgrid.Row = 2
    InsFgrid = "F - Family"
    InsFgrid.Row = 3
    InsFgrid = "G - Group Individual"
    InsFgrid.Row = 4
    InsFgrid = "H - Group Family"
    
    InsFgrid.Col = 0
   
    InsFgrid.ColAlignment(0) = 1
    InsFgrid.ColAlignment(1) = 6
    InsFgrid.ColAlignment(2) = 6
    InsFgrid.ColAlignment(3) = 6
    InsFgrid.ColAlignment(4) = 6
    InsFgrid.ColAlignment(5) = 6
    InsFgrid.ColAlignment(6) = 6
    InsFgrid.ColAlignment(7) = 6
    
    For AA = 1 To 5
        InsFgrid.ColWidth(AA - 1) = 1180
        If AA = 1 Then
            InsFgrid.ColWidth(AA - 1) = 1600
        End If
        InsFgrid.Row = AA
    Next AA
    InsFgrid.Row = 5
    InsFgrid.CellFontBold = True
    InsFgrid.CellFontSize = 10
    InsFgrid = "Total"
End Sub

Private Sub cmdprint_Click()
    frmDailyMBMCASByIns.PrintForm
End Sub

Private Sub Form_Load()
    Call InsInitial
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If
End Sub
