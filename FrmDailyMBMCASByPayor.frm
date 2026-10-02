VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDailyMBMCASByPayor 
   Caption         =   "FrmDailyMBMCASByPayor"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdprint 
      Caption         =   "&Print To File"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   7800
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid PayorFgrid 
      Height          =   6735
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11880
      _Version        =   393216
      Rows            =   3
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
      TabIndex        =   4
      Top             =   0
      Width           =   7335
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
      TabIndex        =   3
      Top             =   480
      Width           =   7455
   End
End
Attribute VB_Name = "FrmDailyMBMCASByPayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PayorInitial()
Dim aa As Integer
    
    PayorFgrid.Row = 0
    PayorFgrid.Col = 0
    PayorFgrid = "Payor"
    PayorFgrid.Col = 1
    PayorFgrid = "Principal"
    PayorFgrid.Col = 2
    PayorFgrid = "Supplementary"
    PayorFgrid.Col = 3
    PayorFgrid = "Total"
    PayorFgrid.Col = 4
    PayorFgrid = "Premium"
    PayorFgrid.Col = 5
    PayorFgrid = "Cases"
    PayorFgrid.Col = 6
    PayorFgrid = "Approved"
    PayorFgrid.Col = 7
    PayorFgrid = "Prof/Loss"
        
    PayorFgrid.Col = 0
    PayorFgrid.ColAlignment(0) = 1
    PayorFgrid.ColAlignment(1) = 7
    PayorFgrid.ColAlignment(2) = 7
    PayorFgrid.ColAlignment(3) = 7
    PayorFgrid.ColAlignment(4) = 7
    PayorFgrid.ColAlignment(5) = 7
    PayorFgrid.ColAlignment(6) = 7
    PayorFgrid.ColAlignment(7) = 7
    
 '   For aa = 1 To EOF
 '       PayorFgrid.ColWidth(aa - 1) = 1180
 '       If aa = 1 Then
 '           PayorFgrid.ColWidth(aa - 1) = 1600
 '       End If
 '       PayorFgrid.Row = aa
 '   Next aa
 '   PayorFgrid.Row = 51
 '   PayorFgrid.CellFontBold = True
 '   PayorFgrid.CellFontSize = 10
 '   PayorFgrid = "Total"
End Sub

Private Sub cmdOK_Click()
    Unload Me
    frmDateSelection.show
End Sub

Private Sub cmdprint_Click()
    FrmDailyMBMCASByPayor.PrintForm
End Sub

Private Sub Form_Load()
    Call PayorInitial
End Sub

