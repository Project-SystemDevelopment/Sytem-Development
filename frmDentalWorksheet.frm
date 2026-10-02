VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDentalWorksheet 
   Caption         =   "DentalWorksheet"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Dental WorkSheet"
      TabPicture(0)   =   "frmDentalWorksheet.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fgrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdExit"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         Height          =   255
         Left            =   10560
         TabIndex        =   1
         Top             =   7920
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid Fgrid 
         Height          =   7455
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   13150
         _Version        =   393216
         Rows            =   20
         Cols            =   9
         FixedCols       =   2
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
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Dental Worksheet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmDentalWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call BNFcombolisting
End Sub

Private Sub BNFcombolisting()
Screen.MousePointer = vbHourglass

Dim subCat As New ADODB.Recordset
Dim Cnt As Integer
Call DoInitialSettings

    subCat.Open "select DentalBNFCATCode , DentalBNFCATDesc   from DentalBNFCategory order by DentalBNFCATCode", MediConnCISDB, adOpenKeyset, adLockOptimistic
    Cnt = 1
    Fgrid.Rows = subCat.RecordCount + 8
    
    Do Until subCat.EOF
         Fgrid.Col = 0
         Fgrid.Row = Cnt
         Fgrid.Text = subCat!DentalBNFCATCode
         
         Fgrid.Col = 1
         Fgrid.Row = Cnt
         Fgrid.Text = subCat!DentalBNFCATDesc
    
         Cnt = Cnt + 1
         subCat.MoveNext
    Loop
Screen.MousePointer = Default
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
' Benefits
Sub DoInitialSettings()
    Dim i, Y    As Integer
    Fgrid.MergeCol(1) = True     ' Allow merge on Columns 0 thru 3
    
    Fgrid.ColAlignment(1) = 1
    Fgrid.MergeCells = flexMergeRestrictColumns
    Fgrid.ColAlignment(3) = 7
    Fgrid.ColAlignment(4) = 7
    Fgrid.ColAlignment(5) = 7
    Fgrid.ColAlignment(6) = 7
    Fgrid.ColAlignment(7) = 7
    Fgrid.ColAlignment(8) = 7
    Fgrid.Row = 0
    'Fgrid.Font = Bold
    Fgrid.ColWidth(0) = 0     ' Set column's width
    Fgrid.ColWidth(1) = 2200
    Fgrid.ColWidth(2) = 1000
    Fgrid.ColWidth(3) = 1400
    Fgrid.ColWidth(4) = 1200
    Fgrid.ColWidth(5) = 1500
    Fgrid.ColWidth(6) = 1000
    Fgrid.ColWidth(7) = 1400
    Fgrid.ColWidth(8) = 1000

    Fgrid.Col = 1
    Fgrid = "Benfits"
    Fgrid.Col = 2
    Fgrid = "Actual"
    Fgrid.Col = 3
    Fgrid = "Contracted Rate"
    Fgrid.Col = 4
    Fgrid = "Payable to CLN"
    Fgrid.Col = 5
    Fgrid = "Not Payable to CLN"
    Fgrid.Col = 6
    Fgrid = "Benefit Lmt"
    Fgrid.Col = 7
    Fgrid = "Approved Benefit"
    Fgrid.Col = 8
    Fgrid = "Excess"
    
    Fgrid.Col = 1
    Fgrid.Row = 14
    Fgrid = "Total"

    Fgrid.Row = 15
    Fgrid = "Amount Not Contracted"
    
    Fgrid.Row = 16
    Fgrid = "Less Co-Payment"
    
    Fgrid.Row = 17
    Fgrid = "Visit Limit"
    
    Fgrid.Row = 18
    Fgrid = "Balance Limit"

    Fgrid.Row = 19
    Fgrid = "Final Approved"
    
End Sub


