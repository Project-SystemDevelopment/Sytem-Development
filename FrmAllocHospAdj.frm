VERSION 5.00
Begin VB.Form FrmAllocHospAdj 
   Caption         =   "Hospital Adjustment"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   80
      TabIndex        =   43
      Top             =   360
      Width           =   11775
      Begin VB.TextBox txtPVNo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         MaxLength       =   2015
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtPVNo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         MaxLength       =   15
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.PictureBox txtPVDate1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         ScaleHeight     =   225
         ScaleWidth      =   1635
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.PictureBox txtPVDate2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         ScaleHeight     =   225
         ScaleWidth      =   1635
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtPVCasNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   15
         TabIndex        =   0
         Text            =   "10400"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtAllocNumber 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   8
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Case No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Alloc. No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "PV Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   47
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "PV Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   46
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   45
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   44
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Frame Frame10 
      Height          =   585
      Left            =   80
      TabIndex        =   38
      Top             =   6840
      Width           =   11775
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         TabIndex        =   21
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Ex&port"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8880
         TabIndex        =   24
         Top             =   200
         Width           =   840
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10440
         TabIndex        =   26
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9720
         TabIndex        =   25
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6600
         TabIndex        =   23
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         TabIndex        =   19
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         TabIndex        =   18
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         TabIndex        =   20
         Top             =   200
         Width           =   720
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         TabIndex        =   22
         Top             =   200
         Width           =   720
      End
      Begin VB.Label Label9 
         Caption         =   "of"
         Height          =   255
         Left            =   840
         TabIndex        =   42
         Top             =   200
         Width           =   255
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   200
         Width           =   525
      End
      Begin VB.Label Label8 
         Caption         =   "Records"
         Height          =   255
         Left            =   1920
         TabIndex        =   40
         Top             =   200
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1200
         TabIndex        =   39
         Top             =   200
         Width           =   645
      End
   End
   Begin VB.Frame FrameDet 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   80
      TabIndex        =   27
      Top             =   5760
      Width           =   11775
      Begin VB.TextBox txtChqNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         MaxLength       =   20
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtPVNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   10080
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   240
         Width           =   1620
      End
      Begin VB.PictureBox txtPVDate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         ScaleHeight     =   225
         ScaleWidth      =   1035
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.PictureBox txtChqDate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         ScaleHeight     =   225
         ScaleWidth      =   1035
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCasNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   15
         TabIndex        =   7
         Text            =   "10400"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtChqAmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         MaxLength       =   12
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox CboACPeriodYr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmAllocHospAdj.frx":0000
         Left            =   8520
         List            =   "FrmAllocHospAdj.frx":0002
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox CboACPeriodMth 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmAllocHospAdj.frx":0004
         Left            =   7320
         List            =   "FrmAllocHospAdj.frx":0006
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtCasVer 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   2
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtAllocNo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   8
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9360
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8115
         TabIndex        =   36
         Top             =   555
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "A/C Month"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   35
         Top             =   495
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Chq. Amt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Chq. No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   33
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "PV Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   32
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "PV No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   31
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "Alloc. No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "Case No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Chq. Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   28
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.PictureBox lstHospAdj 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   75
      ScaleHeight     =   4155
      ScaleWidth      =   11595
      TabIndex        =   6
      Top             =   1440
      Width           =   11655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Hospital Adjustment"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      TabIndex        =   51
      Top             =   0
      Width           =   11895
   End
   Begin VB.Label Label33 
      Caption         =   "Bold - Searchable"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   7440
      Width           =   1575
   End
End
Attribute VB_Name = "FrmAllocHospAdj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim intExit As Integer
Dim EditFlag As Boolean

Private Sub ComboListing()
Dim monthstart
Dim yearstart
        
    For monthstart = 1 To 12
        With CboACPeriodMth
            .AddItem Format(monthstart, "00")
        End With
        
    Next monthstart
    
    For yearstart = 1999 To 3000
        With CboACPeriodYr
            .AddItem yearstart
        End With
    
    Next yearstart

End Sub

Private Sub CboACPeriodMth_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboACPeriodMth.Text, CboACPeriodMth.SelStart) + Chr(KeyAscii) + Mid(CboACPeriodMth.Text, CboACPeriodMth.SelStart + CboACPeriodMth.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboACPeriodMth.ListCount - 1
       
        If NewText = LCase(Left(CboACPeriodMth.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboACPeriodMth.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               CboACPeriodMth.Text = ValidValue
               CboACPeriodMth.SelStart = 0
               CboACPeriodMth.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboACPeriodYr_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboACPeriodYr.Text, CboACPeriodYr.SelStart) + Chr(KeyAscii) + Mid(CboACPeriodYr.Text, CboACPeriodYr.SelStart + CboACPeriodYr.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboACPeriodYr.ListCount - 1
       
        If NewText = LCase(Left(CboACPeriodYr.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboACPeriodYr.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               CboACPeriodYr.Text = ValidValue
               CboACPeriodYr.SelStart = 0
               CboACPeriodYr.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CmdAdd_Click()
    EditFlag = False
    Call ClearDet
    FrameDet.Enabled = True
    CmdSave.Enabled = True
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    txtCasNo.SetFocus
End Sub

Private Sub CmdClear_Click()
Dim TmpCaseNO As String
    TmpCaseNO = txtCasNo
    Call ClearForm(Me)
    txtPVCasNo = "10400"
    txtCasNo = "10400"
    lstHospAdj.ListItems.Clear
    CboACPeriodMth.Text = ""
    CboACPeriodYr.Text = ""
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    lblCurrent = ""
    lblTotal = ""
    txtAllocNumber.SetFocus
End Sub

Private Sub CmdDelete_Click()
Dim rst2 As New ADODB.Recordset
Dim sql2 As String
Dim RecordDel
    
    CmdSearch.Enabled = False
    CmdDelete.Enabled = False
    CmdSave.Enabled = False
    
    If lstHospAdj.ListItems.Count > 0 Then
        If Trim(lstHospAdj.SelectedItem.SubItems(2)) <> "" Then
            RecordDel = MsgBox("Do you want to delete the record?", vbYesNo + vbExclamation)
            
            If RecordDel = vbYes Then
                Screen.MousePointer = vbHourglass
                    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    
                    sql2 = ""
                    sql2 = "Delete From RecHospAdj Where HospAdjAllocNo = '" & Trim(lstHospAdj.SelectedItem.SubItems(2)) & "'"
                    rst2.Open sql2, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
                    MsgBox "Record Deleted... ", vbOKOnly + vbInformation, DialogBoxTitle
                    Call ClearDet
                    medixmasterconnPayment.Close
                    Set medixmasterconnPayment = Nothing
                Screen.MousePointer = vbDefault
            End If
            lstHospAdj.ListItems.Remove (lstHospAdj.SelectedItem.Index)
        End If
    End If
    CmdSearch.Enabled = True
    CmdDelete.Enabled = True
    CmdSave.Enabled = True
End Sub

Private Sub CmdEdit_Click()
    If Trim(txtAllocNo) <> "" Then
        EditFlag = True
        CmdSave.Caption = "&Update"
        CmdSave.Enabled = True
        FrameDet.Enabled = True
        txtPVNo.SetFocus
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
    Call ExportToExcelAnalysis1(lstHospAdj)
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSave_Click()
Dim UpdateRec
Dim HospAdjRec As New ADODB.Recordset
Dim HospAdjAllocSeq As String
Dim HospAdjSql As String
Dim tmpAllocNo As String
Dim CLmSql As String
Dim CLmRec As New ADODB.Recordset
Dim ClmNotFound As Boolean
    
    CmdSave.Enabled = False
    
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    ClmNotFound = False
    If Trim(txtCasNo) <> "10400" And Trim(txtCasNo) <> "" And Trim(txtCasVer) <> "" Then
        CLmSql = "SELECT CASNumber, ClaimVersion FROM CLAIM WHERE CASNumber='" & Trim(txtCasNo) & "' and ClaimVersion='" & Trim(txtCasVer) & "'"
        CLmRec.Open CLmSql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        If CLmRec.recordCount <= 0 Then
           ClmNotFound = True
        End If
        CLmRec.Close
        Set CLmRec = Nothing
    End If
    If Trim(txtPVNo) = "" Then
        MsgBox "Please Keyin PV Number ! ", vbOKOnly
        txtPVNo.SetFocus
    ElseIf txtPVDate = "__/__/____" Then
        MsgBox "Please Keyin PV Date!", vbCritical
        txtPVDate.SetFocus
    ElseIf Trim(txtChqNo) = "" Then
        MsgBox "Please Keyin Cheque Number ! ", vbCritical
        txtChqNo.SetFocus
    ElseIf txtChqDate = "__/__/____" Then
        MsgBox "Please Keyin Cheque Date!", vbCritical
        txtChqDate.SetFocus
    ElseIf Trim(txtChqAmt) = "" Or IsNumeric(txtChqAmt) = False Then
        MsgBox "Please Keyin Cheque Amt!", vbCritical
        txtChqAmt.SetFocus
    ElseIf Trim(CboACPeriodMth) = "" Then
        MsgBox "Please Keyin Account Period - Month !", vbCritical
        CboACPeriodMth.SetFocus
    ElseIf Trim(CboACPeriodYr) = "" Then
        MsgBox "Please Keyin Account Period - Year !", vbCritical
        CboACPeriodYr.SetFocus
    ElseIf Trim(txtCasNo) = "" Then
        MsgBox "Please Keyin Case No ! ", vbCritical
        txtCasNo.SetFocus
    ElseIf Trim(txtCasNo) <> "10400" And Trim(txtCasNo) <> "" And Trim(txtCasVer) = "" Then
        MsgBox "Please Keyin Case Version No. ! ", vbCritical
        txtCasVer.Enabled = True
        txtCasVer.SetFocus
    ElseIf ClmNotFound = True Then
        MsgBox "Invalid Case ! Please try again !", vbCritical
        txtCasNo.SetFocus
    Else
        UpdateRec = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
        If UpdateRec = vbYes Then
            If EditFlag = False Then
                HospAdjAllocSeq = GenerateAllocNo("AH")
                txtAllocNo = ""
                txtAllocNo = HospAdjAllocSeq
            End If
            HospAdjSql = "Select * from RecHospAdj WHERE HOSpAdjAllocno='" & Trim(txtAllocNo) & "'"
            HospAdjRec.Open HospAdjSql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
            If EditFlag = False Then
                HospAdjRec.AddNew
            End If
            With HospAdjRec
                !HospAdjAllocNo = Trim(txtAllocNo)
                !HospAdjCasENo = Trim(txtCasNo)
                !HospAdjVerNo = IIf(Len(Trim(txtCasVer)), Trim(txtCasVer), "")
                !HospAdjRecNo = Trim(txtPVNo)
                !HospAdjRecDate = Trim(txtPVDate)
                !HospAdjChqNo = Trim(txtChqNo)
                !HospAdjChqDate = Trim(txtChqDate)
                !HospAdjChqAmt = Trim(txtChqAmt)
                !HospAdjAcPeriodMth = Format(Trim(CboACPeriodMth), "00")
                !HospAdjAcPeriodYr = Trim(CboACPeriodYr)
                !HospAdjRemarks = IIf(Len(Trim(txtRemark)), Trim(txtRemark), "")
                !HospAdjLastUpdateUser = Logon
                !HospAdjLastUpdateDate = Date
            End With
            HospAdjRec.Update
            HospAdjRec.Close
            Set HospAdjRec = Nothing
            If EditFlag = True Then
                MsgBox "Record Updated..... ! "
            Else
                MsgBox "Record Saved..... ! "
            End If
            CmdSave.Caption = "Sa&ve"
            CmdSave.Enabled = False
            CmdEdit.Enabled = False
            tmpAllocNo = txtAllocNo
            Call CmdClear_Click
            Call MedixHospAdjListing(" And HospAdjAllocNo = '" & Trim(tmpAllocNo) & "'")
        End If
    End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
    CmdSave.Enabled = True

End Sub

Private Sub CmdSearch_Click()
    
    Screen.MousePointer = vbHourglass
        CmdSearch.Enabled = False
        CmdDelete.Enabled = False
        CmdSave.Enabled = False
        medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
            Call SearchRecord
        medixmasterconnPayment.Close
        Set medixmasterconnPayment = Nothing
        CmdSearch.Enabled = True
        CmdDelete.Enabled = True
        CmdSave.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    CmdExit.Enabled = True
End Sub

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

Private Sub Form_Load()
    intExit = 0
    EditFlag = False
    Call ComboListing
End Sub

Private Sub lstHospAdj_Click()
Dim LstREc As ListItem
    FrameDet.Enabled = False
    If lstHospAdj.ListItems.Count > 0 Then
        Call ClearDet
        CmdEdit.Enabled = True
        CmdDelete.Enabled = True
        txtCasNo = lstHospAdj.SelectedItem.Text
        txtCasVer = lstHospAdj.SelectedItem.SubItems(1)
        txtAllocNo = lstHospAdj.SelectedItem.SubItems(2)
        txtPVNo = lstHospAdj.SelectedItem.SubItems(3)
        If IsDate(lstHospAdj.SelectedItem.SubItems(4)) Then
            txtPVDate = lstHospAdj.SelectedItem.SubItems(4)
        End If
        txtChqNo = lstHospAdj.SelectedItem.SubItems(5)
        If IsDate(lstHospAdj.SelectedItem.SubItems(6)) Then
            txtChqDate = lstHospAdj.SelectedItem.SubItems(6)
        End If
        txtChqAmt = IIf(IsNumeric(lstHospAdj.SelectedItem.SubItems(7)), lstHospAdj.SelectedItem.SubItems(7), 0)
        If Len(lstHospAdj.SelectedItem.SubItems(8)) Then
            CboACPeriodMth = Mid(lstHospAdj.SelectedItem.SubItems(8), 1, 2)
            CboACPeriodYr = Mid(lstHospAdj.SelectedItem.SubItems(8), 4, 4)
        End If
        txtRemark = lstHospAdj.SelectedItem.SubItems(10)
    End If
End Sub

Private Sub txtAllocNo_LostFocus()
    txtAllocNo = ChkStr(txtAllocNo)
End Sub

Private Sub txtAllocNumber_LostFocus()
    txtAllocNumber = ChkStr(txtAllocNumber)
End Sub


Private Sub txtCasNo_GotFocus()
    If Trim(txtCasNo) = "10400" Then
        txtCasVer.Enabled = False
        txtCasVer = ""
    Else
        If Trim(txtCasNo) <> "" Then
            txtCasVer.Enabled = True
        End If
    End If
End Sub

Private Sub txtCasNo_LostFocus()
    txtCasNo = ChkStr(txtCasNo)
    If Trim(txtCasNo) = "10400" Then
        txtCasVer.Enabled = False
        txtCasVer = ""
    Else
        If Trim(txtCasNo) <> "" Then
            txtCasVer.Enabled = True
            txtCasVer.SetFocus
        End If
    End If
End Sub

Private Sub txtChqAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtChqNo_LostFocus()
    txtChqNo = ChkStr(txtChqNo)
End Sub

Private Sub txtChqDate_LostFocus()
    If IsDate(txtChqDate) Then
    Else
        If txtChqDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtChqDate.SetFocus
        End If
    End If
End Sub

Private Sub txtPVDate_LostFocus()
    If IsDate(txtPVDate) Then
        CboACPeriodMth.Text = Format(txtPVDate, "mm")
        CboACPeriodYr.Text = Format(txtPVDate, "yyyy")
    Else
        If txtPVDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDate.SetFocus
        End If
    End If
End Sub

Private Sub txtPVDate1_LostFocus()
    If IsDate(txtPVDate1) Then
    Else
        If txtPVDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtPVDate2_LostFocus()
    If IsDate(txtPVDate2) Then
    Else
        If txtPVDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtPVNo_LostFocus()
    txtPVNo = ChkStr(txtPVNo)
End Sub

Private Sub txtPVNo1_LostFocus()
    txtPVNo1 = ChkStr(txtPVNo1)
End Sub

Private Sub txtPVNo2_LostFocus()
    txtPVNo2 = ChkStr(txtPVNo2)
End Sub

Private Sub txtRemark_LostFocus()
    txtRemark = ChkStr(txtRemark)
End Sub

Private Sub SearchRecord()
Dim strAppend As String

    strAppend = ""
    If Len(Trim(txtPVCasNo)) Then
        strAppend = strAppend + " And HospAdjCasENo = '" & Trim(txtPVCasNo) & "'"
    End If
    If Len(Trim(txtAllocNumber)) Then
        strAppend = strAppend + " And HospAdjAllocNo = '" & Trim(txtAllocNumber) & "'"
    End If
    
    If Len(Trim(txtPVNo1)) Then
        strAppend = strAppend + " And HospAdjRecNo >= '" & Trim(txtPVNo1) & "'"
    End If
    If Len(txtPVNo2) Then
        strAppend = strAppend + " and HospAdjRecNo <= '" & Trim(txtPVNo2) & "'"
    End If
    
    If Len(Trim(txtPVDate1)) > 0 And IsDate(txtPVDate1) = True Then
        strAppend = strAppend + " And HospAdjRecDate >= '" & Format(txtPVDate1, "mm/dd/yyyy") & "'"
    End If
    If Len(Trim(txtPVDate2)) > 0 And IsDate(txtPVDate2) = True Then
        strAppend = strAppend + " And HospAdjRecDate <= '" & Format(txtPVDate2, "mm/dd/yyyy") & "'"
    End If
    
    Call MedixHospAdjListing(strAppend)
   
End Sub
Private Sub MedixHospAdjListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim totalREc As String
Dim CurrentRec As String
Dim Check As Boolean
Dim TmpType As String

    totalREc = 0
    CurrentRec = 0
    lstHospAdj.ListItems.Clear

    Sql = "Select * FROM RecHospAdj where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        Sql = Sql + strAppend
    End If
    
    rst2.Open Sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
   
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
    Do
        Do Until rst2.EOF
        totalREc = rst2.recordCount
        lblTotal = totalREc
            
            With rst2
                Set Record = lstHospAdj.ListItems.Add(, , !HospAdjCasENo)
                Record.SubItems(1) = IIf(Len(!HospAdjVerNo), !HospAdjVerNo, "")
                Record.SubItems(2) = !HospAdjAllocNo
                Record.SubItems(3) = IIf(Len(!HospAdjRecNo), !HospAdjRecNo, "")
               
                If Len(!HospAdjRecDate) Then
                    Record.SubItems(4) = !HospAdjRecDate
                End If
                Record.SubItems(5) = IIf(Len(!HospAdjChqNo), !HospAdjChqNo, "")
                If Len(!HospAdjChqDate) Then
                    Record.SubItems(6) = Trim(!HospAdjChqDate)
                End If
                Record.SubItems(7) = Format(IIf(IsNumeric(!HospAdjChqAmt), !HospAdjChqAmt, 0), "#,##0.00")
                If Len(Trim(!HospAdjAcPeriodMth) & Trim(!HospAdjAcPeriodYr)) Then
                    Record.SubItems(8) = Format(!HospAdjAcPeriodMth & "/" & !HospAdjAcPeriodYr, "MM/YYYY")
                End If
                Record.SubItems(9) = IIf(Len(!HospAdjLastUpdateUser), !HospAdjLastUpdateUser, "")
                Record.SubItems(10) = IIf(Len(!HospAdjRemarks), !HospAdjRemarks, "")
                CurrentRec = CurrentRec + 1
                lblCurrent = CurrentRec
            End With
            rst2.MoveNext
            
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
            Loop
        Loop Until Check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub ClearDet()
    txtAllocNo = ""
    txtPVNo = ""
    txtPVDate = "__/__/____"
    txtChqNo = ""
    txtChqDate = "__/__/____"
    txtChqAmt = ""
    CboACPeriodMth.Text = ""
    CboACPeriodYr.Text = ""
    txtRemark = ""
End Sub

