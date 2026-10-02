VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCLMFloatTopupReq 
   Caption         =   " "
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   11790
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CheckAll 
      Caption         =   "Check All"
      Height          =   255
      Left            =   120
      TabIndex        =   59
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1620
      Left            =   120
      TabIndex        =   21
      Top             =   240
      Width           =   11655
      Begin VB.TextBox txtFTNoTo 
         Height          =   285
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   1
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox txtFTNofr 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   0
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox txtTUNoFr 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   2
         Top             =   435
         Width           =   1335
      End
      Begin VB.TextBox txtTUNoTo 
         Height          =   285
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   3
         Top             =   435
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtFTDateFr 
         Height          =   315
         Left            =   7920
         TabIndex        =   8
         Top             =   195
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFTDateTo 
         Height          =   315
         Left            =   9840
         TabIndex        =   9
         Top             =   195
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTUDateFr 
         Height          =   315
         Left            =   7920
         TabIndex        =   10
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTUDateTo 
         Height          =   315
         Left            =   9840
         TabIndex        =   11
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPVNoFr 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   4
         Top             =   655
         Width           =   1335
      End
      Begin VB.TextBox txtPVNoTo 
         Height          =   285
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   5
         Top             =   655
         Width           =   1335
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   915
         Width           =   3255
      End
      Begin VB.ComboBox CboTopupStatus 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   1200
         Width           =   3255
      End
      Begin MSMask.MaskEdBox txtPVDateFr 
         Height          =   315
         Left            =   7920
         TabIndex        =   12
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPVDateTo 
         Height          =   315
         Left            =   9840
         TabIndex        =   13
         Top             =   740
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label17 
         Caption         =   "Req Date from"
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
         Index           =   2
         Left            =   6360
         TabIndex        =   35
         Top             =   500
         Width           =   1455
      End
      Begin VB.Label Label18 
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
         Index           =   2
         Left            =   9360
         TabIndex        =   34
         Top             =   500
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "PV Date from"
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
         Index           =   1
         Left            =   6360
         TabIndex        =   33
         Top             =   780
         Width           =   1455
      End
      Begin VB.Label Label18 
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
         Index           =   1
         Left            =   9360
         TabIndex        =   32
         Top             =   780
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "Req No from"
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
         Index           =   3
         Left            =   240
         TabIndex        =   31
         Top             =   450
         Width           =   1215
      End
      Begin VB.Label Label15 
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
         Index           =   3
         Left            =   3285
         TabIndex        =   30
         Top             =   435
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "FT No from"
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
         Index           =   2
         Left            =   240
         TabIndex        =   29
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label15 
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
         Index           =   2
         Left            =   3285
         TabIndex        =   28
         Top             =   195
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "PV No from"
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
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   680
         Width           =   1215
      End
      Begin VB.Label Label15 
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
         Index           =   1
         Left            =   3285
         TabIndex        =   26
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Payor"
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
         Left            =   240
         TabIndex        =   25
         Top             =   930
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Req Status"
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
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label18 
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
         Index           =   0
         Left            =   9360
         TabIndex        =   23
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "FT Date from"
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
         Index           =   0
         Left            =   6360
         TabIndex        =   22
         Top             =   210
         Width           =   1455
      End
   End
   Begin VB.Frame FrameReceiptPV 
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   6120
      Width           =   11625
      Begin VB.TextBox txtTUNo 
         Height          =   285
         Left            =   885
         MaxLength       =   15
         TabIndex        =   14
         Top             =   200
         Width           =   1215
      End
      Begin VB.TextBox txtRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7320
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   240
         Width           =   4215
      End
      Begin MSMask.MaskEdBox txtTUDate 
         Height          =   285
         Left            =   2925
         TabIndex        =   15
         Top             =   210
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtReqAmt 
         Height          =   285
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   16
         Top             =   225
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "TU Date"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   53
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "TU No"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   52
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Request Amt"
         Height          =   255
         Left            =   4320
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Remark"
         Height          =   255
         Left            =   6600
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView lstFTListing 
      Height          =   3960
      Left            =   120
      TabIndex        =   36
      Top             =   2160
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   6985
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TU No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "TU Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Req Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "FT No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "FT Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "FT Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "PV No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "PV Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Payor"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "NoofBordx"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   37
      Top             =   6840
      Width           =   11655
      Begin VB.TextBox lblSelectedAmt 
         Height          =   285
         Left            =   4800
         TabIndex        =   62
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox lbltotalAmt 
         Height          =   285
         Left            =   2400
         TabIndex        =   61
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox lblcurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdPrintPV 
         Caption         =   "Print P&V"
         Height          =   330
         Left            =   8640
         TabIndex        =   54
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   330
         Left            =   6000
         TabIndex        =   47
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "S&earch"
         Height          =   330
         Left            =   7320
         TabIndex        =   46
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   330
         Left            =   8040
         TabIndex        =   45
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   10920
         TabIndex        =   44
         Top             =   220
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10320
         TabIndex        =   43
         Top             =   220
         Width           =   615
      End
      Begin VB.TextBox TxtPStatus 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         MaxLength       =   15
         TabIndex        =   42
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtPStatus2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   41
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         Height          =   330
         Left            =   6600
         TabIndex        =   40
         Top             =   220
         Width           =   735
      End
      Begin VB.CommandButton cmdEXPtoExcel 
         Caption         =   "&Print to File"
         Height          =   330
         Left            =   9360
         TabIndex        =   39
         Top             =   220
         Width           =   975
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Height          =   330
         Left            =   6600
         TabIndex        =   38
         Top             =   220
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "Bordx Amt"
         Height          =   255
         Left            =   1605
         TabIndex        =   50
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   840
         TabIndex        =   49
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Checked Amt"
         Height          =   255
         Left            =   3705
         TabIndex        =   48
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label lblMG 
      Caption         =   "MedicaGen Float Balance :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   58
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Label lblMGAmt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   57
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "MedicaLife Float Balance :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   56
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label lblMLAmt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   55
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Claim Float Topup Request"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   225
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "frmCLMFloatTopupReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditFlag As Boolean
Dim tmpPrevFloatAmt As Boolean
Dim tmpTUNo As String
Dim intExit As Integer
Dim tmpPayCode As String
Dim tmpChkFloatAmt As Double
Dim m_blnDirection(1 To 12) As Boolean

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayor.Text, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor.Text, CboPayor.SelStart + CboPayor.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboPayor.ListCount - 1
 
        If NewText = LCase(Left(CboPayor.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayor.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
                CboPayor.Text = ValidValue
                CboPayor.SelStart = 0
                CboPayor.SelLength = Len(ValidValue)
             End If
        End If
End Sub


Private Sub CboTopupStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboTopupStatus.Text, CboTopupStatus.SelStart) + Chr(KeyAscii) + Mid(CboTopupStatus.Text, CboTopupStatus.SelStart + CboTopupStatus.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To CboTopupStatus.ListCount - 1
 
        If NewText = LCase(Left(CboTopupStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboTopupStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             If ValidCount = 1 Then
                CboTopupStatus.Text = ValidValue
                CboTopupStatus.SelStart = 0
                CboTopupStatus.SelLength = Len(ValidValue)
             End If
        End If

End Sub

Private Sub CheckAll_Click()
Dim LstCnt As Integer

Dim counttransfered As Integer
counttransfered = 0

    LstCnt = lstFTListing.ListItems.Count
    
    txtReqAmt = CDbl(0)
    lblSelectedAmt = CDbl(0)
    
        If LstCnt = 0 Then
            CheckAll.Value = 0
        Else
            If CheckAll.Value = 1 Then
                For LstCnt = 1 To LstCnt
                    If lstFTListing.ListItems(LstCnt).Checked = False And lstFTListing.ListItems(LstCnt).Text = "Requested" Then
                        'MsgBox "Request has been made for this item! ", vbCritical
                        counttransfered = counttransfered + 1
                    ElseIf lstFTListing.ListItems(LstCnt).Checked = False Then
                        lstFTListing.ListItems(LstCnt).Checked = True
                        lstFTListing.ListItems(LstCnt).Text = "Request"
                        
                        txtReqAmt = CDbl(txtReqAmt) - CDbl(lstFTListing.ListItems(LstCnt).SubItems(6))
                        lblSelectedAmt = CDbl(lblSelectedAmt) - CDbl(lstFTListing.ListItems(LstCnt).SubItems(6))
                
                    ElseIf lstFTListing.ListItems(LstCnt).Checked = True Then
                        lstFTListing.ListItems(LstCnt).Checked = False
                        lstFTListing.ListItems(LstCnt).Text = ""
                    End If
                Next LstCnt
            Else
                For LstCnt = 1 To LstCnt
                    If lstFTListing.ListItems(LstCnt).Checked = True Then
                        lstFTListing.ListItems(LstCnt).Checked = False
                        lstFTListing.ListItems(LstCnt).Text = ""
                    End If
               Next LstCnt
                txtReqAmt = CDbl(0)
                lblSelectedAmt = CDbl(0)
            End If
            
            If lstFTListing.ListItems.Count = counttransfered And counttransfered <> 0 Then
                MsgBox "All items have been transferred! ", vbCritical
            ElseIf lstFTListing.ListItems.Count <> counttransfered And counttransfered <> 0 Then
                MsgBox "Some ( " + counttransfered + " )item have been transferred! ", vbCritical
            End If
        End If

End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    CboTopupStatus.ListIndex = 0
    lstFTListing.ListItems.Clear
    lblSelectedAmt = 0
    txtReqAmt = 0
    EditFlag = False
End Sub

Private Sub cmdEdit_Click()
    EditFlag = True
    CmdEdit.Visible = False
    CmdUpdate.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdEXPtoExcel_Click()
    Screen.MousePointer = vbHourglass
    
    If lstFTListing.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstFTListing)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdPrintPV_Click()
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.WORKSHEET
Dim strsql As String
Dim rstTU As New ADODB.Recordset
Dim rst As New ADODB.Recordset
Dim TUCnt As Integer
Dim PAY As New ADODB.Recordset
Dim strsqlPay As String
Dim tmpAdd1 As String
Dim tmpAdd2 As String
Dim tmpAdd3 As String
Dim tmpPayName As String
Dim strsqlFLT As String
Dim rstFLT As New ADODB.Recordset
Dim sql2 As String
Dim sql1 As String

    Screen.MousePointer = vbHourglass

    If txtTUNo = "" Then
    
        MsgBox "Please Enter or Generate TU No", vbOKOnly + vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
                objExcel.Interactive = False
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
            objExcel.WindowState = xlMaximized
            Set objWorkbook = objExcel.Workbooks.Add(LocCardPath & "PaymentVoucher\FloatReq.xlt")
            Set objWorksheet = objWorkbook.Sheets(1)
            objExcel.DisplayAlerts = False
            Set objWorksheet = objWorkbook.Worksheets.Item(1)
        End If
        
        If Err.Number > 0 Then
            MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
            GoTo ExitFunction
        End If
        
        On Error GoTo HANDLE_ERROR
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
      '  medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        
        tmpPayCode = IIf(Mid(txtTUNo, 1, 3) = "FRL", "XL", "XG")

        strsqlPay = "Select PAYCode, PAYCompanyName,PAYAddress1,PAYAddress2,PAYAddress3 from PAY where PAYCode = '" & tmpPayCode & "'"
        PAY.Open strsqlPay, medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        
    
        Do Until PAY.EOF
            tmpPayName = PAY!PAYCompanyName
            tmpAdd1 = PAY!PAYAddress1
            tmpAdd2 = PAY!PAYAddress2
            tmpAdd3 = PAY!PAYAddress3
        PAY.MoveNext
        Loop
        PAY.Close
        Set PAY = Nothing
        
        strsqlFLT = ""
        strsqlFLT = "Select CLMFTNo, CLMFTDate,CLMFTAmt,CLMNoodBordx,CLMTUNo,CLMTUDate from CLMFloatTransfer where CLMTUNo = '" & Trim(txtTUNo) & "'"
        rstTU.Open strsqlFLT, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
        
        TUCnt = 0
        Do While Not rstTU.EOF
            TUCnt = TUCnt + 1
            'If TUCnt = 1 Then
            '    objWorksheet.Cells(9, "C") = tmpPayName
            '    objWorksheet.Cells(10, "C") = IIf(Trim(tmpAdd1) <> "", Trim(tmpAdd1), "")
            '    objWorksheet.Cells(11, "C") = IIf(Trim(tmpAdd2) <> "", Trim(tmpAdd2), "")
            '    objWorksheet.Cells(12, "C") = IIf(Trim(tmpAdd3) <> "", Trim(tmpAdd3), "")
            '    objWorksheet.Cells(9, "J") = rstTU!CLMTUNo
            '    objWorksheet.Cells(10, "J") = rstTU!CLMTUDate
            '    If tmpPAYCode = "XL" Then
            '        objWorksheet.Cells(35, "C") = Format(lblMLAmt, "#,##0.00")
            '    Else
            '        objWorksheet.Cells(35, "C") = Format(lblMLAmt, "#,##0.00")
            '    End If
            '    objWorksheet.Cells(39, "C") = Logon
            'End If
            
            objWorksheet.Cells(4 + CDbl(TUCnt), "A") = TUCnt
            objWorksheet.Cells(4 + CDbl(TUCnt), "B") = rstTU!CLMTUNo
            objWorksheet.Cells(4 + CDbl(TUCnt), "C") = Format(rstTU!CLMTUDate, "dd/mm/yyyy")
            objWorksheet.Cells(4 + CDbl(TUCnt), "D") = rstTU!CLMFTNo
            objWorksheet.Cells(4 + CDbl(TUCnt), "E") = Format(rstTU!CLMFTDate, "dd/mm/yyyy")
            objWorksheet.Cells(4 + CDbl(TUCnt), "F") = rstTU!CLMNoodBordx
            objWorksheet.Cells(4 + CDbl(TUCnt), "G") = Format(rstTU!CLMFTAmt, "#,##0.00")

            rstTU.MoveNext
        Loop
        rstTU.Close
        Set rstTU = Nothing
        
        Set objWorksheet = objWorkbook.Sheets(2)
        Set objWorksheet = objWorkbook.Worksheets.Item(2)

        sql1 = "SELECT CLMTUNo, CLMTUDate, SUM(CLMFTAmt) AS CLMTUAmt,Sum(CLMNoodBordx) AS CLMNoodBordx From CLMFloatTransfer where CLMTUNo = '" & Trim(txtTUNo) & "'"
        sql2 = " GROUP BY CLMTUNo,CLMTUDate "
        sql1 = sql1 + sql2
        rst.Open sql1, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
        'rstTU.Open strsqlFLT, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
        'TUCnt = 0
        Do While Not rst.EOF
            'TUCnt = TUCnt + 1
            'If TUCnt = 1 Then
                objWorksheet.Cells(9, "C") = tmpPayName
                objWorksheet.Cells(10, "C") = IIf(Trim(tmpAdd1) <> "", Trim(tmpAdd1), "")
                objWorksheet.Cells(11, "C") = IIf(Trim(tmpAdd2) <> "", Trim(tmpAdd2), "")
                objWorksheet.Cells(12, "C") = IIf(Trim(tmpAdd3) <> "", Trim(tmpAdd3), "")
                objWorksheet.Cells(9, "J") = rst!CLMTUNo
                objWorksheet.Cells(10, "J") = rst!CLMTUDate
               If tmpPayCode = "XL" Then
                   objWorksheet.Cells(35, "C") = Format(lblMLAmt, "#,##0.00")
               Else
                   objWorksheet.Cells(35, "C") = Format(lblMLAmt, "#,##0.00")
               End If
               objWorksheet.Cells(39, "C") = Logon
            'End If
            'objWorksheet.Cells(16 + CDbl(TUCnt), "A") = TUCnt
            'objWorksheet.Cells(16 + CDbl(TUCnt), "B") = rst!CLMTUNo
            'objWorksheet.Cells(16 + CDbl(TUCnt), "C") = Format(rst!CLMTUDate, "dd/MM/yyyy")
            objWorksheet.Cells(17, "B") = rst!CLMNoodBordx
            objWorksheet.Cells(17, "J") = Format(rst!CLMTUAmt, "#,##0.00")

            rst.MoveNext
        Loop
        rst.Close
        Set rst = Nothing

        
        Set objWorksheet = objWorkbook.Sheets(3)
        Set objWorksheet = objWorkbook.Worksheets.Item(3)
        
        strsqlFLT = ""
        strsqlFLT = "Select CLMFTNo, CLMFTDate,CLMFTAmt,CLMNoodBordx,CLMTUNo,CLMTUDate,BordxRefNo,BordxDate,BordxClaimAmt,CLMPayorAppCode,CLMPayorAppDate from View_TU_Listing where CLMTUNo = '" & Trim(txtTUNo) & "'"
        rstTU.Open strsqlFLT, medixmasterconnWS, adOpenForwardOnly, adLockBatchOptimistic
        
        TUCnt = 0
        Do While Not rstTU.EOF
            TUCnt = TUCnt + 1
            objWorksheet.Cells(3 + CDbl(TUCnt), "A") = TUCnt
            objWorksheet.Cells(3 + CDbl(TUCnt), "B") = rstTU!BordxRefNo
            objWorksheet.Cells(3 + CDbl(TUCnt), "C") = Format(rstTU!BordxDate, "dd/mm/yyyy")
            objWorksheet.Cells(3 + CDbl(TUCnt), "D") = rstTU!CLMPayorAppCode
            objWorksheet.Cells(3 + CDbl(TUCnt), "E") = Format(rstTU!CLMPayorAppDate, "MM/dd/yyyy")
            objWorksheet.Cells(3 + CDbl(TUCnt), "F") = rstTU!BordxClaimAmt
            rstTU.MoveNext
        Loop
        rstTU.Close
        Set rstTU = Nothing
        
   '     medixmasterconn.Close
    '    Set medixmasterconn = Nothing
     '   medixmasterconnWS.Close
     '   Set medixmasterconnWS = Nothing
        'cmdCredit.Enabled = False
        objExcel.DisplayAlerts = True
        'objWorksheet.Columns.AutoFit
        objExcel.Interactive = True
        Screen.MousePointer = vbDefault
        Exit Sub
ExitFunction:
        objExcel.DisplayAlerts = True
        objExcel.Interactive = True
        objExcel.Quit
        Screen.MousePointer = vbDefault
        Exit Sub
HANDLE_ERROR:
        MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
        Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
        Set objWorksheet = Nothing
        Set objWorkbook = Nothing
      '  medixmasterconn.Close
     '   Set medixmasterconn = Nothing
      '  medixmasterconnWS.Close
      '  Set medixmasterconnWS = Nothing
        GoTo ExitFunction
End Sub

Private Sub CmdSave_Click()
Dim AA As Integer
Dim ItmChk As Boolean
Dim BB As Integer
Dim tmpSql As String
Dim startTrans As Boolean
Dim bSucess  As Boolean
Dim DiffPay As Boolean
Dim tmpFTCode As String
Dim tmpFTNo As String

    ItmChk = False
    DiffPay = False
    tmpFTCode = ""
        AA = lstFTListing.ListItems.Count
        For BB = 1 To lstFTListing.ListItems.Count
            If lstFTListing.ListItems(BB).Checked = True Then
                If ItmChk = False Then
                    ItmChk = True
                End If
                    If tmpFTCode <> "" Then
                        If tmpFTCode <> lstFTListing.ListItems(BB).SubItems(10) Then
                            DiffPay = True
                            Exit For
                        End If
                    End If
                    tmpFTCode = lstFTListing.ListItems(BB).SubItems(10)
            End If
        Next BB
    
    If AA = 0 Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    ElseIf ItmChk = False Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    ElseIf txtTUDate = "__/__/____" Then
        MsgBox "Please Enter FT Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtTUDate.SetFocus
    ElseIf Trim(CDbl(txtReqAmt)) = 0 Then
        MsgBox "Please Enter Request Amount! ", vbOKOnly + vbCritical, DialogBoxTitle
        txtReqAmt.SetFocus
    ElseIf DiffPay = True Then
        MsgBox "Different Products for this transaction is not allowed! ", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf CDbl(lblSelectedAmt) <> CDbl(txtReqAmt) Then
        MsgBox "Please check Request Amount! ", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        Screen.MousePointer = vbHourglass
       ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                
        startTrans = True
        medixmasterconnWS.beginTrans
        
            tmpTUNo = GenerateTUNo(tmpFTCode)
            txtTUNo = tmpTUNo
            tmpPayCode = IIf(Mid(tmpTUNo, 1, 3) = "FRG", "XG", "XL")
            For BB = 1 To lstFTListing.ListItems.Count
                If lstFTListing.ListItems(BB).Checked = True Then
                    tmpFTNo = lstFTListing.ListItems(BB).SubItems(4)
                    Call UpdateFTFloat(tmpFTNo)
                End If
            Next BB
            medixmasterconnWS.commitTrans
            startTrans = False
            'txtTUNo = tmpTUNo
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle
            tmpSql = ""
            tmpSql = " and  CLMTUNo = '" & Trim(txtTUNo) & "'"
            lstFTListing.ListItems.Clear
            Call CLMFloatListing(tmpSql)
        EditFlag = False
        
     '   medixmasterconnWS.Close
     '   Set medixmasterconnWS = Nothing
        'medixmasterconnPayment.Close
        'Set medixmasterconnPayment = Nothing
        Screen.MousePointer = vbDefault
        'Call ClearContraDetail
        'Call ClearRecDetail
        lblSelectedAmt = 0
        lbltotalAmt = 0
    End If
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If bSucess = True Then
        medixmasterconn.RollbackTrans
        medixmasterconnPayment.RollbackTrans
    End If
    
   ' medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
  '  MediConnCISDB.Close
  '  Set MediConnCISDB = Nothing
    MsgBox Err.Description
End Sub

Private Sub CmdSearch_Click()
On Error GoTo errRun
Dim strAppend As String
    strAppend = ""
    CmdSearch.Enabled = False
    'cmdprint.Enabled = False
    CmdClear.Enabled = False
    CmdExit.Enabled = False
    Call ClearFrameCase
    Screen.MousePointer = vbHourglass
    'medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    Call SearchRecords
 '   medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
    Screen.MousePointer = vbDefault
    CmdSearch.Enabled = True
    CmdEdit.Visible = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    Exit Sub

errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    lstFTListing.ListItems.Clear
    lblcurrent = ""
    lbltotalAmt = ""
    lblSelectedAmt = ""
    'cmdprint.Enabled = True
    CmdClear.Enabled = True
    CmdExit.Enabled = True
    CmdSearch.Enabled = True
    CmdEdit.Visible = True
   ' medixmasterconnWS.Close
   ' Set medixmasterconnWS = Nothing
End Sub

Private Function SearchRecords()
Dim sql As String
Dim strAppend As String
Dim tmpPayCode As String
    strAppend = ""
    'GrpCompany = ""
    'HOspitals = ""
    
    If Len(Trim(txtFTNoFr)) Then
        strAppend = strAppend + " AND CLMFTNo >= '" & Trim(txtFTNoFr) & "'"
    End If
    If Len(Trim(txtFTNoTo)) Then
        strAppend = strAppend + " AND CLMFTNo <= '" & Trim(txtFTNoTo) & "'"
    End If
    
    If Len(Trim(txtTUNoFr)) Then
        strAppend = strAppend + " AND  CLMTUNo >= '" & Trim(txtTUNoFr) & "'"
    End If
    
    If Len(Trim(txtTUNoTo)) Then
        strAppend = strAppend + " AND  CLMTUNo <= '" & Trim(txtTUNoTo) & "'"
    End If
            
    If Len(Trim(txtPVNoFr)) Then
        strAppend = strAppend + " AND   CLMPVNo >= '" & Trim(txtPVNoFr) & "'"
    End If
    If Len(Trim(txtPVNoTo)) Then
        strAppend = strAppend + " AND   CLMPVNo <= '" & Trim(txtPVNoTo) & "'"
    End If
    
    If Len(Trim(CboPayor)) Then
        If Mid(CboPayor, 1, 2) = "XL" Then
            tmpPayCode = "TUL"
        Else
            tmpPayCode = "TUG"
        End If
        strAppend = strAppend + " AND  Substring(CLMTUNo,1,3) = '" & Trim(tmpPayCode) & "'"
    End If
    If Len(Trim(txtFTDateFR)) > 0 And IsDate(txtFTDateFR) = True Then
        strAppend = strAppend + " And CLMFTDate >= '" & Format(txtFTDateFR, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtFTDateTo)) > 0 And IsDate(txtFTDateTo) = True Then
        strAppend = strAppend + " And CLMFTDate <= '" & Format(txtFTDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtTUDateFR)) > 0 And IsDate(txtTUDateFR) = True Then
        strAppend = strAppend + " And CLMTUDate >= '" & Format(txtTUDateFR, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtTUDateTo)) > 0 And IsDate(txtTUDateTo) = True Then
        strAppend = strAppend + " And CLMTUDate <= '" & Format(txtTUDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtPVDateFr)) > 0 And IsDate(txtPVDateFr) = True Then
        strAppend = strAppend + " And CLMPVDate >= '" & Format(txtPVDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtPVDateTo)) > 0 And IsDate(txtPVDateTo) = True Then
        strAppend = strAppend + " And CLMPVDate <= '" & Format(txtPVDateTo, "mm/dd/yyyy") & "'"
    End If

    If Len(Trim(CboTopupStatus)) Then
        If Mid(CboTopupStatus, 1, 1) = "0" Then
            strAppend = strAppend + " AND  (CLMTUNo Is null OR CLMTUNo = '') "
        ElseIf Mid(CboTopupStatus, 1, 1) = "1" Then
            strAppend = strAppend + " AND  (CLMTUNo Is NOT null) "
        End If
    End If
    
    Call CLMFloatListing(strAppend)
    Call SearchFloatAmt("")
End Function

Private Sub CmdStop_Click()
    intExit = 1
End Sub

Private Sub cmdUpdate_Click()
Dim AA As Integer
Dim ItmChk As Boolean
Dim BB As Integer
Dim tmpSql As String
Dim startTrans As Boolean
Dim bSucess  As Boolean
Dim tmpFTCode As String
Dim tmpFTNo As String

    
    If txtTUNo = "" Then
        MsgBox "Please select at least one record to be updated! ", vbCritical
    ElseIf txtTUDate = "__/__/____" Then
        MsgBox "Please Enter FT Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtTUDate.SetFocus
    ElseIf Trim(CDbl(txtReqAmt)) = 0 Then
        MsgBox "Please Enter Request Amount! ", vbOKOnly + vbCritical, DialogBoxTitle
        txtReqAmt.SetFocus
    Else
        Screen.MousePointer = vbHourglass
       ' medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
                
        startTrans = True
        medixmasterconnWS.beginTrans
        
            'tmpTUNo = GenerateTUNo(tmpFTCode)
            'txtTUNo = tmpTUNo
            'For BB = 1 To lstFTListing.listitems.Count
                tmpFTNo = lstFTListing.SelectedItem.SubItems(1)
                Call UpdateFTFloat(tmpFTNo)
            'Next BB
            medixmasterconnWS.commitTrans
            startTrans = False
            'txtTUNo = tmpTUNo
            MsgBox "Record Updated...", vbOKOnly + vbInformation, DialogBoxTitle
            tmpSql = ""
            tmpSql = " and  CLMTUNo = '" & Trim(txtTUNo) & "'"
            lstFTListing.ListItems.Clear
            Call CLMFloatListing(tmpSql)
        EditFlag = False
        CmdEdit.Visible = True
        CmdSave.Enabled = True
        
        
    '    medixmasterconnWS.Close
    '    Set medixmasterconnWS = Nothing
        'medixmasterconnPayment.Close
        'Set medixmasterconnPayment = Nothing
        Screen.MousePointer = vbDefault
        'Call ClearContraDetail
        'Call ClearRecDetail
        lblSelectedAmt = 0
        lbltotalAmt = 0
    End If
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If bSucess = True Then
        medixmasterconn.RollbackTrans
        medixmasterconnPayment.RollbackTrans
    End If
    
  '  medixmasterconn.Close
 '   Set medixmasterconn = Nothing
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
  '  MediConnCISDB.Close
  '  Set MediConnCISDB = Nothing
    MsgBox Err.Description
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
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        EditFlag = False
        Call ComboListing
        txtReqAmt = 0
        lblSelectedAmt = 0
        lbltotalAmt = 0
  '  medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset

    PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenForwardOnly, adLockBatchOptimistic
        
    Do Until PAY.EOF
        CboPayor.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing

    With CboTopupStatus
        .AddItem "0 - Not Requested"
        .AddItem "1 - Requested"
        .Text = "0 - Not Requested"
    End With
End Sub

Private Sub CLMFloatListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim sql As String
Dim Current As String
Dim Check As Boolean
Dim strsql As String
Dim RecFound As Boolean
Dim cmd1 As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Record As ListItem
Dim BordxAPPStatus As String

    Current = 0
        lstFTListing.ListItems.Clear
        lblcurrent = 0
        lblSelectedAmt = 0
        lbltotalAmt = 0
    
    strsql = " Select  CLMFTNo,CLMFTDate,CLMFTAmt, CLMPVNo, CLMPVDate,CLMTUAmt,  " & _
                "CLMRecNo, CLMRecDate,CLMTUNo,CLMTUDate,CLMTURemarks from CLMFloatTransfer " & _
                " Where 0=0 "
    strsql = strsql + strAppend
     strsql = strsql + " order by CLMFTNo asc "
    rst2.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    Do
        Do Until rst2.EOF

        RecFound = True
        
        With rst2
            If Len(Trim(!CLMTUNo)) Then
                Set Record = lstFTListing.ListItems.Add(, , "Requested")
            Else
                Set Record = lstFTListing.ListItems.Add(, , "")
            End If
            Record.SubItems(1) = IIf(Len(!CLMTUNo), !CLMTUNo, "")
            Record.SubItems(2) = IIf(Len(!CLMTUDate), Format(!CLMTUDate, "dd/mm/yyyy"), "")
            Record.SubItems(3) = Format(IIf(Len(!CLMTUAmt), !CLMTUAmt, 0), "#,##0.00")
            Record.SubItems(4) = IIf(Len(!CLMFTNo), !CLMFTNo, "")
            Record.SubItems(5) = IIf(Len(!CLMFTDate), Format(!CLMFTDate, "dd/mm/yyyy"), "")
            Record.SubItems(6) = Format(IIf(Len(!CLMFTAmt), Trim(!CLMFTAmt), 0), "#,##0.00")
            Record.SubItems(7) = IIf(Trim(Len(!CLMPVNo)), !CLMPVNo, "")
            Record.SubItems(8) = IIf(Trim(Len(!CLMPVDate)), Format(!CLMPVDate, "dd/mm/yyyy"), "")
            Record.SubItems(9) = IIf(Len(!CLMTURemarks), !CLMTURemarks, "")
            Record.SubItems(10) = IIf(Len(!CLMFTNo), Mid(!CLMFTNo, 1, 3), "")
            lbltotalAmt = CDbl(lbltotalAmt) + Format(CDbl(!CLMFTAmt), "#,##0.00")
            Current = Current + 1
            lblcurrent = Current
        End With
        rst2.MoveNext
            
        DoEvents
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
         
        Loop
       
    Loop Until Check = False
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    
    If RecFound = False Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        lblcurrent = 0
    End If
    CmdExit.Enabled = True
    CmdEdit.Enabled = True
    cmdEXPtoExcel.Enabled = True
    CmdSave.Enabled = True
    CmdClear.Enabled = True
    CmdSearch.Enabled = True
End Sub


Private Sub lstFTListing_Click()
    If lstFTListing.SelectedItem.Text = "Requested" Then
        'tmpPrevFloatAmt = lstFTListing.SelectedItem.SubItems(7)
        Call LvDisplayData
    End If
End Sub

Private Sub lstFTListing_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
    Case 4
        SortListView lstFTListing, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lstFTListing, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstFTListing, ColumnHeader.Index, ldtDateTime, m_blnDirection(9)
    Case 10
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstFTListing, ColumnHeader.Index, ldtString, m_blnDirection(11)
    Case 12
        SortListView lstFTListing, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    End Select

End Sub

Private Sub lstFTListing_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Text <> "Requested" Then
        If Item.Checked = True Then
            Item.Text = "Request"
            txtReqAmt = CDbl(txtReqAmt) + CDbl(Item.SubItems(6))
            lblSelectedAmt = CDbl(lblSelectedAmt) + CDbl(Item.SubItems(6))
        ElseIf Item.Checked = False Then
            Item.Text = ""
            If txtReqAmt > 0 Then
                txtReqAmt = CDbl(txtReqAmt) - CDbl(Item.SubItems(6))
                lblSelectedAmt = CDbl(lblSelectedAmt) - CDbl(Item.SubItems(6))
            End If
        End If
    Else
        Item.Checked = False
        MsgBox "Request has been made for this item", vbOKOnly + vbCritical, DialogBoxTitle
    End If
End Sub

Private Sub LvDisplayData(Optional Item As ListItem, Optional ItemCheck As Boolean)
Dim lstRec As ListItem

    Call ClearFrameCase
    If lstFTListing.ListItems.Count Then
        If ItemCheck = True Then
            Set lstRec = Item
            lstFTListing.SelectedItem.Selected = False
            lstRec.Selected = True
        Else
            Set lstRec = lstFTListing.SelectedItem
        End If
            txtTUNo = IIf(Len(Trim(lstRec.SubItems(1))), lstRec.SubItems(1), "")
            If Len(lstRec.SubItems(2)) Then
                txtTUDate = Format(lstRec.SubItems(2), "DD/MM/YYYY")
            End If
            If Len(lstRec.SubItems(3)) Then
                txtReqAmt = Format(lstRec.SubItems(3), "#,##0.00")
            End If
            
            TxtRemark = IIf(Len(lstRec.SubItems(9)), lstRec.SubItems(9), "")
            CmdSave.Enabled = False
    End If
End Sub

Private Sub ClearFrameCase()
    txtTUNo = ""
    txtTUDate = "__/__/____"
    txtReqAmt = 0
    TxtRemark.Text = ""
End Sub

Private Sub UpdateFTFloat(tmpFTNo As String)
Dim rst As New ADODB.Recordset
Dim strsqlFT As String
Dim savedstatus As Boolean
    strsqlFT = ""
    
    If EditFlag = False Then
        strsqlFT = " Select CLMTUNo, CLMTUDate,CLMTUAmt,CLMTURemarks,CLMTURegUser,CLMTURegDate from CLMFloatTransfer Where CLMFTNo = '" & Trim(tmpFTNo) & "'"
    Else
        strsqlFT = " Select CLMTUNo, CLMTUDate,CLMTUAmt,CLMTURemarks,CLMTURegUser,CLMTURegDate from CLMFloatTransfer Where CLMTUNo = '" & Trim(tmpFTNo) & "'"
    End If
    rst.Open strsqlFT, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        If rst.recordCount = 0 Then
            rst.AddNew
            savedstatus = False
        ElseIf rst.recordCount > 0 Then
            savedstatus = True
        End If
            
            Do While Not rst.EOF
            With rst
                !CLMTUNo = Trim(txtTUNo)
                !CLMTUDate = Format(txtTUDate, "dd/mm/yyyy")
                !CLMTUAmt = txtReqAmt
                !CLMTURemarks = IIf(Len(TxtRemark), TxtRemark, "")
                !CLMTURegUser = Logon
                !CLMTURegDate = Date
            End With
            
            rst.Update
            rst.MoveNext
            Loop
            
        rst.Close
        Set rst = Nothing
End Sub

Private Sub SearchFloatAmt(strsqlSearch As String)
Dim rst As New ADODB.Recordset
Dim strsqlFT As String
    
    strsqlFT = " Select FTPayCode, FTFloatAmt from CLMFloatAmt where 0=0"
    strsqlFT = strsqlFT + strsqlSearch
    rst.Open strsqlFT, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
        
        If rst.recordCount = 1 Then
            tmpPayCode = rst!FTPayCode
            tmpChkFloatAmt = rst!FTFloatAmt
        Else
            Do While Not rst.EOF
                If rst!FTPayCode = "XG" Then
                    lblMGAmt = Format(rst!FTFloatAmt, "#,##0.00")
                Else
                    lblMLAmt = Format(rst!FTFloatAmt, "#,##0.00")
                End If
            rst.MoveNext
            Loop
        End If
    rst.Close
    Set rst = Nothing
End Sub

Private Sub txtFTDateFR_LostFocus()
    If IsDate(txtFTDateFR) Then
    Else
        If txtFTDateFR <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtFTDateFR.SetFocus
        End If
    End If
End Sub

Private Sub txtFTDateTo_LostFocus()
    If IsDate(txtFTDateTo) Then
    Else
        If txtFTDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtFTDateTo.SetFocus
        End If
    End If
End Sub

Private Sub txtPVDateFr_LostFocus()
    If IsDate(txtPVDateFr) Then
    Else
        If txtPVDateFr <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDateFr.SetFocus
        End If
    End If
End Sub


Private Sub txtPVDateTo_LostFocus()
    If IsDate(txtPVDateTo) Then
    Else
        If txtPVDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtPVDateTo.SetFocus
        End If
    End If

End Sub

Private Sub txtTUDateFR_LostFocus()
    If IsDate(txtTUDateFR) Then
    Else
        If txtTUDateFR <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtTUDateFR.SetFocus
        End If
    End If
End Sub

Private Sub txtTUDateTo_LostFocus()
    If IsDate(txtTUDateTo) Then
    Else
        If txtTUDateTo <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtTUDateTo.SetFocus
        End If
    End If
End Sub
