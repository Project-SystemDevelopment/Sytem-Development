VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRecAllocMedixGainLoss 
   Caption         =   "MedixAllocGainLoss(Receipt)"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   11850
   Begin VB.Frame FramePv 
      Height          =   1080
      Left            =   0
      TabIndex        =   42
      Top             =   6000
      Width           =   11820
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
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   9
         Top             =   150
         Width           =   1215
      End
      Begin VB.TextBox TxtPVNo 
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
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   10
         Top             =   405
         Width           =   1215
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
         ItemData        =   "frmRecAllocMedixGainLoss.frx":0000
         Left            =   5760
         List            =   "frmRecAllocMedixGainLoss.frx":0002
         TabIndex        =   15
         Top             =   180
         Width           =   975
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
         ItemData        =   "frmRecAllocMedixGainLoss.frx":0004
         Left            =   7200
         List            =   "frmRecAllocMedixGainLoss.frx":0006
         TabIndex        =   16
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox TxtRemark 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   470
         Width           =   2415
      End
      Begin VB.TextBox txtChequeNo 
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
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   12
         Top             =   150
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtPVDate 
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Top             =   645
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtChequeDate 
         Height          =   285
         Left            =   3480
         TabIndex        =   13
         Top             =   405
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtChequeAmt 
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
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   14
         Top             =   645
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         TabIndex        =   51
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "PV. Remark"
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
         Left            =   4800
         TabIndex        =   50
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label22 
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
         Left            =   2640
         TabIndex        =   49
         Top             =   435
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "A/C Mth"
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
         Left            =   4800
         TabIndex        =   48
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label24 
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
         Left            =   6840
         TabIndex        =   47
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Chq Amt"
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
         Left            =   2640
         TabIndex        =   46
         Top             =   675
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Chq. No"
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
         Left            =   2640
         TabIndex        =   45
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "MBB Rec Date"
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
         TabIndex        =   44
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "MBB Rec No"
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
         TabIndex        =   43
         Top             =   435
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   550
      Left            =   0
      TabIndex        =   31
      Top             =   7080
      Width           =   11820
      Begin VB.TextBox lblMedixAmt 
         Height          =   285
         Left            =   4920
         TabIndex        =   56
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox LblAmt 
         Height          =   285
         Left            =   2520
         TabIndex        =   55
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Sa&ve"
         Height          =   330
         Left            =   6960
         TabIndex        =   18
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   330
         Left            =   6360
         TabIndex        =   19
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print to &File"
         Height          =   330
         Left            =   9600
         TabIndex        =   37
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   330
         Left            =   8160
         TabIndex        =   36
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Edit"
         Height          =   330
         Left            =   7560
         TabIndex        =   35
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   330
         Left            =   11160
         TabIndex        =   34
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   330
         Left            =   10560
         TabIndex        =   33
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   330
         Left            =   8880
         TabIndex        =   32
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
         Height          =   255
         Left            =   720
         TabIndex        =   40
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Checked Amt"
         Height          =   255
         Left            =   1440
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Medix Amt"
         Height          =   255
         Left            =   4080
         TabIndex        =   38
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   0
      TabIndex        =   20
      Top             =   240
      Width           =   11775
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
         Height          =   255
         Left            =   3240
         TabIndex        =   21
         Top             =   180
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox TxtWSNoTo 
         Height          =   310
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   1
         Top             =   160
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.TextBox TxtWSNo 
         Height          =   310
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   0
         Top             =   160
         Width           =   1560
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   400
         Width           =   3615
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   690
         Width           =   3615
      End
      Begin VB.ComboBox CboAllocStatus 
         Height          =   315
         Left            =   8040
         TabIndex        =   4
         Top             =   180
         Width           =   3015
      End
      Begin VB.TextBox txtPVNofr 
         Height          =   285
         Left            =   8040
         MaxLength       =   15
         TabIndex        =   5
         Top             =   465
         Width           =   1215
      End
      Begin VB.TextBox txtPVNoto 
         Height          =   285
         Left            =   9720
         MaxLength       =   15
         TabIndex        =   6
         Top             =   465
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtPVDatefr 
         Height          =   315
         Left            =   8040
         TabIndex        =   7
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPVDateto 
         Height          =   315
         Left            =   9720
         TabIndex        =   8
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label35 
         Caption         =   "Health Type"
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
         TabIndex        =   30
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Clm No"
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
         TabIndex        =   29
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "To"
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
         Left            =   3240
         TabIndex        =   28
         Top             =   160
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "To"
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
         Left            =   9360
         TabIndex        =   27
         Top             =   765
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "PV Date"
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
         Left            =   6840
         TabIndex        =   26
         Top             =   810
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "To"
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
         Left            =   9360
         TabIndex        =   25
         Top             =   525
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "PV No"
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
         Left            =   6840
         TabIndex        =   24
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Alloc Status"
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
         Left            =   6840
         TabIndex        =   23
         Top             =   225
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Payor"
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
         TabIndex        =   22
         Top             =   795
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lstMedixGainLoss 
      Height          =   4620
      Left            =   0
      TabIndex        =   41
      Top             =   1440
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8149
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
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Alloc Status"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Claim No"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Rec No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Rec Date"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Gain/Loss"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "MedixRec"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Alloc No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Alloc Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Alloc By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Chq No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Chq Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Chq Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "AC Period"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Receipt Medix Allocation (Gain/Loss)"
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
      TabIndex        =   53
      Top             =   0
      Width           =   11805
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
      TabIndex        =   52
      Top             =   7680
      Width           =   1575
   End
End
Attribute VB_Name = "frmRecAllocMedixGainLoss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intMBMAccount As Integer
Dim intExit As Integer
Public searchCriteria As String
Dim one As Boolean
Dim ScanCode As String
Dim listloop1 As Integer
Dim listloop2 As Integer
Dim Check As Boolean
Dim RecordSelected As Integer
Dim ItemCheck As Boolean
Dim USRRec As New ADODB.Recordset
Dim USRSql As String
Dim Authorize As String
Dim EditFlag As Boolean
Private m_blnDirection(1 To 15) As Boolean

Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim monthstart
Dim yearstart
Dim HLT As New ADODB.Recordset
Dim cmd As New ADODB.Command

    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHLTType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    HLT.CursorLocation = adUseClient
    HLT.CursorType = adOpenForwardOnly
    HLT.CacheSize = 100
    Set HLT = cmd.Execute

    
    Do Until HLT.EOF
        With HLT
            CboHltCode.AddItem !HLTCode & " - " & !HLTDescription
            HLT.MoveNext
        End With
    Loop
    HLT.Close
    Set HLT = Nothing

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute

    Do Until PAY.EOF
        With PAY
            CboPayor.AddItem !PAYCode & " - " & !PAYCompanyName
            PAY.MoveNext
        End With
    Loop
    PAY.Close
    Set PAY = Nothing
    
    With CboAllocStatus
        .AddItem "N - No"
        .AddItem "Y - Yes"
    End With
    
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

Private Sub CboAllocStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboAllocStatus, CboAllocStatus.SelStart) + Chr(KeyAscii) + Mid(CboAllocStatus, CboAllocStatus.SelStart + CboAllocStatus.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboAllocStatus.ListCount - 1
        
        If NewText = LCase(Left(CboAllocStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboAllocStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               CboAllocStatus = ValidValue
               CboAllocStatus.SelStart = 0
               CboAllocStatus.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub cboHLTCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboHltCode.Text, CboHltCode.SelStart) + Chr(KeyAscii) + Mid(CboHltCode.Text, CboHltCode.SelStart + CboHltCode.SelLength + 1))
    
    ValidCount = 0
    ValidValue = ""
    
    For i = 0 To CboHltCode.ListCount - 1
        
        If NewText = LCase(Left(CboHltCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboHltCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
               CboHltCode.Text = ValidValue
               CboHltCode.SelStart = 0
               CboHltCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboPayor_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPayor, CboPayor.SelStart) + Chr(KeyAscii) + Mid(CboPayor, CboPayor.SelStart + CboPayor.SelLength + 1))
   
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
               CboPayor = ValidValue
               CboPayor.SelStart = 0
               CboPayor.SelLength = Len(ValidValue)
            End If
        End If

End Sub

Private Sub ChkScan_Click()
    lstMedixGainLoss.ListItems.Clear
    If ChkScan.Value = 1 Then
        Label25.Visible = False
        txtWSNoTo.Visible = False
        ChkScan.Left = 3240
        TxtWSNo = ""
        TxtWSNo.SetFocus
    Else
        Label25.Visible = True
        txtWSNoTo.Visible = True
        ChkScan.Left = 5280
        TxtWSNo.SetFocus
    End If
End Sub

Private Sub CmdAdd_Click()
Dim ZZ As Integer
Dim ItemChk As Boolean
    
    ItemChk = False
    For ZZ = 1 To lstMedixGainLoss.ListItems.Count
        If lstMedixGainLoss.ListItems(ZZ).Checked = True Then
            ItemChk = True
        Exit For
        End If
    Next ZZ

    CmdUpdate.Enabled = False
    Call ClearFrameCase
    If ItemChk Then
        CmdSave.Enabled = True
        FramePV.Enabled = True
        txtPVNo.SetFocus
    Else
        MsgBox "Please checked at least 1 Claim  ! ", vbCritical + vbOKOnly, DialogBoxTitle
    End If
End Sub

Private Sub cmdClear_Click()
    Call ClearForm(Me)
    lstMedixGainLoss.ListItems.Clear
    lblCurrent = 0
    LblAmt = 0
    lblMedixAmt = 0
    one = False
    ChkScan.Value = 1
    EditFlag = False
End Sub

Private Sub CmdEdit_Click()
Dim ZZ As Integer
Dim ItemChk As Boolean
    EditFlag = True
    ItemChk = False
    For ZZ = 1 To lstMedixGainLoss.ListItems.Count
        If lstMedixGainLoss.ListItems(ZZ).Checked = True Then
            ItemChk = True
        Exit For
        End If
    Next ZZ

    CmdSave.Enabled = False
    CmdUpdate.Enabled = False
    FramePV.Enabled = False
    If ItemChk = False And Len(Trim(txtAllocNo)) Then
        CmdUpdate.Enabled = True
        FramePV.Enabled = True
        txtPVNo.SetFocus
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    'Screen.MousePointer = vbHourglass
        'Call ExportToExcelAnalysis1(lstMedixGainLoss)
    'Screen.MousePointer = vbDefault
    
    Screen.MousePointer = vbHourglass
        
    If lstMedixGainLoss.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstMedixGainLoss)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdSave_Click()
Dim CASNo As String
Dim CASVer As String
Dim status As String
Dim InvoiceNo As String
Dim Update
Dim ItemChecked As Boolean
Dim i As Integer
Dim MedixPaymentAmt As Double
Dim PVSql As String
Dim AllocNo  As String
    ItemChecked = False
    CASNo = ""
    CASVer = ""
    status = ""
    MedixPaymentAmt = 0
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    
    If lstMedixGainLoss.ListItems.Count = 0 Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
    ElseIf txtPVDate = "__/__/____" Then
        MsgBox "Please Enter PV Date!", vbCritical
        txtPVDate.SetFocus
    ElseIf IsDate(txtPVDate) = False Then
        MsgBox "Invalid PV Date!", vbCritical
        txtPVDate.SetFocus
    ElseIf Trim(CboACPeriodMth) = "" Then
        MsgBox "Please Keyin Account Month ! ", vbCritical
        CboACPeriodMth.SetFocus
    ElseIf Trim(CboACPeriodYr) = "" Then
        MsgBox "Please Keyin Account Year ! ", vbCritical
        CboACPeriodYr.SetFocus
    ElseIf Trim(txtChequeAmt) = "" Or IsNumeric(txtChequeAmt) = False Then
        MsgBox "Please Enter Cheque Amt!", vbCritical
        txtChequeAmt.SetFocus
    Else
        Update = MsgBox("Do you want to Save record(s) ?", vbYesNo + vbExclamation)
        If Update = vbYes Then
            
                AllocNo = GenerateAllocNo("AM")
                txtAllocNo = ""
                txtAllocNo = AllocNo
           
            For i = 1 To lstMedixGainLoss.ListItems.Count
                If lstMedixGainLoss.ListItems(i).Checked = True Then
                    ItemChecked = True
                Exit For
                End If
            Next i
            If ItemChecked = True Then
                For i = 1 To lstMedixGainLoss.ListItems.Count
                    If lstMedixGainLoss.ListItems(i).Checked = True Then
                        CASNo = Trim(Mid(lstMedixGainLoss.ListItems(i).SubItems(2), 1, IIf(InStr(1, lstMedixGainLoss.ListItems(i).SubItems(2), "-") = 0, 10, InStr(1, lstMedixGainLoss.ListItems(i).SubItems(2), "-") - 1)))
                        CASVer = Trim(Mid(lstMedixGainLoss.ListItems(i).SubItems(2), InStr(1, lstMedixGainLoss.ListItems(i).SubItems(2), "-") + 1, Len(lstMedixGainLoss.ListItems(i).SubItems(2)) - InStr(1, lstMedixGainLoss.ListItems(i).SubItems(2), "-")))
                        MedixPaymentAmt = Format(IIf(IsNumeric(lstMedixGainLoss.ListItems(i).SubItems(6)), lstMedixGainLoss.ListItems(i).SubItems(6), 0), "#,##0.00")
                        
                        Call UpdateMedixGL(CASNo, CInt(CASVer), MedixPaymentAmt)
                                                
                        CASNo = ""
                        CASVer = ""
                        status = ""
                        MedixPaymentAmt = 0
                    End If
                Next i
            MsgBox "Record Saved...", vbOKOnly + vbInformation, DialogBoxTitle

            lstMedixGainLoss.ListItems.Clear
            End If
            EditFlag = False
        End If
    End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
End Sub

Private Sub CmdSearch_Click()
Dim listrecord As Integer
    
    Screen.MousePointer = vbHourglass
    'one = False
    LblAmt = 0
    lblMedixAmt = 0
    CmdSearch.Enabled = False
        medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
        
        If ChkScan.Value = 1 Then
            If TxtWSNo = "" Then
                MsgBox "Please Enter Worksheet No", vbOKOnly + vbCritical, DialogBoxTitle
                TxtWSNo.SetFocus
            Else
            
                Screen.MousePointer = vbHourglass
                
                listrecord = lstMedixGainLoss.ListItems.Count
                If listrecord <> 0 Then
                    
                    For listloop1 = 1 To listrecord
                        Call CaseNoCheck
                    Next listloop1
                End If
                    If Check = True Then
                        MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                    ElseIf Check = False Then
                        Call SearchRecord
                    End If
                    Check = False
                Screen.MousePointer = Default
                TxtWSNo = ""
                TxtWSNo.SetFocus
            End If
            TxtWSNo.SetFocus
        Else
            Call SearchRecord
        End If
        medixmasterconnPayment.Close
        Set medixmasterconnPayment = Nothing
    CmdSearch.Enabled = True
    Screen.MousePointer = Default
End Sub

Private Sub CmdStop_Click()

End Sub

Private Sub CmdUpdate_Click()
Dim RecordSaved
Dim strsql As String
Dim ZZ As Integer
Dim ItemChk As Boolean
Dim AmtPaid As Variant
Dim MedixPaymentAmt As Double
Dim CASRemark As String
Dim TempCase As String
Dim TempVer As String
Dim TempPVno As String
Dim PVSql As String

    ItemChk = False
    medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    
    If lstMedixGainLoss.ListItems.Count = 0 Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
           
        ElseIf txtPVNo = "" Then
            MsgBox "Please Confirm PV No!", vbCritical
        ElseIf IsDate(txtPVDate) = False Then
            MsgBox "Invalid PV Date!", vbCritical
            txtPVDate.SetFocus
        ElseIf txtPVDate = "__/__/____" Then
            MsgBox "Please Enter PV Date! ", vbCritical
            txtPVDate.SetFocus
        ElseIf Trim(txtChequeAmt) = "" Or IsNumeric(txtChequeAmt) = False Then
            MsgBox "Please Enter Cheque Amt!", vbCritical
            txtChequeAmt.SetFocus
        Else
            RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                TempCase = Trim(Mid(lstMedixGainLoss.SelectedItem.SubItems(2), 1, IIf(InStr(1, lstMedixGainLoss.SelectedItem.SubItems(2), "-") = 0, 10, InStr(1, lstMedixGainLoss.SelectedItem.SubItems(2), "-") - 1)))
                TempVer = Trim(Mid(lstMedixGainLoss.SelectedItem.SubItems(2), InStr(1, lstMedixGainLoss.SelectedItem.SubItems(2), "-") + 1, Len(lstMedixGainLoss.SelectedItem.SubItems(2)) - InStr(1, lstMedixGainLoss.SelectedItem.SubItems(2), "-")))
                Call UpdateMedixGL(TempCase, CInt(TempVer), "0")
            End If
            MsgBox "Record updated !", vbOKOnly
            EditFlag = False
            CmdUpdate.Enabled = False
            Call ClearFrameCase
            lstMedixGainLoss.ListItems.Clear
    End If
    medixmasterconnPayment.Close
    Set medixmasterconnPayment = Nothing
End Sub

Private Sub Form_Load()
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call ComboListing
    medixmasterconn.Close
    Set medixmasterconn = Nothing
End Sub

Private Sub ClearFrameCase()
    txtAllocNo = ""
    txtPVNo = ""
    txtPVDate = "__/__/____"
    txtChequeAmt = ""
    txtChequeNo = ""
    txtChequeDate.Text = "__/__/____"
    txtRemark = ""
    CboACPeriodMth.Text = ""
    CboACPeriodYr.Text = ""
    FramePV.Enabled = False
End Sub

Private Sub lstMedixGainLoss_Click()
Dim rst3 As New ADODB.Recordset
Dim strsql As String
    ItemCheck = False
    CmdSave.Enabled = False
    EditFlag = False
    Call LvDisplayData(, False)

End Sub

Private Sub lstMedixGainLoss_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
    Case 5
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
    Case 9
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
    Case 12
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
    Case 13
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstMedixGainLoss, ColumnHeader.Index, ldtString, m_blnDirection(14)
    End Select
End Sub

Private Sub lstMedixGainLoss_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    CmdSave.Enabled = True
    one = False
    ItemCheck = True
    Call ProcessCheckedItem(Item)
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

Private Function CaseNoCheck()
    If Trim(ChkStr(TxtWSNo)) = Trim(Mid(lstMedixGainLoss.ListItems(listloop1).SubItems(2), 1, IIf(InStr(1, lstMedixGainLoss.ListItems(listloop1).SubItems(2), "-") = 0, 10, InStr(1, lstMedixGainLoss.ListItems(listloop1).SubItems(2), "-") - 1))) Then
        Check = True
    End If
End Function

Private Function SearchRecord()
Dim strsql As String
Dim StrAppend As String
Dim sql As String
    
    StrAppend = ""
        
    If Len(Trim(CboHltCode)) Then
        StrAppend = StrAppend + " And HLTCode = '" & Mid(CboHltCode, 1, 1) & "'"
    End If
    
    If ChkScan.Value = 1 Then
        StrAppend = StrAppend + " AND CASNumber = '" & Trim(TxtWSNo) & "'"
    Else
        If Len(TxtWSNo) Then
            StrAppend = StrAppend + " AND CASNumber >= '" & Trim(TxtWSNo) & "'"
        End If
    
        If Len(txtWSNoTo) Then
            StrAppend = StrAppend + " AND CASNumber <= '" & Trim(txtWSNoTo) & "'"
        End If
    End If

    If Len(Trim(CboPayor)) Then
        StrAppend = StrAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If Len(txtPVNoFr) Then
        StrAppend = StrAppend + " and PAYPVNo >= '" & Trim(txtPVNoFr) & "'"
    End If
        
    If Len(txtPVNoTo) Then
        StrAppend = StrAppend + " and PAYPVNo <= '" & Trim(txtPVNoTo) & "'"
    End If
    
    If (txtPVDateFr) <> "__/__/____" And IsDate(txtPVDateFr) = True _
        And (txtPVDateTo) <> "__/__/____" And IsDate(txtPVDateTo) = True Then
        sql = ""
        sql = " AND PayPVDate >= '" & Format(Trim(txtPVDateFr), "mm/dd/yyyy") & _
              "' And  PayPVDate <= '" & Format(Trim(txtPVDateTo), "mm/dd/yyyy") & "'"
        StrAppend = StrAppend + sql
    End If
    
    If Len(Trim(txtPVDateFr)) > 0 And IsDate(txtPVDateFr) = True Then
        StrAppend = StrAppend + " And PayPVDate >= '" & Format(txtPVDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtPVDateTo)) > 0 And IsDate(txtPVDateTo) = True Then
        StrAppend = StrAppend + " And PayPVDate <= '" & Format(txtPVDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(CboAllocStatus) Then
        If Mid(CboAllocStatus, 1, 1) = "Y" Then
            StrAppend = StrAppend + " And PAYAllocStatus = '" & Mid(CboAllocStatus, 1, 1) & "'"
        Else
            StrAppend = StrAppend + " And (PAYAllocStatus = '' or PAYAllocStatus is null) "
        End If
    End If

    searchCriteria = StrAppend
    Call MedixGLListing(StrAppend)
   
End Function

Private Sub MedixGLListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim strsqlPAY As String
Dim PAY As New ADODB.Recordset
Dim INVAmt As Variant
Dim CasNumber As String
Dim WSVersion As String
Dim Record As ListItem
Dim sql As String
Dim total As String
Dim Current As String
Dim Check As Boolean
Dim RECFull As Boolean
Dim RECPartial As String
Dim MedixPaidMBM As Double
Dim MedixPaidHOS As Double
Dim Paid2Medix  As Double
Dim Paid2Others As Double
Dim TotalPaid As Double
Dim BankCharges As Double
Dim RecMBM As Double
Dim RecMNRB As Double
Dim RecOthers As Double
Dim totalREc As Double
Dim CheckTotal As Double
Dim RecMedix As Double
Dim Contra As Double
Dim CashPos As Double
Dim Pay2HOS As Double
Dim Pay2MBM As Double
Dim TotalApp As Double
Dim Pay2Medix As Double
Dim Pay2Others As Double
Dim GainLoss As Double
    total = 0
    Current = 0
    lstMedixGainLoss.ListItems.Clear
    sql = " Select CASNumber, ClaimVersion, MBMAmt, MNRBAmt, CLMTotalApproved," & _
            " HOSPaidAmt, MBMPaidAmt, CLMPayableByMedix2H, CLMPayableByMedix2M, " & _
            "ContraCaseAmt, PaymentBankCharge, PVHOSDate,PVDate, MBMReceiptDate, " & _
            "MNRBReceiptDate, HLTCode, PAYAllocStatus, PAYAllocNo, PAYPVNo, PAYPVDate, " & _
            "PAYChqNo, PAYChqDate, PAYChqAmt, MedixRecAmt, PAYRemarks, " & _
            "PAYAccPeriodMth, PAYAccPeriodYr, PAYLastUpdateUser, PAYLastUpdateDate " & _
            "From VIEW_RecMedix_GL where 0 = 0 "

    If Len(Trim(StrAppend)) Then
        sql = sql + StrAppend
    End If
    
    rst2.Open sql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
   
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
        Do Until rst2.EOF
        
            With rst2
                
                If Len(rst2!MNRBAmt) Then
                    RecMNRB = rst2!MNRBAmt
                Else
                    RecMNRB = 0
                End If

                If Len(rst2!MBMAmt) Then
                    RecMBM = rst2!MBMAmt
                Else
                    RecMBM = 0
                End If
                
                If Len(rst2!MedixRecAmt) Then
                    RecMedix = Format(rst2!MedixRecAmt, "#,#.00;(#,#.00)")
                Else
                    RecMedix = 0
                End If
                
                'If Len(rst2!OthersReceipt) Then
                '    Record.SubItems(16) = Format(rst2!OthersReceipt, "#,#.00;(#,#.00)")
                    'RecMedix = Format(rst2!OthersReceipt, "#,#.00;(#,#.00)")
                'Else
                    RecOthers = 0
                'End If
                
                totalREc = CDbl(RecMNRB) + CDbl(RecMBM) + CDbl(RecMedix) + CDbl(RecOthers)
                
                
                TotalApp = IIf(Len(rst2!CLMTotalApproved), rst2!CLMTotalApproved, 0)
                GainLoss = Format(CDbl(totalREc) - CDbl(TotalApp), "#,##0.00;(#,##0.00)")
                
                If CInt(GainLoss) <> "0" Or Len(rst2!MedixRecAmt) Then
                    Set Record = lstMedixGainLoss.ListItems.Add(, , "")
                    
                    If Trim(Len(!PAYAllocStatus)) Then
                        Record.SubItems(1) = "Allocated"
                    Else
                        Record.SubItems(1) = ""
                    End If
                    
                    If Len(rst2!CasNumber & "-" & rst2!claimversion) Then
                        Record.SubItems(2) = Trim(rst2!CasNumber & "-" & rst2!claimversion)
                    End If
                    
                    If Len(rst2!PayPVNo) Then
                        Record.SubItems(3) = Trim(rst2!PayPVNo)
                    End If
                    
                    If Len(rst2!PayPVDate) Then
                        Record.SubItems(4) = Trim(rst2!PayPVDate)
                    End If
                    If Trim(Len(rst2!PAYAllocStatus)) Then
                        Record.SubItems(5) = 0
                    Else
                        If Len(GainLoss) Then
                            Record.SubItems(5) = Format(GainLoss, "#,##0.00;(#,##0.00)")
                        End If
                    End If
                    If Len(rst2!MedixRecAmt) Then
                        Record.SubItems(6) = Format(rst2!MedixRecAmt, "#,##0.00;(#,##0.00)")
                        lblMedixAmt = CDbl(lblMedixAmt) + CDbl(rst2!MedixRecAmt)
                    End If
                    
                    If Len(rst2!PAYAllocNo) Then
                        Record.SubItems(7) = rst2!PAYAllocNo
                    End If
                    
                    If Len(rst2!PayLastUpdateDate) Then
                        Record.SubItems(8) = rst2!PayLastUpdateDate
                    End If
                    
                    If Len(rst2!PAYLastUpdateUser) Then
                        Record.SubItems(9) = rst2!PAYLastUpdateUser
                    End If
                    
                    If Len(rst2!PAYChqNo) Then
                        Record.SubItems(10) = rst2!PAYChqNo
                    End If
                    
                    If Len(rst2!PAYChqDate) Then
                        Record.SubItems(11) = rst2!PAYChqDate
                    End If
                    
                    If Len(rst2!PAYChqAmt) Then
                        Record.SubItems(12) = rst2!PAYChqAmt
                    End If
                    
                    If Len(rst2!PayAccPeriodMth) Then
                        Record.SubItems(13) = rst2!PayAccPeriodMth & "/" & rst2!PayAccPeriodYr
                    End If
                    
                    If Len(rst2!PayRemarks) Then
                        Record.SubItems(14) = rst2!PayRemarks
                    End If
                    Current = Current + 1
                    
                End If
                
                lblCurrent = Current
            End With
            rst2.MoveNext

        Loop

    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
End Sub

Public Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
    
    If Item.Checked = True Then
        If Trim(Item.SubItems(1)) = "Allocated" Then
            MsgBox "Record has been allocated!", vbCritical
            Item.Checked = False
        Else
            Item.SubItems(1) = "Allocated"
            Item.SubItems(6) = Format(Item.SubItems(5) * -1, "#,##0.00;(#,##0.00)")
            LblAmt = LblAmt + CDbl(Format(Item.SubItems(5), "#,##0.00"))
        End If
    Else
        LblAmt = LblAmt - CDbl(Format(Item.SubItems(5), "#,##0.00"))
        Item.SubItems(1) = ""
        Item.SubItems(6) = ""
    End If
End Function

Private Sub UpdateMedixGL(CASNo As String, CASVer As Integer, MedixPay As Double)
Dim Update
Dim strsql As String
Dim PAY As New ADODB.Recordset
Dim strsqlPAY As String
Dim CLM As New ADODB.Recordset
Dim strsqlCLM As String
Dim AddNewRec As Boolean
    
    AddNewRec = False
    strsqlPAY = " Select PAYIndex, * from ReceiptMedixGL " & _
                " Where PAYClaimNo = '" & Trim(CASNo) & "' AND PAYCLMVersion = '" & Trim(CASVer) & "'"
    PAY.Open strsqlPAY, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    
        If PAY.recordCount = 0 Then
             PAY.AddNew
        End If
            With PAY
                !PAYClaimNo = Trim(CASNo)
                !PAYCLMVersion = Trim(CASVer)
                !PayPVNo = Trim(txtPVNo)
                !PayPVDate = Format(txtPVDate, "dd/mm/yyyy")
                If txtChequeDate <> "__/__/____" Then
                    !PAYChqDate = Format(txtChequeDate, "dd/mm/yyyy")
                Else
                    !PAYChqDate = ""
                End If
                !PAYChqAmt = Format(IIf(IsNumeric(Trim(txtChequeAmt)), Trim(txtChequeAmt), 0), "#,##0.00")
                !PAYChqNo = IIf(Len(Trim(txtChequeNo)), Trim(txtChequeNo), "")
                !PAYAllocNo = IIf(Len(Trim(txtAllocNo)), Trim(txtAllocNo), "")
                If EditFlag = False Then
                    !PAYAllocStatus = "Y"
                    !MedixRecAmt = MedixPay
                End If
                !PayRemarks = IIf(Len(Trim(txtRemark)), Trim(txtRemark), "")
                !PayAccPeriodMth = IIf(Len(Trim(CboACPeriodMth)), Trim(CboACPeriodMth), "")
                !PayAccPeriodYr = IIf(Len(Trim(CboACPeriodYr)), Trim(CboACPeriodYr), "")
                !PayLastUpdateDate = Date
                !PAYLastUpdateUser = Logon
            End With
            PAY.Update
        
        PAY.Close
        Set PAY = Nothing
    LblAmt = 0
End Sub

Private Sub LvDisplayData(Optional Item As ListItem, Optional ItemCheck As Boolean)
Dim lstRec As ListItem
Dim TmpAmtPaid As Double

    Call ClearFrameCase
    If lstMedixGainLoss.ListItems.Count Then
        If ItemCheck = True Then
            Set lstRec = Item
            lstMedixGainLoss.SelectedItem.Selected = False
            lstRec.Selected = True
        Else
            Set lstRec = lstMedixGainLoss.SelectedItem
        End If
            txtPVNo = IIf(Len(Trim(lstRec.SubItems(3))), lstRec.SubItems(3), "")
            If Len(lstRec.SubItems(4)) Then
                txtPVDate = Format(lstRec.SubItems(4), "DD/MM/YYYY")
            End If
            
            If Len(lstRec.SubItems(7)) Then
                txtAllocNo = lstRec.SubItems(7)
            End If
            If Len(lstRec.SubItems(10)) Then
                txtChequeNo = lstRec.SubItems(10)
            End If
            If IsDate(lstRec.SubItems(11)) Then
                txtChequeDate = Format(lstRec.SubItems(11), "DD/MM/YYYY")
            End If
            If IsNumeric(lstRec.SubItems(12)) Then
                txtChequeAmt = lstRec.SubItems(12)
            End If
            
            If Len(lstRec.SubItems(13)) Then
                CboACPeriodMth = Mid(lstRec.SubItems(13), 1, 2)
                CboACPeriodYr = Mid(lstRec.SubItems(13), 4, 4)
            End If
            txtRemark = IIf(Len(lstRec.SubItems(14)), lstRec.SubItems(14), "")
            CmdSave.Enabled = False
    End If
End Sub
