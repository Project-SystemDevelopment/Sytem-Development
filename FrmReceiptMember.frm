VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReceiptMember 
   Caption         =   "Receipt From Member"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10530
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   10530
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1980
      Left            =   0
      TabIndex        =   55
      Top             =   240
      Width           =   11775
      Begin VB.CheckBox ChkScan 
         Caption         =   "Use Scan"
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
         Left            =   5400
         TabIndex        =   2
         Top             =   280
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtCaseNo2 
         Height          =   285
         Left            =   3840
         TabIndex        =   1
         Top             =   300
         Width           =   1510
      End
      Begin VB.TextBox TxtCaseNo1 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   300
         Width           =   1510
      End
      Begin VB.TextBox txtRECNoTo 
         Height          =   285
         Left            =   9960
         MaxLength       =   20
         TabIndex        =   9
         Top             =   300
         Width           =   1215
      End
      Begin VB.CheckBox ChkPeriodYr 
         Height          =   255
         Left            =   11235
         TabIndex        =   25
         Top             =   1665
         Width           =   375
      End
      Begin VB.CheckBox chkRecpDate 
         Height          =   255
         Left            =   11235
         TabIndex        =   13
         Top             =   540
         Width           =   375
      End
      Begin VB.CheckBox chkDRNo 
         Height          =   255
         Left            =   11235
         TabIndex        =   16
         Top             =   810
         Width           =   375
      End
      Begin VB.CheckBox chkDRAmt 
         Height          =   255
         Left            =   11235
         TabIndex        =   19
         Top             =   1095
         Width           =   375
      End
      Begin VB.CheckBox chkPeriodMth 
         Height          =   255
         Left            =   11235
         TabIndex        =   22
         Top             =   1350
         Width           =   375
      End
      Begin VB.CheckBox ChkRecpNo 
         Height          =   255
         Left            =   11235
         TabIndex        =   10
         Top             =   270
         Width           =   375
      End
      Begin VB.TextBox txtMBMNo 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   3
         Top             =   540
         Width           =   3560
      End
      Begin VB.TextBox TxtMBMName 
         Height          =   285
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   4
         Top             =   795
         Width           =   3560
      End
      Begin VB.TextBox txtRECNoFr 
         Height          =   285
         Left            =   8280
         MaxLength       =   20
         TabIndex        =   8
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox txtRECDateFr 
         Height          =   285
         Left            =   8280
         MaxLength       =   20
         TabIndex        =   11
         Top             =   555
         Width           =   1215
      End
      Begin VB.TextBox txtDRNoFr 
         Height          =   285
         Left            =   8280
         MaxLength       =   20
         TabIndex        =   14
         Top             =   810
         Width           =   1215
      End
      Begin VB.TextBox txtRECDateTo 
         Height          =   285
         Left            =   9960
         MaxLength       =   20
         TabIndex        =   12
         Top             =   555
         Width           =   1215
      End
      Begin VB.TextBox txtDrNoTo 
         Height          =   285
         Left            =   9960
         MaxLength       =   20
         TabIndex        =   15
         Top             =   810
         Width           =   1215
      End
      Begin VB.TextBox txtDRAmtFr 
         Height          =   285
         Left            =   8280
         MaxLength       =   20
         TabIndex        =   17
         Top             =   1065
         Width           =   1215
      End
      Begin VB.TextBox txtDRAmtTo 
         Height          =   285
         Left            =   9960
         MaxLength       =   20
         TabIndex        =   18
         Top             =   1065
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodMthFr 
         Height          =   315
         ItemData        =   "FrmReceiptMember.frx":0000
         Left            =   8280
         List            =   "FrmReceiptMember.frx":0002
         TabIndex        =   20
         Top             =   1305
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodMthTo 
         Height          =   315
         ItemData        =   "FrmReceiptMember.frx":0004
         Left            =   9960
         List            =   "FrmReceiptMember.frx":0006
         TabIndex        =   21
         Top             =   1305
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodYRTo 
         Height          =   315
         ItemData        =   "FrmReceiptMember.frx":0008
         Left            =   9960
         List            =   "FrmReceiptMember.frx":000A
         TabIndex        =   24
         Top             =   1590
         Width           =   1215
      End
      Begin VB.ComboBox cboACPeriodYRFr 
         Height          =   315
         ItemData        =   "FrmReceiptMember.frx":000C
         Left            =   8280
         List            =   "FrmReceiptMember.frx":000E
         TabIndex        =   23
         Top             =   1590
         Width           =   1215
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         ItemData        =   "FrmReceiptMember.frx":0010
         Left            =   1800
         List            =   "FrmReceiptMember.frx":0012
         TabIndex        =   5
         Top             =   1050
         Width           =   3560
      End
      Begin VB.ComboBox CboPayor 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   1290
         Width           =   3560
      End
      Begin VB.ComboBox CboPayment 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1575
         Width           =   3560
      End
      Begin VB.Label Label38 
         Caption         =   "Health Type"
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
         Left            =   240
         TabIndex        =   90
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label30 
         Caption         =   "To"
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
         Left            =   3480
         TabIndex        =   89
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label29 
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
         Left            =   240
         TabIndex        =   88
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   11235
         TabIndex        =   87
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label19 
         Caption         =   "To"
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
         Left            =   9600
         TabIndex        =   86
         Top             =   1110
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "DR Amt From"
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
         Left            =   6720
         TabIndex        =   85
         Top             =   1110
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "A/C Period Mth"
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
         Left            =   6720
         TabIndex        =   82
         Top             =   1395
         Width           =   1815
      End
      Begin VB.Label Label23 
         Caption         =   "To"
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
         Left            =   9600
         TabIndex        =   81
         Top             =   1395
         Width           =   375
      End
      Begin VB.Label Label24 
         Caption         =   "To"
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
         Left            =   9600
         TabIndex        =   80
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label25 
         Caption         =   "A/C Period Yr"
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
         Left            =   6720
         TabIndex        =   79
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "To"
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
         Index           =   0
         Left            =   9600
         TabIndex        =   78
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "Recp No from"
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
         Left            =   6720
         TabIndex        =   77
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Recp Date from"
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
         Left            =   6720
         TabIndex        =   76
         Top             =   555
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "To"
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
         Left            =   9600
         TabIndex        =   75
         Top             =   615
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "DR No From"
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
         Left            =   6720
         TabIndex        =   74
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "To"
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
         Left            =   9600
         TabIndex        =   73
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Payor"
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
         Left            =   240
         TabIndex        =   72
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Payment Status"
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
         Left            =   240
         TabIndex        =   71
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Mem No"
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
         Left            =   240
         TabIndex        =   57
         Top             =   585
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Mem Name"
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
         Left            =   240
         TabIndex        =   56
         Top             =   840
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lstReceiptMBM 
      Height          =   2985
      Left            =   0
      TabIndex        =   63
      Top             =   2280
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5265
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   22
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Check"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Reference No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Pymt Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "WKS No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amt Rec'd"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "OS Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Amt Receiving"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ReceiptIndex"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Rept No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Rept Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Acc Period"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Bank Charge"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Contra Case No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "Contra Amt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "Clm Gain Loss"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Misc Diff"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Text            =   "Discount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Text            =   "Amt Paid"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Text            =   "Write Off"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Text            =   "Medix Receipt"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Remarks"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   1695
      Left            =   0
      TabIndex        =   58
      Top             =   5280
      Width           =   6555
      Begin VB.TextBox txtReceiptNo 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtReceiptDate 
         Height          =   285
         Left            =   1320
         TabIndex        =   27
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtAmt 
         Height          =   285
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   28
         Top             =   740
         Width           =   1215
      End
      Begin VB.TextBox txtChequeNo 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   29
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox cboAccPeriodMth 
         Height          =   315
         ItemData        =   "FrmReceiptMember.frx":0014
         Left            =   3720
         List            =   "FrmReceiptMember.frx":0016
         TabIndex        =   30
         Top             =   200
         Width           =   1095
      End
      Begin VB.ComboBox cboAccPeriodYR 
         Height          =   315
         ItemData        =   "FrmReceiptMember.frx":0018
         Left            =   5640
         List            =   "FrmReceiptMember.frx":001A
         TabIndex        =   31
         Top             =   200
         Width           =   855
      End
      Begin VB.TextBox TxtRecRemarks 
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   3720
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   480
         Width           =   2745
      End
      Begin VB.Label Label12 
         Caption         =   "A/C Period Mth"
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
         TabIndex        =   84
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label26 
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
         Left            =   5160
         TabIndex        =   83
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
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
         Left            =   2640
         TabIndex        =   68
         Top             =   520
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Amount"
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
         TabIndex        =   62
         Top             =   800
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Cheque No"
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
         TabIndex        =   61
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Receipt Date"
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
         TabIndex        =   60
         Top             =   520
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Receipt No"
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
         TabIndex        =   59
         Top             =   250
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1695
      Left            =   6600
      TabIndex        =   91
      Top             =   5280
      Width           =   5200
      Begin VB.CommandButton CmdAdd 
         Caption         =   "Update &Case"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3960
         TabIndex        =   45
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox TxtContraVer 
         Height          =   285
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   36
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox TxtClaimNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   33
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox TxtBankCharge 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   34
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxtMisc 
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   40
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox TxtDiscount 
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   41
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxtAmtPaid 
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   42
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TxtRemarks 
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2400
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Top             =   1200
         Width           =   1320
      End
      Begin VB.TextBox TxtContraCase 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   35
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TxtContraAmt 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   37
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TxtGainLoss 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   38
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TxtWriteOff 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   39
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox TxtMedixRec 
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   43
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label45 
         Caption         =   "Medix Rec."
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
         Left            =   3120
         TabIndex        =   104
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label44 
         Caption         =   "Write Off"
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
         TabIndex        =   102
         Top             =   1400
         Width           =   1095
      End
      Begin VB.Label Label43 
         Caption         =   "--"
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
         Left            =   2350
         TabIndex        =   101
         Top             =   650
         Width           =   135
      End
      Begin VB.Label Label42 
         Caption         =   "Contra Case"
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
         TabIndex        =   100
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label41 
         Caption         =   "Claim No"
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
         TabIndex        =   99
         Top             =   165
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Bank Chrg"
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
         TabIndex        =   98
         Top             =   405
         Width           =   975
      End
      Begin VB.Label Label32 
         Caption         =   "Contra Amt"
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
         TabIndex        =   97
         Top             =   920
         Width           =   1095
      End
      Begin VB.Label Label34 
         Caption         =   "Clm Gain/Loss"
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
         TabIndex        =   96
         Top             =   1120
         Width           =   1095
      End
      Begin VB.Label Label35 
         Caption         =   "Misc. Diff."
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
         Left            =   3120
         TabIndex        =   95
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label36 
         Caption         =   "Discount"
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
         Left            =   3120
         TabIndex        =   94
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label37 
         Caption         =   "Amt Paid"
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
         Left            =   3120
         TabIndex        =   93
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label40 
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
         Left            =   2400
         TabIndex        =   92
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   795
      Left            =   0
      TabIndex        =   64
      Top             =   6960
      Width           =   11775
      Begin VB.TextBox lblReceivedAmt 
         Height          =   285
         Left            =   3600
         TabIndex        =   108
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtCHKAmt 
         Height          =   285
         Left            =   3600
         TabIndex        =   107
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox lblCurRecord 
         Height          =   375
         Left            =   240
         TabIndex        =   106
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdPrintPV 
         Caption         =   "Print &OR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9600
         TabIndex        =   52
         Top             =   300
         Width           =   855
      End
      Begin VB.CommandButton cmdEXPtoExcel 
         Caption         =   "Print to &File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8640
         TabIndex        =   51
         Top             =   300
         Width           =   975
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6600
         TabIndex        =   48
         Top             =   300
         Width           =   735
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6120
         TabIndex        =   47
         Top             =   300
         Width           =   495
      End
      Begin VB.TextBox txtPVStatus 
         Height          =   285
         Left            =   120
         TabIndex        =   69
         Top             =   600
         Visible         =   0   'False
         Width           =   375
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
         Height          =   330
         Left            =   8040
         TabIndex        =   50
         Top             =   300
         Width           =   615
      End
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
         Height          =   330
         Left            =   10440
         TabIndex        =   53
         Top             =   300
         Width           =   615
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
         Height          =   330
         Left            =   11040
         TabIndex        =   54
         Top             =   300
         Width           =   615
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "S&earch"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7320
         TabIndex        =   49
         Top             =   300
         Width           =   735
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5520
         TabIndex        =   46
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label46 
         Caption         =   "Received Amt"
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
         Left            =   2400
         TabIndex        =   105
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Checked Amt"
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
         Left            =   2400
         TabIndex        =   70
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1200
         TabIndex        =   65
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Records"
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
         Left            =   1560
         TabIndex        =   67
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "of"
         Height          =   255
         Left            =   960
         TabIndex        =   66
         Top             =   840
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Receipt from Member"
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
      TabIndex        =   103
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmReceiptMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CASNo As String
Dim CASVer As String
Dim CaseNoArray() As String
Public intExit As Integer
Public Source As Integer
Public intMBMAccount As Integer
Dim one As Boolean
Dim two As Boolean
Dim ScanCode As String
Dim listloop1 As Integer
Dim RecNumber As String
Dim RcdCheck As Boolean
Private m_blnDirection(1 To 22) As Boolean


Private Sub ChkScan_Click()
lstReceiptMBM.ListItems.Clear
If ChkScan = 1 Then
    Label30.Visible = False
    TxtCaseNo2.Visible = False
    ChkScan.Left = 3480
    TxtCaseNo1 = ""
    TxtCaseNo1.SetFocus
Else
    Label30.Visible = True
    TxtCaseNo2.Visible = True
    ChkScan.Left = 5400
    TxtCaseNo1.SetFocus
End If

End Sub

Private Sub CmdAdd_Click()
Dim AskRec
Dim AmtPaid As Double
Dim PartialAmt As Double

If IsNumeric(TxtAmtPaid) = False Then
    AmtPaid = 0
Else
    AmtPaid = TxtAmtPaid
End If

If lstReceiptMBM.ListItems.Count <> 0 Then
    If lstReceiptMBM.SelectedItem.Checked = True Then
        If (Trim(TxtContraCase) <> "" And Trim(TxtContraVer) = "") Or (Trim(TxtContraCase) = "" And Trim(TxtContraVer) <> "") Then
            MsgBox "Invalid Contra Case..., Please Keyin Contra Case No. or Contra Version ", vbCritical + vbOKOnly, DialogBoxTitle
            TxtContraCase.SetFocus
        ElseIf (Trim(TxtContraCase) <> "" And Trim(TxtContraVer) <> "" And Trim(TxtContraAmt) = "") Then
            MsgBox "Please Keyin Contra Case Amount ", vbCritical + vbOKOnly, DialogBoxTitle
            TxtContraAmt.SetFocus
        ElseIf IsNumeric(AmtPaid) = True And CDbl(AmtPaid) > CDbl(lstReceiptMBM.SelectedItem.SubItems(4)) Then
                MsgBox "Invalid Amount Paid ! ", vbCritical + vbOKOnly, DialogBoxTitle
                TxtAmtPaid.SetFocus
        Else
            AskRec = MsgBox("Update " & TxtClaimNo & " ?", vbYesNo + vbExclamation)
            
            If AskRec = vbYes Then
                If IsNumeric(TxtAmtPaid) Then
                    txtCHKAmt = txtCHKAmt - CDbl(Format(lstReceiptMBM.SelectedItem.SubItems(7), "#,##0.00"))
                    PartialAmt = CDbl(lstReceiptMBM.SelectedItem.SubItems(7))
                    lstReceiptMBM.SelectedItem.SubItems(7) = TxtAmtPaid
                    
                    If lstReceiptMBM.SelectedItem.SubItems(4) = lstReceiptMBM.SelectedItem.SubItems(7) Then
                        lstReceiptMBM.SelectedItem.SubItems(2) = "Full"
                    Else
                        lstReceiptMBM.SelectedItem.SubItems(2) = "Partial"
                        lstReceiptMBM.SelectedItem.SubItems(6) = CDbl(PartialAmt) - CDbl(lstReceiptMBM.SelectedItem.SubItems(7))
                        If CDbl(PartialAmt) = CDbl(TxtAmtPaid) Then
                            lstReceiptMBM.SelectedItem.SubItems(2) = "Full"
                        End If
                    End If
                    txtCHKAmt = txtCHKAmt + CDbl(Format(lstReceiptMBM.SelectedItem.SubItems(7), "#,##0.00"))
                End If
                lstReceiptMBM.SelectedItem.SubItems(12) = TxtBankCharge
                lstReceiptMBM.SelectedItem.SubItems(13) = (Trim(TxtContraCase) & "-" & Trim(TxtContraVer))
                lstReceiptMBM.SelectedItem.SubItems(14) = TxtContraAmt
                lstReceiptMBM.SelectedItem.SubItems(15) = TxtGainLoss
                lstReceiptMBM.SelectedItem.SubItems(16) = TxtMisc
                lstReceiptMBM.SelectedItem.SubItems(17) = TxtDiscount
                lstReceiptMBM.SelectedItem.SubItems(19) = TxtWriteOff
                lstReceiptMBM.SelectedItem.SubItems(20) = TxtMedixRec
                lstReceiptMBM.SelectedItem.SubItems(21) = TxtRemarks
                MsgBox "Updated .... ", vbInformation + vbOKOnly, DialogBoxTitle
            End If
            TxtClaimNo = ""
            TxtBankCharge = ""
            TxtContraCase = ""
            TxtContraVer = ""
            TxtContraAmt = ""
            TxtGainLoss = ""
            TxtMisc = ""
            TxtDiscount = ""
            TxtAmtPaid = ""
            TxtWriteOff = ""
            TxtMedixRec = ""
            TxtRemarks = ""
        End If
    Else
        MsgBox "Please select a record ! ", vbOKOnly, DialogBoxTitle
    End If
End If
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

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    ReDim CaseNoArray(0)
    CboPayment.ListIndex = 0
    CmdSave.Enabled = True
    one = False
    two = False
    lstReceiptMBM.ListItems.Clear
    txtCHKAmt = 0
    lblCurRecord = 0
    lblTotal = 0
    lblReceivedAmt = 0
End Sub

Private Sub CmdEdit_Click()
    one = True
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

'************************Start ComboListing**************************

Private Sub ComboListing()
Dim PAY As New Adodb.Recordset
Dim monthstart
Dim yearstart
Dim cmd As New Adodb.Command

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
    
    CboHltCode.AddItem "N - Non-Sihat"
    CboHltCode.AddItem "S - Sihat"
    
    With CboPayment
        .AddItem "0 - All"
        .AddItem "1 - Paid"
        .AddItem "2 - Unpaid"
        .AddItem "3 - Partial Paid"
        .AddItem "4 - Paid & Partial Paid"
        .AddItem "5 - Unpaid & Partial Paid"
    End With
    
    For monthstart = 1 To 12
        With cboAccPeriodMth
            .AddItem monthstart
        End With
        With cboACPeriodMthFr
            .AddItem monthstart
        End With
        With cboACPeriodMthTo
            .AddItem monthstart
        End With
    Next monthstart
    
    For yearstart = 1999 To 3000
        With cboAccPeriodYR
            .AddItem yearstart
        End With
        With cboACPeriodYRFr
            .AddItem yearstart
        End With
        With cboACPeriodYRTo
            .AddItem yearstart
        End With
    Next yearstart
    
End Sub
'************************Start ComboListing**************************

Private Function SearchRecord()
Dim StrSql As String
Dim strAppend As String
Dim sql As String
    
    strAppend = ""
    If ChkScan.Value = 1 Then
        If Len(Trim(TxtCaseNo1)) Then
            strAppend = strAppend + " And CASNUMBER ='" & ChkStr(TxtCaseNo1) & "'"
        End If
    Else
        If Len(Trim(TxtCaseNo1)) Then
            strAppend = strAppend + " And CASNUMBER >='" & ChkStr(TxtCaseNo1) & "'"
        End If
    
        If Len(Trim(TxtCaseNo2)) Then
            strAppend = strAppend + " And CASNUMBER <='" & ChkStr(TxtCaseNo2) & "'"
        End If
    End If
    
    If Len(Trim(TxtMBMNo)) Then
        strAppend = strAppend + " And MBMNumber like '%" & ChkStr(TxtMBMNo) & "%'"
    End If
    
    If Len(Trim(TxtMBMName)) Then
        strAppend = strAppend + " And MBMName like '%" & ChkStr(TxtMBMName) & "%'"
    End If
    
    If Len(Trim(CboPayor)) Then
        strAppend = strAppend + " And PAYCode = '" & Trim(Mid(CboPayor, 1, IIf(InStr(1, CboPayor, "-") = 0, 2, InStr(1, CboPayor, "-") - 1))) & "'"
    End If
    
    If Mid(CboPayment, 1, 1) = 1 Then
        strAppend = strAppend + " and PVRecMBMStatus = 'F'"
    End If
 
    If Mid(CboPayment, 1, 1) = 2 Then
      strAppend = strAppend + " and (PVRecMBMStatus = '' or PVRecMBMStatus is null) "
    End If
    
    If Mid(CboPayment, 1, 1) = 3 Then
      strAppend = strAppend + " and PVRecMBMStatus = 'P'"
    End If
    
    If Mid(CboPayment, 1, 1) = 4 Then
      strAppend = strAppend + " and (PVRecMBMStatus = 'P' or PVRecMBMStatus = 'F') "
    End If
    
    If Mid(CboPayment, 1, 1) = 5 Then
      strAppend = strAppend + " and (PVRecMBMStatus = '' or PVRecMBMStatus is null or PVRecMBMStatus = 'P') "
    End If
        
    If Len(Trim(txtRecNoFr)) Then
        strAppend = strAppend + " AND ReceiptNo >= '" & RemStr(Trim(txtRecNoFr)) & "'"
    End If
    
    If Len(Trim(txtRecNoFr)) Then
        strAppend = strAppend + " AND ReceiptNo <= '" & RemStr(Trim(txtRecNoTo)) & "'"
    End If
    
    If ChkRecpNo.Value = 1 Then
        sql = ""
        sql = " AND (ReceiptNo Is null OR ReceiptNo = '')" & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtRecDateFr)) > 0 And IsDate(txtRecDateFr) = True Then
        strAppend = strAppend + " And ReceiptDate >= '" & Format(txtRecDateFr, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtRecDateTo)) > 0 And IsDate(txtRecDateTo) = True Then
        strAppend = strAppend + " And ReceiptDate <= '" & Format(txtRecDateTo, "mm/dd/yyyy") & "'"
    End If
    
    If Len(txtDRNoFr) Then
        strAppend = strAppend + " AND DebitCreditRefNo >= '" & Trim(txtDRNoFr) & "'"
    End If

    If Len(txtDrNoTo) Then
        strAppend = strAppend + " AND DebitCreditRefNo <= '" & Trim(txtDrNoTo) & "'"
    End If
    
    If chkDRNo.Value = 1 Then
        sql = ""
        sql = " AND (DebitCreditRefNo Is null OR DebitCreditRefNo = '')" & ""
        strAppend = strAppend + sql
    End If

    If Len(txtDRAmtFr) Then
        strAppend = strAppend + " AND CLMPayableByMedix2M >= '" & Trim(Format(txtDRAmtFr, "#,##0.00")) & "'"
    End If
    
    If Len(txtDRAmtTo) Then
        strAppend = strAppend + " AND CLMPayableByMedix2M <= '" & Trim(Format(txtDRAmtTo, "#,##0.00")) & "'"
    End If
    
    If chkDRAmt.Value = 1 Then
        sql = ""
        sql = " AND (CLMPayableByMedix2M Is null OR CLMPayableByMedix2M = '')" & ""
        strAppend = strAppend + sql
    End If

    If Len(cboACPeriodMthFr) Then
        strAppend = strAppend + " AND RECMBMAccPeriodMth >= '" & Trim(cboACPeriodMthFr) & "'"
    End If
    
    If Len(cboACPeriodMthTo) Then
        strAppend = strAppend + " AND RECMBMAccPeriodMth <= '" & Trim(cboACPeriodMthTo) & "'"
    End If
    
    If chkPeriodMth.Value = 1 Then
        sql = ""
        sql = " AND RECMBMAccPeriodMth Is null " & ""
        strAppend = strAppend + sql
    End If

    If Len(cboACPeriodYRFr) Then
        strAppend = strAppend + " AND RECMBMAccPeriodYr >= '" & Trim(cboACPeriodYRFr) & "'"
    End If
    
    If Len(cboACPeriodYRTo) Then
        strAppend = strAppend + " AND RECMBMAccPeriodYr <= '" & Trim(cboACPeriodYRTo) & "'"
    End If
    
    If ChkPeriodYr.Value = 1 Then
        sql = ""
        sql = " AND RECMBMAccPeriodYr Is null " & ""
        strAppend = strAppend + sql
    End If
       
    searchCriteria = strAppend
    Call ReceiptFrMBMListing(strAppend)
   
End Function

Private Sub ReceiptFrMBMListing(Optional strAppend As String)
Dim rst2 As New Adodb.Recordset
Dim REC As New Adodb.Recordset
Dim Record As ListItem
Dim sql As String
Dim strsqlREC As String
Dim Total As String
Dim Current As String
Dim Check As Boolean
Dim CasNumber As String
Dim WSVersion As String
Dim INVAmt As Variant


    Total = 0
    Current = 0
    If ChkScan.Value = 0 Then
        lstReceiptMBM.ListItems.Clear
    End If
    sql = " Select * From VIEW_ReceiptFrMBM where CLMPayableByMedix2M < 0  "
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterconnPayment, adOpenForwardOnly, adLockBatchOptimistic
   
    If rst2.EOF = True Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
    Do
        Do Until rst2.EOF
        'total = rst2.recordCount
        'lblTotal = total
        txtPVStatus = ""
        If Trim(rst2!PVRecMBMStatus) <> "" Then
            txtPVStatus = rst2!PVRecMBMStatus
        End If
            
            With rst2
                Set Record = lstReceiptMBM.ListItems.Add(, , "")
                
                If Len(rst2!DebitCreditRefNo) Then
                    Record.SubItems(1) = rst2!DebitCreditRefNo
                End If
                
                If txtPVStatus = "F" Then
                    Record.SubItems(2) = "Full"
                ElseIf txtPVStatus = "P" Then
                    Record.SubItems(2) = "Partial"
                Else
                    Record.SubItems(2) = ""
                End If
                                
                If Len(rst2!CasNumber & "-" & rst2!claimversion) Then
                    Record.SubItems(3) = (rst2!CasNumber & "-" & rst2!claimversion)
                End If
                
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(4) = Format(rst2!CLMPayableByMedix2M, "#,##0.00") * (-1)
                End If
                
                Record.SubItems(5) = IIf(IsNumeric(!ReceiptAmt), !ReceiptAmt, 0)
                                               
                If txtPVStatus = "P" Then
                    Record.SubItems(6) = Round(CDbl(Record.SubItems(4) - CDbl(Record.SubItems(5))), 2)
                ElseIf txtPVStatus = "F" Then
                    Record.SubItems(6) = "0.00"
                Else
                    Record.SubItems(6) = CDbl(Round(Record.SubItems(4), 2))
                End If
                
                If Len(rst2!ReceiptIndex) Then
                    Record.SubItems(8) = rst2!ReceiptIndex
                End If
                
                If Len(rst2!RECEIPTNO) Then
                    Record.SubItems(9) = rst2!RECEIPTNO
                End If
                
                If Len(rst2!ReceiptDate) Then
                    Record.SubItems(10) = rst2!ReceiptDate
                End If
                
                If Len(rst2!RECMBMAccPeriodMth) Or Len(rst2!RECMBMAccPeriodYr) Then
                    Record.SubItems(11) = rst2!RECMBMAccPeriodMth & "/" & rst2!RECMBMAccPeriodYr
                End If
                
                If Len(rst2!MBMBankCharge) Then
                    Record.SubItems(12) = Format(rst2!MBMBankCharge, "#,##0.00")
                End If
                
                If Len(rst2!ContraCaseNo & "-" & rst2!ContraVersionNo) Then
                    Record.SubItems(13) = (Trim(rst2!ContraCaseNo) & "-" & Trim(rst2!ContraVersionNo))
                End If
                
                If Len(rst2!ContraCaseAmt) Then
                    Record.SubItems(14) = Format(rst2!ContraCaseAmt, "#,##0.00")
                End If
                
                If Len(rst2!MBMGainLoss) Then
                    Record.SubItems(15) = Format(rst2!MBMGainLoss, "#,##0.00")
                End If
                
                If Len(rst2!MBMMisc) Then
                    Record.SubItems(16) = Format(rst2!MBMMisc, "#,##0.00")
                End If
                
                If Len(rst2!MBMDiscount) Then
                    Record.SubItems(17) = Format(rst2!MBMDiscount, "#,##0.00")
                End If
                
                If Len(rst2!MBMAmtPaid) Then
                    Record.SubItems(18) = Format(rst2!MBMAmtPaid, "#,##0.00")
                End If
                
                If Len(rst2!MBMWriteOff) Then
                    Record.SubItems(19) = Format(rst2!MBMWriteOff, "#,##0.00")
                End If
                
                If Len(rst2!MBMMedixReceipt) Then
                    Record.SubItems(20) = Format(rst2!MBMMedixReceipt, "#,##0.00")
                End If
                
                If Len(rst2!MBMRemarks) Then
                    Record.SubItems(21) = Format(rst2!MBMRemarks, "#,##0.00")
                End If
                
                Current = Current + 1
                lblCurRecord = Current
                lblReceivedAmt = lblReceivedAmt + CDbl(Format(IIf(IsNumeric(!ReceiptAmt) = True, !ReceiptAmt, 0), "#,##0.00"))
                
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

Private Sub cmdEXPtoExcel_Click()
    'Screen.MousePointer = vbHourglass
        'Call ExportToExcelReceiptMBM(lstReceiptMBM)
    'Screen.MousePointer = vbDefault
    
    Screen.MousePointer = vbHourglass
        
    If lstReceiptMBM.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(lstReceiptMBM)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
        
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdSave_Click()
Dim CASNo As String
Dim CASVer As String
Dim status As String
Dim BankCharge As String
Dim ContraCase As String
Dim ContraVer As String
Dim ContraAmt As String
Dim GainLoss As String
Dim MiscDiff As String
Dim Discount As String
Dim AmtPaid As String
Dim WriteOff As String
Dim MedixRec As String
Dim Remarks As String
Dim ii As Integer
Dim Cnt As Integer
Dim delimiter1, delimiter2, delimiter3, delimiter4 As Integer
Dim strsqlCLM As String
Dim CLM As New Adodb.Recordset
Dim Update
Dim ItemChecked As Boolean

    Cnt = 0
    ItemChecked = False
    CmdSave.Enabled = False
  '  medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
  '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    If TxtReceiptNo = "" Then
        MsgBox "Please Enter Receipt No", vbOKOnly + vbCritical, DialogBoxTitle
        TxtReceiptNo.SetFocus
    ElseIf txtReceiptDate = "__/__/____" Then
        MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
        txtReceiptDate.SetFocus
    ElseIf txtAmt = "" Then
        MsgBox "Please Enter Amount", vbOKOnly + vbCritical, DialogBoxTitle
        txtAmt.SetFocus
    ElseIf txtChequeNo = "" Then
        MsgBox "Please Enter Cheque No", vbOKOnly + vbCritical, DialogBoxTitle
        txtChequeNo.SetFocus
    Else
        Update = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
        
        If Update = vbYes Then
        
            For ii = 1 To lstReceiptMBM.ListItems.Count
                If lstReceiptMBM.ListItems(ii).Checked = True Then
                    ItemChecked = True
                Exit For
                End If
            Next ii
            
            If ItemChecked = True Then
            
                For ii = 1 To lstReceiptMBM.ListItems.Count
                    If lstReceiptMBM.ListItems(ii).Checked = True Then
                        
                        CASNo = Trim(Mid(lstReceiptMBM.ListItems(ii).SubItems(3), 1, IIf(InStr(1, lstReceiptMBM.ListItems(ii).SubItems(3), "-") = 0, 10, InStr(1, lstReceiptMBM.ListItems(ii).SubItems(3), "-") - 1)))
                        CASVer = Trim(Mid(lstReceiptMBM.ListItems(ii).SubItems(3), InStr(1, lstReceiptMBM.ListItems(ii).SubItems(3), "-") + 1, Len(lstReceiptMBM.ListItems(ii).SubItems(3)) - InStr(1, lstReceiptMBM.ListItems(ii).SubItems(3), "-")))
                        status = Trim(Mid(lstReceiptMBM.ListItems(ii).SubItems(2), 1, 1))
                        PaymentAmt = Trim(lstReceiptMBM.ListItems(ii).SubItems(7))
                        BankCharge = Trim(lstReceiptMBM.ListItems(ii).SubItems(12))
                        ContraCase = Trim(Mid(lstReceiptMBM.ListItems(ii).SubItems(13), 1, IIf(InStr(1, lstReceiptMBM.ListItems(ii).SubItems(13), "-") = 0, 10, InStr(1, lstReceiptMBM.ListItems(ii).SubItems(13), "-") - 1)))
                        ContraVer = Trim(Mid(lstReceiptMBM.ListItems(ii).SubItems(13), InStr(1, lstReceiptMBM.ListItems(ii).SubItems(13), "-") + 1, Len(lstReceiptMBM.ListItems(ii).SubItems(13)) - InStr(1, lstReceiptMBM.ListItems(ii).SubItems(13), "-")))
                        ContraAmt = Trim(lstReceiptMBM.ListItems(ii).SubItems(14))
                        GainLoss = Trim(lstReceiptMBM.ListItems(ii).SubItems(15))
                        MiscDiff = Trim(lstReceiptMBM.ListItems(ii).SubItems(16))
                        Discount = Trim(lstReceiptMBM.ListItems(ii).SubItems(17))
                        AmtPaid = Trim(lstReceiptMBM.ListItems(ii).SubItems(18))
                        MedixRec = Trim(lstReceiptMBM.ListItems(ii).SubItems(19))
                        Remarks = ChkStr(Trim(lstReceiptMBM.ListItems(ii).SubItems(20)))
                        
                        Call UpdateRecord(CASNo, CInt(CASVer), status, PaymentAmt, BankCharge, GainLoss, MiscDiff, Discount, AmtPaid, WriteOff, MedixRec, Remarks)
                        Call UpdateContra(CASNo, CInt(CASVer), PaymentAmt, ContraCase, ContraVer, ContraAmt)
                    End If
                Next ii
                
                MsgBox "Record updated...", vbOKOnly + vbInformation, DialogBoxTitle
                If ChkScan.Value = 0 Then
                    Call SearchRecord
                ElseIf ChkScan.Value = 1 Then
                    lstReceiptMBM.ListItems.Clear
                    TxtCaseNo1.SetFocus
                End If
            End If
        End If
    End If
   ' medixmasterconnPayment.Close
   ' Set medixmasterconnPayment = Nothing
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
    CmdSave.Enabled = True
End Sub

Private Sub cmdSearch_Click()

    txtCHKAmt = 0
    lblReceivedAmt = 0
    
    ScanCode = TxtCaseNo1
    Screen.MousePointer = vbHourglass
    CmdSearch.Enabled = False
   ' medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"

    If Mid(ScanCode, 1, 2) = "LO" Then
        TxtCaseNo1 = ""
        Call SearchForm(ScanCode)
    Else
        one = False
        If ChkScan.Value = 1 Then
            If TxtCaseNo1 = "" Then
                MsgBox "Please Enter Worksheet No", vbOKOnly + vbCritical, DialogBoxTitle
                TxtCaseNo1.SetFocus
            Else
            
                
                listrecord = lstReceiptMBM.ListItems.Count
                If listrecord <> 0 Then
                    
                    For listloop1 = 1 To listrecord
                        Call CaseNoCheck
                    Next listloop1
                End If
                    If RcdCheck = True Then
                        MsgBox "Data has been uploaded!", vbOKOnly + vbCritical
                    ElseIf RcdCheck = False Then
                        Call SearchRecord
                    End If
                    RcdCheck = False
                TxtCaseNo1 = ""
                TxtCaseNo1.SetFocus
            End If
            TxtCaseNo1.SetFocus
        Else
                Call SearchRecord
        End If
    End If
'    medixmasterconnPayment.Close
 '   Set medixmasterconnPayment = Nothing
    CmdSearch.Enabled = True
    Screen.MousePointer = Default

End Sub

Private Sub cmdStop_Click()
    intExit = 1
    CmdExit.Enabled = True
End Sub

Private Sub CmdUpdate_Click()
Dim RecordSaved
Dim RptMBM As New Adodb.Recordset
Dim StrSql As String
Dim strsqlContra As String
Dim RptContra As New Adodb.Recordset
Dim CaseNo
Dim VersionNo
    
    CmdUpdate.Enabled = False
 '   medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
'    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    If lstReceiptMBM.ListItems.Count = 0 Then
        MsgBox "Please select the record", vbOKOnly + vbCritical
           
        ElseIf TxtReceiptNo = "" Then
            MsgBox "Please Enter Receipt No!", vbCritical
            TxtReceiptNo.SetFocus
        ElseIf IsDate(txtReceiptDate) = False Then
            MsgBox "Invalid Receipt Date!", vbCritical
            txtReceiptDate.SetFocus
        ElseIf txtReceiptDate = "__/__/____" Then
            MsgBox "Please Enter Receipt Date", vbOKOnly + vbCritical, DialogBoxTitle
            txtReceiptDate.SetFocus
        ElseIf txtChequeNo = "" Then
            MsgBox "Please Enter Cheque No!", vbCritical
            txtChequeNo.SetFocus
        ElseIf txtAmt = "" Then
            MsgBox "Please Enter Cheque Amt!", vbCritical
            txtAmt.SetFocus
               
        Else
            StrSql = "Select * From ReceiptFrMBM Where ReceiptIndex = '" & Trim(lstReceiptMBM.SelectedItem.SubItems(8)) & "'"
                     
            RptMBM.Open StrSql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
            
            RecordSaved = MsgBox("Do you want to update the record?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                If RptMBM.recordCount = 0 Then
                    RptMBM.AddNew
                End If
                
                With RptMBM
                    !RECEIPTNO = TxtReceiptNo
                    !ReceiptDate = txtReceiptDate
                    !ReceiptAmt = txtAmt
                    !ReceiptChequeNo = txtChequeNo
                    !ReceiptRemarks = Trim(TxtRecRemarks)
                    !ReceiptLastUpdateDate = Date
                    !ReceiptLastUpdateUser = Logon
                    If cboAccPeriodMth <> "" Then
                        !RECMBMAccPeriodMth = cboAccPeriodMth
                    End If
                    If cboAccPeriodYR <> "" Then
                        !RECMBMAccPeriodYr = cboAccPeriodYR
                    End If
                    '============================================
                    If TxtBankCharge <> "" Then
                        !MBMBankCharge = Format(TxtBankCharge, "#,##0.00")
                    Else
                        !MBMBankCharge = Null
                    End If
                    If TxtGainLoss <> "" Then
                        !MBMGainLoss = Format(TxtGainLoss, "#,##0.00")
                    Else
                        !MBMGainLoss = Null
                    End If
                    If TxtMisc <> "" Then
                        !MBMMisc = Format(TxtMisc, "#,##0.00")
                    Else
                        !MBMMisc = Null
                    End If
                    If TxtDiscount <> "" Then
                        !MBMDiscount = Format(TxtDiscount, "#,##0.00")
                    Else
                        !MBMDiscount = Null
                    End If
                    If TxtAmtPaid <> "" Then
                        !MBMAmtPaid = Format(TxtAmtPaid, "#,##0.00")
                    Else
                        !MBMAmtPaid = Null
                    End If
                    If TxtWriteOff <> "" Then
                        !MBMWriteOff = Format(TxtWriteOff, "#,##0.00")
                    Else
                        !MBMWriteOff = Null
                    End If
                    If TxtMedixRec <> "" Then
                        !MBMMedixReceipt = Format(TxtMedixRec, "#,##0.00")
                    Else
                        !MBMMedixReceipt = Null
                    End If
                    If TxtRemarks <> "" Then
                        !MBMRemarks = ChkStr(Trim(TxtRemarks))
                    Else
                        !MBMRemarks = Null
                    End If
                    '============================================
            
                    .Update
                End With
                
                CaseNo = Trim(Mid(lstReceiptMBM.SelectedItem.SubItems(3), 1, IIf(InStr(1, lstReceiptMBM.SelectedItem.SubItems(3), "-") = 0, 10, InStr(1, lstReceiptMBM.SelectedItem.SubItems(3), "-") - 1)))
                VersionNo = Trim(Mid(lstReceiptMBM.SelectedItem.SubItems(3), InStr(1, lstReceiptMBM.SelectedItem.SubItems(3), "-") + 1, Len(lstReceiptMBM.SelectedItem.SubItems(3)) - InStr(1, lstReceiptMBM.SelectedItem.SubItems(3), "-")))
                
                                
                strsqlContra = " Select * From ContraAllocation Where RefNo = '" & Trim(lstReceiptMBM.SelectedItem.SubItems(9)) & "'" & _
                               " AND CaseNo = '" & CaseNo & "'" & _
                               " AND VersionNo = '" & VersionNo & "'"
                               
                RptContra.Open strsqlContra, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
                
                If RptContra.recordCount = 0 Then
                    RptContra.AddNew
                End If
                
                With RptContra
                    !RefNo = TxtReceiptNo
                    If TxtContraCase <> "" Then
                        !ContraCaseNo = Trim(TxtContraCase)
                    End If
                    If TxtContraVer <> "" Then
                        !ContraVersionNo = Trim(TxtContraVer)
                    End If
                    If TxtContraAmt <> "" Then
                        !ContraCaseAmt = Format(TxtContraAmt, "#,##0.00")
                    End If
                               
                    .Update
                End With
                
                
                MsgBox "Record Updated...", vbOKOnly + vbInformation, "Information"
          
                Call SearchRecord
                one = False
            End If
            RptMBM.Close
            Set RptMBM = Nothing
            
        TxtReceiptNo.Text = ""
        txtReceiptDate = "__/__/____"
        txtAmt.Text = ""
        txtChequeNo.Text = ""
        TxtRecRemarks.Text = ""
        cboAccPeriodMth = ""
        cboAccPeriodYR = ""
        '===============================
        TxtContraCase.Text = ""
        TxtContraVer.Text = ""
        TxtBankCharge.Text = ""
        TxtContraAmt.Text = ""
        TxtGainLoss.Text = ""
        TxtMisc.Text = ""
        TxtDiscount.Text = ""
        TxtAmtPaid.Text = ""
        TxtWriteOff.Text = ""
        TxtMedixRec.Text = ""
        TxtRemarks.Text = ""
        '===============================
        CmdSave.Enabled = True
    End If
  '  medixmasterconnPayment.Close
  '  Set medixmasterconnPayment = Nothing
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    CmdUpdate.Enabled = True
End Sub

Private Sub Form_Load()
    one = False
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        Call ComboListing
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing

    CboPayment.ListIndex = 0
    ReDim CaseNoArray(0)
    txtCHKAmt = 0
    lblReceivedAmt = 0
    intExit = 0
    intMBMAccount = 1
    Label30.Visible = False
    TxtCaseNo2.Visible = False
    ChkScan.Value = 1
    ChkScan.Left = 3480
End Sub

Private Sub lstReceiptMBM_Click()
Dim rst3 As New Adodb.Recordset
Dim StrSql As String


    If one = True Then
        'medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
        CmdSave.Enabled = False
        If lstReceiptMBM.ListItems.Count Then
            TxtReceiptNo = ""
            txtReceiptDate.mask = ""
            txtReceiptDate.Text = ""
            txtReceiptDate.mask = "##/##/####"
            txtAmt = ""
            txtChequeNo = ""
            TxtRecRemarks = ""
            cboAccPeriodMth = ""
            cboAccPeriodYR = ""
            '======================
            TxtClaimNo = ""
            TxtBankCharge = ""
            TxtContraCase = ""
            TxtContraVer = ""
            TxtContraAmt = ""
            TxtGainLoss = ""
            TxtMisc = ""
            TxtDiscount = ""
            TxtAmtPaid = ""
            TxtWriteOff = ""
            TxtMedixRec = ""
            TxtRemarks = ""
            '=======================
            StrSql = " SELECT * From VIEW_ReceiptFrMBM Where ReceiptIndex = '" & lstReceiptMBM.SelectedItem.SubItems(8) & "'"
            rst3.Open StrSql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
        
            If rst3.recordCount = 0 Then
                MsgBox "Payment hasn't been made!", vbOKOnly + vbCritical, DialogBoxTitle
            Else
                TxtReceiptNo.SetFocus
                If Len(rst3!RECEIPTNO) Then
                    TxtReceiptNo = rst3!RECEIPTNO
                End If
                If Len(rst3!ReceiptDate) Then
                    txtReceiptDate = rst3!ReceiptDate
                End If
                If Len(rst3!ReceiptAmt) Then
                    txtAmt = rst3!ReceiptAmt
                End If
                If Len(rst3!ReceiptChequeNo) Then
                    txtChequeNo = rst3!ReceiptChequeNo
                End If
                If Len(rst3!ReceiptRemarks) Then
                    TxtRecRemarks = rst3!ReceiptRemarks
                End If
                If Len(rst3!RECMBMAccPeriodMth) Then
                    cboAccPeriodMth = rst3!RECMBMAccPeriodMth
                End If
                If Len(rst3!RECMBMAccPeriodYr) Then
                    cboAccPeriodYR = rst3!RECMBMAccPeriodYr
                End If
                '======================================================
                If Len(rst3!CasNumber & "-" & rst3!claimversion) Then
                    TxtClaimNo = (Trim(rst3!CasNumber) & "-" & Trim(rst3!claimversion))
                End If
                
                If Len(rst3!MBMBankCharge) Then
                    TxtBankCharge = rst3!MBMBankCharge
                End If
                If Len(rst3!ContraCaseNo) <> "" Then
                    TxtContraCase = Trim(rst3!ContraCaseNo)
                End If
                If Len(rst3!ContraVersionNo) <> "" Then
                    TxtContraVer = Trim(rst3!ContraVersionNo)
                End If
                If Len(rst3!ContraCaseAmt) Then
                    TxtContraAmt = rst3!ContraCaseAmt
                End If
                If Len(rst3!MBMGainLoss) Then
                    TxtGainLoss = rst3!MBMGainLoss
                End If
                If Len(rst3!MBMMisc) Then
                    TxtMisc = rst3!MBMMisc
                End If
                If Len(rst3!MBMDiscount) Then
                    TxtDiscount = rst3!MBMDiscount
                End If
                If Len(rst3!MBMAmtPaid) Then
                    TxtAmtPaid = rst3!MBMAmtPaid
                End If
                If Len(rst3!MBMWriteOff) Then
                    TxtWriteOff = rst3!MBMWriteOff
                End If
                If Len(rst3!MBMMedixReceipt) Then
                    TxtMedixRec = rst3!MBMMedixReceipt
                End If
                If Len(rst3!MBMRemarks) Then
                    TxtRemarks = rst3!MBMRemarks
                End If
                '======================================================
                CmdSave.Enabled = False
                rst3.Close
                Set rst3 = Nothing
            End If
        End If
        one = False
       ' medixmasterconnPayment.Close
        'Set medixmasterconnPayment = Nothing
    End If
    
    If lstReceiptMBM.ListItems.Count <> 0 Then
        TxtClaimNo = Trim(lstReceiptMBM.SelectedItem.SubItems(3))
        TxtBankCharge.SetFocus
    End If

End Sub

Private Sub lstReceiptMBM_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ProcessCheckedItem(Item)
End Sub

Public Function ProcessCheckedItem(Item As ListItem, Optional Partial As Boolean)
Dim i, newbound  As Integer
Dim tempArray() As String
Dim Found As Boolean
Dim temp
Dim listloop As Integer
Dim ListCount As Integer
Dim TotalAmt As Variant
Dim RECStatus As String
Dim RecAmt As Integer
Dim strsqlCLM As String
Dim CLM As New Adodb.Recordset
Dim strsqlREC As String
Dim REC As New Adodb.Recordset
Dim CaseNo As String
Dim caseversion As String
Dim RECEIPTNO As String
Static Arraycnt As Integer
Dim strsqlUSR As String
Dim USR As New Adodb.Recordset
       
    'medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

    CaseNo = Trim(Mid(Item.SubItems(3), 1, IIf(InStr(1, Item.SubItems(3), "-") = 0, 10, InStr(1, Item.SubItems(3), "-") - 1)))
    caseversion = Trim(Mid(Item.SubItems(3), InStr(1, Item.SubItems(3), "-") + 1, Len(Item.SubItems(3)) - InStr(1, Item.SubItems(3), "-")))
    'status = Trim(Item.SubItems(2))
    RECEIPTNO = Trim(Item.SubItems(9))
     
    strsqlUSR = "Select USRCode, AuthorisedStatus From USR where USRCode = '" & Trim(Logon) & "'"
    USR.Open strsqlUSR, medixmasterconnMaint, adOpenKeyset, adLockPessimistic
    
    If Trim(USR!AuthorisedStatus) = "N" Then
    
        If Item.Checked = True Then
                        
            lstReceiptMBM.SelectedItem.Selected = False
            Item.Selected = True
            
            strsqlREC = " Select * From VIEW_ReceiptFrMBM where CASNumber = '" & Trim(CaseNo) & "'" & _
                        " And ClaimVersion = '" & Trim(caseversion) & "'" & _
                        " And ReceiptNo = '" & Trim(RECEIPTNO) & "'"
            REC.Open strsqlREC, medixmasterconnPayment, adOpenKeyset, adLockPessimistic
            
            If REC.recordCount <> 0 Then
                If REC!PVRecMBMStatus = "F" Then
                    MsgBox "Payment has been made!", vbCritical
                    Item.Checked = False
                ElseIf REC!PVRecMBMStatus = "P" Then
                    Item.SubItems(2) = "Full"
                    Item.SubItems(7) = CDbl(Item.SubItems(6))
                    Item.SubItems(6) = "0"
                    'txtCHKAmt = txtCHKAmt + CDbl(Format(Item.SubItems(7), "#,##0.00")) '* (-1)
                    txtCHKAmt = Format(CDbl(txtCHKAmt) + CDbl(Item.SubItems(7)), "#,##0.00")
                End If
            Else
                If Partial = False Then
                    Item.SubItems(2) = "Full"
                    Item.SubItems(6) = "0"
                    Item.SubItems(7) = CDbl(Item.SubItems(4))
                    'txtCHKAmt = txtCHKAmt + CDbl(Format(Item.SubItems(7), "#,##0.00")) '* (-1)
                    txtCHKAmt = Format(CDbl(txtCHKAmt) + CDbl(Item.SubItems(7)), "#,##0.00")
                    'ReDim Preserve CaseNoArray(Arraycnt)
                    'CaseNoArray(Arraycnt) = Item.SubItems(3) & "~" & Item.SubItems(2) & "~" & Item.SubItems(7) '& "~" & item.SubItems(8)
                    'Arraycnt = Arraycnt + 1
                Else
                    Item.SubItems(2) = "Partial"
                    Item.SubItems(6) = CDbl(Item.SubItems(6)) - CDbl(Item.SubItems(7))
                    'Item.SubItems(7) = CDbl(Item.SubItems(4))
                    'txtCHKAmt = txtCHKAmt + CDbl(Format(Item.SubItems(7), "#,##0.00")) '* (-1)
                    txtCHKAmt = Format(CDbl(txtCHKAmt) + CDbl(Item.SubItems(7)), "#,##0.00")
                    'ReDim Preserve CaseNoArray(Arraycnt)
                    'CaseNoArray(Arraycnt) = Item.SubItems(3) & "~" & Item.SubItems(2) & "~" & Item.SubItems(7) '& "~" & item.SubItems(8)
                    'Arraycnt = Arraycnt + 1
                End If
            End If
            REC.Close
            Set REC = Nothing
        Else
        
            ' Remove from the array
            'newbound = 0
            
            CaseNo = Trim(Mid(Item.SubItems(3), 1, IIf(InStr(1, Item.SubItems(3), "-") = 0, 10, InStr(1, Item.SubItems(3), "-") - 1)))
            caseversion = Trim(Mid(Item.SubItems(3), InStr(1, Item.SubItems(3), "-") + 1, Len(Item.SubItems(3)) - InStr(1, Item.SubItems(3), "-")))
            RECEIPTNO = Trim(Item.SubItems(9))
            
            lstReceiptMBM.SelectedItem.Selected = False
            Item.Selected = True
                        
            strsqlREC = " Select * From VIEW_ReceiptFrMBM where CASNumber = '" & Trim(CaseNo) & "'" & _
                        " And ClaimVersion = '" & Trim(caseversion) & "'" & _
                        " And ReceiptNo = '" & Trim(RECEIPTNO) & "'"
            REC.Open strsqlREC, medixmasterconnPayment, adOpenKeyset, adLockPessimistic
            
            If REC.recordCount <> 0 Then
        
                If REC!PVRecMBMStatus = "P" Then
                    Item.SubItems(2) = "Partial"
                    Item.SubItems(6) = CDbl(Item.SubItems(4)) - CDbl(Item.SubItems(5))
                    If Len(Item.SubItems(7)) Then
                        txtCHKAmt = Format(CDbl(txtCHKAmt) - CDbl(Item.SubItems(7)), "#,##0.00")
                    End If
                    Item.SubItems(7) = ""
                ElseIf REC!PVRecMBMStatus = "F" Then
                    txtCHKAmt = Format(CDbl(txtCHKAmt) - CDbl(Item.SubItems(4)), "#,##0.00")
                    Item.SubItems(2) = "Full"
                    Item.SubItems(5) = REC!ReceiptAmt
                    Item.SubItems(6) = ""
                    Item.SubItems(7) = ""
                End If
            Else
                Item.SubItems(2) = ""
                Item.SubItems(6) = CDbl(Item.SubItems(4))
                If Len(Item.SubItems(7)) Then
                    txtCHKAmt = Format(CDbl(txtCHKAmt) - CDbl(Item.SubItems(7)), "#,##0.00")
                    Item.SubItems(7) = ""
                End If
            End If
            
            Item.SubItems(12) = ""
            Item.SubItems(13) = ""
            Item.SubItems(14) = ""
            Item.SubItems(15) = ""
            Item.SubItems(16) = ""
            Item.SubItems(17) = ""
            Item.SubItems(18) = ""
            Item.SubItems(19) = ""
            Item.SubItems(20) = ""
            Item.SubItems(21) = ""
            
        REC.Close
        Set REC = Nothing
        End If
Else
    If Item.Checked = True Then
    
        lstReceiptMBM.SelectedItem.Selected = False
        Item.Selected = True
 
        strsqlREC = " Select * From VIEW_ReceiptFrMBM where CASNumber = '" & Trim(CaseNo) & "'" & _
                    " And ClaimVersion = '" & Trim(caseversion) & "'" & _
                    " And ReceiptNo = '" & Trim(RECEIPTNO) & "'"
        REC.Open strsqlREC, medixmasterconnPayment, adOpenKeyset, adLockPessimistic
            
        If REC.recordCount <> 0 Then
            If REC!PVRecMBMStatus = "F" Then
                Item.SubItems(2) = "Full"
                Item.SubItems(7) = CDbl(Item.SubItems(4))
                Item.SubItems(6) = ""
                Item.SubItems(5) = ""
                txtCHKAmt = Format(CDbl(txtCHKAmt) + CDbl(Item.SubItems(7)), "#,##0.00")
            ElseIf REC!PVRecMBMStatus = "P" Then
                    Item.SubItems(2) = "Full"
                    Item.SubItems(7) = CDbl(Item.SubItems(6))
                    Item.SubItems(6) = "0"
                    txtCHKAmt = Format(CDbl(txtCHKAmt) + CDbl(Item.SubItems(7)), "#,##0.00")
            End If
        Else
            If Partial = False Then
                Item.SubItems(2) = "Full"
                Item.SubItems(6) = "0"
                Item.SubItems(7) = CDbl(Item.SubItems(4))
                txtCHKAmt = Format(CDbl(txtCHKAmt) + CDbl(Item.SubItems(7)), "#,##0.00")
            Else
                Item.SubItems(2) = "Partial"
                Item.SubItems(6) = CDbl(Item.SubItems(6)) - CDbl(Item.SubItems(7))
                txtCHKAmt = Format(CDbl(txtCHKAmt) + CDbl(Item.SubItems(7)), "#,##0.00")
            End If
        End If
        REC.Close
        Set REC = Nothing
    Else
        ' Remove from the array
        'newbound = 0
        
        CaseNo = Trim(Mid(Item.SubItems(3), 1, IIf(InStr(1, Item.SubItems(3), "-") = 0, 10, InStr(1, Item.SubItems(3), "-") - 1)))
        caseversion = Trim(Mid(Item.SubItems(3), InStr(1, Item.SubItems(3), "-") + 1, Len(Item.SubItems(3)) - InStr(1, Item.SubItems(3), "-")))
        RECEIPTNO = Trim(Item.SubItems(9))
        
        lstReceiptMBM.SelectedItem.Selected = False
        Item.Selected = True
            
        strsqlREC = " Select * From VIEW_ReceiptFrMBM where CASNumber = '" & Trim(CaseNo) & "'" & _
                    " And ClaimVersion = '" & Trim(caseversion) & "'" & _
                    " And ReceiptNo = '" & Trim(RECEIPTNO) & "'"
        REC.Open strsqlREC, medixmasterconnPayment, adOpenKeyset, adLockPessimistic
        
        If REC.recordCount <> 0 Then
        
            If REC!PVRecMBMStatus = "P" Then
                Item.SubItems(2) = "Partial"
                Item.SubItems(6) = CDbl(Item.SubItems(4)) - CDbl(Item.SubItems(5))
                If Len(Item.SubItems(7)) Then
                    txtCHKAmt = Format(CDbl(txtCHKAmt) - CDbl(Item.SubItems(7)), "#,##0.00")
                End If
                Item.SubItems(7) = ""
            ElseIf REC!PVRecMBMStatus = "F" Then
                txtCHKAmt = Format(CDbl(txtCHKAmt) - CDbl(Item.SubItems(4)), "#,##0.00")
                Item.SubItems(2) = "Full"
                Item.SubItems(5) = REC!ReceiptAmt
                Item.SubItems(6) = ""
                Item.SubItems(7) = ""
            End If
        Else
            Item.SubItems(2) = ""
            Item.SubItems(6) = CDbl(Item.SubItems(4))
            If Len(Item.SubItems(7)) Then
                txtCHKAmt = Format(CDbl(txtCHKAmt) - CDbl(Item.SubItems(7)), "#,##0.00")
                Item.SubItems(7) = ""
            End If
        End If
            Item.SubItems(12) = ""
            Item.SubItems(13) = ""
            Item.SubItems(14) = ""
            Item.SubItems(15) = ""
            Item.SubItems(16) = ""
            Item.SubItems(17) = ""
            Item.SubItems(18) = ""
            Item.SubItems(19) = ""
            Item.SubItems(20) = ""
            Item.SubItems(21) = ""
            
        REC.Close
        Set REC = Nothing
        End If
    End If
 '   medixmasterconnPayment.Close
  '  Set medixmasterconnPayment = Nothing
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing

End Function

Private Sub UpdateRecord(CASNo As String, CASVer As String, status As String, PaymentAmt As Variant, BankCharge As Variant, GainLoss As Variant, MiscDiff As Variant, Discount As Variant, AmtPaid As Variant, WriteOff As Variant, MedixRec As Variant, Remarks As String)
Dim REC As New Adodb.Recordset
Dim strsqlREC As String
Dim USR As New Adodb.Recordset
Dim strsqlUSR As String
Dim REC2 As New Adodb.Recordset
Dim strsqlREC2 As String
    
strsqlUSR = "Select USRCode, AuthorisedStatus From USR where USRCode = '" & Trim(Logon) & "'"
USR.Open strsqlUSR, medixmasterconnMaint, adOpenKeyset, adLockPessimistic

If Trim(USR!AuthorisedStatus) = "N" Then
    
    strsqlREC = " Select * from ReceiptfrMBM where CASNumber = '" & Trim(CASNo) & "'" & _
                " And ClaimVersion = '" & Trim(CASVer) & "' AND RECEIPTNO = '" & Trim(TxtReceiptNo) & "'"
                
    REC.Open strsqlREC, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    
        If REC.recordCount = 0 Then
            REC.AddNew
        End If
        
        With REC
            !RECEIPTNO = ChkStr(Trim(TxtReceiptNo))
            !ReceiptDate = txtReceiptDate
            !ReceiptChequeNo = ChkStr(txtChequeNo)
            !ReceiptLastUpdateDate = Date
            !ReceiptLastUpdateUser = Logon
            !CasNumber = ChkStr(CASNo)
            !claimversion = CASVer
            !ReceiptRemarks = ChkStr(TxtRecRemarks)
            
            If cboAccPeriodMth <> "" Then
                !RECMBMAccPeriodMth = cboAccPeriodMth
            End If
            If cboAccPeriodYR <> "" Then
                !RECMBMAccPeriodYr = cboAccPeriodYR
            End If
            REC!PVRecMBMStatus = status
            
            !ReceiptAmt = PaymentAmt
                        
            '======================================================
            !MBMBankCharge = IIf(IsNumeric(BankCharge) = True, BankCharge, 0)
            !MBMGainLoss = IIf(IsNumeric(GainLoss) = True, GainLoss, 0)
            !MBMMisc = IIf(IsNumeric(MiscDiff) = True, MiscDiff, 0)
            !MBMDiscount = IIf(IsNumeric(Discount) = True, Discount, 0)
            !MBMAmtPaid = IIf(IsNumeric(AmtPaid) = True, AmtPaid, 0)
            !MBMWriteOff = IIf(IsNumeric(WriteOff) = True, WriteOff, 0)
            !MBMMedixReceipt = IIf(IsNumeric(MedixRec) = True, MedixRec, 0)
            If Len(Remarks) Then
                !MBMRemarks = ChkStr(Trim(Remarks))
            End If
            '======================================================
            
        End With
        REC.Update
        REC.Close
        Set REC = Nothing
        
        strsqlREC2 = " Select * from ReceiptfrMBM where CASNumber = '" & Trim(CASNo) & "'" & _
                     " And ClaimVersion = '" & Trim(CASVer) & "'"
                
        REC2.Open strsqlREC2, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    
        With REC2
            
            If status = "F" Then
                REC2!PVRecMBMStatus = status
            End If
            
        End With
        REC2.Update
        REC2.Close
        Set REC2 = Nothing
        
Else

    
    strsqlREC = " Select * from ReceiptfrMBM "
                
    REC.Open strsqlREC, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    
        'If REC.recordCount = 0 Then
            REC.AddNew
        'End If
        
        With REC
            !RECEIPTNO = ChkStr(TxtReceiptNo)
            !ReceiptDate = txtReceiptDate
            !ReceiptChequeNo = ChkStr(txtChequeNo)
            !ReceiptLastUpdateDate = Date
            !ReceiptLastUpdateUser = Logon
            !CasNumber = ChkStr(CASNo)
            !claimversion = CASVer
            !ReceiptRemarks = ChkStr(TxtRecRemarks)
            
            If cboAccPeriodMth <> "" Then
                !RECMBMAccPeriodMth = cboAccPeriodMth
            End If
            If cboAccPeriodYR <> "" Then
                !RECMBMAccPeriodYr = cboAccPeriodYR
            End If
            REC!PVRecMBMStatus = status
            
            !ReceiptAmt = PaymentAmt
                        
            '======================================================
            !MBMBankCharge = IIf(IsNumeric(BankCharge) = True, BankCharge, 0)
            !MBMGainLoss = IIf(IsNumeric(GainLoss) = True, GainLoss, 0)
            !MBMMisc = IIf(IsNumeric(MiscDiff) = True, MiscDiff, 0)
            !MBMDiscount = IIf(IsNumeric(Discount) = True, Discount, 0)
            !MBMAmtPaid = IIf(IsNumeric(AmtPaid) = True, AmtPaid, 0)
            !MBMWriteOff = IIf(IsNumeric(WriteOff) = True, WriteOff, 0)
            !MBMMedixReceipt = IIf(IsNumeric(MedixRec) = True, MedixRec, 0)
            If Len(Remarks) Then
                !MBMRemarks = ChkStr(Trim(Remarks))
            End If
            '======================================================
            
        End With
        REC.Update
        REC.Close
        Set REC = Nothing
        
        strsqlREC2 = " Select * from ReceiptfrMBM where CASNumber = '" & Trim(CASNo) & "'" & _
                     " And ClaimVersion = '" & Trim(CASVer) & "'"
                
        REC2.Open strsqlREC2, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
    
        With REC2
            
            If status = "F" Then
                REC2!PVRecMBMStatus = status
            End If
            
        End With
        REC2.Update
        REC2.Close
        Set REC2 = Nothing

End If

End Sub

Private Sub UpdateContra(CASNo As String, CASVer As String, PaymentAmt As Variant, ContraCase As String, ContraVer As String, ContraAmt As Variant)
Dim strsqlContra As String
Dim Contra As New Adodb.Recordset
    
            
    If ContraCase <> "" Then
        
        strsqlContra = " Select * from ContraAllocation " & _
                       " Where CaseNo = '" & Trim(CASNo) & "' And VersionNo = '" & Trim(CASVer) & "'"
                
        Contra.Open strsqlContra, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
        
        If Contra.recordCount = 0 Then
            Contra.AddNew
        End If
            
        With Contra
        
            !RefNo = ChkStr(TxtReceiptNo)
            !CaseNo = Trim(CASNo)
            !VersionNo = Trim(CASVer)
            !CaseAmt = Format(PaymentAmt, "#,##0.00")
            !ContraCaseNo = Trim(ContraCase)
            !ContraVersionNo = Trim(ContraVer)
            !ContraCaseAmt = Format(ContraAmt, "#,##0.00")
        
        End With
        Contra.Update
        Contra.Close
        Set Contra = Nothing
            
    End If
    
    txtCHKAmt = 0
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtAmtPaid_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtBankCharge_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtCaseNo1_GotFocus()
    CmdSearch.Default = True
End Sub

Private Sub TxtCaseNo1_LostFocus()
    TxtCaseNo1 = ChkStr(TxtCaseNo1)
    CmdSearch.Default = False
End Sub

Private Sub txtContraAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtDiscount_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtGainLoss_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtMBMNo_LostFocus()
TxtMBMNo = ChkStr(TxtMBMNo)
End Sub

Private Sub TxtMisc_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtMedixRec_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtReceiptDate_LostFocus()
    If IsDate(txtReceiptDate) Then
        cboAccPeriodMth.Text = Format(txtReceiptDate, "mm")
        cboAccPeriodYR.Text = Format(txtReceiptDate, "yyyy")
    Else
        If txtReceiptDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtReceiptDate.SetFocus
        End If
    End If
End Sub

Private Sub lstReceiptMBM_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    Case 7
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    Case 9
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(10)
    Case 11
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
    Case 12
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
    Case 14
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
    Case 16
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
    Case 17
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
    Case 18
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(18)
    Case 19
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(19)
    Case 20
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(20)
    Case 21
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtNumber, m_blnDirection(21)
    Case 22
        SortListView lstReceiptMBM, ColumnHeader.Index, ldtString, m_blnDirection(22)
    End Select
End Sub

Private Sub CboPayor_Change()
If CboPayor.Text = "" Then
    If InStr(CboPayor.Text, "null") Then
       CboPayor.Enabled = False
    End If
End If
End Sub

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

Private Function CaseNoCheck()
    If Trim(ChkStr(TxtCaseNo1)) = Trim(Mid(lstReceiptMBM.ListItems(listloop1).SubItems(3), 1, IIf(InStr(1, lstReceiptMBM.ListItems(listloop1).SubItems(3), "-") = 0, 3, InStr(1, lstReceiptMBM.ListItems(listloop1).SubItems(3), "-") - 1))) Then
        RcdCheck = True
    End If
End Function

Private Sub cmdPrintPV_Click()
Dim objExcel As excel.Application
Dim objWorkbook As excel.Workbook
Dim objWorksheet As excel.WORKSHEET
Dim i As Integer
Dim StrSql As String
Dim REC As New Adodb.Recordset
    
    Screen.MousePointer = vbHourglass
    cmdPrintPV.Enabled = False
  '  medixmasterconnPayment.Open "File Name=" & BOSPath & "Conn\mediconnectPayment.UDL"

    If TxtReceiptNo = "" Then
    
        MsgBox "Please Enter or Generate Receipt No", vbOKOnly + vbCritical
        medixmasterconnPayment.Close
        Set medixmasterconnPayment = Nothing
        Screen.MousePointer = vbDefault
    Else
        If objExcel Is Nothing Then
            Set objExcel = CreateObject("Excel.application")
        End If
        
        If Err.Number > 0 Then
            MsgBox "Microsoft Excel is Not loaded On this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
            GoTo ExitFunction
        End If
        
        On Error GoTo HANDLE_ERROR
        ' Don't allow user to affect workbook
        objExcel.Interactive = False
    
        StrSql = " Select CASNumber, ClaimVersion, ReceiptNo, ReceiptDate, HLTCode, " & _
                 " ReceiptChequeNo, Receiptamt, DebitCreditRefNo, MBMName, PatientName " & _
                 " From View_ReceiptFrMBMRec2 Where ReceiptNo = '" & Trim(TxtReceiptNo) & "'"
        
        REC.Open StrSql, medixmasterconnPayment, adOpenKeyset, adLockOptimistic
                
        If REC.recordCount Then
            RecordSelected = REC.recordCount
            If objExcel.Visible = False Then
                objExcel.Visible = True
            End If
        
            objExcel.WindowState = xlMaximized
    
        If Trim(REC!HLTCode) = "S" Then
            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\OR-MBMS.xlt")
        ElseIf Trim(REC!HLTCode) = "N" Then
            Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\PaymentVoucher\OR-MBMN.xlt")
        End If
        
        Set objWorksheet = objWorkbook.Sheets(1)
        'Turn off the alerts, otherwise user will have to confirm my actions.
        objExcel.DisplayAlerts = False
        'Do While objWorkbook.Worksheets.Count > 1
         '   Set objWorksheet = objWorkbook.Worksheets.Item(objWorkbook.Worksheets.Count)
          '  objWorksheet.Delete
        'Loop
        Set objWorksheet = objWorkbook.Worksheets.Item(1)
            
            objWorksheet.Cells(7, "I") = Trim(REC!RECEIPTNO)
            objWorksheet.Cells(9, "D") = Trim(REC!MBMName)
            objWorksheet.Cells(8, "O") = Format(REC!ReceiptDate, "dd/mm/yyyy")
            objWorksheet.Cells(11, "K") = Format(REC!ReceiptAmt, "#,##0.00")
            objWorksheet.Cells(13, "O") = Trim(REC!ReceiptChequeNo)
            
            For i = 1 To RecordSelected
            
                objWorksheet.Cells(14, "D") = Trim(REC!PatientName)
                objWorksheet.Cells(16, "D") = Trim(REC!DebitCreditRefNo)
                objWorksheet.Cells(18, "D") = (REC!CasNumber & "-" & REC!claimversion)
                
                'objWorksheet.Cells(15 + CDbl(i), "D") = i
                'objWorksheet.Cells(16 + CDbl(i), "D") = Trim(REC!PatientName)
                'objWorksheet.Cells(18 + CDbl(i), "D") = Trim(REC!DebitCreditRefNo)
                'objWorksheet.Cells(20 + CDbl(i), "D") = (REC!CASNumber & "-" & REC!claimversion)
            
            REC.MoveNext
            
            Next i
            
            End If
            'cmdCredit.Enabled = False
            objExcel.DisplayAlerts = True
            'objWorksheet.Columns.AutoFit
            objExcel.Interactive = True
           ' medixmasterconnPayment.Close
           ' Set medixmasterconnPayment = Nothing
            cmdPrintPV.Enabled = True
            Screen.MousePointer = vbDefault
            Exit Sub
ExitFunction:
            objExcel.DisplayAlerts = True
            objExcel.Interactive = True
            objExcel.Quit
            cmdPrintPV.Enabled = True
           ' medixmasterconnPayment.Close
           ' Set medixmasterconnPayment = Nothing
            Screen.MousePointer = vbDefault
            Exit Sub
HANDLE_ERROR:
            MsgBox "Export To Excel failed. Encountered the following Error" & vbCrLf & vbCrLf & _
            Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
            Set objWorksheet = Nothing
            Set objWorkbook = Nothing
            GoTo ExitFunction
    End If
    CmdSearch.Enabled = False
End Sub

Private Sub CboHLTCode_KeyPress(KeyAscii As Integer)
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

Private Sub TxtWriteOff_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub
