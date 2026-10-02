VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAnalysisReport 
   Caption         =   "Analysis Reports"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Analysis Reports"
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
         TabIndex        =   66
         Top             =   120
         Width           =   3285
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   56
      Top             =   7200
      Width           =   11655
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   375
         Left            =   10320
         TabIndex        =   62
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   6840
         TabIndex        =   61
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   6000
         TabIndex        =   60
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print to File"
         Height          =   375
         Left            =   9360
         TabIndex        =   59
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Default         =   -1  'True
         Height          =   375
         Left            =   7680
         TabIndex        =   58
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   375
         Left            =   8520
         TabIndex        =   57
         Top             =   160
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView11 
         Height          =   375
         Left            =   2760
         TabIndex        =   63
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Payor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Admission  Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label39 
         Caption         =   "* - Search Empty Data"
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
         TabIndex        =   64
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame17 
      Height          =   6735
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   6975
      Begin VB.TextBox TxtBordxNo2 
         Height          =   315
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   139
         Top             =   4080
         Width           =   1125
      End
      Begin VB.CheckBox ChkJoin 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   5160
         TabIndex        =   134
         Top             =   5880
         Width           =   495
      End
      Begin VB.ComboBox CboMbmType 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   132
         Top             =   3960
         Width           =   1365
      End
      Begin VB.ComboBox cboPrincipalType 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   130
         Top             =   3600
         Width           =   1365
      End
      Begin VB.ComboBox cboHOSGrp 
         Height          =   315
         Left            =   1200
         TabIndex        =   127
         Top             =   1440
         Width           =   5715
      End
      Begin VB.ComboBox cboHOS 
         Height          =   315
         Left            =   960
         TabIndex        =   126
         Top             =   960
         Width           =   5955
      End
      Begin VB.CheckBox ChkFinalDiag 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   4440
         TabIndex        =   125
         Top             =   4800
         Width           =   495
      End
      Begin VB.CheckBox ChkFirstDiag 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   4440
         TabIndex        =   124
         Top             =   4440
         Width           =   495
      End
      Begin VB.TextBox TxtFinalDiag1 
         Height          =   315
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   121
         Top             =   4800
         Width           =   645
      End
      Begin VB.TextBox TxtFinalDiag2 
         Height          =   315
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   120
         Top             =   4800
         Width           =   645
      End
      Begin VB.TextBox TxtFirstDiag1 
         Height          =   315
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   117
         Top             =   4440
         Width           =   645
      End
      Begin VB.TextBox TxtFirstDiag2 
         Height          =   315
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   116
         Top             =   4440
         Width           =   645
      End
      Begin VB.TextBox TxtAttDoc 
         Height          =   315
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   114
         Top             =   3000
         Width           =   3525
      End
      Begin VB.ComboBox CboClaimNature 
         Height          =   315
         Left            =   1440
         TabIndex        =   112
         Top             =   2640
         Width           =   3525
      End
      Begin VB.TextBox TxtBordxNo1 
         Height          =   315
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   109
         Top             =   4080
         Width           =   1125
      End
      Begin VB.TextBox TxtCaseNo1 
         Height          =   315
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   33
         Top             =   3435
         Width           =   1125
      End
      Begin VB.ComboBox CboGrpCom 
         Height          =   315
         Left            =   1440
         TabIndex        =   32
         Top             =   2280
         Width           =   3525
      End
      Begin VB.ComboBox CboClaimability 
         Height          =   315
         Left            =   5760
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   600
         Width           =   1125
      End
      Begin VB.ComboBox CboStatus 
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
         Left            =   3240
         TabIndex        =   30
         Top             =   600
         Width           =   1125
      End
      Begin VB.ComboBox cboPayorCode 
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
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   1875
         Width           =   3525
      End
      Begin VB.ComboBox cboDataStatus 
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
         Left            =   3240
         TabIndex        =   28
         Top             =   240
         Width           =   1125
      End
      Begin VB.ComboBox CboCaseStatus 
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
         Left            =   5760
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox TxtCaseNo2 
         Height          =   315
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   24
         Top             =   3435
         Width           =   1125
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   960
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   555
         Width           =   1125
      End
      Begin VB.CheckBox ChkDOA 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   4440
         TabIndex        =   14
         Top             =   5160
         Width           =   495
      End
      Begin VB.CheckBox ChkDOD 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   4440
         TabIndex        =   13
         Top             =   5520
         Width           =   495
      End
      Begin VB.CheckBox chkExpDate 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   4440
         TabIndex        =   12
         Top             =   6375
         Width           =   615
      End
      Begin VB.CheckBox chkDOR 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   4440
         TabIndex        =   11
         Top             =   5880
         Width           =   615
      End
      Begin VB.CheckBox ChkEffDate 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   280
         Left            =   4440
         TabIndex        =   10
         Top             =   6120
         Width           =   495
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   960
         Sorted          =   -1  'True
         TabIndex        =   9
         Text            =   "S - Sihat"
         Top             =   200
         Width           =   1125
      End
      Begin VB.TextBox TxtPlan1 
         Height          =   315
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   8
         Top             =   3720
         Width           =   1125
      End
      Begin VB.TextBox TxtPlan2 
         Height          =   315
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   7
         Top             =   3720
         Width           =   1125
      End
      Begin MSMask.MaskEdBox TxtPayEffDate1 
         Height          =   315
         Left            =   1440
         TabIndex        =   15
         Top             =   6075
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtExpDate1 
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Top             =   6360
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOR1 
         Height          =   315
         Left            =   1440
         TabIndex        =   17
         Top             =   5835
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPayEffDate2 
         Height          =   315
         Left            =   3240
         TabIndex        =   18
         Top             =   6120
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtExpDate2 
         Height          =   315
         Left            =   3240
         TabIndex        =   19
         Top             =   6360
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOR2 
         Height          =   315
         Left            =   3240
         TabIndex        =   20
         Top             =   5835
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOD1 
         Height          =   315
         Left            =   1440
         TabIndex        =   22
         Top             =   5475
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOA1 
         Height          =   315
         Left            =   1440
         TabIndex        =   23
         Top             =   5160
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOA2 
         Height          =   315
         Left            =   3240
         TabIndex        =   25
         Top             =   5160
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOD2 
         Height          =   315
         Left            =   3240
         TabIndex        =   26
         Top             =   5475
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtJoinDate1 
         Height          =   315
         Left            =   5160
         TabIndex        =   135
         Top             =   4920
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtJoinDate2 
         Height          =   315
         Left            =   5160
         TabIndex        =   136
         Top             =   5520
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         Caption         =   "Joined Date"
         Height          =   255
         Left            =   5160
         TabIndex        =   138
         ToolTipText     =   "MBM First Joined Date"
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "To"
         Height          =   255
         Left            =   5160
         TabIndex        =   137
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label34 
         Caption         =   "Mem Type"
         Height          =   255
         Left            =   4440
         TabIndex        =   133
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label27 
         Caption         =   "Supp. Type"
         Height          =   255
         Left            =   4440
         TabIndex        =   131
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Hospitals Grp"
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Hospitals"
         Height          =   255
         Left            =   120
         TabIndex        =   128
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Final Diag."
         Height          =   255
         Left            =   120
         TabIndex        =   123
         Top             =   4965
         Width           =   975
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   2640
         TabIndex        =   122
         Top             =   4845
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "First Diag."
         Height          =   255
         Left            =   120
         TabIndex        =   119
         Top             =   4605
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   2640
         TabIndex        =   118
         Top             =   4485
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Attend Doctor"
         Height          =   255
         Left            =   120
         TabIndex        =   115
         Top             =   3165
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Nature of Claims"
         Height          =   255
         Left            =   120
         TabIndex        =   113
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   2640
         TabIndex        =   111
         Top             =   4125
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   4245
         Width           =   735
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   2640
         TabIndex        =   55
         Top             =   5520
         Width           =   375
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   2640
         TabIndex        =   54
         Top             =   5160
         Width           =   375
      End
      Begin VB.Label Label48 
         Caption         =   "DOD"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         ToolTipText     =   "Date of Discharged"
         Top             =   5520
         Width           =   615
      End
      Begin VB.Label Label47 
         Caption         =   "DOA"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         ToolTipText     =   "Date of Admission"
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Label42 
         Caption         =   "Case No"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   2640
         TabIndex        =   50
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   2760
         TabIndex        =   49
         Top             =   5880
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "To"
         Height          =   255
         Left            =   2760
         TabIndex        =   48
         Top             =   6360
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "To"
         Height          =   255
         Left            =   2760
         TabIndex        =   47
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label Label21 
         Caption         =   "DOR"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         ToolTipText     =   "Case Reg. Date"
         Top             =   5880
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Payor Exp Date"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         ToolTipText     =   "Payor Expiry Date"
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "Payor Eff Date"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         ToolTipText     =   "Payor Effective Date"
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label73 
         Caption         =   "HLT Type"
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
         Top             =   250
         Width           =   705
      End
      Begin VB.Label Label26 
         Caption         =   "INS Type"
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
         TabIndex        =   42
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label38 
         Caption         =   "PAY Code"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1935
         Width           =   975
      End
      Begin VB.Label Label72 
         Caption         =   "Data Status"
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
         Left            =   2280
         TabIndex        =   40
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label46 
         Caption         =   "To"
         Height          =   255
         Left            =   2760
         TabIndex        =   39
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label Label51 
         Caption         =   "Plan Code"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label52 
         Caption         =   "Policy Status"
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
         Left            =   2280
         TabIndex        =   37
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label53 
         Caption         =   "Claimability"
         Height          =   255
         Left            =   4800
         TabIndex        =   36
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label54 
         Caption         =   "Case Status"
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
         TabIndex        =   35
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label4 
         Caption         =   "Grp Company"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2280
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reports Type"
      Height          =   1815
      Left            =   8760
      TabIndex        =   0
      Top             =   5400
      Width           =   2175
      Begin VB.OptionButton Option3 
         Caption         =   "Hospital"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Principals and Supp."
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Policy Effective Date"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Analysis By Post Code"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton Option7 
         Caption         =   "By Repudiated"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   1920
         Width           =   1935
      End
   End
   Begin MSComctlLib.ListView lstHospitalList 
      Height          =   4695
      Left            =   7320
      TabIndex        =   67
      Top             =   2520
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hospitals"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "No of Cases"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Total Actual Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Total Approved Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Actual Average"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Approved Average"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstPrincipalsList 
      Height          =   4695
      Left            =   7320
      TabIndex        =   68
      Top             =   2520
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Insured Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Member Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Total Cases"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Total Duration"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total Actual"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Total Approved"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstPolicyList 
      Height          =   4695
      Left            =   7320
      TabIndex        =   69
      Top             =   2520
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Case No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Final Diag"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "DOA"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Policy Eff. Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Date Joined"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Duration Fr EffDate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Duration Fr Date Joined"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ListView lstPCodeList 
      Height          =   4695
      Left            =   7320
      TabIndex        =   79
      Top             =   2520
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Post Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "No of Cases"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Total Actual Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Total Approved Amt"
         Object.Width           =   2469
      EndProperty
   End
   Begin MSComctlLib.ListView lstRepList 
      Height          =   4695
      Left            =   7320
      TabIndex        =   80
      Top             =   2520
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Member No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Case No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Patient Name"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Reason for Repudiation"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "DOA"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Frame Frame6 
      Caption         =   "Search By"
      Height          =   2055
      Left            =   7320
      TabIndex        =   75
      Top             =   480
      Width           =   4455
      Begin VB.TextBox TxtFinal 
         Height          =   315
         Left            =   1080
         TabIndex        =   77
         Top             =   720
         Width           =   1560
      End
      Begin VB.CommandButton CmdFinal 
         Caption         =   ">"
         Height          =   255
         Left            =   2760
         TabIndex        =   76
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "Final Diag."
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Search By"
      Height          =   2055
      Left            =   7320
      TabIndex        =   70
      Top             =   480
      Width           =   4455
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   1560
         Width           =   1995
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Hospital"
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Hospital Group"
         Height          =   195
         Left            =   1320
         TabIndex        =   71
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Selection"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Search By"
      Height          =   2055
      Left            =   7320
      TabIndex        =   104
      Top             =   480
      Width           =   4455
      Begin VB.ComboBox CboInsCode 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   435
         Width           =   2955
      End
      Begin VB.ComboBox CboMember 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   1035
         Width           =   3315
      End
      Begin VB.Label Label6 
         Caption         =   "Insured Type"
         Height          =   255
         Left            =   120
         TabIndex        =   108
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Member Type"
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Search By"
      Height          =   2055
      Left            =   7320
      TabIndex        =   99
      Top             =   480
      Width           =   4455
      Begin VB.TextBox TxtPCode1 
         Height          =   285
         Left            =   960
         MaxLength       =   5
         TabIndex        =   101
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   100
         Top             =   840
         Width           =   2355
      End
      Begin VB.Label Label22 
         Caption         =   "Post Code"
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Selection"
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   900
         Width           =   735
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "Search By"
      Height          =   2055
      Left            =   7320
      TabIndex        =   81
      Top             =   480
      Width           =   4455
      Begin VB.ComboBox CboOCP 
         Height          =   315
         Left            =   1080
         TabIndex        =   90
         Top             =   510
         Width           =   3135
      End
      Begin VB.ComboBox CboAdultChild 
         Height          =   315
         Left            =   3120
         TabIndex        =   89
         Top             =   900
         Width           =   1095
      End
      Begin VB.ComboBox CboSex 
         Height          =   315
         Left            =   3120
         TabIndex        =   88
         Top             =   1185
         Width           =   1095
      End
      Begin VB.TextBox TxtPostCode 
         Height          =   315
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   87
         Top             =   1500
         Width           =   975
      End
      Begin VB.TextBox TxtPolYear 
         Height          =   315
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   86
         Top             =   1500
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   ">"
         Height          =   255
         Left            =   1800
         TabIndex        =   85
         Top             =   885
         Width           =   255
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         TabIndex        =   84
         Top             =   825
         Width           =   615
      End
      Begin VB.ComboBox CboHOS2 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   83
         Top             =   180
         Width           =   3135
      End
      Begin VB.ComboBox CboRace 
         Height          =   315
         Left            =   1080
         TabIndex        =   82
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label33 
         Caption         =   "Occupation"
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Sex"
         Height          =   255
         Left            =   2160
         TabIndex        =   96
         Top             =   1275
         Width           =   735
      End
      Begin VB.Label Label30 
         Caption         =   "Race Code"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "Adult/Child"
         Height          =   255
         Left            =   2160
         TabIndex        =   94
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label35 
         Caption         =   "Post Code"
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label36 
         Caption         =   "Policy Year"
         Height          =   255
         Left            =   2160
         TabIndex        =   92
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label37 
         Caption         =   "ICD Code"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   915
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmAnalysisReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DiagCode, PostCode, InsureCode As String
Public HOSCode, HOSGRPCode As String
Public TotalActual, TotalApproved As String
Public Total, TotalDay, TotalDay2 As String
Public TotalDuration As String
Public Description As String
Public strAppend As String
Private DurationEffDate, DurationJoinedDate As String
Private m_blnDirection(1 To 8) As Boolean

'================ Load Payor ===============================
Private Sub LoadPayor(Optional HLTCode As String)
    Dim PAY As New ADODB.Recordset
    If Mid(ChkStr(HLTCode), 1, 1) = "N" Then
        sqlPAY = "select PAYCode, PAYCompanyName from PAY "
        sqlPAY = sqlPAY + " Where HLTCODE = 'N' "
    Else
        sqlPAY = "select PAYCode, PAYCompanyName from PAY "
        sqlPAY = sqlPAY + " Where HLTCODE = 'S' "
    End If
    
    cboPayorCode.Clear
    PAY.Open sqlPAY, medixmasterconn, adOpenKeyset, adLockOptimistic
    Do Until PAY.EOF
        cboPayorCode.AddItem ChkStr(PAY!PAYCode) & " - " & ChkStr(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    
    PAY.Close
    Set PAY = Nothing
End Sub
'================ Load Payor ===============================


Private Sub CmdBack_Click()
    If Option3.value = True Then
        Frame1.Visible = True
        Frame3.Visible = True
        Frame17.Visible = True
        lstHospitalList.Visible = True
        lstHospitalList.Height = 4695
        lstHospitalList.Left = 7320
        lstHospitalList.Top = 2520
        lstHospitalList.Width = 4455
   ElseIf Option4.value = True Then
        Frame1.Visible = True
        Frame4.Visible = True
        Frame17.Visible = True
        lstPrincipalsList.Visible = True
        lstPrincipalsList.Height = 4695
        lstPrincipalsList.Left = 7320
        lstPrincipalsList.Top = 2520
        lstPrincipalsList.Width = 4455
   ElseIf Option5.value = True Then
        Frame1.Visible = True
        Frame6.Visible = True
        Frame17.Visible = True
        lstPolicyList.Visible = True
        lstPolicyList.Height = 4695
        lstPolicyList.Left = 7320
        lstPolicyList.Top = 2520
        lstPolicyList.Width = 4455
   ElseIf Option6.value = True Then
        Frame1.Visible = True
        Frame9.Visible = True
        Frame17.Visible = True
        lstPCodeList.Visible = True
        lstPCodeList.Height = 4695
        lstPCodeList.Left = 7320
        lstPCodeList.Top = 2520
        lstPCodeList.Width = 4455
  ElseIf Option7.value = True Then
        Frame1.Visible = True
        Frame14.Visible = True
        Frame17.Visible = True
        lstRepList.Visible = True
        lstRepList.Height = 4695
        lstRepList.Left = 7320
        lstRepList.Top = 2520
        lstRepList.Width = 4455
  End If
End Sub

Private Sub CmdView_Click()
    If Option3.value = True Then
        Frame1.Visible = False
        Frame3.Visible = False
        Frame17.Visible = False
        lstHospitalList.Visible = True
        lstHospitalList.Height = 6720
        lstHospitalList.Left = 0
        lstHospitalList.Top = 480
        lstHospitalList.Width = 11775
   ElseIf Option4.value = True Then
        Frame1.Visible = False
        Frame4.Visible = False
        Frame17.Visible = False
        lstPrincipalsList.Visible = True
        lstPrincipalsList.Height = 6720
        lstPrincipalsList.Left = 0
        lstPrincipalsList.Top = 480
        lstPrincipalsList.Width = 11775
   ElseIf Option5.value = True Then
        Frame1.Visible = False
        Frame6.Visible = False
        Frame17.Visible = False
        lstPolicyList.Visible = True
        lstPolicyList.Height = 6720
        lstPolicyList.Left = 0
        lstPolicyList.Top = 480
        lstPolicyList.Width = 11775
   ElseIf Option6.value = True Then
        Frame1.Visible = False
        Frame9.Visible = False
        Frame17.Visible = False
        lstPCodeList.Visible = True
        lstPCodeList.Height = 6720
        lstPCodeList.Left = 0
        lstPCodeList.Top = 480
        lstPCodeList.Width = 11775
  ElseIf Option7.value = True Then
        Frame1.Visible = False
        Frame14.Visible = False
        Frame17.Visible = False
        lstRepList.Visible = True
        lstRepList.Height = 6720
        lstRepList.Left = 0
        lstRepList.Top = 480
        lstRepList.Width = 11775
  End If
End Sub

Private Sub Form_Load()
    Label5.Caption = "Hospitals"
    Frame14.Visible = True
    'cboHOS.Visible = True
    'cboHOSGrp.Visible = False
    Call ComboListing
    Call LoadPayor
End Sub

'============== ComboBox Listing =================================
Private Sub ComboListing()
Dim HOS As New ADODB.Recordset
Dim HOSGrp As New ADODB.Recordset
Dim PAY As New ADODB.Recordset
Dim INS As New ADODB.Recordset
Dim OCP As New ADODB.Recordset
Dim INSCode As New ADODB.Recordset
Dim PolStatus As New ADODB.Recordset
Dim ClaimNature As New ADODB.Recordset
    
    HOS.Open "select HosName from HOS", medixmasterconn, adOpenKeyset, adLockOptimistic
    HOSGrp.Open "select HosGrpName from HOSGRP", medixmasterconn, adOpenKeyset, adLockOptimistic
    PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
    INS.Open "Select INSCode, INSDescription from INS", medixmasterconn, adOpenKeyset, adLockOptimistic
    OCP.Open "Select OCPCode, OCPDescription from OCP", medixmasterconn, adOpenKeyset, adLockOptimistic
    PolStatus.Open "Select pstatuscode ,pstatusDesc from PolicyStatus", medixmasterconn, adOpenKeyset, adLockOptimistic
    INSCode.Open " Select INSCode, INSDescription from INS ", medixmasterconn, adOpenKeyset, adLockOptimistic
    ClaimNature.Open " Select ClaimNatureCode, ClaimNatureName from ClaimNature ", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until INSCode.EOF
        Combo2.AddItem ChkStr(INSCode!INSCode) & " - " & ChkStr(INSCode!INSDescription)
        INSCode.MoveNext
    Loop
    INSCode.Close
    Set INSCode = Nothing
    
    Do Until ClaimNature.EOF
        CboClaimNature.AddItem ChkStr(ClaimNature!ClaimNatureCode) & " - " & ChkStr(ClaimNature!ClaimNatureName)
        ClaimNature.MoveNext
    Loop
    ClaimNature.Close
    Set ClaimNature = Nothing
        
    Do Until HOS.EOF
        cboHOS.AddItem Trim(HOS!HOSName)
        CboHOS2.AddItem Trim(HOS!HOSName)
        HOS.MoveNext
    Loop
    HOS.Close
    Set HOS = Nothing
    
    Do Until HOSGrp.EOF
        cboHOSGrp.AddItem Trim(HOSGrp!HOSGRPName)
        HOSGrp.MoveNext
    Loop
    HOSGrp.Close
    Set HOSGrp = Nothing
    
    Do Until PAY.EOF
        cboPayorCode.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
    'Do Until INS.EOF
        'CboInsCode.AddItem ChkStr(INS!INSCode) & " - " & ChkStr(INS!INSDescription)
        'INS.MoveNext
    'Loop
    'INS.Close
    'Set INS = Nothing
    
    Do Until OCP.EOF
        CboOCP.AddItem ChkStr(OCP!OCPCode) & " - " & ChkStr(OCP!OCPDescription)
        OCP.MoveNext
    Loop
    
    OCP.Close
    Set OCP = Nothing
    
    CboRace.AddItem "C - Chinese"
    CboRace.AddItem "I - Indian"
    CboRace.AddItem "M - Malay"
    CboRace.AddItem "O - Other"
        
    CboAdultChild.AddItem "A - Adult"
    CboAdultChild.AddItem "M - Child"
    
    CboSex.AddItem "M" & " - " & "Male"
    CboSex.AddItem "F" & " - " & "Female"
    
                
    CboInsCode.AddItem "A" & " - " & "All"
    CboInsCode.AddItem "I" & " - " & "Individual"
    CboInsCode.AddItem "G" & " - " & "Group Individual"
    CboInsCode.AddItem "F" & " - " & "Family"
    CboInsCode.AddItem "H" & " - " & "Group Family"
    CboInsCode.Text = "A" & " - " & "All"
    
    If Option1.value = True Then
        Combo1.AddItem "0" & " - " & "Top 10"
        Combo1.AddItem "1" & " - " & "Top 20"
        Combo1.AddItem "2" & " - " & "Selected Records"
        Combo1.Text = "2" & " - " & "Selected Records"
    End If
    
    If Option2.value = True Then
        Combo1.AddItem "2" & " - " & "Selected Records"
        Combo1.Text = "2" & " - " & "Selected Records"
    End If
    
    
    Combo3.AddItem "0" & " - " & "Top 10"
    Combo3.AddItem "1" & " - " & "Top 20"
    Combo3.AddItem "2" & " - " & "Selected Records"
    Combo3.Text = "2" & " - " & "Selected Records"
    
    cboHLTCode.AddItem "S" & " - " & "Sihat"
    cboHLTCode.AddItem "N" & " - " & "Non-Sihat"
    
    With CboCaseStatus
        .AddItem "O - Open"
        .AddItem "C - Close"
        .AddItem "D - Delete"
        .AddItem "A - All"
    End With
    
    With CboClaimability
        .AddItem "A - Accepted"
        .AddItem "R - Rejected"
    End With
    
    CboMbmType.AddItem "Both"
    CboMbmType.AddItem "P" & " - " & "Principal"
    CboMbmType.AddItem "S" & " - " & "Supplementary"
    CboMbmType.Text = "Both"
    
    cboPrincipalType.AddItem "Both"
    cboPrincipalType.AddItem "C" & " - " & "Combined Limit"
    cboPrincipalType.AddItem "S" & " - " & "Separated Limit"
    cboPrincipalType.Text = "C" & " - " & "Combined Limit"
    
    cboDataStatus.AddItem "A - Active"
    cboDataStatus.AddItem "D - Delete"
    
        
    Do Until PolStatus.EOF
        CboStatus.AddItem (PolStatus!pstatuscode) & " - " & (PolStatus!pstatusDesc)
        PolStatus.MoveNext
    Loop
    
    PolStatus.Close
    Set PolStatus = Nothing
       
End Sub
'============== ComboBox Listing =================================

' =================== Search Diag. Selected Results ======================

Private Function SearchFinalDiag()
'Dim strAppend As String
    
    strAppend = ""
    
    If Len(Trim(txtFinalICD)) Then
        strAppend = strAppend + " And FinalDiag = '" & Trim(txtFinalICD) & "'"
    End If
    
    Call FinalDiagListing(strAppend)

End Function
' =================== Search Diag. Selected Results ======================


' =================== Search HOS Selected Results ======================
Private Function SearchHOS()
Dim strAppend As String
    
    strAppend = ""
    
    If Option1.value = True Then
        If Len(Trim(cboHOS)) Then
            strAppend = strAppend + " And HOSName = '" & Trim(cboHOS) & "'"
        End If
    End If
    
    If Option2.value = True Then
        If Len(Trim(cboHOSGrp)) Then
            strAppend = strAppend + " And HOSGRPName = '" & Trim(cboHOSGrp) & "'"
        End If
    End If
    
    If Len(Trim(cboPrincipalType)) And InStr(1, cboPrincipalType, "-") <> 0 Then
           If Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(Mid(CboMbmType, 1, 1)) = "P" Then
               strAppend = strAppend & " AND (MemberType = 'P')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(Mid(CboMbmType, 1, 1)) = "S" Then
               strAppend = strAppend & " AND (MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(CboMbmType) = "Both" Then
               strAppend = strAppend & " AND (MemberType = 'P' or MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(Mid(CboMbmType, 1, 1)) = "P" Then
               strAppend = strAppend & " AND (MemberType = 'Q')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(Mid(CboMbmType, 1, 1)) = "S" Then
               strAppend = strAppend & " AND (MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(CboMbmType) = "Both" Then
               strAppend = strAppend & " AND (MemberType = 'Q' or MemberType = 'C')"
           End If
      End If
    
    If Len(Trim(TxtCaseNo1)) Then
       strAppend = strAppend + " AND CASNumber >= '" & RemStr(Trim(TxtCaseNo1)) & "'"
    End If
    
    If Len(Trim(TxtCaseNo2)) Then
       strAppend = strAppend + " AND CASNumber <= '" & RemStr(Trim(TxtCaseNo2)) & "'"
    End If
    
    If Len(Trim(TxtBordxNo1)) Then
       strAppend = strAppend + " AND BordxRefNo >= '" & RemStr(Trim(TxtBordxNo1)) & "'"
    End If
    
    If Len(Trim(TxtBordxNo2)) Then
       strAppend = strAppend + " AND BordxRefNo <= '" & RemStr(Trim(TxtBordxNo2)) & "'"
    End If
    
    If Len(Trim(TxtAttDoc)) Then
       strAppend = strAppend + " And AttendDoctor like '%" & ChkStr(TxtAttDoc) & "%' "
    End If
    
    If Len(Trim(CboClaimNature)) Then
        strAppend = strAppend + " And ClaimNatureCode = '" & Trim(Mid(CboClaimNature, 1, IIf(InStr(1, CboClaimNature, "-") = 0, 4, InStr(1, CboClaimNature, "-") - 1))) & "'"
    End If
                
    If Mid(CboCaseStatus, 1, 1) = "C" Then
        strAppend = strAppend + " And CASStatus = 'C'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "D" Then
        strAppend = strAppend + " And CASStatus = 'D'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "O" Then
        strAppend = strAppend + " And CASStatus = 'O'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "A" Then
        strAppend = strAppend + " And (CASStatus = 'C' OR CASStatus = 'O')"
    End If
    
    If Mid(CboClaimability, 1, 1) = "R" Then
        strAppend = strAppend + " AND CASClaimability = 'R'"
    ElseIf Mid(CboClaimability, 1, 1) = "A" Then
        strAppend = strAppend + " AND CASClaimability = 'A'"
    End If
    
    If ChkDOA.value = 1 Then
        sql = ""
        sql = "AND CasDateAdmitted Is null " & ""
        strAppend = strAppend + sql
    End If
    
    
    If ChkDOD.value = 1 Then
        sql = ""
        sql = "AND CasDischargeDate Is null " & ""
        strAppend = strAppend + sql
    End If
    
    If (TxtDOA1) <> "__/__/____" And IsDate(TxtDOA1) = True _
        And (TxtDOA2) <> "__/__/____" And IsDate(TxtDOA2) = True Then
        sql = ""
        sql = " AND CasDateAdmitted >= '" & Format(Trim(TxtDOA1), "mm/dd/yyyy") & _
              "' And CasDateAdmitted <= '" & Format(Trim(TxtDOA2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtDOA1)) > 0 And IsDate(TxtDOA1) = True Then
        strAppend = strAppend + " And CasDateAdmitted >= '" & Format(TxtDOA1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOA2)) > 0 And IsDate(TxtDOA2) = True Then
        strAppend = strAppend + " And CasDateAdmitted <= '" & Format(TxtDOA2, "mm/dd/yyyy") & "'"
    End If
    
    If (txtDOD1) <> "__/__/____" And IsDate(txtDOD1) = True _
        And (txtDOD2) <> "__/__/____" And IsDate(txtDOD2) = True Then
        sql = ""
        sql = " AND CasDischargeDate >= '" & Format(Trim(txtDOD1), "mm/dd/yyyy") & _
              "' And CasDischargeDate <= '" & Format(Trim(txtDOD2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
        
    If Len(Trim(txtDOD1)) > 0 And IsDate(txtDOD1) = True Then
        strAppend = strAppend + " And CasDischargeDate >= '" & Format(txtDOD1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOD2)) > 0 And IsDate(txtDOD2) = True Then
        strAppend = strAppend + " And CasDischargeDate <= '" & Format(txtDOD2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Mid(cboHLTCode, 1, 1)) Then
        strAppend = strAppend & "AND HLTCode = '" & Mid(cboHLTCode, 1, 1) & "' "
    End If
    
    If Len(Trim(cboDataStatus)) Then
        strAppend = strAppend & "AND MBMDataStatus = '" & Trim(Mid(cboDataStatus, 1, InStr(1, cboDataStatus, "-") - 1)) & "' "
    End If
    
    If Len(Trim(TxtPlan1)) Then
        strAppend = strAppend + " AND PLNCode >= '" & Trim(TxtPlan1) & "'"
    End If
    
    If Len(Trim(TxtPlan2)) Then
        strAppend = strAppend + " AND PLNCode <= '" & Trim(TxtPlan2) & "'"
    End If
    
    If Len(Mid(cboPayorCode, 1, 2)) Then
        strAppend = strAppend & " AND PAYCode = '" & Mid(cboPayorCode, 1, 2) & "' "
    End If
    
    If (TxtPayEffDate1) <> "__/__/____" And IsDate(TxtPayEffDate1) = True _
        And (TxtPayEffDate2) <> "__/__/____" And IsDate(TxtPayEffDate2) = True Then
        sql = ""
        sql = " AND MBMPayorEffDate >= '" & Format(Trim(TxtPayEffDate1), "mm/dd/yyyy") & _
              "' And MBMPayorEffDate <= '" & Format(Trim(TxtPayEffDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtPayEffDate1)) > 0 And IsDate(TxtPayEffDate1) = True Then
        strAppend = strAppend + " And MBMPayorEffDate >= '" & Format(TxtEffDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtPayEffDate2)) > 0 And IsDate(TxtPayEffDate2) = True Then
        strAppend = strAppend + " And MBMPayorEffDate <= '" & Format(TxtPayEffDate2, "mm/dd/yyyy") & "'"
    End If
      
    If ChkEffDate.value = 1 Then
        sql = ""
        sql = "AND MBMPayorEffDate Is null " & ""
        strAppend = strAppend + sql
    End If
      
    If (TxtExpDate1) <> "__/__/____" And IsDate(TxtExpDate1) = True _
        And (TxtExpDate2) <> "__/__/____" And IsDate(TxtExpDate2) = True Then
        sql = ""
        sql = " AND MBMPayorExpDate >= '" & Format(Trim(TxtExpDate1), "mm/dd/yyyy") & _
              "' And MBMPayorExpDate <= '" & Format(Trim(TxtExpDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtExpDate1)) > 0 And IsDate(TxtExpDate1) = True Then
        strAppend = strAppend + " And MBMPayorExpDate >= '" & Format(TxtExpDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtExpDate2)) > 0 And IsDate(TxtExpDate2) = True Then
        strAppend = strAppend + " And MBMPayorExpDate <= '" & Format(TxtExpDate2, "mm/dd/yyyy") & "'"
    End If
      
    If chkExpDate.value = 1 Then
        sql = ""
        sql = "AND MBMPayorExpDate Is null " & ""
        strAppend = strAppend + sql
    End If
      
    If (TxtDOR1) <> "__/__/____" And IsDate(TxtDOR1) = True _
        And (TxtDOR2) <> "__/__/____" And IsDate(TxtDOR2) = True Then
        sql = ""
        sql = " AND CaseRegDate >= '" & Format(Trim(TxtDOR1), "mm/dd/yyyy") & _
              "' And CaseRegDate <= '" & Format(Trim(TxtDOR2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtDOR1)) > 0 And IsDate(TxtDOR1) = True Then
        strAppend = strAppend + " And CaseRegDate >= '" & Format(TxtDOR1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOR2)) > 0 And IsDate(TxtDOR2) = True Then
        strAppend = strAppend + " And CaseRegDate <= '" & Format(TxtDOR2, "mm/dd/yyyy") & "'"
    End If
      
    If chkDOR.value = 1 Then
        sql = ""
        sql = "AND CaseRegDate Is null " & ""
        strAppend = strAppend + sql
    End If
      
    If Len(Trim(Combo2)) Then
           strAppend = strAppend & " AND INSCode = '" & Trim(Mid(Combo2, 1, IIf(InStr(1, Combo2, "-") = 0, 1, InStr(1, Combo2, "-") - 1))) & "' "
    End If
      
    If Len(Trim(CboStatus)) Then
          strAppend = strAppend & "AND MBMStatus = '" & Trim(Mid(CboStatus, 1, InStr(1, CboStatus, "-") - 1)) & "' "
    End If
      
    'If (TxtBordx1) <> "__/__/____" And IsDate(TxtBordx1) = True _
        And (TxtBordx2) <> "__/__/____" And IsDate(TxtBordx2) = True Then
        'sql = ""
        'sql = " AND MBMBordxDate >= '" & Format(Trim(TxtBordx1), "mm/dd/yyyy") & _
              "' And MBMBordxDate <= '" & Format(Trim(TxtBordx2), "mm/dd/yyyy") & "'"
        'strAppend = strAppend + sql
    'End If
    
    'If Len(Trim(TxtBordx1)) > 0 And IsDate(TxtBordx1) = True Then
        'strAppend = strAppend + " And MBMBordxDate >= '" & Format(TxtBordx1, "mm/dd/yyyy") & "'"
    'End If
    
    'If Len(Trim(TxtBordx2)) > 0 And IsDate(TxtBordx2) = True Then
        'strAppend = strAppend + " And MBMBordxDate <= '" & Format(TxtBordx2, "mm/dd/yyyy") & "'"
    'End If
      
    'If chkBordxDate.value = 1 Then
        'sql = ""
        'sql = "AND MBMBordxDate Is null " & ""
        'strAppend = strAppend + sql
    'End If
      
    'If Len(Trim(TxtBatch1)) Then
        'strAppend = strAppend + " AND MBMBatchNo >= '" & Trim(TxtBatch1) & "'"
    'End If
    
    'If Len(Trim(TxtBatch2)) Then
        'strAppend = strAppend + " AND MBMBatchNo <= '" & Trim(TxtBatch2) & "'"
    'End If
      
    'If chkBatch.value = 1 Then
        'sql = ""
        'sql = "AND MBMBatchNo Is null " & ""
        'strAppend = strAppend + sql
    'End If
      
    'If (TxtEndtBordx1) <> "__/__/____" And IsDate(TxtEndtBordx1) = True _
        And (TxtEndtBordx2) <> "__/__/____" And IsDate(TxtEndtBordx2) = True Then
        'sql = ""
        'sql = " AND MBMAEndtBordxDate >= '" & Format(Trim(TxtEndtBordx1), "mm/dd/yyyy") & _
              "' And MBMAEndtBordxDate <= '" & Format(Trim(TxtEndtBordx2), "mm/dd/yyyy") & "'"
        'strAppend = strAppend + sql
    'End If
    
    'If Len(Trim(TxtEndtBordx1)) > 0 And IsDate(TxtEndtBordx1) = True Then
        'strAppend = strAppend + " And MBMAEndtBordxDate >= '" & Format(TxtEndtBordx1, "mm/dd/yyyy") & "'"
    'End If
    
    'If Len(Trim(TxtEndtBordx2)) > 0 And IsDate(TxtEndtBordx2) = True Then
        'strAppend = strAppend + " And MBMAEndtBordxDate <= '" & Format(TxtEndtBordx2, "mm/dd/yyyy") & "'"
    'End If
      
    'If chkEndtBordx.value = 1 Then
        'sql = ""
        'sql = "AND MBMAEndtBordxDate Is null " & ""
        'strAppend = strAppend + sql
    'End If
      
    'If Len(Trim(TxtEndtBatch1)) Then
        'strAppend = strAppend + " AND MBMAEndtBatchNo >= '" & Trim(TxtBatch1) & "'"
    'End If
    
    'If Len(Trim(TxtEndtBatch2)) Then
        'strAppend = strAppend + " AND MBMAEndtBatchNo <= '" & Trim(TxtBatch2) & "'"
    'End If
      
    'If chkEndtBatch.value = 1 Then
        'sql = ""
        'sql = "AND MBMAEndtBatchNo Is null " & ""
        'strAppend = strAppend + sql
    'End If
    
    If Len(Trim(CboGrpCom)) Then
        strAppend = strAppend + " AND MBMGroupCompany like '" & RemStr(Mid(CboGrpCom, 1, IIf(InStr(1, CboGrpCom, "-") = 0, 4, InStr(1, CboGrpCom, "-") - 1))) & "%'"
    End If
    
    
    If Option1.value = True And Mid(Combo1, 1, 1) = "2" Then
        Call HOSListing(strAppend)
    ElseIf Option2.value = True And Mid(Combo1, 1, 1) = "2" Then
        Call HOSGRPListing(strAppend)
    ElseIf Option1.value = True And Mid(Combo1, 1, 1) = "0" Or Mid(Combo1, 1, 1) = "1" Then
        Call HOSTop1020(strAppend)
    End If
    
    

End Function
' =================== Search HOS Selected Results ======================

' =================== Search Post Code Selected Results ======================
Private Function SearchPCode()
Dim strAppend As String
    
    strAppend = ""
    
    If Len(Trim(TxtCaseNo1)) Then
       strAppend = strAppend + " AND CASNumber >= '" & RemStr(Trim(TxtCaseNo1)) & "'"
    End If
    
    If Len(Trim(TxtCaseNo2)) Then
       strAppend = strAppend + " AND CASNumber <= '" & RemStr(Trim(TxtCaseNo2)) & "'"
    End If
    
    If Len(Trim(cboPrincipalType)) And InStr(1, cboPrincipalType, "-") <> 0 Then
           If Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(Mid(CboMbmType, 1, 1)) = "P" Then
               strAppend = strAppend & " AND (MemberType = 'P')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(Mid(CboMbmType, 1, 1)) = "S" Then
               strAppend = strAppend & " AND (MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(CboMbmType) = "Both" Then
               strAppend = strAppend & " AND (MemberType = 'P' or MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(Mid(CboMbmType, 1, 1)) = "P" Then
               strAppend = strAppend & " AND (MemberType = 'Q')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(Mid(CboMbmType, 1, 1)) = "S" Then
               strAppend = strAppend & " AND (MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(CboMbmType) = "Both" Then
               strAppend = strAppend & " AND (MemberType = 'Q' or MemberType = 'C')"
           End If
      End If
    
    If Len(Trim(TxtAttDoc)) Then
       strAppend = strAppend + " And AttendDoctor like '%" & ChkStr(TxtAttDoc) & "%' "
    End If
    
    If Len(Trim(CboClaimNature)) Then
        strAppend = strAppend + " And ClaimNatureCode = '" & Trim(Mid(CboClaimNature, 1, IIf(InStr(1, CboClaimNature, "-") = 0, 4, InStr(1, CboClaimNature, "-") - 1))) & "'"
    End If
    
    If Len(Trim(TxtBordxNo1)) Then
       strAppend = strAppend + " AND BordxRefNo >= '" & RemStr(Trim(TxtBordxNo1)) & "'"
    End If
    
    If Len(Trim(TxtBordxNo2)) Then
       strAppend = strAppend + " AND BordxRefNo <= '" & RemStr(Trim(TxtBordxNo2)) & "'"
    End If
                
    If Mid(CboCaseStatus, 1, 1) = "C" Then
        strAppend = strAppend + " And CASStatus = 'C'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "D" Then
        strAppend = strAppend + " And CASStatus = 'D'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "O" Then
        strAppend = strAppend + " And CASStatus = 'O'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "A" Then
        strAppend = strAppend + " And (CASStatus = 'C' OR CASStatus = 'O')"
    End If
    
    If Mid(CboClaimability, 1, 1) = "R" Then
        strAppend = strAppend + " AND CASClaimability = 'R'"
    ElseIf Mid(CboClaimability, 1, 1) = "A" Then
        strAppend = strAppend + " AND CASClaimability = 'A'"
    End If
    
    If ChkDOA.value = 1 Then
        sql = ""
        sql = "AND CasDateAdmitted Is null " & ""
        strAppend = strAppend + sql
    End If
    
    
    If ChkDOD.value = 1 Then
        sql = ""
        sql = "AND CasDischargeDate Is null " & ""
        strAppend = strAppend + sql
    End If
    
    If (TxtDOA1) <> "__/__/____" And IsDate(TxtDOA1) = True _
        And (TxtDOA2) <> "__/__/____" And IsDate(TxtDOA2) = True Then
        sql = ""
        sql = " AND CasDateAdmitted >= '" & Format(Trim(TxtDOA1), "mm/dd/yyyy") & _
              "' And CasDateAdmitted <= '" & Format(Trim(TxtDOA2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtDOA1)) > 0 And IsDate(TxtDOA1) = True Then
        strAppend = strAppend + " And CasDateAdmitted >= '" & Format(TxtDOA1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOA2)) > 0 And IsDate(TxtDOA2) = True Then
        strAppend = strAppend + " And CasDateAdmitted <= '" & Format(TxtDOA2, "mm/dd/yyyy") & "'"
    End If
    
    If (txtDOD1) <> "__/__/____" And IsDate(txtDOD1) = True _
        And (txtDOD2) <> "__/__/____" And IsDate(txtDOD2) = True Then
        sql = ""
        sql = " AND CasDischargeDate >= '" & Format(Trim(txtDOD1), "mm/dd/yyyy") & _
              "' And CasDischargeDate <= '" & Format(Trim(txtDOD2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDOD1)) > 0 And IsDate(txtDOD1) = True Then
        strAppend = strAppend + " And CasDischargeDate >= '" & Format(txtDOD1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOD2)) > 0 And IsDate(txtDOD2) = True Then
        strAppend = strAppend + " And CasDischargeDate <= '" & Format(txtDOD2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Mid(cboHLTCode, 1, 1)) Then
        strAppend = strAppend & "AND HLTCode = '" & Mid(cboHLTCode, 1, 1) & "' "
    End If
    
    If Len(Trim(cboDataStatus)) Then
        strAppend = strAppend & "AND MBMDataStatus = '" & Trim(Mid(cboDataStatus, 1, InStr(1, cboDataStatus, "-") - 1)) & "' "
    End If
    
    If Len(Trim(TxtPlan1)) Then
        strAppend = strAppend + " AND PLNCode >= '" & Trim(TxtPlan1) & "'"
    End If
    
    If Len(Trim(TxtPlan2)) Then
        strAppend = strAppend + " AND PLNCode <= '" & Trim(TxtPlan2) & "'"
    End If
    
    If Len(Mid(cboPayorCode, 1, 2)) Then
        strAppend = strAppend & " AND PAYCode = '" & Mid(cboPayorCode, 1, 2) & "' "
    End If
    
    If (TxtPayEffDate1) <> "__/__/____" And IsDate(TxtPayEffDate1) = True _
        And (TxtPayEffDate2) <> "__/__/____" And IsDate(TxtPayEffDate2) = True Then
        sql = ""
        sql = " AND MBMPayorEffDate >= '" & Format(Trim(TxtPayEffDate1), "mm/dd/yyyy") & _
              "' And MBMPayorEffDate <= '" & Format(Trim(TxtPayEffDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtPayEffDate1)) > 0 And IsDate(TxtPayEffDate1) = True Then
        strAppend = strAppend + " And MBMPayorEffDate >= '" & Format(TxtEffDate1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtPayEffDate2)) > 0 And IsDate(TxtPayEffDate2) = True Then
        strAppend = strAppend + " And MBMPayorEffDate <= '" & Format(TxtPayEffDate2, "mm/dd/yyyy") & "'"
      End If
      
      If ChkEffDate.value = 1 Then
        sql = ""
        sql = "AND MBMPayorEffDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If (TxtExpDate1) <> "__/__/____" And IsDate(TxtExpDate1) = True _
        And (TxtExpDate2) <> "__/__/____" And IsDate(TxtExpDate2) = True Then
        sql = ""
        sql = " AND MBMPayorExpDate >= '" & Format(Trim(TxtExpDate1), "mm/dd/yyyy") & _
              "' And MBMPayorExpDate <= '" & Format(Trim(TxtExpDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtExpDate1)) > 0 And IsDate(TxtExpDate1) = True Then
        strAppend = strAppend + " And MBMPayorExpDate >= '" & Format(TxtExpDate1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtExpDate2)) > 0 And IsDate(TxtExpDate2) = True Then
        strAppend = strAppend + " And MBMPayorExpDate <= '" & Format(TxtExpDate2, "mm/dd/yyyy") & "'"
      End If
      
      If chkExpDate.value = 1 Then
        sql = ""
        sql = "AND MBMPayorExpDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If (TxtDOR1) <> "__/__/____" And IsDate(TxtDOR1) = True _
        And (TxtDOR2) <> "__/__/____" And IsDate(TxtDOR2) = True Then
        sql = ""
        sql = " AND CaseRegDate >= '" & Format(Trim(TxtDOR1), "mm/dd/yyyy") & _
              "' And CaseRegDate <= '" & Format(Trim(TxtDOR2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtDOR1)) > 0 And IsDate(TxtDOR1) = True Then
        strAppend = strAppend + " And CaseRegDate >= '" & Format(TxtDOR1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtDOR2)) > 0 And IsDate(TxtDOR2) = True Then
        strAppend = strAppend + " And CaseRegDate <= '" & Format(TxtDOR2, "mm/dd/yyyy") & "'"
      End If
      
      If chkDOR.value = 1 Then
        sql = ""
        sql = "AND CaseRegDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If Len(Trim(Combo2)) Then
           strAppend = strAppend & " AND INSCode = '" & Trim(Mid(Combo2, 1, IIf(InStr(1, Combo2, "-") = 0, 1, InStr(1, Combo2, "-") - 1))) & "' "
      End If
      
      If Len(Trim(CboStatus)) Then
          strAppend = strAppend & "AND MBMStatus = '" & Trim(Mid(CboStatus, 1, InStr(1, CboStatus, "-") - 1)) & "' "
      End If
      
      'If (TxtBordx1) <> "__/__/____" And IsDate(TxtBordx1) = True _
        And (TxtBordx2) <> "__/__/____" And IsDate(TxtBordx2) = True Then
        'sql = ""
        'sql = " AND MBMBordxDate >= '" & Format(Trim(TxtBordx1), "mm/dd/yyyy") & _
              "' And MBMBordxDate <= '" & Format(Trim(TxtBordx2), "mm/dd/yyyy") & "'"
        'strAppend = strAppend + sql
      'End If
    
      'If Len(Trim(TxtBordx1)) > 0 And IsDate(TxtBordx1) = True Then
        'strAppend = strAppend + " And MBMBordxDate >= '" & Format(TxtBordx1, "mm/dd/yyyy") & "'"
      'End If
    
      'If Len(Trim(TxtBordx2)) > 0 And IsDate(TxtBordx2) = True Then
        'strAppend = strAppend + " And MBMBordxDate <= '" & Format(TxtBordx2, "mm/dd/yyyy") & "'"
      'End If
      
      'If chkBordxDate.value = 1 Then
        'sql = ""
        'sql = "AND MBMBordxDate Is null " & ""
        'strAppend = strAppend + sql
      'End If
      
      'If Len(Trim(TxtBatch1)) Then
        'strAppend = strAppend + " AND MBMBatchNo >= '" & Trim(TxtBatch1) & "'"
      'End If
    
      'If Len(Trim(TxtBatch2)) Then
        'strAppend = strAppend + " AND MBMBatchNo <= '" & Trim(TxtBatch2) & "'"
      'End If
      
      'If chkBatch.value = 1 Then
        'sql = ""
        'sql = "AND MBMBatchNo Is null " & ""
        'strAppend = strAppend + sql
      'End If
      
      'If (TxtEndtBordx1) <> "__/__/____" And IsDate(TxtEndtBordx1) = True _
        And (TxtEndtBordx2) <> "__/__/____" And IsDate(TxtEndtBordx2) = True Then
        'sql = ""
        'sql = " AND MBMAEndtBordxDate >= '" & Format(Trim(TxtEndtBordx1), "mm/dd/yyyy") & _
              '"' And MBMAEndtBordxDate <= '" & Format(Trim(TxtEndtBordx2), "mm/dd/yyyy") & "'"
        'strAppend = strAppend + sql
      'End If
    
      'If Len(Trim(TxtEndtBordx1)) > 0 And IsDate(TxtEndtBordx1) = True Then
        'strAppend = strAppend + " And MBMAEndtBordxDate >= '" & Format(TxtEndtBordx1, "mm/dd/yyyy") & "'"
      'End If
    
      'If Len(Trim(TxtEndtBordx2)) > 0 And IsDate(TxtEndtBordx2) = True Then
        'strAppend = strAppend + " And MBMAEndtBordxDate <= '" & Format(TxtEndtBordx2, "mm/dd/yyyy") & "'"
      'End If
      
      'If chkEndtBordx.value = 1 Then
        'sql = ""
        'sql = "AND MBMAEndtBordxDate Is null " & ""
        'strAppend = strAppend + sql
      'End If
      
      'If Len(Trim(TxtEndtBatch1)) Then
        'strAppend = strAppend + " AND MBMAEndtBatchNo >= '" & Trim(TxtBatch1) & "'"
      'End If
    
      'If Len(Trim(TxtEndtBatch2)) Then
        'strAppend = strAppend + " AND MBMAEndtBatchNo <= '" & Trim(TxtBatch2) & "'"
      'End If
      
      'If chkEndtBatch.value = 1 Then
        'sql = ""
        'sql = "AND MBMAEndtBatchNo Is null " & ""
        'strAppend = strAppend + sql
      'End If
    
    If Len(Trim(TxtPCode1)) Then
        strAppend = strAppend + " AND MBMPostCode = '" & Trim(TxtPCode1) & "'"
    End If
    
    If Len(Trim(CboGrpCom)) Then
        strAppend = strAppend + " AND MBMGroupCompany like '" & RemStr(Mid(CboGrpCom, 1, IIf(InStr(1, CboGrpCom, "-") = 0, 4, InStr(1, CboGrpCom, "-") - 1))) & "%'"
    End If
        
    If Mid(Combo3, 1, 1) = "2" Then
        Call PCodeListing(strAppend)
    ElseIf Mid(Combo3, 1, 1) = "0" Or Mid(Combo3, 1, 1) = "1" Then
        Call PCodeTop1020(strAppend)
    End If
    

End Function
' =================== Search Post Code Selected Results ======================

' =================== Search Repudiation Results ======================
Private Function SearchRepudiation()
Dim strAppend As String
    
    strAppend = ""
    
    If Len(Trim(TxtCaseNo1)) Then
       strAppend = strAppend + " AND CASNumber >= '" & RemStr(Trim(TxtCaseNo1)) & "'"
    End If
    
    If Len(Trim(TxtCaseNo2)) Then
       strAppend = strAppend + " AND CASNumber <= '" & RemStr(Trim(TxtCaseNo2)) & "'"
    End If
    
    If Len(Trim(cboPrincipalType)) And InStr(1, cboPrincipalType, "-") <> 0 Then
           If Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(Mid(CboMbmType, 1, 1)) = "P" Then
               strAppend = strAppend & " AND (MemberType = 'P')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(Mid(CboMbmType, 1, 1)) = "S" Then
               strAppend = strAppend & " AND (MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(CboMbmType) = "Both" Then
               strAppend = strAppend & " AND (MemberType = 'P' or MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(Mid(CboMbmType, 1, 1)) = "P" Then
               strAppend = strAppend & " AND (MemberType = 'Q')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(Mid(CboMbmType, 1, 1)) = "S" Then
               strAppend = strAppend & " AND (MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(CboMbmType) = "Both" Then
               strAppend = strAppend & " AND (MemberType = 'Q' or MemberType = 'C')"
           End If
      End If
    
    If Len(Trim(TxtAttDoc)) Then
       strAppend = strAppend + " And AttendDoctor like '%" & ChkStr(TxtAttDoc) & "%' "
    End If
    
    If Len(Trim(CboClaimNature)) Then
        strAppend = strAppend + " And ClaimNatureCode = '" & Trim(Mid(CboClaimNature, 1, IIf(InStr(1, CboClaimNature, "-") = 0, 4, InStr(1, CboClaimNature, "-") - 1))) & "'"
    End If
    
    If Len(Trim(TxtBordxNo1)) Then
       strAppend = strAppend + " AND BordxRefNo >= '" & RemStr(Trim(TxtBordxNo1)) & "'"
    End If
    
    If Len(Trim(TxtBordxNo2)) Then
       strAppend = strAppend + " AND BordxRefNo <= '" & RemStr(Trim(TxtBordxNo2)) & "'"
    End If
                
    If Mid(CboCaseStatus, 1, 1) = "C" Then
        strAppend = strAppend + " And CASStatus = 'C'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "D" Then
        strAppend = strAppend + " And CASStatus = 'D'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "O" Then
        strAppend = strAppend + " And CASStatus = 'O'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "A" Then
        strAppend = strAppend + " And (CASStatus = 'C' OR CASStatus = 'O')"
    End If
    
    If Mid(CboClaimability, 1, 1) = "R" Then
        strAppend = strAppend + " AND CASClaimability = 'R'"
    ElseIf Mid(CboClaimability, 1, 1) = "A" Then
        strAppend = strAppend + " AND CASClaimability = 'A'"
    End If
    
    If ChkDOA.value = 1 Then
        sql = ""
        sql = "AND CasDateAdmitted Is null " & ""
        strAppend = strAppend + sql
    End If
    
    
    If ChkDOD.value = 1 Then
        sql = ""
        sql = "AND CasDischargeDate Is null " & ""
        strAppend = strAppend + sql
    End If
    
    If (TxtDOA1) <> "__/__/____" And IsDate(TxtDOA1) = True _
        And (TxtDOA2) <> "__/__/____" And IsDate(TxtDOA2) = True Then
        sql = ""
        sql = " AND CasDateAdmitted >= '" & Format(Trim(TxtDOA1), "mm/dd/yyyy") & _
              "' And CasDateAdmitted <= '" & Format(Trim(TxtDOA2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtDOA1)) > 0 And IsDate(TxtDOA1) = True Then
        strAppend = strAppend + " And CasDateAdmitted >= '" & Format(TxtDOA1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOA2)) > 0 And IsDate(TxtDOA2) = True Then
        strAppend = strAppend + " And CasDateAdmitted <= '" & Format(TxtDOA2, "mm/dd/yyyy") & "'"
    End If
    
    If (txtDOD1) <> "__/__/____" And IsDate(txtDOD1) = True _
        And (txtDOD2) <> "__/__/____" And IsDate(txtDOD2) = True Then
        sql = ""
        sql = " AND CasDischargeDate >= '" & Format(Trim(txtDOD1), "mm/dd/yyyy") & _
              "' And CasDischargeDate <= '" & Format(Trim(txtDOD2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDOD1)) > 0 And IsDate(txtDOD1) = True Then
        strAppend = strAppend + " And CasDischargeDate >= '" & Format(txtDOD1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOD2)) > 0 And IsDate(txtDOD2) = True Then
        strAppend = strAppend + " And CasDischargeDate <= '" & Format(txtDOD2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Mid(cboHLTCode, 1, 1)) Then
        strAppend = strAppend & "AND HLTCode = '" & Mid(cboHLTCode, 1, 1) & "' "
    End If
    
    If Len(Trim(cboDataStatus)) Then
        strAppend = strAppend & "AND MBMDataStatus = '" & Trim(Mid(cboDataStatus, 1, InStr(1, cboDataStatus, "-") - 1)) & "' "
    End If
    
    If Len(Trim(TxtPlan1)) Then
        strAppend = strAppend + " AND PLNCode >= '" & Trim(TxtPlan1) & "'"
    End If
    
    If Len(Trim(TxtPlan2)) Then
        strAppend = strAppend + " AND PLNCode <= '" & Trim(TxtPlan2) & "'"
    End If
    
    If Len(Mid(cboPayorCode, 1, 2)) Then
        strAppend = strAppend & " AND PAYCode = '" & Mid(cboPayorCode, 1, 2) & "' "
    End If
    
    If (TxtPayEffDate1) <> "__/__/____" And IsDate(TxtPayEffDate1) = True _
        And (TxtPayEffDate2) <> "__/__/____" And IsDate(TxtPayEffDate2) = True Then
        sql = ""
        sql = " AND MBMPayorEffDate >= '" & Format(Trim(TxtPayEffDate1), "mm/dd/yyyy") & _
              "' And MBMPayorEffDate <= '" & Format(Trim(TxtPayEffDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtPayEffDate1)) > 0 And IsDate(TxtPayEffDate1) = True Then
        strAppend = strAppend + " And MBMPayorEffDate >= '" & Format(TxtEffDate1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtPayEffDate2)) > 0 And IsDate(TxtPayEffDate2) = True Then
        strAppend = strAppend + " And MBMPayorEffDate <= '" & Format(TxtPayEffDate2, "mm/dd/yyyy") & "'"
      End If
      
      If ChkEffDate.value = 1 Then
        sql = ""
        sql = "AND MBMPayorEffDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If (TxtExpDate1) <> "__/__/____" And IsDate(TxtExpDate1) = True _
        And (TxtExpDate2) <> "__/__/____" And IsDate(TxtExpDate2) = True Then
        sql = ""
        sql = " AND MBMPayorExpDate >= '" & Format(Trim(TxtExpDate1), "mm/dd/yyyy") & _
              "' And MBMPayorExpDate <= '" & Format(Trim(TxtExpDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtExpDate1)) > 0 And IsDate(TxtExpDate1) = True Then
        strAppend = strAppend + " And MBMPayorExpDate >= '" & Format(TxtExpDate1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtExpDate2)) > 0 And IsDate(TxtExpDate2) = True Then
        strAppend = strAppend + " And MBMPayorExpDate <= '" & Format(TxtExpDate2, "mm/dd/yyyy") & "'"
      End If
      
      If chkExpDate.value = 1 Then
        sql = ""
        sql = "AND MBMPayorExpDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If (TxtDOR1) <> "__/__/____" And IsDate(TxtDOR1) = True _
        And (TxtDOR2) <> "__/__/____" And IsDate(TxtDOR2) = True Then
        sql = ""
        sql = " AND CaseRegDate >= '" & Format(Trim(TxtDOR1), "mm/dd/yyyy") & _
              "' And CaseRegDate <= '" & Format(Trim(TxtDOR2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtDOR1)) > 0 And IsDate(TxtDOR1) = True Then
        strAppend = strAppend + " And CaseRegDate >= '" & Format(TxtDOR1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtDOR2)) > 0 And IsDate(TxtDOR2) = True Then
        strAppend = strAppend + " And CaseRegDate <= '" & Format(TxtDOR2, "mm/dd/yyyy") & "'"
      End If
      
      If chkDOR.value = 1 Then
        sql = ""
        sql = "AND CaseRegDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If Len(Trim(Combo2)) Then
           strAppend = strAppend & " AND INSCode = '" & Trim(Mid(Combo2, 1, IIf(InStr(1, Combo2, "-") = 0, 1, InStr(1, Combo2, "-") - 1))) & "' "
      End If
      
      If Len(Trim(CboStatus)) Then
          strAppend = strAppend & "AND MBMStatus = '" & Trim(Mid(CboStatus, 1, InStr(1, CboStatus, "-") - 1)) & "' "
      End If
      
      'If (TxtBordx1) <> "__/__/____" And IsDate(TxtBordx1) = True _
        And (TxtBordx2) <> "__/__/____" And IsDate(TxtBordx2) = True Then
        'sql = ""
        'sql = " AND MBMBordxDate >= '" & Format(Trim(TxtBordx1), "mm/dd/yyyy") & _
              "' And MBMBordxDate <= '" & Format(Trim(TxtBordx2), "mm/dd/yyyy") & "'"
        'strAppend = strAppend + sql
      'End If
    
      'If Len(Trim(TxtBordx1)) > 0 And IsDate(TxtBordx1) = True Then
        'strAppend = strAppend + " And MBMBordxDate >= '" & Format(TxtBordx1, "mm/dd/yyyy") & "'"
      'End If
    
      'If Len(Trim(TxtBordx2)) > 0 And IsDate(TxtBordx2) = True Then
        'strAppend = strAppend + " And MBMBordxDate <= '" & Format(TxtBordx2, "mm/dd/yyyy") & "'"
      'End If
      
      'If chkBordxDate.value = 1 Then
        'sql = ""
        'sql = "AND MBMBordxDate Is null " & ""
        'strAppend = strAppend + sql
      'End If
      
      'If Len(Trim(TxtBatch1)) Then
        'strAppend = strAppend + " AND MBMBatchNo >= '" & Trim(TxtBatch1) & "'"
      'End If
    
      'If Len(Trim(TxtBatch2)) Then
        'strAppend = strAppend + " AND MBMBatchNo <= '" & Trim(TxtBatch2) & "'"
      'End If
      
      'If chkBatch.value = 1 Then
        'sql = ""
        'sql = "AND MBMBatchNo Is null " & ""
        'strAppend = strAppend + sql
      'End If
      
      'If (TxtEndtBordx1) <> "__/__/____" And IsDate(TxtEndtBordx1) = True _
        And (TxtEndtBordx2) <> "__/__/____" And IsDate(TxtEndtBordx2) = True Then
        'sql = ""
        'sql = " AND MBMAEndtBordxDate >= '" & Format(Trim(TxtEndtBordx1), "mm/dd/yyyy") & _
              "' And MBMAEndtBordxDate <= '" & Format(Trim(TxtEndtBordx2), "mm/dd/yyyy") & "'"
        'strAppend = strAppend + sql
      'End If
    
      'If Len(Trim(TxtEndtBordx1)) > 0 And IsDate(TxtEndtBordx1) = True Then
        'strAppend = strAppend + " And MBMAEndtBordxDate >= '" & Format(TxtEndtBordx1, "mm/dd/yyyy") & "'"
      'End If
    
      'If Len(Trim(TxtEndtBordx2)) > 0 And IsDate(TxtEndtBordx2) = True Then
        'strAppend = strAppend + " And MBMAEndtBordxDate <= '" & Format(TxtEndtBordx2, "mm/dd/yyyy") & "'"
      'End If
      
      'If chkEndtBordx.value = 1 Then
        'sql = ""
        'sql = "AND MBMAEndtBordxDate Is null " & ""
        'strAppend = strAppend + sql
      'End If
      
      'If Len(Trim(TxtEndtBatch1)) Then
        'strAppend = strAppend + " AND MBMAEndtBatchNo >= '" & Trim(TxtBatch1) & "'"
      'End If
    
      'If Len(Trim(TxtEndtBatch2)) Then
        'strAppend = strAppend + " AND MBMAEndtBatchNo <= '" & Trim(TxtBatch2) & "'"
      'End If
      
      'If chkEndtBatch.value = 1 Then
        'sql = ""
        'sql = "AND MBMAEndtBatchNo Is null " & ""
        'strAppend = strAppend + sql
      'End If
    
        
    If Len(Trim(CboHOS2)) Then
        strAppend = strAppend + " And HOSName = '" & Trim(CboHOS2) & "'"
    End If
    
    If Len(Trim(Mid(CboOCP, 1, 2))) Then
          strAppend = strAppend & "AND MBMOccupation = '" & Trim(Mid(CboOCP, 1, 2)) & "' "
    End If
    
            
    If Len(Trim(Mid(CboSex, 1, 1))) Then
          strAppend = strAppend & "AND MBMSex = '" & Trim(Mid(CboSex, 1, 1)) & "' "
    End If
    
    If Len(Trim(Mid(CboAdultChild, 1, 1))) Then
          strAppend = strAppend & "AND MBMAdultChild = '" & Trim(Mid(CboAdultChild, 1, 1)) & "' "
    End If
    
    If Len(Trim(Mid(CboRace, 1, 1))) Then
          strAppend = strAppend & "AND RACCode = '" & Trim(Mid(CboRace, 1, 1)) & "' "
    End If
    
    If Len(Trim(CboGrpCom)) Then
        strAppend = strAppend + " AND MBMGroupCompany like '" & RemStr(Mid(CboGrpCom, 1, IIf(InStr(1, CboGrpCom, "-") = 0, 4, InStr(1, CboGrpCom, "-") - 1))) & "%'"
    End If
        
    If Len(Trim(TxtPolYear)) Then
        strAppend = strAppend + " AND MBMPolicyYear = '" & Trim(TxtPolYear) & "'"
    End If
    
    If Len(Trim(TxtPostCode)) Then
        strAppend = strAppend + " AND MBMPostCode = '" & Trim(TxtPostCode) & "'"
    End If
            
    If Len(Trim(Text3)) Then
        strAppend = strAppend + " AND ICDCode1 = '" & Trim(Text3) & "'"
    End If
    
    Call RepudiaListing(strAppend)

End Function
' =================== Search Repudiation Results ======================

Private Function SearchInsure()
    
    strAppend = ""
        
    'If Len(Mid(CboInsCode, 1, 1)) Then
        'strAppend = strAppend + " And INSCode = '" & Mid(CboInsCode, 1, 1) & "'"
    'End If
    
    If Mid(CboInsCode, 1, 1) = "I" Then
        strAppend = strAppend + " And INSCode = 'I' "
    ElseIf Mid(CboInsCode, 1, 1) = "F" Then
        strAppend = strAppend + " And INSCode = 'F' "
    ElseIf Mid(CboInsCode, 1, 1) = "G" Then
        strAppend = strAppend + " And INSCode = 'G' "
    ElseIf Mid(CboInsCode, 1, 1) = "H" Then
        strAppend = strAppend + " And INSCode = 'H' "
    ElseIf Mid(CboInsCode, 1, 1) = "A" Then
        strAppend = strAppend + " And (INSCode = 'I' OR INSCode = 'F' OR INSCode = 'G' OR INSCode = 'H') "
    End If
    
    If Len(Trim(cboPrincipalType)) And InStr(1, cboPrincipalType, "-") <> 0 Then
           If Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(Mid(CboMbmType, 1, 1)) = "P" Then
               strAppend = strAppend & " AND (MemberType = 'P')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(Mid(CboMbmType, 1, 1)) = "S" Then
               strAppend = strAppend & " AND (MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(CboMbmType) = "Both" Then
               strAppend = strAppend & " AND (MemberType = 'P' or MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(Mid(CboMbmType, 1, 1)) = "P" Then
               strAppend = strAppend & " AND (MemberType = 'Q')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(Mid(CboMbmType, 1, 1)) = "S" Then
               strAppend = strAppend & " AND (MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(CboMbmType) = "Both" Then
               strAppend = strAppend & " AND (MemberType = 'Q' or MemberType = 'C')"
           End If
      End If
    
    If Len(Trim(TxtCaseNo1)) Then
       strAppend = strAppend + " AND CASNumber >= '" & RemStr(Trim(TxtCaseNo1)) & "'"
    End If
    
    If Len(Trim(TxtCaseNo2)) Then
       strAppend = strAppend + " AND CASNumber <= '" & RemStr(Trim(TxtCaseNo2)) & "'"
    End If
    
    If Len(Trim(TxtAttDoc)) Then
       strAppend = strAppend + " And AttendDoctor like '%" & ChkStr(TxtAttDoc) & "%' "
    End If
    
    If Len(Trim(CboClaimNature)) Then
        strAppend = strAppend + " And ClaimNatureCode = '" & Trim(Mid(CboClaimNature, 1, IIf(InStr(1, CboClaimNature, "-") = 0, 4, InStr(1, CboClaimNature, "-") - 1))) & "'"
    End If
    
    If Len(Trim(TxtBordxNo1)) Then
       strAppend = strAppend + " AND BordxRefNo >= '" & RemStr(Trim(TxtBordxNo1)) & "'"
    End If
    
    If Len(Trim(TxtBordxNo2)) Then
       strAppend = strAppend + " AND BordxRefNo <= '" & RemStr(Trim(TxtBordxNo2)) & "'"
    End If
                
    If Mid(CboCaseStatus, 1, 1) = "C" Then
        strAppend = strAppend + " And CASStatus = 'C'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "D" Then
        strAppend = strAppend + " And CASStatus = 'D'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "O" Then
        strAppend = strAppend + " And CASStatus = 'O'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "A" Then
        strAppend = strAppend + " And (CASStatus = 'C' OR CASStatus = 'O')"
    End If
    
    If Mid(CboClaimability, 1, 1) = "R" Then
        strAppend = strAppend + " AND CASClaimability = 'R'"
    ElseIf Mid(CboClaimability, 1, 1) = "A" Then
        strAppend = strAppend + " AND CASClaimability = 'A'"
    End If
    
    If ChkDOA.value = 1 Then
        sql = ""
        sql = "AND CasDateAdmitted Is null " & ""
        strAppend = strAppend + sql
    End If
    
    
    If ChkDOD.value = 1 Then
        sql = ""
        sql = "AND CasDischargeDate Is null " & ""
        strAppend = strAppend + sql
    End If
    
    If (TxtDOA1) <> "__/__/____" And IsDate(TxtDOA1) = True _
        And (TxtDOA2) <> "__/__/____" And IsDate(TxtDOA2) = True Then
        sql = ""
        sql = " AND CasDateAdmitted >= '" & Format(Trim(TxtDOA1), "mm/dd/yyyy") & _
              "' And CasDateAdmitted <= '" & Format(Trim(TxtDOA2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtDOA1)) > 0 And IsDate(TxtDOA1) = True Then
        strAppend = strAppend + " And CasDateAdmitted >= '" & Format(TxtDOA1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOA2)) > 0 And IsDate(TxtDOA2) = True Then
        strAppend = strAppend + " And CasDateAdmitted <= '" & Format(TxtDOA2, "mm/dd/yyyy") & "'"
    End If
    
    If (txtDOD1) <> "__/__/____" And IsDate(txtDOD1) = True _
        And (txtDOD2) <> "__/__/____" And IsDate(txtDOD2) = True Then
        sql = ""
        sql = " AND CasDischargeDate >= '" & Format(Trim(txtDOD1), "mm/dd/yyyy") & _
              "' And CasDischargeDate <= '" & Format(Trim(txtDOD2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDOD1)) > 0 And IsDate(txtDOD1) = True Then
        strAppend = strAppend + " And CasDischargeDate >= '" & Format(txtDOD1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOD2)) > 0 And IsDate(txtDOD2) = True Then
        strAppend = strAppend + " And CasDischargeDate <= '" & Format(txtDOD2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Mid(cboHLTCode, 1, 1)) Then
        strAppend = strAppend & "AND HLTCode = '" & Mid(cboHLTCode, 1, 1) & "' "
    End If
    
    If Len(Trim(cboDataStatus)) Then
        strAppend = strAppend & "AND MBMDataStatus = '" & Trim(Mid(cboDataStatus, 1, InStr(1, cboDataStatus, "-") - 1)) & "' "
    End If
    
    If Len(Trim(TxtPlan1)) Then
        strAppend = strAppend + " AND PLNCode >= '" & Trim(TxtPlan1) & "'"
    End If
    
    If Len(Trim(TxtPlan2)) Then
        strAppend = strAppend + " AND PLNCode <= '" & Trim(TxtPlan2) & "'"
    End If
    
    If Len(Mid(cboPayorCode, 1, 2)) Then
        strAppend = strAppend & " AND PAYCode = '" & Mid(cboPayorCode, 1, 2) & "' "
    End If
    
    If (TxtPayEffDate1) <> "__/__/____" And IsDate(TxtPayEffDate1) = True _
        And (TxtPayEffDate2) <> "__/__/____" And IsDate(TxtPayEffDate2) = True Then
        sql = ""
        sql = " AND MBMPayorEffDate >= '" & Format(Trim(TxtPayEffDate1), "mm/dd/yyyy") & _
              "' And MBMPayorEffDate <= '" & Format(Trim(TxtPayEffDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtPayEffDate1)) > 0 And IsDate(TxtPayEffDate1) = True Then
        strAppend = strAppend + " And MBMPayorEffDate >= '" & Format(TxtEffDate1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtPayEffDate2)) > 0 And IsDate(TxtPayEffDate2) = True Then
        strAppend = strAppend + " And MBMPayorEffDate <= '" & Format(TxtPayEffDate2, "mm/dd/yyyy") & "'"
      End If
      
      If ChkEffDate.value = 1 Then
        sql = ""
        sql = "AND MBMPayorEffDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If (TxtExpDate1) <> "__/__/____" And IsDate(TxtExpDate1) = True _
        And (TxtExpDate2) <> "__/__/____" And IsDate(TxtExpDate2) = True Then
        sql = ""
        sql = " AND MBMPayorExpDate >= '" & Format(Trim(TxtExpDate1), "mm/dd/yyyy") & _
              "' And MBMPayorExpDate <= '" & Format(Trim(TxtExpDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtExpDate1)) > 0 And IsDate(TxtExpDate1) = True Then
        strAppend = strAppend + " And MBMPayorExpDate >= '" & Format(TxtExpDate1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtExpDate2)) > 0 And IsDate(TxtExpDate2) = True Then
        strAppend = strAppend + " And MBMPayorExpDate <= '" & Format(TxtExpDate2, "mm/dd/yyyy") & "'"
      End If
      
      If chkExpDate.value = 1 Then
        sql = ""
        sql = "AND MBMPayorExpDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If (TxtDOR1) <> "__/__/____" And IsDate(TxtDOR1) = True _
        And (TxtDOR2) <> "__/__/____" And IsDate(TxtDOR2) = True Then
        sql = ""
        sql = " AND CaseRegDate >= '" & Format(Trim(TxtDOR1), "mm/dd/yyyy") & _
              "' And CaseRegDate <= '" & Format(Trim(TxtDOR2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtDOR1)) > 0 And IsDate(TxtDOR1) = True Then
        strAppend = strAppend + " And CaseRegDate >= '" & Format(TxtDOR1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtDOR2)) > 0 And IsDate(TxtDOR2) = True Then
        strAppend = strAppend + " And CaseRegDate <= '" & Format(TxtDOR2, "mm/dd/yyyy") & "'"
      End If
      
      If chkDOR.value = 1 Then
        sql = ""
        sql = "AND CaseRegDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If Len(Trim(Combo2)) Then
           strAppend = strAppend & " AND INSCode = '" & Trim(Mid(Combo2, 1, IIf(InStr(1, Combo2, "-") = 0, 1, InStr(1, Combo2, "-") - 1))) & "' "
      End If
      
      If Len(Trim(CboStatus)) Then
          strAppend = strAppend & "AND MBMStatus = '" & Trim(Mid(CboStatus, 1, InStr(1, CboStatus, "-") - 1)) & "' "
      End If
      
      'If (TxtBordx1) <> "__/__/____" And IsDate(TxtBordx1) = True _
        And (TxtBordx2) <> "__/__/____" And IsDate(TxtBordx2) = True Then
        'sql = ""
        'sql = " AND MBMBordxDate >= '" & Format(Trim(TxtBordx1), "mm/dd/yyyy") & _
              "' And MBMBordxDate <= '" & Format(Trim(TxtBordx2), "mm/dd/yyyy") & "'"
        'strAppend = strAppend + sql
      'End If
    
      'If Len(Trim(TxtBordx1)) > 0 And IsDate(TxtBordx1) = True Then
        'strAppend = strAppend + " And MBMBordxDate >= '" & Format(TxtBordx1, "mm/dd/yyyy") & "'"
      'End If
    
      'If Len(Trim(TxtBordx2)) > 0 And IsDate(TxtBordx2) = True Then
        'strAppend = strAppend + " And MBMBordxDate <= '" & Format(TxtBordx2, "mm/dd/yyyy") & "'"
      'End If
      
      'If chkBordxDate.value = 1 Then
        'sql = ""
        'sql = "AND MBMBordxDate Is null " & ""
        'strAppend = strAppend + sql
      'End If
      
      'If Len(Trim(TxtBatch1)) Then
        'strAppend = strAppend + " AND MBMBatchNo >= '" & Trim(TxtBatch1) & "'"
      'End If
    
      'If Len(Trim(TxtBatch2)) Then
        'strAppend = strAppend + " AND MBMBatchNo <= '" & Trim(TxtBatch2) & "'"
      'End If
      
      'If chkBatch.value = 1 Then
        'sql = ""
        'sql = "AND MBMBatchNo Is null " & ""
        'strAppend = strAppend + sql
      'End If
      
      'If (TxtEndtBordx1) <> "__/__/____" And IsDate(TxtEndtBordx1) = True _
        And (TxtEndtBordx2) <> "__/__/____" And IsDate(TxtEndtBordx2) = True Then
        'sql = ""
        'sql = " AND MBMAEndtBordxDate >= '" & Format(Trim(TxtEndtBordx1), "mm/dd/yyyy") & _
              "' And MBMAEndtBordxDate <= '" & Format(Trim(TxtEndtBordx2), "mm/dd/yyyy") & "'"
        'strAppend = strAppend + sql
      'End If
    
      'If Len(Trim(TxtEndtBordx1)) > 0 And IsDate(TxtEndtBordx1) = True Then
        'strAppend = strAppend + " And MBMAEndtBordxDate >= '" & Format(TxtEndtBordx1, "mm/dd/yyyy") & "'"
      'End If
    
      'If Len(Trim(TxtEndtBordx2)) > 0 And IsDate(TxtEndtBordx2) = True Then
        'strAppend = strAppend + " And MBMAEndtBordxDate <= '" & Format(TxtEndtBordx2, "mm/dd/yyyy") & "'"
      'End If
      
      'If chkEndtBordx.value = 1 Then
        'sql = ""
        'sql = "AND MBMAEndtBordxDate Is null " & ""
        'strAppend = strAppend + sql
      'End If
      
      'If Len(Trim(TxtEndtBatch1)) Then
        'strAppend = strAppend + " AND MBMAEndtBatchNo >= '" & Trim(TxtBatch1) & "'"
      'End If
    
      'If Len(Trim(TxtEndtBatch2)) Then
        'strAppend = strAppend + " AND MBMAEndtBatchNo <= '" & Trim(TxtBatch2) & "'"
      'End If
      
      'If chkEndtBatch.value = 1 Then
        'sql = ""
        'sql = "AND MBMAEndtBatchNo Is null " & ""
        'strAppend = strAppend + sql
      'End If
    
    
    If Len(Mid(CboMember, 1, 1)) Then
        If Mid(CboMember, 1, 1) = "P" Then
            strAppend = strAppend + " And MemberType = 'P'"
        ElseIf Mid(CboMember, 1, 1) = "S" Then
            strAppend = strAppend + " And (MemberType = 'C' OR MemberType = 'Q')"
        ElseIf Mid(CboMember, 1, 1) = "B" Then
            If Mid(CboInsCode, 1, 1) = "F" Then
                strAppend = strAppend + " AND (MemberType = 'P' OR MemberType = 'Q' OR MemberType = 'C')"
            ElseIf Mid(CboInsCode, 1, 1) = "H" Then
                strAppend = strAppend + " AND (MemberType = 'P' OR MemberType = 'Q' OR MemberType = 'C')"
            End If
        ElseIf Mid(CboMember, 1, 1) = "A" Then
            strAppend = strAppend + " AND (MemberType = 'P' OR MemberType = 'Q' OR MemberType = 'C')"
        End If
    End If
    
    If Len(Trim(CboGrpCom)) Then
        strAppend = strAppend + " AND MBMGroupCompany like '" & RemStr(Mid(CboGrpCom, 1, IIf(InStr(1, CboGrpCom, "-") = 0, 4, InStr(1, CboGrpCom, "-") - 1))) & "%'"
    End If
    
    
    If Mid(CboInsCode, 1, 1) = "A" Then
        Call SearchAllInsure(strAppend)
    Else
        Call SearchInsuredType(strAppend)
    End If

End Function

Private Sub SearchInsuredType(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
    
    Total = 0
    TotalActual = 0
    TotalApproved = 0
    
    lstPrincipalsList.listitems.Clear
    
    sql = " SELECT * From VIEW_PrinNDep Where 0 = 0 "
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
        
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
                      
            With rst2
                Set Record = lstPrincipalsList.listitems.Add(, , rst2!INSCode)
                InsureCode = Trim(rst2!INSCode)
                                               
                ' ===== Member Type ======
                If Mid(CboMember, 1, 1) = "P" Then
                    Record.SubItems(1) = "Principals"
                ElseIf Mid(CboMember, 1, 1) = "S" Then
                    Record.SubItems(1) = "Supplementary"
                ElseIf Mid(CboMember, 1, 1) = "B" Then
                    Record.SubItems(1) = "Both"
                End If
                ' ===== Member Type ======
                
                Call SearchMemberType(strAppend)
                
                If Len(Total) Then
                   Record.SubItems(2) = Total
                End If
                
                If Len(TotalDay2) Then
                   Record.SubItems(3) = TotalDay2
                End If
                
                If Len(TotalActual) Then
                    Record.SubItems(4) = Format(TotalActual, "#,##0.00")
                End If
                                
                If Len(TotalApproved) Then
                   Record.SubItems(5) = Format(TotalApproved, "#,##0.00")
                End If
            End With
            rst2.MoveNext
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub SearchMemberType(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
    
    Total = 0
    TotalDay = 0
    TotalDay2 = 0
    TotalActual = 0
    TotalApproved = 0
            
    ListView11.listitems.Clear
            
    sql = " Select * From VIEW_PrinNDep Where 0 = 0 "
                  
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
       
        Do Until rst2.EOF
            
            Total = rst2.recordCount
            
            TotalDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TotalDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TotalDay = "0" Then
                TotalDay = TotalDay + 1
            End If
                    
            With rst2
            
                If Len(rst2!MemberType) Then
                    Set Record = ListView11.listitems.Add(, , rst2!MemberType)
                End If
                
                If Len(Total) Then
                    Record.SubItems(1) = Total
                End If
                
                If Len(rst2!CLMTotalActual) Then
                       Record.SubItems(2) = rst2!CLMTotalActual
                       TotalActual = TotalActual + rst2!CLMTotalActual
                End If
                
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(3) = rst2!CLMTotalApproved
                    TotalApproved = TotalApproved + rst2!CLMTotalApproved
                End If
                
                If Len(TotalDay) Then
                    TotalDay2 = TotalDay2 + TotalDay
                End If
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub

'============ Search All Insured ==========================
Private Sub SearchAllInsure(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
    
    TotalActual = 0
    TotalApproved = 0
        
    lstPrincipalsList.listitems.Clear
            
    
    sql = " SELECT Distinct InsCode, Count(InsCode) as Total From VIEW_PrinNDep Where 0 = 0 "
    
    sql2 = " Group by InsCode "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend + sql2
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    
        Do Until rst2.EOF
    
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = lstPrincipalsList.listitems.Add(, , rst2!INSCode)
                    InsureCode = Trim(rst2!INSCode)
                                
                    ' ===== Member Type ======
                    If InsureCode = "I" And CboMember.ListIndex = 1 Then
                        Record.SubItems(1) = "Principals"
                    ElseIf InsureCode = "F" And CboMember.ListIndex = 1 Then
                        Record.SubItems(1) = "Principals"
                    ElseIf InsureCode = "G" And CboMember.ListIndex = 1 Then
                        Record.SubItems(1) = "Principals"
                    ElseIf InsureCode = "H" And CboMember.ListIndex = 1 Then
                        Record.SubItems(1) = "Principals"
                    ElseIf InsureCode = "F" And CboMember.ListIndex = 2 Then
                        Record.SubItems(1) = "Supplementary"
                    ElseIf InsureCode = "H" And CboMember.ListIndex = 2 Then
                        Record.SubItems(1) = "Supplementary"
                    ElseIf InsureCode = "I" And CboMember.ListIndex = 0 Then
                        Record.SubItems(1) = "Principals"
                    ElseIf InsureCode = "F" And CboMember.ListIndex = 0 Then
                        Record.SubItems(1) = "All"
                    ElseIf InsureCode = "G" And CboMember.ListIndex = 0 Then
                        Record.SubItems(1) = "Principals"
                    ElseIf InsureCode = "H" And CboMember.ListIndex = 0 Then
                        Record.SubItems(1) = "All"
                    End If
                    ' ===== Member Type ======
                    
                    Call SearchAllMemberType(strAppend)
                    
                    If Len(rst2!Total) Then
                       Record.SubItems(2) = rst2!Total
                    End If
                    
                    If Len(TotalDay2) Then
                       Record.SubItems(3) = TotalDay2
                    End If
                    
                    If Len(TotalActual) Then
                        Record.SubItems(4) = Format(TotalActual, "#,##0.00")
                    End If
                                    
                    If Len(TotalApproved) Then
                       Record.SubItems(5) = Format(TotalApproved, "#,##0.00")
                    End If
                End If
            
            End With
            rst2.MoveNext
        Loop
            
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub SearchAllMemberType(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim Current1 As String
Dim Current2 As String

    Current1 = 0
    Current2 = 0
    
    Total = 0
    TotalDay = 0
    TotalDay2 = 0
    TotalActual = 0
    TotalApproved = 0
            
    ListView11.listitems.Clear
            
    sql = " Select INSCode, MemberType, CLMTotalActual, " & _
          " CLMTotalApproved, CASDischargeDate, CasDateAdmitted " & _
          " From VIEW_PrinNDep Where INSCode = '" & Trim(InsureCode) & "'"
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
        
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
            
            Total = rst2.recordCount
            
            TotalDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TotalDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TotalDay = "0" Then
                TotalDay = TotalDay + 1
            End If
                    
            With rst2
            
                If Len(rst2!MemberType) Then
                    Set Record = ListView11.listitems.Add(, , rst2!MemberType)
                End If
                
                If Len(Total) Then
                    Record.SubItems(1) = Total
                End If
                
                If Len(rst2!CLMTotalActual) Then
                       Record.SubItems(2) = rst2!CLMTotalActual
                       TotalActual = TotalActual + rst2!CLMTotalActual
                End If
                
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(3) = rst2!CLMTotalApproved
                    TotalApproved = TotalApproved + rst2!CLMTotalApproved
                End If
                
                If Len(TotalDay) Then
                    TotalDay2 = TotalDay2 + TotalDay
                End If
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
' ================== Search All Insure ===========================

Private Sub cboINSCode_Change()
    
    CboMember.Clear
    
    If InStr(CboInsCode.Text, "null") Then
       CboInsCode.Enabled = False
    End If
    
    If Mid(CboInsCode, 1, 1) = "I" Then
        CboMember.AddItem "P" & " - " & "Principal"
        CboMember.Text = "P" & " - " & "Principal"
    ElseIf Mid(CboInsCode, 1, 1) = "G" Then
        CboMember.AddItem "P" & " - " & "Principal"
        CboMember.Text = "P" & " - " & "Principal"
    ElseIf Mid(CboInsCode, 1, 1) = "F" Then
        CboMember.AddItem "P" & " - " & "Principal"
        CboMember.AddItem "S" & " - " & "Supplementary"
        CboMember.AddItem "B" & " - " & "Both"
        CboMember.Text = "B" & " - " & "Both"
    ElseIf Mid(CboInsCode, 1, 1) = "H" Then
        CboMember.AddItem "P" & " - " & "Principal"
        CboMember.AddItem "S" & " - " & "Supplementary"
        CboMember.AddItem "B" & " - " & "Both"
        CboMember.Text = "B" & " - " & "Both"
    ElseIf Mid(CboInsCode, 1, 1) = "A" Then
        CboMember.AddItem "A" & " - " & "All"
        CboMember.AddItem "P" & " - " & "Principal"
        CboMember.AddItem "S" & " - " & "Supplementary"
        CboMember.Text = "A" & " - " & "All"
    End If
    
End Sub


Private Sub CboInsCode_Click()
    
    CboMember.Clear
    
    If InStr(CboInsCode.Text, "null") Then
       CboInsCode.Enabled = False
    End If
           
    If Mid(CboInsCode, 1, 1) = "I" Then
        CboMember.AddItem "P" & " - " & "Principal"
        CboMember.Text = "P" & " - " & "Principal"
    ElseIf Mid(CboInsCode, 1, 1) = "G" Then
        CboMember.AddItem "P" & " - " & "Principal"
        CboMember.Text = "P" & " - " & "Principal"
    ElseIf Mid(CboInsCode, 1, 1) = "F" Then
        CboMember.AddItem "P" & " - " & "Principal"
        CboMember.AddItem "S" & " - " & "Supplementary"
        CboMember.AddItem "B" & " - " & "Both"
        CboMember.Text = "B" & " - " & "Both"
    ElseIf Mid(CboInsCode, 1, 1) = "H" Then
        CboMember.AddItem "P" & " - " & "Principal"
        CboMember.AddItem "S" & " - " & "Supplementary"
        CboMember.AddItem "B" & " - " & "Both"
        CboMember.Text = "B" & " - " & "Both"
    ElseIf Mid(CboInsCode, 1, 1) = "A" Then
        CboMember.AddItem "A" & " - " & "All"
        CboMember.AddItem "P" & " - " & "Principal"
        CboMember.AddItem "S" & " - " & "Supplementary"
        CboMember.Text = "A" & " - " & "All"
    End If
End Sub


'==================== List Final Diag. Listing Resullts ==========================
Private Sub FinalDiagListing()
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
       
    lstFinalDiagList.listitems.Clear
            
    sql = " SELECT FinalDiag, COUNT(FinalDiag) AS Total" & _
          " From VIEW_AnalysisReport1 GROUP BY FinalDiag ORDER BY Total DESC "
    
        
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
    
            With rst2
                Set Record = lstFinalDiagList.listitems.Add(, , rst2!FinalDiag)
                DiagCode = Trim(rst2!FinalDiag)
                
                Call SearchRecords
                
                If Len(Description) Then
                    Record.SubItems(1) = Trim(Description)
                End If
                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
                
                If Len(TotalDuration) Then
                   Record.SubItems(3) = TotalDuration
                End If
                
                If Len(TotalActual) Then
                   Record.SubItems(4) = TotalActual
                End If
                
                If Len(TotalApproved) Then
                    Record.SubItems(5) = TotalApproved
                End If
                                
                Record.SubItems(6) = Format(Round(TotalActual / rst2!Total, 0), "#,##0.00")
                Record.SubItems(7) = Format(Round(TotalApproved / rst2!Total, 0), "#,##0.00")
            
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub


Private Sub lstHospitalList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    
    Case 1
        SortListView lstHospitalList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstHospitalList, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
    Case 3
        SortListView lstHospitalList, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    Case 4
        SortListView lstHospitalList, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstHospitalList, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstHospitalList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    End Select
End Sub


Private Sub lstPCodeList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    
    Case 1
        SortListView lstPCodeList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstPCodeList, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
    Case 3
        SortListView lstPCodeList, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    Case 4
        SortListView lstPCodeList, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    End Select
End Sub


Private Sub lstPolicyList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    
    Case 1
        SortListView lstPolicyList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstPolicyList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstPolicyList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstPolicyList, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
    Case 5
        SortListView lstPolicyList, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
    Case 6
        SortListView lstPolicyList, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
    Case 7
        SortListView lstPolicyList, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstPolicyList, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
    End Select
End Sub


Private Sub lstPrincipalsList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    
    Case 1
        SortListView lstPrincipalsList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstPrincipalsList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstPrincipalsList, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    Case 4
        SortListView lstPrincipalsList, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstPrincipalsList, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstPrincipalsList, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
    End Select
End Sub


Private Sub lstRepList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    
    Case 1
        SortListView lstRepList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstRepList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstRepList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstRepList, ColumnHeader.Index, ldtString, m_blnDirection(4)
    Case 5
        SortListView lstRepList, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
    End Select
End Sub


Private Sub Option1_Click()

    Combo1.Clear
    
    If Option1.value = True Then
        'cboHOS.Visible = True
        'cboHOSGrp.Visible = False
        'Label5.Caption = "Hospitals"
        Combo1.AddItem "0" & " - " & "Top 10"
        Combo1.AddItem "1" & " - " & "Top 20"
        Combo1.AddItem "2" & " - " & "Selected Records"
        Combo1.Text = "2" & " - " & "Selected Records"
    End If
    
End Sub

Private Sub Option2_Click()
    
    Combo1.Clear
    
    If Option2.value = True Then
        'cboHOSGrp.Visible = True
        'cboHOS.Visible = False
        'Label5.Caption = "Hospitals Group"
        Combo1.AddItem "0" & " - " & "Top 10"
        Combo1.AddItem "1" & " - " & "Top 20"
        Combo1.AddItem "2" & " - " & "Selected Records"
        Combo1.Text = "2" & " - " & "Selected Records"
    End If
    
End Sub


Private Sub HOSListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
    
    lstHospitalList.listitems.Clear
            
    sql = " SELECT HOSName, COUNT(HOSName) AS Total" & _
          " From VIEW_AnalysisReport4 Where 0 = 0 "
    
    sql2 = " GROUP BY HOSName Order by HOSName "
    
    'If Len(Trim(strAppend)) Then
        sql = sql + strAppend + sql2
    'End If
                     
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
            
            Total = rst2.recordCount
            Do Until rst2.EOF
                
            With rst2
                Set Record = lstHospitalList.listitems.Add(, , rst2!HOSName)
                HOSCode = Trim(rst2!HOSName)
                
                Call SearchHOSRecords(strAppend)
                
                If Len(Total) Then
                    Record.SubItems(1) = Total
                End If
                
                If Len(TotalActual) Then
                   Record.SubItems(2) = TotalActual
                End If
                
                If Len(TotalApproved) Then
                    Record.SubItems(3) = TotalApproved
                End If
                                
                Record.SubItems(4) = Format(Round(TotalActual / Total, 0), "#,##0.00")
                Record.SubItems(5) = Format(Round(TotalApproved / Total, 0), "#,##0.00")
                           
            End With
            rst2.MoveNext
        Loop
        
    rst2.Close
    Set rst2 = Nothing
End Sub

' =================== Search HOS Top 10 / 20 Results ======================
Private Sub HOSTop1020(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
       
    lstHospitalList.listitems.Clear
            
    If Mid(Combo1, 1, 1) = "0" Then
    
        sql = " SELECT TOP 10 HOSName, COUNT(HOSName) AS Total" & _
              " From VIEW_AnalysisReport4 Where 0 = 0 "
          
    ElseIf Mid(Combo1, 1, 1) = "1" Then
    
        sql = " SELECT TOP 20 HOSName, COUNT(HOSName) AS Total" & _
              " From VIEW_AnalysisReport4 Where 0 = 0 "
    
    End If
    
    sql2 = " GROUP BY HOSName ORDER BY Total DESC "
    
    sql = sql + strAppend + sql2
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
    
            With rst2
                Set Record = lstHospitalList.listitems.Add(, , rst2!HOSName)
                HOSCode = Trim(rst2!HOSName)
                
                Call SearchHOSRecords(strAppend)
                
                If Len(rst2!Total) Then
                    Record.SubItems(1) = rst2!Total
                End If
                
                If Len(TotalActual) Then
                   Record.SubItems(2) = TotalActual
                End If
                
                If Len(TotalApproved) Then
                    Record.SubItems(3) = TotalApproved
                End If
                                
                Record.SubItems(4) = Format(Round(TotalActual / rst2!Total, 0), "#,##0.00")
                Record.SubItems(5) = Format(Round(TotalApproved / rst2!Total, 0), "#,##0.00")
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
' =================== Search HOS Top 10 / 20 Results ======================

' =================== Count HOS Records ======================
Private Sub SearchHOSRecords(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
        
    Total = 0
    TotalActual = 0
    TotalApproved = 0
    TotalDuration = 0
    
    ListView11.listitems.Clear
            
    sql = " Select * From VIEW_AnalysisReport4 where HOSName = '" & Trim(HOSCode) & "'"
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        
        Do Until rst2.EOF
                    
            With rst2
                
                Set Record = ListView11.listitems.Add(, , rst2!HOSName)
                
                If Len(rst2!CLMTotalActual) Then
                    Record.SubItems(1) = Format(rst2!CLMTotalActual, "#,##0.00")
                End If
                
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(2) = Format(rst2!CLMTotalApproved, "#,##0.00")
                End If
                                                               
                If Len(rst2!CLMTotalActual) Then
                    TotalActual = TotalActual + rst2!CLMTotalActual
                End If
                                
                If Len(rst2!CLMTotalApproved) Then
                    TotalApproved = TotalApproved + rst2!CLMTotalApproved
                End If
                
                Total = Total + 1
                
            End With
            rst2.MoveNext
        Loop
    
    rst2.Close
    Set rst2 = Nothing
End Sub
' =================== Count HOS Records ======================

Private Sub HOSGRPListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String

    lstHospitalList.listitems.Clear
    
    sql = " SELECT HOSGRPName, COUNT(HOSGRPName) AS Total" & _
          " From VIEW_AnalysisReport4 Where 0 = 0 "
    
    sql2 = " GROUP BY HOSGRPName Order by HOSGRPName "
    
    'If Len(Trim(strAppend)) Then
        sql = sql + strAppend + sql2
    'End If
    
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                
            With rst2
                Set Record = lstHospitalList.listitems.Add(, , rst2!HOSGRPName)
                HOSGRPCode = Trim(rst2!HOSGRPName)
                
                Call SearchHOSGRPRecords(strAppend)
                                               
                If Len(Total) Then
                    Record.SubItems(1) = Total
                End If
                                                
                If Len(TotalActual) Then
                   Record.SubItems(2) = TotalActual
                End If
                
                If Len(TotalApproved) Then
                    Record.SubItems(3) = TotalApproved
                End If
                                
                Record.SubItems(4) = Format(Round(TotalActual / Total, 0), "#,##0.00")
                Record.SubItems(5) = Format(Round(TotalApproved / Total, 0), "#,##0.00")
                           
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub


Private Sub SearchHOSGRPRecords(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
        
    Total = 0
    TotalActual = 0
    TotalApproved = 0
    TotalDuration = 0
    
    ListView11.listitems.Clear
            
    sql = " Select * From VIEW_AnalysisReport4 where HOSGRPName = '" & Trim(HOSGRPCode) & "'"
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
        
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                    
            With rst2
                
                Set Record = ListView11.listitems.Add(, , rst2!HOSGRPName)
                
                If Len(rst2!CLMTotalActual) Then
                    Record.SubItems(1) = Format(rst2!CLMTotalActual, "#,##0.00")
                End If
                
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(2) = Format(rst2!CLMTotalApproved, "#,##0.00")
                End If
                                                
                If Len(rst2!CLMTotalActual) Then
                    TotalActual = TotalActual + rst2!CLMTotalActual
                End If
                                
                If Len(rst2!CLMTotalApproved) Then
                    TotalApproved = TotalApproved + rst2!CLMTotalApproved
                End If
                
            Total = Total + 1
                
            End With
            rst2.MoveNext
        Loop
    
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub PolEffDateListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String

    DurationEffDate = 0
    DurationJoinedDate = 0
       
    lstPolicyList.listitems.Clear
            
    sql = " SELECT * From VIEW_AnalysisReport5 Where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
    
        Do Until rst2.EOF
        
            DurationEffDate = ""
            If IsDate(rst2!CASDateAdmitted) And Len(rst2!MBMPayorEffDate) Then
                DurationEffDate = DateValue(rst2!CASDateAdmitted) - DateValue((rst2!MBMPayorEffDate))
            End If
            
            If DurationEffDate = "0" Then
                DurationEffDate = DurationEffDate + 1
            End If
            
            DurationJoinedDate = ""
            If IsDate(rst2!CASDateAdmitted) And Len(rst2!MBMFirstJoinedDate) Then
                DurationJoinedDate = DateValue(rst2!CASDateAdmitted) - DateValue((rst2!MBMFirstJoinedDate))
            End If
            
            If DurationJoinedDate = "0" Then
                DurationJoinedDate = DurationJoinedDate + 1
            End If
    
            With rst2
                Set Record = lstPolicyList.listitems.Add(, , rst2!CASNumber)
                                                
                If Len(rst2!FinalDiag) Then
                    Record.SubItems(1) = Trim(rst2!FinalDiag)
                End If
                
                If Len(rst2!ICDDescription) Then
                    Record.SubItems(2) = Trim(rst2!ICDDescription)
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(3) = rst2!CASDateAdmitted
                End If
                
                If Len(rst2!MBMPayorEffDate) Then
                   Record.SubItems(4) = rst2!MBMPayorEffDate
                End If
                
                If Len(rst2!MBMFirstJoinedDate) Then
                   Record.SubItems(5) = rst2!MBMFirstJoinedDate
                End If
                
                If Len(DurationEffDate) Then
                    Record.SubItems(6) = DurationEffDate
                End If
                                
                If Len(DurationJoinedDate) Then
                    Record.SubItems(7) = DurationJoinedDate
                End If
            
            End With
            rst2.MoveNext
        Loop
    End If
    rst2.Close
    Set rst2 = Nothing
End Sub


Private Function SearchDateRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
Dim CAS As New ADODB.Recordset
    
    strAppend = ""
    
    If Len(Trim(TxtCaseNo1)) Then
       strAppend = strAppend + " AND CASNumber >= '" & RemStr(Trim(TxtCaseNo1)) & "'"
    End If
    
    If Len(Trim(TxtCaseNo2)) Then
       strAppend = strAppend + " AND CASNumber <= '" & RemStr(Trim(TxtCaseNo2)) & "'"
    End If
    
    If Len(Trim(cboPrincipalType)) And InStr(1, cboPrincipalType, "-") <> 0 Then
           If Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(Mid(CboMbmType, 1, 1)) = "P" Then
               strAppend = strAppend & " AND (MemberType = 'P')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(Mid(CboMbmType, 1, 1)) = "S" Then
               strAppend = strAppend & " AND (MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(CboMbmType) = "Both" Then
               strAppend = strAppend & " AND (MemberType = 'P' or MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(Mid(CboMbmType, 1, 1)) = "P" Then
               strAppend = strAppend & " AND (MemberType = 'Q')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(Mid(CboMbmType, 1, 1)) = "S" Then
               strAppend = strAppend & " AND (MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(CboMbmType) = "Both" Then
               strAppend = strAppend & " AND (MemberType = 'Q' or MemberType = 'C')"
           End If
      End If
    
    If Len(Trim(TxtAttDoc)) Then
       strAppend = strAppend + " And AttendDoctor like '%" & ChkStr(TxtAttDoc) & "%' "
    End If
    
    If Len(Trim(CboClaimNature)) Then
        strAppend = strAppend + " And ClaimNatureCode = '" & Trim(Mid(CboClaimNature, 1, IIf(InStr(1, CboClaimNature, "-") = 0, 4, InStr(1, CboClaimNature, "-") - 1))) & "'"
    End If
    
    If Len(Trim(TxtBordxNo1)) Then
       strAppend = strAppend + " AND BordxRefNo >= '" & RemStr(Trim(TxtBordxNo1)) & "'"
    End If
    
    If Len(Trim(TxtBordxNo2)) Then
       strAppend = strAppend + " AND BordxRefNo <= '" & RemStr(Trim(TxtBordxNo2)) & "'"
    End If
                
    If Mid(CboCaseStatus, 1, 1) = "C" Then
        strAppend = strAppend + " And CASStatus = 'C'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "D" Then
        strAppend = strAppend + " And CASStatus = 'D'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "O" Then
        strAppend = strAppend + " And CASStatus = 'O'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "A" Then
        strAppend = strAppend + " And (CASStatus = 'C' OR CASStatus = 'O')"
    End If
    
    If Mid(CboClaimability, 1, 1) = "R" Then
        strAppend = strAppend + " AND CASClaimability = 'R'"
    ElseIf Mid(CboClaimability, 1, 1) = "A" Then
        strAppend = strAppend + " AND CASClaimability = 'A'"
    End If
    
    If ChkDOA.value = 1 Then
        sql = ""
        sql = "AND CasDateAdmitted Is null " & ""
        strAppend = strAppend + sql
    End If
    
    
    If ChkDOD.value = 1 Then
        sql = ""
        sql = "AND CasDischargeDate Is null " & ""
        strAppend = strAppend + sql
    End If
    
    If (TxtDOA1) <> "__/__/____" And IsDate(TxtDOA1) = True _
        And (TxtDOA2) <> "__/__/____" And IsDate(TxtDOA2) = True Then
        sql = ""
        sql = " AND CasDateAdmitted >= '" & Format(Trim(TxtDOA1), "mm/dd/yyyy") & _
              "' And CasDateAdmitted <= '" & Format(Trim(TxtDOA2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtDOA1)) > 0 And IsDate(TxtDOA1) = True Then
        strAppend = strAppend + " And CasDateAdmitted >= '" & Format(TxtDOA1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOA2)) > 0 And IsDate(TxtDOA2) = True Then
        strAppend = strAppend + " And CasDateAdmitted <= '" & Format(TxtDOA2, "mm/dd/yyyy") & "'"
    End If
    
    If (txtDOD1) <> "__/__/____" And IsDate(txtDOD1) = True _
        And (txtDOD2) <> "__/__/____" And IsDate(txtDOD2) = True Then
        sql = ""
        sql = " AND CasDischargeDate >= '" & Format(Trim(txtDOD1), "mm/dd/yyyy") & _
              "' And CasDischargeDate <= '" & Format(Trim(txtDOD2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDOD1)) > 0 And IsDate(txtDOD1) = True Then
        strAppend = strAppend + " And CasDischargeDate >= '" & Format(txtDOD1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOD2)) > 0 And IsDate(txtDOD2) = True Then
        strAppend = strAppend + " And CasDischargeDate <= '" & Format(txtDOD2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Mid(cboHLTCode, 1, 1)) Then
        strAppend = strAppend & "AND HLTCode = '" & Mid(cboHLTCode, 1, 1) & "' "
    End If
    
    If Len(Trim(cboDataStatus)) Then
        strAppend = strAppend & "AND MBMDataStatus = '" & Trim(Mid(cboDataStatus, 1, InStr(1, cboDataStatus, "-") - 1)) & "' "
    End If
    
    If Len(Trim(TxtPlan1)) Then
        strAppend = strAppend + " AND PLNCode >= '" & Trim(TxtPlan1) & "'"
    End If
    
    If Len(Trim(TxtPlan2)) Then
        strAppend = strAppend + " AND PLNCode <= '" & Trim(TxtPlan2) & "'"
    End If
    
    If Len(Mid(cboPayorCode, 1, 2)) Then
        strAppend = strAppend & " AND PAYCode = '" & Mid(cboPayorCode, 1, 2) & "' "
    End If
    
    If (TxtPayEffDate1) <> "__/__/____" And IsDate(TxtPayEffDate1) = True _
        And (TxtPayEffDate2) <> "__/__/____" And IsDate(TxtPayEffDate2) = True Then
        sql = ""
        sql = " AND MBMPayorEffDate >= '" & Format(Trim(TxtPayEffDate1), "mm/dd/yyyy") & _
              "' And MBMPayorEffDate <= '" & Format(Trim(TxtPayEffDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtPayEffDate1)) > 0 And IsDate(TxtPayEffDate1) = True Then
        strAppend = strAppend + " And MBMPayorEffDate >= '" & Format(TxtEffDate1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtPayEffDate2)) > 0 And IsDate(TxtPayEffDate2) = True Then
        strAppend = strAppend + " And MBMPayorEffDate <= '" & Format(TxtPayEffDate2, "mm/dd/yyyy") & "'"
      End If
      
      If ChkEffDate.value = 1 Then
        sql = ""
        sql = "AND MBMPayorEffDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If (TxtExpDate1) <> "__/__/____" And IsDate(TxtExpDate1) = True _
        And (TxtExpDate2) <> "__/__/____" And IsDate(TxtExpDate2) = True Then
        sql = ""
        sql = " AND MBMPayorExpDate >= '" & Format(Trim(TxtExpDate1), "mm/dd/yyyy") & _
              "' And MBMPayorExpDate <= '" & Format(Trim(TxtExpDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtExpDate1)) > 0 And IsDate(TxtExpDate1) = True Then
        strAppend = strAppend + " And MBMPayorExpDate >= '" & Format(TxtExpDate1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtExpDate2)) > 0 And IsDate(TxtExpDate2) = True Then
        strAppend = strAppend + " And MBMPayorExpDate <= '" & Format(TxtExpDate2, "mm/dd/yyyy") & "'"
      End If
      
      If chkExpDate.value = 1 Then
        sql = ""
        sql = "AND MBMPayorExpDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If (TxtDOR1) <> "__/__/____" And IsDate(TxtDOR1) = True _
        And (TxtDOR2) <> "__/__/____" And IsDate(TxtDOR2) = True Then
        sql = ""
        sql = " AND CaseRegDate >= '" & Format(Trim(TxtDOR1), "mm/dd/yyyy") & _
              "' And CaseRegDate <= '" & Format(Trim(TxtDOR2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtDOR1)) > 0 And IsDate(TxtDOR1) = True Then
        strAppend = strAppend + " And CaseRegDate >= '" & Format(TxtDOR1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtDOR2)) > 0 And IsDate(TxtDOR2) = True Then
        strAppend = strAppend + " And CaseRegDate <= '" & Format(TxtDOR2, "mm/dd/yyyy") & "'"
      End If
      
      If chkDOR.value = 1 Then
        sql = ""
        sql = "AND CaseRegDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If Len(Trim(Combo2)) Then
           strAppend = strAppend & " AND INSCode = '" & Trim(Mid(Combo2, 1, IIf(InStr(1, Combo2, "-") = 0, 1, InStr(1, Combo2, "-") - 1))) & "' "
      End If
      
      If Len(Trim(CboStatus)) Then
          strAppend = strAppend & "AND MBMStatus = '" & Trim(Mid(CboStatus, 1, InStr(1, CboStatus, "-") - 1)) & "' "
      End If
      
      'If (TxtBordx1) <> "__/__/____" And IsDate(TxtBordx1) = True _
        And (TxtBordx2) <> "__/__/____" And IsDate(TxtBordx2) = True Then
        'sql = ""
        'sql = " AND MBMBordxDate >= '" & Format(Trim(TxtBordx1), "mm/dd/yyyy") & _
              "' And MBMBordxDate <= '" & Format(Trim(TxtBordx2), "mm/dd/yyyy") & "'"
        'strAppend = strAppend + sql
      'End If
    
      'If Len(Trim(TxtBordx1)) > 0 And IsDate(TxtBordx1) = True Then
        'strAppend = strAppend + " And MBMBordxDate >= '" & Format(TxtBordx1, "mm/dd/yyyy") & "'"
      'End If
    
      'If Len(Trim(TxtBordx2)) > 0 And IsDate(TxtBordx2) = True Then
        'strAppend = strAppend + " And MBMBordxDate <= '" & Format(TxtBordx2, "mm/dd/yyyy") & "'"
      'End If
      
      'If chkBordxDate.value = 1 Then
        'sql = ""
        'sql = "AND MBMBordxDate Is null " & ""
        'strAppend = strAppend + sql
      'End If
      
      'If Len(Trim(TxtBatch1)) Then
        'strAppend = strAppend + " AND MBMBatchNo >= '" & Trim(TxtBatch1) & "'"
      'End If
    
      'If Len(Trim(TxtBatch2)) Then
        'strAppend = strAppend + " AND MBMBatchNo <= '" & Trim(TxtBatch2) & "'"
      'End If
      
      'If chkBatch.value = 1 Then
        'sql = ""
        'sql = "AND MBMBatchNo Is null " & ""
        'strAppend = strAppend + sql
      'End If
      
      'If (TxtEndtBordx1) <> "__/__/____" And IsDate(TxtEndtBordx1) = True _
        And (TxtEndtBordx2) <> "__/__/____" And IsDate(TxtEndtBordx2) = True Then
        'sql = ""
        'sql = " AND MBMAEndtBordxDate >= '" & Format(Trim(TxtEndtBordx1), "mm/dd/yyyy") & _
              "' And MBMAEndtBordxDate <= '" & Format(Trim(TxtEndtBordx2), "mm/dd/yyyy") & "'"
        'strAppend = strAppend + sql
      'End If
    
      'If Len(Trim(TxtEndtBordx1)) > 0 And IsDate(TxtEndtBordx1) = True Then
        'strAppend = strAppend + " And MBMAEndtBordxDate >= '" & Format(TxtEndtBordx1, "mm/dd/yyyy") & "'"
      'End If
    
      'If Len(Trim(TxtEndtBordx2)) > 0 And IsDate(TxtEndtBordx2) = True Then
        'strAppend = strAppend + " And MBMAEndtBordxDate <= '" & Format(TxtEndtBordx2, "mm/dd/yyyy") & "'"
      'End If
      
      'If chkEndtBordx.value = 1 Then
        'sql = ""
        'sql = "AND MBMAEndtBordxDate Is null " & ""
        'strAppend = strAppend + sql
      'End If
      
      'If Len(Trim(TxtEndtBatch1)) Then
        'strAppend = strAppend + " AND MBMAEndtBatchNo >= '" & Trim(TxtBatch1) & "'"
      'End If
    
      'If Len(Trim(TxtEndtBatch2)) Then
        'strAppend = strAppend + " AND MBMAEndtBatchNo <= '" & Trim(TxtBatch2) & "'"
      'End If
      
    'If chkEndtBatch.value = 1 Then
        'sql = ""
        'sql = "AND MBMAEndtBatchNo Is null " & ""
        'strAppend = strAppend + sql
    'End If
        
    If ChkJoin.value = 1 Then
        sql = ""
        sql = "AND MBMFirstJoinedDate Is null " & ""
        strAppend = strAppend + sql
    End If
    
    If (txtJoinDate1) <> "__/__/____" And IsDate(txtJoinDate1) = True _
        And (txtJoinDate2) <> "__/__/____" And IsDate(txtJoinDate2) = True Then
        sql = ""
        sql = " AND MBMFirstJoinedDate >= '" & Format(Trim(txtJoinDate1), "mm/dd/yyyy") & _
              "' And MBMFirstJoinedDate <= '" & Format(Trim(txtJoinDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtJoinDate1)) > 0 And IsDate(txtJoinDate1) = True Then
        strAppend = strAppend + " And MBMFirstJoinedDate >= '" & Format(txtJoinDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtJoinDate2)) > 0 And IsDate(txtJoinDate2) = True Then
        strAppend = strAppend + " And MBMFirstJoinedDate <= '" & Format(txtJoinDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtFinal)) Then
        strAppend = strAppend + " AND FinalDiag = '" & Trim(TxtFinal) & "'"
    End If
    
    If Len(Trim(CboGrpCom)) Then
        strAppend = strAppend + " AND MBMGroupCompany like '" & RemStr(Mid(CboGrpCom, 1, IIf(InStr(1, CboGrpCom, "-") = 0, 4, InStr(1, CboGrpCom, "-") - 1))) & "%'"
    End If
        
    Call PolEffDateListing(strAppend)
   
End Function

' =================== Search Post Code Listing (Selected Records) ======================
Private Sub PCodeListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String

    
    lstPCodeList.listitems.Clear
            
    sql = " SELECT Distinct MBMPostCode From VIEW_AnalysisReport6 Where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    

            Do Until rst2.EOF
            
            With rst2
                If Len(rst2!MBMPostCode) Then
                    Set Record = lstPCodeList.listitems.Add(, , rst2!MBMPostCode)
                    PostCode = Trim(rst2!MBMPostCode)
                
                    Call SearchPCodeRecords(strAppend)
                                    
                    If Len(Total) Then
                        Record.SubItems(1) = Total
                    End If
                    
                    If Len(TotalActual) Then
                       Record.SubItems(2) = Format(TotalActual, "#,##0.00")
                    End If
                    
                    If Len(TotalApproved) Then
                        Record.SubItems(3) = Format(TotalApproved, "#,##0.00")
                    End If
                End If
            End With
            rst2.MoveNext
        Loop

    rst2.Close
    Set rst2 = Nothing
End Sub
' =================== Search Post Code Listing (Selected Records) ======================


' =================== Search Post Code Top 10 / 20 Results ======================
Private Sub PCodeTop1020(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
       
    lstPCodeList.listitems.Clear
        
    If Mid(Combo3, 1, 1) = "0" Then
    
        sql = " SELECT TOP 10 MBMPostCode, COUNT(MBMPostCode) AS Total" & _
              " From VIEW_AnalysisReport6 Where 0 = 0 "
          
    ElseIf Mid(Combo3, 1, 1) = "1" Then
    
        sql = " SELECT TOP 20 MBMPostCode, COUNT(MBMPostCode) AS Total" & _
              " From VIEW_AnalysisReport6 Where 0 = 0 "
    
    End If
    
    sql2 = " GROUP BY MBMPostCode ORDER BY Total DESC "
    
    sql = sql + strAppend + sql2
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                    
            With rst2
                Set Record = lstPCodeList.listitems.Add(, , rst2!MBMPostCode)
                PostCode = Trim(rst2!MBMPostCode)
                
                Call SearchPCodeRecords(strAppend)
                                
                If Len(rst2!Total) Then
                    Record.SubItems(1) = rst2!Total
                End If
                
                If Len(TotalActual) Then
                   Record.SubItems(2) = Format(TotalActual, "#,##0.00")
                End If
                If Len(TotalApproved) Then
                    Record.SubItems(3) = Format(TotalApproved, "#,##0.00")
                End If
                
                End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
' =================== Search Post Code Top 10 / 20 Results ======================

' =================== Count Post Code Records ======================
Private Sub SearchPCodeRecords(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
       
    Total = 0
    TotalActual = 0
    TotalApproved = 0
            
    ListView11.listitems.Clear
            
    sql = " Select * From VIEW_AnalysisReport6 where MBMPostCode = '" & Trim(PostCode) & "'"
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                    
            With rst2
                Set Record = ListView11.listitems.Add(, , rst2!MBMPostCode)
                              
                If Len(rst2!CLMTotalActual) Then
                   Record.SubItems(1) = rst2!CLMTotalActual
                   TotalActual = TotalActual + rst2!CLMTotalActual
                End If
                
                If Len(rst2!CLMTotalApproved) Then
                   Record.SubItems(2) = rst2!CLMTotalApproved
                   TotalApproved = TotalApproved + rst2!CLMTotalApproved
                End If
                
                Total = Total + 1
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
' =================== Count Post Code Records ======================

Private Sub Option3_Click()
    Combo2.Enabled = True
    lstHospitalList.Visible = True
    lstPrincipalsList.Visible = False
    lstPolicyList.Visible = False
    lstPCodeList.Visible = False
    lstRepList.Visible = False
    Frame3.Visible = True
    Frame4.Visible = False
    Frame6.Visible = False
    Frame9.Visible = False
    Frame14.Visible = False
End Sub

Private Sub Option4_Click()
    Combo2.Enabled = False
    lstHospitalList.Visible = False
    lstPrincipalsList.Visible = True
    lstPolicyList.Visible = False
    lstPCodeList.Visible = False
    lstRepList.Visible = False
    Frame3.Visible = False
    Frame4.Visible = True
    Frame6.Visible = False
    Frame9.Visible = False
    Frame14.Visible = False
End Sub

Private Sub Option5_Click()
    Combo2.Enabled = True
    lstHospitalList.Visible = False
    lstPrincipalsList.Visible = False
    lstPolicyList.Visible = True
    lstPCodeList.Visible = False
    lstRepList.Visible = False
    Frame6.Visible = True
    Frame3.Visible = False
    Frame4.Visible = False
    Frame9.Visible = False
    Frame14.Visible = False
End Sub

Private Sub Option6_Click()
    Combo2.Enabled = True
    lstHospitalList.Visible = False
    lstPrincipalsList.Visible = False
    lstPolicyList.Visible = False
    lstPCodeList.Visible = True
    lstRepList.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    Frame6.Visible = False
    Frame9.Visible = True
    Frame14.Visible = False
End Sub

Private Sub Option7_Click()
    Combo2.Enabled = True
    lstHospitalList.Visible = False
    lstPrincipalsList.Visible = False
    lstPolicyList.Visible = False
    lstPCodeList.Visible = False
    lstRepList.Visible = True
    Frame3.Visible = False
    Frame4.Visible = False
    Frame6.Visible = False
    Frame14.Visible = True
    Frame9.Visible = False
End Sub


'================= Check Date Format =============================

Private Sub TxtDOA1_LostFocus()
    If IsDate(TxtDOA1) Then
    Else
        If TxtDOA1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDOA1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOA2_LostFocus()
    If IsDate(TxtDOA2) Then
    Else
        If TxtDOA2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDOA2.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOD1_LostFocus()
    If IsDate(txtDOD1) Then
    Else
        If txtDOD1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOD1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOD2_LostFocus()
    If IsDate(txtDOD2) Then
    Else
        If txtDOD2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOD2.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOR1_LostFocus()
    If IsDate(TxtDOR1) Then
    Else
        If TxtDOR1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDOR1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOR2_LostFocus()
    If IsDate(TxtDOR2) Then
    Else
        If TxtDOR2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtDOR2.SetFocus
        End If
    End If
End Sub

Private Sub TxtExpDate1_LostFocus()
    If IsDate(TxtExpDate1) Then
    Else
        If TxtExpDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtExpDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtExpDate2_LostFocus()
    If IsDate(TxtExpDate2) Then
    Else
        If TxtExpDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtExpDate2.SetFocus
        End If
    End If
End Sub

Private Sub TxtJoinDate1_LostFocus()
    If IsDate(txtJoinDate1) Then
    Else
        If txtJoinDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtJoinDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtJoinDate2_LostFocus()
    If IsDate(txtJoinDate2) Then
    Else
        If txtJoinDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtJoinDate2.SetFocus
        End If
    End If
End Sub

Private Sub TxtPayEffDate1_LostFocus()
    If IsDate(TxtPayEffDate1) Then
    Else
        If TxtPayEffDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPayEffDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtPayEffDate2_LostFocus()
    If IsDate(TxtPayEffDate2) Then
    Else
        If TxtPayEffDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPayEffDate2.SetFocus
        End If
    End If
End Sub
'================= Check Date Format =============================

' ================ Command Button Click / Change =================

Private Sub CmdPrint_Click()
    Screen.MousePointer = vbHourglass
        'Select Case SSTab1.Tab
            'Case 0
            If Option3.value = True Then
                Call ExportToExcelAnalysis1(lstHospitalList)
            'Case 1
            ElseIf Option4.value = True Then
                Call ExportToExcelAnalysis1(lstPrincipalsList)
            'Case 2
            ElseIf Option5.value = True Then
                Call ExportToExcelAnalysis1(lstPolicyList)
            'Case 3
            ElseIf Option6.value = True Then
                Call ExportToExcelAnalysis1(lstPCodeList)
            'Case 4
            ElseIf Option7.value = True Then
                Call ExportToExcelAnalysis1(lstRepList)
            End If
            'End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdExit_Click()
Dim RecordExit
    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        bExit = True
        Unload Me
    End If
End Sub

Private Sub cmdSearch_Click()
   
    Screen.MousePointer = vbHourglass
        'Select Case SSTab1.Tab
            'Case 0
            If Option3.value = True Then
                Frame1.Visible = False
                Frame3.Visible = False
                Frame17.Visible = False
                lstHospitalList.Visible = True
                lstHospitalList.Height = 6720
                lstHospitalList.Left = 0
                lstHospitalList.Top = 480
                lstHospitalList.Width = 11775
                Call SearchHOS
            'Case 1
            ElseIf Option4.value = True Then
                Frame1.Visible = False
                Frame4.Visible = False
                Frame17.Visible = False
                lstPrincipalsList.Visible = True
                lstPrincipalsList.Height = 6720
                lstPrincipalsList.Left = 0
                lstPrincipalsList.Top = 480
                lstPrincipalsList.Width = 11775
                Call SearchInsure
            'Case 2
            ElseIf Option5.value = True Then
                Frame1.Visible = False
                Frame6.Visible = False
                Frame17.Visible = False
                lstPolicyList.Visible = True
                lstPolicyList.Height = 6720
                lstPolicyList.Left = 0
                lstPolicyList.Top = 480
                lstPolicyList.Width = 11775
                Call SearchDateRecord
            'Case 3
            ElseIf Option6.value = True Then
                Frame1.Visible = False
                Frame9.Visible = False
                Frame17.Visible = False
                lstPCodeList.Visible = True
                lstPCodeList.Height = 6720
                lstPCodeList.Left = 0
                lstPCodeList.Top = 480
                lstPCodeList.Width = 11775
                Call SearchPCode
            'Case 4
            ElseIf Option7.value = True Then
                Frame1.Visible = False
                Frame14.Visible = False
                Frame17.Visible = False
                lstRepList.Visible = True
                lstRepList.Height = 6720
                lstRepList.Left = 0
                lstRepList.Top = 480
                lstRepList.Width = 11775
                Call SearchRepudiation
            End If
        'End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lstHospitalList.listitems.Clear
    lstPrincipalsList.listitems.Clear
    lstPolicyList.listitems.Clear
    lstPCodeList.listitems.Clear
    lstRepList.listitems.Clear
End Sub

Private Sub CmdFinal_Click()
    frmSearchICD.Source = 24
    frmSearchICD.show vbModal
End Sub

Private Sub Command2_Click()
    frmSearchICD.Source = 26
    frmSearchICD.show vbModal
End Sub
' ================ Command Button Click / Change =================


'================ Repudiated Resons Listing ===========================
Private Sub RepudiaListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
       
    lstRepList.listitems.Clear
            
    sql = " SELECT * From VIEW_RepudiatedMBM Where 0 = 0 "
    
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
    MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    
    Else
        
        Do Until rst2.EOF
                    
            With rst2
                Set Record = lstRepList.listitems.Add(, , rst2!MBMNumber)
                                                               
                If Len(rst2!CASNumber) Then
                    Record.SubItems(1) = Trim(rst2!CASNumber)
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(2) = rst2!PatientName
                End If
                               
                If Len(rst2!repdescription) Then
                    Record.SubItems(3) = Trim(rst2!repdescription)
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                   Record.SubItems(4) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                
            End With
            rst2.MoveNext
        Loop
    End If
    rst2.Close
    Set rst2 = Nothing
End Sub
'================ Repudiated Resons Listing ===========================

Public Function ClearForm(frmCall)
On Error Resume Next
Dim ctl As Control

For Each ctl In frmCall.Controls
    If TypeOf ctl Is TextBox Then
        ctl.Text = ""
    ElseIf TypeOf ctl Is ListBox Then
            ctl.Clear
    ElseIf TypeOf ctl Is ComboBox Then
            ctl.Text = ""
    ElseIf TypeOf ctl Is CheckBox Then
            ctl.value = False
    ElseIf TypeOf ctl Is MaskEdBox Then
        ctl.Text = "__/__/____"
    
    End If
Next ctl
End Function

'============== Combo Box Click/KeyPress ======================
Private Sub CboAdultChild_Change()
    If InStr(CboAdultChild.Text, "null") Then
       CboAdultChild.Enabled = False
    End If
End Sub

Private Sub CboAdultChild_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboAdultChild.Text, CboAdultChild.SelStart) + Chr(KeyAscii) + Mid(CboAdultChild.Text, CboAdultChild.SelStart + CboAdultChild.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboAdultChild.ListCount - 1
 
        If NewText = LCase(Left(CboAdultChild.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboAdultChild.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboAdultChild.Text = ValidValue
                CboAdultChild.SelStart = 0
                CboAdultChild.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboCaseStatus_Change()
    If InStr(CboCaseStatus.Text, "null") Then
       CboCaseStatus.Enabled = False
    End If
End Sub

Private Sub CboCaseStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCaseStatus.Text, CboCaseStatus.SelStart) + Chr(KeyAscii) + Mid(CboCaseStatus.Text, CboCaseStatus.SelStart + CboCaseStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCaseStatus.ListCount - 1
 
        If NewText = LCase(Left(CboCaseStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCaseStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCaseStatus.Text = ValidValue
                CboCaseStatus.SelStart = 0
                CboCaseStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboClaimability_Change()
    If InStr(CboClaimability.Text, "null") Then
       CboClaimability.Enabled = False
    End If
End Sub

Private Sub CboClaimability_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClaimability.Text, CboClaimability.SelStart) + Chr(KeyAscii) + Mid(CboClaimability.Text, CboClaimability.SelStart + CboClaimability.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboClaimability.ListCount - 1
 
        If NewText = LCase(Left(CboClaimability.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClaimability.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboClaimability.Text = ValidValue
                CboClaimability.SelStart = 0
                CboClaimability.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboDataStatus_Change()
    If InStr(cboDataStatus.Text, "null") Then
       cboDataStatus.Enabled = False
    End If
End Sub

Private Sub cboDataStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboDataStatus.Text, cboDataStatus.SelStart) + Chr(KeyAscii) + Mid(cboDataStatus.Text, cboDataStatus.SelStart + cboDataStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboDataStatus.ListCount - 1
 
        If NewText = LCase(Left(cboDataStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboDataStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboDataStatus.Text = ValidValue
                cboDataStatus.SelStart = 0
                cboDataStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboHLTCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHLTCode.Text, cboHLTCode.SelStart) + Chr(KeyAscii) + Mid(cboHLTCode.Text, cboHLTCode.SelStart + cboHLTCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboHLTCode.ListCount - 1
 
        If NewText = LCase(Left(cboHLTCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHLTCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboHLTCode.Text = ValidValue
                cboHLTCode.SelStart = 0
                cboHLTCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboHLTCode_Change()
    If InStr(cboHLTCode.Text, "null") Then
       cboHLTCode.Enabled = False
    End If
    
    If Len(Trim(cboHLTCode)) Then
        LoadPayor (cboHLTCode)
    End If
End Sub

Private Sub cboHLTCode_Click()
    If Len(Trim(cboHLTCode)) Then
        LoadPayor (cboHLTCode)
    End If
End Sub

Private Sub cboHOS_Change()
    If InStr(cboHOS.Text, "null") Then
       cboHOS.Enabled = False
    End If
End Sub

Private Sub cboHOS_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHOS.Text, cboHOS.SelStart) + Chr(KeyAscii) + Mid(cboHOS.Text, cboHOS.SelStart + cboHOS.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboHOS.ListCount - 1
 
        If NewText = LCase(Left(cboHOS.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHOS.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboHOS.Text = ValidValue
                cboHOS.SelStart = 0
                cboHOS.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboHOS2_Change()
    If InStr(CboHOS2.Text, "null") Then
       CboHOS2.Enabled = False
    End If
End Sub

Private Sub CboHOS2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboHOS2.Text, CboHOS2.SelStart) + Chr(KeyAscii) + Mid(CboHOS2.Text, CboHOS2.SelStart + CboHOS2.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboHOS2.ListCount - 1
 
        If NewText = LCase(Left(CboHOS2.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboHOS2.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboHOS2.Text = ValidValue
                CboHOS2.SelStart = 0
                CboHOS2.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboHOSGrp_Change()
    If InStr(cboHOSGrp.Text, "null") Then
       cboHOSGrp.Enabled = False
    End If
End Sub

Private Sub cboHOSGrp_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHOSGrp.Text, cboHOSGrp.SelStart) + Chr(KeyAscii) + Mid(cboHOSGrp.Text, cboHOSGrp.SelStart + cboHOSGrp.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboHOSGrp.ListCount - 1
 
        If NewText = LCase(Left(cboHOSGrp.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHOSGrp.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboHOSGrp.Text = ValidValue
                cboHOSGrp.SelStart = 0
                cboHOSGrp.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboOCP_Change()
    If InStr(CboOCP.Text, "null") Then
       CboOCP.Enabled = False
    End If
End Sub

Private Sub CboOCP_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboOCP.Text, CboOCP.SelStart) + Chr(KeyAscii) + Mid(CboOCP.Text, CboOCP.SelStart + CboOCP.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboOCP.ListCount - 1
 
        If NewText = LCase(Left(CboOCP.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboOCP.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboOCP.Text = ValidValue
                CboOCP.SelStart = 0
                CboOCP.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboPayorCode_Change()
    If InStr(cboPayorCode.Text, "null") Then
       cboPayorCode.Enabled = False
    End If
End Sub

Private Sub cboPayorCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPayorCode.Text, cboPayorCode.SelStart) + Chr(KeyAscii) + Mid(cboPayorCode.Text, cboPayorCode.SelStart + cboPayorCode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPayorCode.ListCount - 1
 
        If NewText = LCase(Left(cboPayorCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPayorCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPayorCode.Text = ValidValue
                cboPayorCode.SelStart = 0
                cboPayorCode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboRace_Change()
    If InStr(CboRace.Text, "null") Then
       CboRace.Enabled = False
    End If
End Sub

Private Sub CboRace_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboRace.Text, CboRace.SelStart) + Chr(KeyAscii) + Mid(CboRace.Text, CboRace.SelStart + CboRace.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboRace.ListCount - 1
 
        If NewText = LCase(Left(CboRace.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboRace.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboRace.Text = ValidValue
                CboRace.SelStart = 0
                CboRace.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboSex_Change()
    If InStr(CboSex.Text, "null") Then
       CboSex.Enabled = False
    End If
End Sub

Private Sub CboSex_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboSex.Text, CboSex.SelStart) + Chr(KeyAscii) + Mid(CboSex.Text, CboSex.SelStart + CboSex.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboSex.ListCount - 1
 
        If NewText = LCase(Left(CboSex.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboSex.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboSex.Text = ValidValue
                CboSex.SelStart = 0
                CboSex.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboStatus_Change()
    If InStr(CboStatus.Text, "null") Then
       CboStatus.Enabled = False
    End If
End Sub

Private Sub CboStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboStatus.Text, CboStatus.SelStart) + Chr(KeyAscii) + Mid(CboStatus.Text, CboStatus.SelStart + CboStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboStatus.ListCount - 1
 
        If NewText = LCase(Left(CboStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboStatus.Text = ValidValue
                CboStatus.SelStart = 0
                CboStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub Combo3_Change()
    If InStr(Combo3.Text, "null") Then
       Combo3.Enabled = False
    End If
    
    If Mid(Combo3, 1, 1) = "0" Or Mid(Combo3, 1, 1) = "1" Then
        TxtPCode1.Text = ""
        TxtPCode1.Enabled = False
    Else
        TxtPCode1.Text = ""
        TxtPCode1.Enabled = True
    End If
End Sub

Private Sub Combo3_Click()
    If InStr(Combo3.Text, "null") Then
       Combo3.Enabled = False
    End If
    
    If Mid(Combo3, 1, 1) = "0" Or Mid(Combo3, 1, 1) = "1" Then
        TxtPCode1.Text = ""
        TxtPCode1.Enabled = False
    Else
        TxtPCode1.Text = ""
        TxtPCode1.Enabled = True
    End If
End Sub

Private Sub Combo1_Change()

    If InStr(Combo1.Text, "null") Then
       Combo1.Enabled = False
    End If
    
    If Mid(Combo1, 1, 1) = "0" Or Mid(Combo1, 1, 1) = "1" Then
        cboHOS.Text = ""
        cboHOS.Enabled = False
        cboHOSGrp.Enabled = False
    Else
        cboHOS.Text = ""
        cboHOS.Enabled = True
        cboHOSGrp.Enabled = True
    End If

End Sub

Private Sub Combo1_Click()
    
    If InStr(Combo1.Text, "null") Then
       Combo1.Enabled = False
    End If
    
    If Mid(Combo1, 1, 1) = "0" Or Mid(Combo1, 1, 1) = "1" Then
        cboHOS.Text = ""
        cboHOS.Enabled = False
        cboHOSGrp.Enabled = False
    Else
        cboHOS.Text = ""
        cboHOS.Enabled = True
        cboHOSGrp.Enabled = True
    End If
    
End Sub

Private Sub Combo2_Change()
    If InStr(Combo2.Text, "null") Then
       Combo2.Enabled = False
    End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(Combo2.Text, Combo2.SelStart) + Chr(KeyAscii) + Mid(Combo2.Text, Combo2.SelStart + Combo2.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To Combo2.ListCount - 1
 
        If NewText = LCase(Left(Combo2.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = Combo2.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                Combo2.Text = ValidValue
                Combo2.SelStart = 0
                Combo2.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboGrpCom_Change()
    If InStr(CboGrpCom.Text, "null") Then
       CboGrpCom.Enabled = False
    End If
End Sub

Private Sub CboGrpCom_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboGrpCom.Text, CboGrpCom.SelStart) + Chr(KeyAscii) + Mid(CboGrpCom.Text, CboGrpCom.SelStart + CboGrpCom.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboGrpCom.ListCount - 1
 
        If NewText = LCase(Left(CboGrpCom.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboGrpCom.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboGrpCom.Text = ValidValue
                CboGrpCom.SelStart = 0
                CboGrpCom.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboPayorCode_Click()
Dim strsql As String
Dim PAYGRP As New ADODB.Recordset

    CboGrpCom.Clear

    strsql = "select MBMGroupCompany from View_PAYGrpCom where PAYCode= '" & Mid(cboPayorCode, 1, 2) & "'"
    PAYGRP.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If PAYGRP.recordCount Then
        Do Until PAYGRP.EOF
           CboGrpCom.AddItem Trim(PAYGRP!MBMGroupCompany)
           PAYGRP.MoveNext
        Loop
    End If
    PAYGRP.Close
    Set PAYGRP = Nothing
End Sub
'============== Combo Box Click/KeyPress ======================


