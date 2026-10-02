VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewMBMAdjustment 
   Caption         =   "Member Adjustment"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   11880
   Begin TabDlg.SSTab SSTab1 
      Height          =   4965
      Left            =   60
      TabIndex        =   96
      Top             =   2520
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8758
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Principal &Details"
      TabPicture(0)   =   "frmNewMBMAdjustment.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Supplementary"
      TabPicture(1)   =   "frmNewMBMAdjustment.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstCovAdd"
      Tab(1).Control(1)=   "Frame6"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Adjustment History"
      TabPicture(2)   =   "frmNewMBMAdjustment.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstAdjustmentHis"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Fees"
      TabPicture(3)   =   "frmNewMBMAdjustment.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstFeesView"
      Tab(3).Control(1)=   "Frame2"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "MBM History"
      TabPicture(4)   =   "frmNewMBMAdjustment.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lstMBMHis"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Mgmt Note"
      TabPicture(5)   =   "frmNewMBMAdjustment.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtMgmtNotes"
      Tab(5).ControlCount=   1
      Begin VB.TextBox txtMgmtNotes 
         Height          =   4425
         Left            =   -74880
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   236
         Top             =   400
         Width           =   11460
      End
      Begin VB.Frame Frame6 
         Height          =   3195
         Left            =   -74880
         TabIndex        =   133
         Top             =   360
         Width           =   11580
         Begin VB.ComboBox Cmb_Cstatus 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmNewMBMAdjustment.frx":00A8
            Left            =   10080
            List            =   "frmNewMBMAdjustment.frx":00B5
            Sorted          =   -1  'True
            TabIndex        =   247
            Top             =   2760
            Width           =   615
         End
         Begin VB.ComboBox Cmd_HRInd 
            Height          =   315
            ItemData        =   "frmNewMBMAdjustment.frx":00C2
            Left            =   8640
            List            =   "frmNewMBMAdjustment.frx":00CC
            Sorted          =   -1  'True
            TabIndex        =   246
            Top             =   2760
            Width           =   615
         End
         Begin VB.ComboBox cmbAgeLmt 
            Height          =   315
            ItemData        =   "frmNewMBMAdjustment.frx":00D6
            Left            =   6840
            List            =   "frmNewMBMAdjustment.frx":00E0
            Sorted          =   -1  'True
            TabIndex        =   241
            Text            =   "Y-Yes"
            Top             =   2760
            Width           =   1095
         End
         Begin VB.TextBox txtCovAllergic 
            Height          =   855
            Left            =   4680
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   81
            Top             =   1800
            Width           =   2415
         End
         Begin VB.TextBox txtCovExclusion 
            Height          =   855
            Left            =   7200
            MaxLength       =   2500
            MultiLine       =   -1  'True
            TabIndex        =   82
            Top             =   1800
            Width           =   2775
         End
         Begin VB.ComboBox cboCovSalutation 
            Height          =   315
            Left            =   1320
            TabIndex        =   54
            Text            =   "MS"
            Top             =   195
            Width           =   975
         End
         Begin VB.TextBox txtCovName 
            Height          =   330
            Left            =   3480
            MaxLength       =   60
            TabIndex        =   55
            Top             =   180
            Width           =   3260
         End
         Begin VB.Frame Frame9 
            Height          =   2055
            Left            =   10540
            TabIndex        =   134
            Top             =   120
            Width           =   980
            Begin VB.CommandButton CmdCancelSupp 
               Caption         =   "Cancel"
               Enabled         =   0   'False
               Height          =   315
               Left            =   80
               TabIndex        =   139
               Top             =   1640
               Width           =   840
            End
            Begin VB.CommandButton cmdRemove 
               Caption         =   "&Remove"
               Enabled         =   0   'False
               Height          =   315
               Left            =   80
               TabIndex        =   136
               Top             =   560
               Width           =   840
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "&Add"
               Height          =   315
               Left            =   80
               TabIndex        =   135
               Top             =   200
               Width           =   840
            End
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "Update"
               Enabled         =   0   'False
               Height          =   315
               Left            =   80
               TabIndex        =   137
               Top             =   920
               Width           =   840
            End
            Begin VB.CommandButton cmdClear 
               Caption         =   "Clear"
               Height          =   315
               Left            =   80
               TabIndex        =   138
               Top             =   1280
               Width           =   840
            End
         End
         Begin VB.TextBox txtCovIcBcPp 
            Height          =   285
            Left            =   7560
            MaxLength       =   30
            TabIndex        =   56
            Top             =   200
            Width           =   2895
         End
         Begin VB.ComboBox cboCovRelationship 
            Height          =   315
            Left            =   1320
            TabIndex        =   57
            Text            =   "W -Wife"
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox cboCovAdultChild 
            Height          =   315
            Left            =   1320
            TabIndex        =   62
            Text            =   "A"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtAvailableLimit 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   65
            Top             =   1010
            Width           =   975
         End
         Begin VB.TextBox txtRBAmt 
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   67
            Top             =   1290
            Width           =   975
         End
         Begin VB.TextBox txtSuppPayPremGross 
            Height          =   285
            Left            =   1320
            TabIndex        =   69
            Top             =   1545
            Width           =   975
         End
         Begin VB.TextBox txtSuppGenWP 
            Height          =   285
            Left            =   11280
            MaxLength       =   9
            TabIndex        =   89
            Top             =   2280
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.TextBox txtSuppPREWP 
            Height          =   285
            Left            =   10875
            MaxLength       =   9
            TabIndex        =   91
            Top             =   2280
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox TxtSuppSpeWP 
            Height          =   300
            Left            =   10875
            MaxLength       =   9
            TabIndex        =   93
            Top             =   2280
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtSuppPayMCOGross 
            Height          =   285
            Left            =   1320
            TabIndex        =   71
            Top             =   1805
            Width           =   975
         End
         Begin MSMask.MaskEdBox txtCovBirthdate 
            Height          =   285
            Left            =   3480
            TabIndex        =   58
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtSuppClmNonPayFr 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   9360
            TabIndex        =   87
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtSuppClmNonPayTo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   9360
            TabIndex        =   88
            Top             =   1440
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSuppGENWPEnd 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   255
            Left            =   10920
            TabIndex        =   90
            Top             =   2640
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSuppPREWPEnd 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Left            =   10920
            TabIndex        =   92
            Top             =   2760
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSuppSPEWPEnd 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   270
            Left            =   11040
            TabIndex        =   94
            Top             =   2760
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboCovSex 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3480
            TabIndex        =   63
            Text            =   "M"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtCovAnnualLimit 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   66
            Top             =   1005
            Width           =   1095
         End
         Begin VB.ComboBox cboSuppStaffDisc 
            Height          =   315
            Left            =   3480
            Sorted          =   -1  'True
            TabIndex        =   68
            Top             =   1260
            Width           =   1095
         End
         Begin VB.TextBox txtSuppPayPremNet 
            Height          =   285
            Left            =   3480
            TabIndex        =   70
            Top             =   1550
            Width           =   1095
         End
         Begin VB.TextBox txtSuppPayMCONet 
            Height          =   285
            Left            =   3480
            TabIndex        =   72
            Top             =   1805
            Width           =   1095
         End
         Begin VB.TextBox txtPolicyDisc 
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   73
            Top             =   2055
            Width           =   975
         End
         Begin VB.TextBox txtPolicyDiscAmt 
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   75
            Top             =   2295
            Width           =   975
         End
         Begin VB.TextBox txtLoadingPrem 
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   77
            Top             =   2550
            Width           =   975
         End
         Begin VB.TextBox txtRenewalDisc 
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   74
            Top             =   2055
            Width           =   1095
         End
         Begin VB.TextBox txtRenewalDiscAmt 
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   76
            Top             =   2280
            Width           =   1095
         End
         Begin VB.TextBox txtSuppPLNCode 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   5640
            MaxLength       =   4
            TabIndex        =   59
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtCovUndExcess 
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   78
            Top             =   2535
            Width           =   1095
         End
         Begin VB.TextBox txtPrevINSCom 
            Height          =   330
            Left            =   3480
            MaxLength       =   60
            TabIndex        =   80
            Top             =   2760
            Width           =   1095
         End
         Begin VB.TextBox txtPrevPLNCode 
            Height          =   285
            Left            =   1320
            TabIndex        =   79
            Top             =   2800
            Width           =   975
         End
         Begin MSMask.MaskEdBox txtSuppEffDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   7560
            TabIndex        =   60
            Top             =   440
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtSuppExpDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   9360
            TabIndex        =   61
            Top             =   440
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtCOVPayRemarks 
            Height          =   285
            Left            =   5640
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   64
            Top             =   720
            Width           =   4815
         End
         Begin VB.TextBox txtPayMBMNo 
            Height          =   285
            Left            =   5640
            MaxLength       =   8
            TabIndex        =   83
            Top             =   980
            Width           =   1095
         End
         Begin VB.TextBox txtClientNo 
            Height          =   285
            Left            =   5640
            MaxLength       =   8
            TabIndex        =   85
            Top             =   1230
            Width           =   1095
         End
         Begin VB.TextBox txtPaySeqID 
            Height          =   285
            Left            =   7560
            MaxLength       =   2
            TabIndex        =   84
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtClientID 
            Height          =   285
            Left            =   7560
            MaxLength       =   2
            TabIndex        =   86
            Top             =   1220
            Width           =   1095
         End
         Begin VB.Label lbl_HRInd 
            Caption         =   "HR Ind"
            Height          =   255
            Left            =   8040
            TabIndex        =   245
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Age Limit For Cancellation"
            Height          =   255
            Index           =   0
            Left            =   4680
            TabIndex        =   242
            Top             =   2760
            Width           =   2055
         End
         Begin VB.Label Label62 
            Caption         =   "EffDate"
            Height          =   255
            Index           =   1
            Left            =   6800
            TabIndex        =   218
            Top             =   495
            Width           =   645
         End
         Begin VB.Label Label62 
            Caption         =   "ExpDate"
            Height          =   255
            Index           =   2
            Left            =   8700
            TabIndex        =   217
            Top             =   495
            Width           =   645
         End
         Begin VB.Label Label45 
            Caption         =   "Pay MBM No"
            Height          =   255
            Index           =   4
            Left            =   4680
            TabIndex        =   216
            Top             =   1030
            Width           =   975
         End
         Begin VB.Label Label45 
            Caption         =   "Pay Seq"
            Height          =   255
            Index           =   1
            Left            =   6800
            TabIndex        =   215
            Top             =   1040
            Width           =   735
         End
         Begin VB.Label Label45 
            Caption         =   "Client No"
            Height          =   255
            Index           =   2
            Left            =   4680
            TabIndex        =   214
            Top             =   1290
            Width           =   975
         End
         Begin VB.Label Label45 
            Caption         =   "Client ID"
            Height          =   255
            Index           =   3
            Left            =   6800
            TabIndex        =   213
            Top             =   1305
            Width           =   615
         End
         Begin VB.Label Label68 
            Caption         =   "Und Excess"
            Height          =   255
            Left            =   2400
            TabIndex        =   209
            Top             =   2565
            Width           =   885
         End
         Begin VB.Label Label66 
            Caption         =   "PLN Code"
            Height          =   255
            Left            =   4680
            TabIndex        =   204
            Top             =   520
            Width           =   1005
         End
         Begin VB.Label Label32 
            Caption         =   "Policy Disc."
            Height          =   255
            Left            =   120
            TabIndex        =   170
            Top             =   2040
            Width           =   1365
         End
         Begin VB.Label Label28 
            Caption         =   "Renewal Disc."
            Height          =   255
            Left            =   2355
            TabIndex        =   169
            Top             =   2055
            Width           =   1485
         End
         Begin VB.Label Label23 
            Caption         =   "Pay MCO Net"
            Height          =   255
            Left            =   2355
            TabIndex        =   168
            Top             =   1815
            Width           =   1365
         End
         Begin VB.Label Label13 
            Caption         =   "Pay MCO Gross"
            Height          =   255
            Left            =   120
            TabIndex        =   167
            Top             =   1815
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Pay Prem Net"
            Height          =   255
            Index           =   0
            Left            =   2355
            TabIndex        =   166
            Top             =   1560
            Width           =   1365
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            Caption         =   "To"
            Height          =   255
            Left            =   8880
            TabIndex        =   165
            Top             =   1440
            Width           =   285
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            Caption         =   "From"
            Height          =   255
            Left            =   8880
            TabIndex        =   164
            Top             =   1245
            Width           =   405
         End
         Begin VB.Label Label111 
            Caption         =   "Pay Prem Gross"
            Height          =   255
            Left            =   120
            TabIndex        =   163
            Top             =   1560
            Width           =   1365
         End
         Begin VB.Label Label37 
            Caption         =   "Clm Non-Pay Date"
            Height          =   255
            Left            =   8880
            TabIndex        =   162
            Top             =   1005
            Width           =   1365
         End
         Begin VB.Label Label53 
            Caption         =   "EOP"
            Height          =   255
            Left            =   10920
            TabIndex        =   161
            Top             =   2400
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label104 
            Caption         =   "Spe. Ill. W. P."
            Height          =   255
            Left            =   8760
            TabIndex        =   160
            Top             =   1920
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label lbl_CStatus 
            Caption         =   "Status"
            Height          =   255
            Left            =   9480
            TabIndex        =   159
            Top             =   2760
            Width           =   450
         End
         Begin VB.Label Label78 
            Caption         =   "Gen W. P."
            Height          =   255
            Left            =   10200
            TabIndex        =   158
            Top             =   2040
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Label33 
            Caption         =   "EOP"
            Height          =   255
            Left            =   10920
            TabIndex        =   157
            Top             =   2520
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label45 
            Caption         =   "EOP"
            Height          =   255
            Index           =   0
            Left            =   11040
            TabIndex        =   156
            Top             =   2520
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label36 
            Caption         =   "Disc Req"
            Height          =   255
            Left            =   2355
            TabIndex        =   155
            Top             =   1280
            Width           =   1065
         End
         Begin VB.Label Label92 
            Caption         =   "Pay Remark"
            Height          =   255
            Left            =   4680
            TabIndex        =   154
            Top             =   760
            Width           =   1095
         End
         Begin VB.Label Label91 
            Caption         =   "T/O Ins "
            Height          =   255
            Left            =   2400
            TabIndex        =   153
            Top             =   2820
            Width           =   1020
         End
         Begin VB.Label Label90 
            Caption         =   "R&&B Amt"
            Height          =   255
            Left            =   120
            TabIndex        =   152
            Top             =   1320
            Width           =   1140
         End
         Begin VB.Label Label82 
            Caption         =   "P. Pln Code"
            Height          =   255
            Left            =   120
            TabIndex        =   151
            Top             =   2800
            Width           =   1125
         End
         Begin VB.Label Label38 
            Caption         =   "Name"
            Height          =   255
            Left            =   2355
            TabIndex        =   150
            Top             =   195
            Width           =   1140
         End
         Begin VB.Label Label51 
            Caption         =   "Adult/Child"
            Height          =   255
            Left            =   120
            TabIndex        =   149
            Top             =   795
            Width           =   825
         End
         Begin VB.Label Label50 
            Caption         =   "Sex "
            Height          =   255
            Left            =   2355
            TabIndex        =   148
            Top             =   735
            Width           =   825
         End
         Begin VB.Label Label48 
            Caption         =   "Allergic"
            Height          =   255
            Left            =   4680
            TabIndex        =   147
            Top             =   1560
            Width           =   600
         End
         Begin VB.Label Label39 
            Caption         =   "Salutation"
            Height          =   255
            Left            =   120
            TabIndex        =   146
            Top             =   225
            Width           =   1140
         End
         Begin VB.Label Label41 
            Caption         =   "IC Num"
            Height          =   255
            Left            =   6800
            TabIndex        =   145
            Top             =   165
            Width           =   1125
         End
         Begin VB.Label Label49 
            Caption         =   "DOB"
            Height          =   255
            Left            =   2355
            TabIndex        =   144
            Top             =   495
            Width           =   1020
         End
         Begin VB.Label Label46 
            Caption         =   "Supp Lmt"
            Height          =   255
            Left            =   2355
            TabIndex        =   143
            Top             =   1000
            Width           =   855
         End
         Begin VB.Label Label52 
            Caption         =   "R'Ship"
            Height          =   255
            Left            =   120
            TabIndex        =   142
            Top             =   525
            Width           =   1005
         End
         Begin VB.Label Label47 
            Caption         =   "Supplementary Warranty "
            Height          =   255
            Left            =   7200
            TabIndex        =   141
            Top             =   1560
            Width           =   2100
         End
         Begin VB.Label Label6 
            Caption         =   "Bal. Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   140
            Top             =   1020
            Width           =   855
         End
         Begin VB.Label Label44 
            Caption         =   "Loading Prem"
            Height          =   255
            Left            =   120
            TabIndex        =   202
            Top             =   2550
            Width           =   1365
         End
         Begin VB.Label Label43 
            Caption         =   "Pol. Disc. Amt"
            Height          =   255
            Left            =   120
            TabIndex        =   201
            Top             =   2325
            Width           =   1365
         End
         Begin VB.Label Label7 
            Caption         =   "Ren. Disc. Amt"
            Height          =   255
            Left            =   2355
            TabIndex        =   200
            Top             =   2295
            Width           =   1485
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4530
         Left            =   120
         TabIndex        =   97
         Top             =   360
         Width           =   11565
         Begin VB.CheckBox chkCheck 
            Caption         =   "Check1"
            Height          =   255
            Left            =   11280
            TabIndex        =   227
            Top             =   180
            Width           =   255
         End
         Begin VB.TextBox txtMBMExclusion 
            Height          =   705
            Left            =   240
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   99
            Top             =   3720
            Width           =   4860
         End
         Begin VB.TextBox txtRemarks 
            Height          =   705
            Left            =   5280
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   98
            Top             =   3720
            Width           =   5295
         End
         Begin VB.ComboBox cboINSCode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6600
            Sorted          =   -1  'True
            TabIndex        =   35
            Text            =   "I"
            Top             =   160
            Width           =   1680
         End
         Begin VB.TextBox txtMBMAdd1 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   17
            Top             =   160
            Width           =   3705
         End
         Begin VB.TextBox txtMBMAdd2 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   18
            Top             =   420
            Width           =   3705
         End
         Begin VB.TextBox txtMBMAdd3 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   19
            Top             =   680
            Width           =   3705
         End
         Begin VB.ComboBox cboCTYCode 
            Height          =   315
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   21
            Top             =   930
            Width           =   1450
         End
         Begin VB.TextBox txtMBMPostCode 
            Height          =   285
            Left            =   1320
            MaxLength       =   6
            TabIndex        =   20
            Top             =   930
            Width           =   1095
         End
         Begin VB.ComboBox cboSTACode 
            Height          =   315
            Left            =   1320
            TabIndex        =   22
            Top             =   1180
            Width           =   3735
         End
         Begin VB.ComboBox cboMBMSex 
            Height          =   315
            Left            =   1320
            Sorted          =   -1  'True
            TabIndex        =   23
            Text            =   "M"
            Top             =   1475
            Width           =   1125
         End
         Begin VB.ComboBox cboMBMTakeOver 
            Height          =   315
            Left            =   3600
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1475
            Width           =   1455
         End
         Begin VB.ComboBox cboMBMMaritalStatus 
            Height          =   315
            Left            =   3600
            TabIndex        =   26
            Text            =   "M"
            Top             =   1760
            Width           =   1455
         End
         Begin VB.ComboBox cboNATCode 
            Height          =   315
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   28
            Text            =   "MY - Malaysia"
            Top             =   2040
            Width           =   1455
         End
         Begin VB.ComboBox cboPLNCode 
            Height          =   315
            Left            =   9480
            Sorted          =   -1  'True
            TabIndex        =   36
            Top             =   160
            Width           =   1755
         End
         Begin MSMask.MaskEdBox txtMBMBirthDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   25
            Top             =   1760
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMBMPayorEffDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   6600
            TabIndex        =   37
            Top             =   450
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMBMPayorExpDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   9480
            TabIndex        =   38
            Top             =   450
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboRACCode 
            Height          =   315
            Left            =   1320
            Sorted          =   -1  'True
            TabIndex        =   27
            Text            =   "C - Chinese"
            Top             =   2040
            Width           =   1095
         End
         Begin VB.ComboBox cboBRNCode 
            Height          =   315
            Left            =   1320
            Sorted          =   -1  'True
            TabIndex        =   29
            Top             =   2300
            Width           =   3735
         End
         Begin VB.TextBox txtMBMAgentCode 
            Height          =   315
            Left            =   3480
            MaxLength       =   15
            TabIndex        =   31
            Top             =   2580
            Width           =   1560
         End
         Begin VB.ComboBox cboMBMRenewal 
            Height          =   315
            ItemData        =   "frmNewMBMAdjustment.frx":00F1
            Left            =   1320
            List            =   "frmNewMBMAdjustment.frx":00F3
            Sorted          =   -1  'True
            TabIndex        =   30
            Top             =   2580
            Width           =   1050
         End
         Begin VB.ComboBox cboSRVCode 
            Height          =   315
            Left            =   6720
            Sorted          =   -1  'True
            TabIndex        =   53
            Top             =   3720
            Visible         =   0   'False
            Width           =   4035
         End
         Begin MSMask.MaskEdBox txtFinalExpDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   6600
            TabIndex        =   39
            Top             =   710
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtMBMPolicyYear 
            Enabled         =   0   'False
            Height          =   285
            Left            =   9480
            MaxLength       =   25
            TabIndex        =   44
            Top             =   1240
            Width           =   555
         End
         Begin VB.ComboBox CboGuaRenewal 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmNewMBMAdjustment.frx":00F5
            Left            =   3480
            List            =   "frmNewMBMAdjustment.frx":00F7
            TabIndex        =   33
            Top             =   2870
            Width           =   1560
         End
         Begin VB.ComboBox CboInsMode 
            Height          =   315
            Left            =   1320
            TabIndex        =   32
            Top             =   2870
            Width           =   1065
         End
         Begin VB.TextBox txtCostCenter 
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
            Left            =   9480
            TabIndex        =   46
            Top             =   1515
            Width           =   1995
         End
         Begin VB.TextBox txtSubDivision 
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
            Left            =   3480
            TabIndex        =   34
            Top             =   3120
            Width           =   1545
         End
         Begin MSMask.MaskEdBox txtNxDueDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   9480
            TabIndex        =   238
            Top             =   705
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDOR 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   9480
            TabIndex        =   40
            Top             =   960
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMBMBordxDate 
            Height          =   285
            Left            =   6600
            TabIndex        =   41
            Top             =   960
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMBMDateJoin 
            Height          =   285
            Left            =   6600
            TabIndex        =   43
            Top             =   1240
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtMBMEmployeeNo 
            Height          =   300
            Left            =   6600
            MaxLength       =   15
            TabIndex        =   45
            Top             =   1515
            Width           =   1680
         End
         Begin VB.TextBox txtDeptName 
            Height          =   300
            Left            =   6600
            MaxLength       =   50
            TabIndex        =   47
            Top             =   1785
            Width           =   4875
         End
         Begin VB.TextBox txtMBMGroupCompany 
            Height          =   285
            Left            =   6600
            MaxLength       =   150
            TabIndex        =   48
            Top             =   2060
            Width           =   4875
         End
         Begin VB.TextBox txtDivision 
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
            Left            =   9360
            TabIndex        =   51
            Top             =   2310
            Width           =   2145
         End
         Begin VB.TextBox txtSubsidiary 
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
            Left            =   9360
            TabIndex        =   52
            Top             =   2580
            Width           =   2145
         End
         Begin VB.TextBox txtMBMPosition 
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
            TabIndex        =   49
            Top             =   2310
            Width           =   1680
         End
         Begin VB.TextBox txtGrpCat 
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
            TabIndex        =   50
            Top             =   2580
            Width           =   1680
         End
         Begin MSMask.MaskEdBox txtRtDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   6600
            TabIndex        =   243
            Top             =   2880
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label11 
            Caption         =   "Retire date"
            Height          =   255
            Left            =   5280
            TabIndex        =   244
            Top             =   2925
            Width           =   1005
         End
         Begin VB.Label Label16 
            Caption         =   "Next Due Date"
            Height          =   255
            Index           =   2
            Left            =   8400
            TabIndex        =   239
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Subsidiary"
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
            Index           =   36
            Left            =   8400
            TabIndex        =   237
            Top             =   2640
            Width           =   1080
         End
         Begin VB.Label Label1 
            Caption         =   "Cost Center"
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
            Index           =   29
            Left            =   8400
            TabIndex        =   235
            Top             =   1545
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "Grp Cat."
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
            Index           =   30
            Left            =   5280
            TabIndex        =   234
            Top             =   2625
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Position"
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
            Index           =   31
            Left            =   5280
            TabIndex        =   233
            Top             =   2340
            Width           =   960
         End
         Begin VB.Label Label1 
            Caption         =   "Sub Division"
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
            Index           =   32
            Left            =   2520
            TabIndex        =   232
            Top             =   3195
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Division"
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
            Index           =   33
            Left            =   8400
            TabIndex        =   231
            Top             =   2400
            Width           =   960
         End
         Begin VB.Label Label16 
            Caption         =   "Final Exp Date"
            Height          =   255
            Index           =   1
            Left            =   5280
            TabIndex        =   226
            Top             =   750
            Width           =   1275
         End
         Begin VB.Label Label65 
            Caption         =   "DOR"
            Height          =   255
            Left            =   8400
            TabIndex        =   225
            Top             =   960
            Width           =   1005
         End
         Begin VB.Label Label67 
            Caption         =   "Gua. Renew"
            Height          =   255
            Left            =   2520
            TabIndex        =   206
            Top             =   2880
            Width           =   1260
         End
         Begin VB.Label Label40 
            Caption         =   "Principal Warranty "
            Height          =   255
            Left            =   240
            TabIndex        =   129
            Top             =   3480
            Width           =   1590
         End
         Begin VB.Label Label84 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   5280
            TabIndex        =   128
            Top             =   3480
            Width           =   1215
         End
         Begin VB.Label Label42 
            Caption         =   "Instl mode"
            Height          =   255
            Left            =   240
            TabIndex        =   127
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label26 
            Caption         =   "Renewal "
            Height          =   240
            Left            =   240
            TabIndex        =   126
            Top             =   2640
            Width           =   1065
         End
         Begin VB.Label Label25 
            Caption         =   "Dept. Name"
            Height          =   255
            Left            =   5280
            TabIndex        =   125
            Top             =   1770
            Width           =   945
         End
         Begin VB.Label Label72 
            Caption         =   "PostCode"
            Height          =   255
            Left            =   240
            TabIndex        =   124
            Top             =   960
            Width           =   1140
         End
         Begin VB.Label Label69 
            Caption         =   "Birthdate"
            Height          =   255
            Left            =   240
            TabIndex        =   123
            Top             =   1800
            Width           =   1005
         End
         Begin VB.Label Label60 
            Caption         =   "State"
            Height          =   255
            Left            =   240
            TabIndex        =   122
            Top             =   1240
            Width           =   465
         End
         Begin VB.Label Label57 
            Caption         =   "City/Town"
            Height          =   255
            Left            =   2520
            TabIndex        =   121
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label54 
            Caption         =   "Address"
            Height          =   255
            Left            =   240
            TabIndex        =   120
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label Label56 
            Caption         =   "Sex "
            Height          =   255
            Left            =   240
            TabIndex        =   119
            Top             =   1505
            Width           =   825
         End
         Begin VB.Label Label9 
            Caption         =   "Pol. Year"
            Height          =   255
            Left            =   8400
            TabIndex        =   118
            Top             =   1275
            Width           =   690
         End
         Begin VB.Label Label63 
            Caption         =   "TakeOver"
            Height          =   255
            Left            =   2520
            TabIndex        =   117
            Top             =   1560
            Width           =   825
         End
         Begin VB.Label Label12 
            Caption         =   "Branch"
            Height          =   255
            Left            =   240
            TabIndex        =   116
            Top             =   2400
            Width           =   960
         End
         Begin VB.Label Label70 
            Caption         =   "Race"
            Height          =   255
            Left            =   240
            TabIndex        =   115
            Top             =   2080
            Width           =   420
         End
         Begin VB.Label Label71 
            Caption         =   "Nationality"
            Height          =   255
            Left            =   2520
            TabIndex        =   114
            Top             =   2080
            Width           =   960
         End
         Begin VB.Label Label27 
            Caption         =   "Marital Status"
            Height          =   255
            Left            =   2520
            TabIndex        =   113
            Top             =   1820
            Width           =   1095
         End
         Begin VB.Label Label29 
            Caption         =   "Age Band"
            Height          =   255
            Left            =   240
            TabIndex        =   112
            Top             =   3150
            Width           =   735
         End
         Begin VB.Label Label20 
            Caption         =   "Ins Type"
            Height          =   255
            Left            =   5280
            TabIndex        =   111
            Top             =   180
            Width           =   960
         End
         Begin VB.Label Label21 
            Caption         =   "Plan Code"
            Height          =   255
            Left            =   8400
            TabIndex        =   110
            Top             =   180
            Width           =   825
         End
         Begin VB.Label Label10 
            Caption         =   "Date Join"
            Height          =   255
            Left            =   5280
            TabIndex        =   109
            Top             =   1275
            Width           =   780
         End
         Begin VB.Label Label19 
            Caption         =   "E'yee Num"
            Height          =   255
            Left            =   5280
            TabIndex        =   108
            Top             =   1515
            Width           =   1545
         End
         Begin VB.Label Label18 
            Caption         =   "Grp. Company"
            Height          =   255
            Left            =   5280
            TabIndex        =   107
            Top             =   2055
            Width           =   1185
         End
         Begin VB.Label Label14 
            Caption         =   "Agent Code"
            Height          =   255
            Left            =   2520
            TabIndex        =   106
            Top             =   2620
            Width           =   1140
         End
         Begin VB.Label Label17 
            Caption         =   "Bordx. Date"
            Height          =   255
            Left            =   5280
            TabIndex        =   104
            Top             =   1020
            Width           =   1275
         End
         Begin VB.Label Label8 
            Caption         =   "Bat Num"
            Height          =   255
            Left            =   10200
            TabIndex        =   103
            Top             =   1275
            Width           =   645
         End
         Begin VB.Label Label15 
            Caption         =   "Eff Date"
            Height          =   255
            Left            =   5280
            TabIndex        =   102
            Top             =   480
            Width           =   1275
         End
         Begin VB.Label Label16 
            Caption         =   "Exp Date"
            Height          =   255
            Index           =   0
            Left            =   8400
            TabIndex        =   101
            Top             =   480
            Width           =   915
         End
         Begin VB.Label txtMBMBatchNo 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   10920
            TabIndex        =   42
            Top             =   1200
            Width           =   555
         End
         Begin VB.Label txtAGECode 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1320
            TabIndex        =   100
            Top             =   3150
            Width           =   1035
         End
         Begin VB.Label Label22 
            Caption         =   "Serv. Provider"
            Height          =   255
            Left            =   6480
            TabIndex        =   105
            Top             =   3840
            Visible         =   0   'False
            Width           =   1005
         End
      End
      Begin MSComctlLib.ListView lstCovAdd 
         Height          =   1305
         Left            =   -74880
         TabIndex        =   171
         Top             =   3540
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   2302
         SortKey         =   10
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   49
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Salutation"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "IC No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Birthdate"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sex"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Adult/Child"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Relationship"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "PlnCode"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Status"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Exclusion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "ID"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "EndtStatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Annual Limit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "EffDate"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "ExpDate"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "UndExcess"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "MBMCMco"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Annual Limit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Premium"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   " Allergic"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "PrevPLNCode"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "RB Amt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "Payor Remarks"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Prev Insurance Com"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Gen WP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "EOP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "Pre WP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "EOP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "Spe WP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "EOP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "StaffDis"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Text            =   "Clm Non-Payment Fr"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Text            =   "Clm Non-Payment To"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   33
            Text            =   "BordxDate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   34
            Text            =   "PayorPrem"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Text            =   "PayorPremNet"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Text            =   "PayorMCOGross"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Text            =   "PayorMCONet"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   38
            Text            =   "RenewalDisc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   39
            Text            =   "PolDisc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   40
            Text            =   "PolDiscAmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   41
            Text            =   "RenDiscAmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   42
            Text            =   "LoadingPrem"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   43
            Text            =   "PayMBMNo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   44
            Text            =   "PayMBMSeqNo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   45
            Text            =   "PayClientNo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   46
            Text            =   "PayClientID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   47
            Text            =   "2ndIC"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   48
            Key             =   "Student"
            Text            =   "HR Ind"
            Object.Width           =   1235
         EndProperty
      End
      Begin MSComctlLib.ListView lstAdjustmentHis 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   172
         Top             =   360
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Endt Bordx Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Batch no"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Adjustment Date"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Endt Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Endt Eff Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Adjusted By"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstMBMHis 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   205
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Mem No"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Topup"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "TopUp MBM Number"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Pol No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Eff Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Exp Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Cancel Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Data Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Pol Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Bordx Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Endt Bordx Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "MBM Reg Date"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstFeesView 
         Height          =   4275
         Left            =   -74880
         TabIndex        =   219
         Top             =   480
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   7541
         SortKey         =   4
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   29
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Rel"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ID"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "PlnCode"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "EffDate"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ExpDate"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "CancelDate"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "AnnLmt"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "AvLmt"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Text            =   "PerDisability"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Text            =   "BasicPrem"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Text            =   "Prem"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   15
            Text            =   "BasicMco"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Mco"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "InvoiceNo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "InvoiceDate"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "CR Note"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "CR Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "RefInvNo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "RefInvDate"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   25
            Text            =   "EffMco"
            Object.Width           =   1409
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   26
            Text            =   "RefMco"
            Object.Width           =   1409
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   27
            Text            =   "EffPrem"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   28
            Text            =   "PremRed"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Frame Frame2 
         Height          =   525
         Left            =   -74880
         TabIndex        =   130
         Top             =   960
         Width           =   11145
         Begin VB.Label txtMCOFees 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3960
            TabIndex        =   221
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label lblMCOFee 
            Caption         =   "MCO Fees"
            Height          =   255
            Left            =   2760
            TabIndex        =   220
            Top             =   195
            Width           =   870
         End
         Begin VB.Label txtMBMPremium 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   132
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label lblPremium 
            Caption         =   "Premium"
            Height          =   255
            Left            =   120
            TabIndex        =   131
            Top             =   180
            Width           =   960
         End
      End
   End
   Begin VB.Frame Frame7 
      Height          =   525
      Left            =   60
      TabIndex        =   173
      Top             =   7440
      Width           =   11655
      Begin VB.TextBox txtOriAnnLmt 
         Height          =   285
         Left            =   3720
         TabIndex        =   230
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAnnLmt 
         Height          =   285
         Left            =   3120
         TabIndex        =   229
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtMBMBasicPremium 
         Height          =   285
         Left            =   2520
         TabIndex        =   228
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton CmdMBMOthers 
         Caption         =   "Others Details"
         Enabled         =   0   'False
         Height          =   315
         Left            =   7395
         TabIndex        =   180
         Top             =   150
         Width           =   1200
      End
      Begin VB.TextBox txtcount 
         Height          =   285
         Left            =   2040
         TabIndex        =   179
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   10500
         TabIndex        =   178
         Top             =   150
         Width           =   960
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9555
         TabIndex        =   95
         Top             =   120
         Width           =   960
      End
      Begin VB.CommandButton cmdEditMBM 
         Caption         =   "E&dit Status"
         Enabled         =   0   'False
         Height          =   315
         Left            =   8590
         TabIndex        =   177
         Top             =   150
         Width           =   960
      End
      Begin VB.TextBox txtMBMAdultChild 
         Height          =   285
         Left            =   120
         TabIndex        =   176
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtLastEndStatus 
         Height          =   285
         Left            =   840
         TabIndex        =   175
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtLastEdtDate 
         Height          =   285
         Left            =   1440
         TabIndex        =   174
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame16 
      Height          =   555
      Left            =   60
      TabIndex        =   181
      Top             =   240
      Width           =   11775
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&Pre"
         Height          =   285
         Left            =   4640
         TabIndex        =   2
         Top             =   160
         Width           =   810
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
         Height          =   285
         Left            =   5450
         TabIndex        =   3
         Top             =   160
         Width           =   810
      End
      Begin VB.CommandButton cmdSearchMBM 
         Caption         =   "&Search"
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
         Left            =   6260
         TabIndex        =   4
         Top             =   160
         Width           =   810
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Go"
         Height          =   285
         Left            =   3840
         TabIndex        =   1
         Top             =   160
         Width           =   810
      End
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   0
         Top             =   160
         Width           =   2655
      End
      Begin VB.Label lblTopupStatus 
         Caption         =   "This Membership has a Topup Policy"
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
         Height          =   315
         Left            =   7560
         TabIndex        =   224
         Top             =   180
         Visible         =   0   'False
         Width           =   3555
      End
      Begin VB.Label Label76 
         Caption         =   "Mem Num."
         Height          =   255
         Left            =   120
         TabIndex        =   182
         Top             =   160
         Width           =   1275
      End
   End
   Begin VB.Frame Frame10 
      Height          =   855
      Left            =   60
      TabIndex        =   183
      Top             =   780
      Width           =   11730
      Begin VB.CheckBox ChkEffInd 
         Caption         =   "Check1"
         Height          =   315
         Left            =   3120
         TabIndex        =   240
         Top             =   480
         Width           =   255
      End
      Begin VB.ComboBox cboExpStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmNewMBMAdjustment.frx":00F9
         Left            =   7560
         List            =   "frmNewMBMAdjustment.frx":00FB
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   184
         Top             =   180
         Width           =   1080
      End
      Begin VB.ComboBox cboEndorseStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmNewMBMAdjustment.frx":00FD
         Left            =   1080
         List            =   "frmNewMBMAdjustment.frx":00FF
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   180
         Width           =   1935
      End
      Begin VB.ComboBox cboStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmNewMBMAdjustment.frx":0101
         Left            =   4800
         List            =   "frmNewMBMAdjustment.frx":0103
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   180
         Width           =   1680
      End
      Begin VB.ComboBox cboDataStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmNewMBMAdjustment.frx":0105
         Left            =   7560
         List            =   "frmNewMBMAdjustment.frx":0107
         TabIndex        =   10
         Text            =   "A - Active"
         Top             =   435
         Width           =   1065
      End
      Begin VB.TextBox txtBatchNo 
         Height          =   285
         Left            =   10200
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "1"
         Top             =   180
         Width           =   480
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   465
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtEndtBordxDate 
         Height          =   285
         Left            =   4800
         TabIndex        =   9
         Top             =   465
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtClnInsType 
         Height          =   285
         Left            =   10920
         TabIndex        =   207
         Top             =   500
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtClnPlanCode 
         Height          =   285
         Left            =   11260
         TabIndex        =   208
         Top             =   500
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblLimitType 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   8760
         TabIndex        =   212
         Top             =   495
         Width           =   2895
      End
      Begin VB.Label Label81 
         Caption         =   "H&&C"
         Height          =   255
         Left            =   10800
         TabIndex        =   211
         Top             =   200
         Width           =   375
      End
      Begin VB.Label LblHnS 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11160
         TabIndex        =   210
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label89 
         Caption         =   "Exp. Status"
         Height          =   255
         Left            =   6600
         TabIndex        =   192
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label59 
         Caption         =   "Bat. Num"
         Height          =   255
         Left            =   8760
         TabIndex        =   191
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Endt Bordx Date"
         Height          =   255
         Left            =   3480
         TabIndex        =   190
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblAdj 
         Caption         =   "Endt Eff Date"
         Height          =   255
         Left            =   120
         TabIndex        =   189
         Top             =   480
         Width           =   960
      End
      Begin VB.Label txtLastBatch 
         BackColor       =   &H80000009&
         Height          =   240
         Left            =   9600
         TabIndex        =   188
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label62 
         Caption         =   "Data Status"
         Height          =   255
         Index           =   0
         Left            =   6600
         TabIndex        =   187
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label24 
         Caption         =   "Endt. Status"
         Height          =   255
         Left            =   120
         TabIndex        =   186
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label55 
         Caption         =   "Pol. Status"
         Height          =   255
         Left            =   3480
         TabIndex        =   185
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   60
      TabIndex        =   193
      Top             =   1620
      Width           =   11775
      Begin VB.ComboBox cboMBMSalutation 
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Text            =   "MR"
         Top             =   165
         Width           =   690
      End
      Begin VB.TextBox txtMBMIcBcPP 
         Height          =   285
         Left            =   7500
         MaxLength       =   30
         TabIndex        =   13
         Top             =   180
         Width           =   1935
      End
      Begin VB.TextBox txtMBMName 
         Height          =   330
         Left            =   2580
         MaxLength       =   60
         TabIndex        =   12
         Top             =   165
         Width           =   4095
      End
      Begin VB.ComboBox cboHLTCode 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7500
         Sorted          =   -1  'True
         TabIndex        =   16
         Text            =   "S"
         Top             =   420
         Width           =   1935
      End
      Begin VB.TextBox txtMBMPolNo 
         Height          =   330
         Left            =   2580
         MaxLength       =   25
         TabIndex        =   15
         Top             =   470
         Width           =   4095
      End
      Begin VB.ComboBox cboPAYCode 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         Style           =   1  'Simple Combo
         TabIndex        =   14
         Top             =   440
         Width           =   675
      End
      Begin MSMask.MaskEdBox txtMBMTopUpJoinDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   10380
         TabIndex        =   222
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "TopUp Join Date"
         Height          =   435
         Index           =   1
         Left            =   9480
         TabIndex        =   223
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label80 
         Caption         =   "Mem Name"
         Height          =   255
         Left            =   1680
         TabIndex        =   199
         Top             =   210
         Width           =   1140
      End
      Begin VB.Label Label73 
         Caption         =   "Policy Num"
         Height          =   255
         Left            =   1680
         TabIndex        =   198
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label31 
         Caption         =   "Salutation"
         Height          =   255
         Left            =   120
         TabIndex        =   197
         Top             =   195
         Width           =   1140
      End
      Begin VB.Label Label64 
         Caption         =   "IC. Num"
         Height          =   255
         Left            =   6780
         TabIndex        =   196
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label75 
         Caption         =   "Hlt Type"
         Height          =   255
         Left            =   6780
         TabIndex        =   195
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label74 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   194
         Top             =   510
         Width           =   915
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Member Adjustment"
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
      Height          =   255
      Left            =   -60
      TabIndex        =   203
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmNewMBMAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bExit As Boolean
Dim ScanCode As String
Dim txtTempPrem As Double
Dim txtTempMCOFees As Double
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset
Dim Disc As Boolean
Dim HisCisEffDate, HISCISExpDate As String
Dim CisEffDate, CisExpDate As String
Public PolDics As String
Public RenDisc As String
Dim TempLifeTimeStatus As String
Dim PrevLifeMBMNo As String
Dim ADOrstMBM As New ADODB.Recordset
Dim SptLmtInd As String
Dim WLifeStatus As Boolean
Dim OverAge As Boolean
Dim PlnAgeLmt As Integer
Dim PLNProRate As String
Dim RoundUpStatus As String
Dim PrevRenewKey1 As String
Dim tmpPerDisabilitAmt As Double
Public MBMPayorEXPDate As Integer
Dim tmpCRNote As String
Dim tmpCRDate As String
Dim tmpCancelDate As String
Dim tmpUpdateAnnLmt As Boolean
Dim tmpServerDate As String
Dim tmpPolHistory As String
Dim tmpPAYGRPCode As String
Dim tmpEmail As String

Private Sub cboCovAdultChild_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboCovAdultChild, KeyAscii)
End Sub

Private Sub cboCovAdultChild_LostFocus()
    Clipboard.Clear
End Sub

Private Sub cboCovRelationship_LostFocus()
    Clipboard.Clear
End Sub

Private Sub cboCovSex_LostFocus()
    Clipboard.Clear
End Sub

Private Sub cboEndorseStatus_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboEndorseStatus, KeyAscii)
End Sub

Private Sub cboEndorseStatus_LostFocus()
Dim tmpEndtStatus As String
Dim RecordSaved

    tmpEndtStatus = Mid(cboEndorseStatus, 1, 1)
    Call EnabledFrame(False)
    cboStatus.Enabled = False
    Call EnableSupplementary(False)
    Call TabEnabled(False)
    CmdCancelSupp.Enabled = False
    If tmpEndtStatus = "S" Or tmpEndtStatus = "T" Then
        If Trim(Mid(cboINSCode, 1, 1)) = "F" Or Trim(Mid(cboINSCode, 1, 1)) = "H" Or _
            Trim(txtClnInsType) = "F" Or Trim(txtClnInsType) = "H" Then
        Else
            MsgBox "This Plan does not have supplementary", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        End If
    End If
    If tmpEndtStatus = "P" And Trim(cboINSCode) = "" And Trim(cboPLNCode) = "" Then
        MsgBox "No Plan to Change !", vbOKOnly + vbCritical, DialogBoxTitle
        Exit Sub
    End If
    Call TabEnabled(True)
    Frame10.Enabled = True
    
    If tmpEndtStatus = "S" Then
        Frame6.Enabled = True
        Call EnableSupplementary(True)
        txtSuppPLNCode.Enabled = True
        txtSuppEffDate.Enabled = True
        txtSuppPLNCode = cboPLNCode.Text
        txtSuppEffDate = txtMBMPayorEffDate
        txtSuppExpDate = txtMBMPayorExpDate
        cmdAdd.Enabled = True
        cmdRemove.Enabled = True
    End If
    If tmpEndtStatus = "T" Then
        Frame6.Enabled = True
        Call EnableSupplementary(True)
        cmdRemove.Enabled = True
    End If
    If tmpEndtStatus = "A" Or tmpEndtStatus = "P" Then
        Frame1.Enabled = True
        Frame4.Enabled = True
        Frame6.Enabled = True
        Call EnableSupplementary(True)
        cboHLTCode.Enabled = False
        txtMBMPayorEffDate.Enabled = True
        txtMBMPayorExpDate.Enabled = True
        cmdAdd.Enabled = False
        cmdRemove.Enabled = False
        cmdUpdate.Enabled = True
        If Trim(cboINSCode) = "" And Trim(cboPLNCode) = "" Then
            txtMBMPayorEffDate.Enabled = False
            txtMBMPayorExpDate.Enabled = False
        End If
    End If
    If tmpEndtStatus = "K" Then
        Frame6.Enabled = True
        Call EnableSupplementary(True)
        CmdCancelSupp.Enabled = True
    End If
    If (tmpEndtStatus = "C" Or tmpEndtStatus = "X" Or tmpEndtStatus = "D" Or tmpEndtStatus = "Y") And Mid(cboStatus, 1, 1) = "A" Then
        If Trim(cboINSCode) = "" And Trim(cboPLNCode) = "" Then
            MsgBox "This is Clinical Member, pls use Clinical Adjustment !", vbOKOnly + vbCritical, DialogBoxTitle
            Exit Sub
        Else
            cboStatus.Enabled = True
        End If
    End If
End Sub

Private Sub CboInsMode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboInsMode, KeyAscii)
End Sub

Private Sub cboMBMRenewal_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(CboInsMode, KeyAscii)
End Sub

Private Sub cboMBMRenewal_LostFocus()
    Clipboard.Clear
End Sub

Private Sub cboMBMSalutation_LostFocus()
    cboMBMSalutation.Text = ChkStr(cboMBMSalutation.Text)
    If cboMBMSalutation.Text = "MR" Then
        cboMBMSex = "M"
    Else
        cboMBMSex = "F"
    End If
End Sub

Private Sub cboMBMSex_LostFocus()
    Clipboard.Clear
End Sub

Private Sub cboNATCode_LostFocus()
    Clipboard.Clear
End Sub
Private Sub cboPLNCode_LostFocus()
Dim tmpGetAnnLmt As String
Dim Confirm1
    Clipboard.Clear
    Call MemberCallAge
    If Len(Trim(txtAGECode)) <> 0 Then
        Call MemberCallFees
    End If
    
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    tmpGetAnnLmt = Format(GetAnnualLimit(Trim(cboINSCode), Trim(cboPLNCode), Trim(cboHLTCode), CDate(txtMBMPayorEffDate)), "#,###,##0.00")
    If CDbl(tmpGetAnnLmt) <> CDbl(txtOriAnnLmt) Then
        Confirm1 = MsgBox(" Do you want to change Annual Limit?", vbCritical + vbYesNo)
        If Confirm1 = vbYes Then
            txtAnnLmt = tmpGetAnnLmt
            tmpUpdateAnnLmt = True
        Else
            tmpUpdateAnnLmt = False
        End If
    End If
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
End Sub

Private Sub cboRACCode_LostFocus()
    Clipboard.Clear
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboStatus, KeyAscii)
End Sub

Private Sub chkCheck_Click()
Dim strsql As String
Dim rst As New ADODB.Recordset

    If chkCheck.Value = 1 Then
        strsql = ""
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        strsql = ""
        strsql = "SELECT PLNCode, PLNDescription,PAYCode from PLN where HLTCode ='" & Trim(cboHLTCode) & "' AND PAYCode ='" & Trim(cboPAYCode) & "' "
        rst.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        cboPLNCode.Clear
        Do While Not rst.EOF
            cboPLNCode.AddItem rst!PLNCode
        rst.MoveNext
        Loop
        rst.Close
        Set rst = Nothing
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
    Else
        cboPLNCode.Clear
    End If
End Sub

Private Sub CmdMBMOthers_Click()
    frmNewMBMOthers.Source = 1
    If frmNewMBMOthers.Visible = False Then
        Load frmNewMBMOthers
    End If
    frmNewMBMOthers.Show
    frmNewMBMOthers.Frame5.Enabled = True
End Sub

Private Sub Form_Load()
    frmBeginScan.txtScanCode.SetFocus
    cboEndorseStatus.Enabled = False
    cboStatus.Enabled = False
    cboExpStatus.Enabled = False
    txtDate.Enabled = False
    Frame10.Enabled = False
    Call ComboListing
    Call EnableEntries(True)
    cmdSave.Enabled = False
'    cmdRemove.Enabled = False
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    bExit = False
'    SSTab1.Tab = 0
    cboINSCode.Enabled = False
    cboPLNCode.Enabled = False
    txtTempMCOFees = 0
    txtTempPrem = 0
    PLNProRate = "Y"
    lblTopupStatus.Visible = False
    Cmd_HRInd.Text = "" 'Add Student/OKU Indicator to cover > 21yrs old - By CHEng 1 Dec 2016'
    Cmb_Cstatus.Text = "" 'Dependent Status change to Editable - By CHEng 3 Dec 2016
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call EnterToTab(KeyAscii)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
    
    '--- Enable PrintScreen feature
    '--- IMPORTANT - Don't forget to set .KeyPreview = True for this form at design time!!
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF6
                KeyCode = 0
                PrintScreenSnapShot
        End Select
    End If

    i = SSTab1.Tab
    'handle ctrl+tab to move to the next tab
    If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
        If i = 2 Then
            'last tab so we need to wrap to tab 1
            SSTab1.Tab = 3
        ElseIf i = 1 Then
            'last tab so we need to wrap to tab 1
            SSTab1.Tab = 2
        ElseIf i = 0 Then
            'increment the tab
            SSTab1.Tab = 1
        Else
            SSTab1.Tab = 0
        End If
    ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
        If i = 3 Then
            'last tab so we need to wrap to tab 1
            SSTab1.Tab = 2
        
        ElseIf i = 2 Then
            'last tab so we need to wrap to tab 1
            SSTab1.Tab = 1
        ElseIf i = 1 Then
            'increment the tab
            SSTab1.Tab = 0
        Else
            SSTab1.Tab = 1
        End If
    End If
'    SSTab1.SetFocus
End Sub

Private Sub ComboListing()
Dim SRV As New ADODB.Recordset
Dim RAC As New ADODB.Recordset
Dim NAT As New ADODB.Recordset
Dim REL As New ADODB.Recordset
Dim STA As New ADODB.Recordset
Dim CTY As New ADODB.Recordset
Dim Cmd As New ADODB.Command

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    Set Cmd.ActiveConnection = medixmasterconn
    Cmd.CommandText = "SearchState"
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandTimeout = 15
    STA.CursorLocation = adUseClient
    STA.CursorType = adOpenForwardOnly
    STA.CacheSize = 100
    Set STA = Cmd.Execute
    
    Do Until STA.EOF
        With STA
            cboSTACode.AddItem !STADescription
            STA.MoveNext
        End With
    Loop
    STA.Close
    Set STA = Nothing

    Set Cmd.ActiveConnection = medixmasterconn
    Cmd.CommandText = "SearchSRV"
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandTimeout = 15
    SRV.CursorLocation = adUseClient
    SRV.CursorType = adOpenForwardOnly
    SRV.CacheSize = 100
    Set SRV = Cmd.Execute

    Do Until SRV.EOF
        With SRV
            cboSRVCode.AddItem !SRVCode
            SRV.MoveNext
        End With
    Loop
    SRV.Close
    Set SRV = Nothing

    Set Cmd.ActiveConnection = medixmasterconn
    Cmd.CommandText = "SearchRace"
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandTimeout = 15
    RAC.CursorLocation = adUseClient
    RAC.CursorType = adOpenForwardOnly
    RAC.CacheSize = 100
    Set RAC = Cmd.Execute

    Do Until RAC.EOF
        With RAC
            cboRACCode.AddItem !RACCode
            RAC.MoveNext
        End With
    Loop
    RAC.Close
    Set RAC = Nothing
    
    Set Cmd.ActiveConnection = medixmasterconn
    Cmd.CommandText = "SearchNAT"
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandTimeout = 15
    NAT.CursorLocation = adUseClient
    NAT.CursorType = adOpenForwardOnly
    NAT.CacheSize = 100
    Set NAT = Cmd.Execute

    Do Until NAT.EOF
        With NAT
            cboNATCode.AddItem !NATCode
            NAT.MoveNext
        End With
    Loop
    NAT.Close
    Set NAT = Nothing
    
   ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    Set Cmd.ActiveConnection = medixmasterconnMaint
    Cmd.CommandText = "SearchCity"
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandTimeout = 15
    CTY.CursorLocation = adUseClient
    CTY.CursorType = adOpenForwardOnly
    CTY.CacheSize = 100
    Set CTY = Cmd.Execute

    Do Until CTY.EOF
        With CTY
            cboBRNCode.AddItem !CTYDescription
            cboCTYCode.AddItem !CTYDescription
            CTY.MoveNext
        End With
    Loop
    CTY.Close
    Set CTY = Nothing
  '  medixmasterconnMaint.Close
  '  Set medixmasterconnMaint = Nothing
    
    Set Cmd.ActiveConnection = medixmasterconn
    Cmd.CommandText = "SearchREL"
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandTimeout = 15
    REL.CursorLocation = adUseClient
    REL.CursorType = adOpenForwardOnly
    REL.CacheSize = 100
    Set REL = Cmd.Execute

    Do Until REL.EOF
        With REL
            cboCovRelationship.AddItem !RELCode
            REL.MoveNext
        End With
    Loop
    REL.Close
    Set REL = Nothing
    
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing

    cboMBMSex.AddItem "M"
    cboMBMSex.AddItem "F"
    
    cboCovSex.AddItem "M"
    cboCovSex.AddItem "F"
    
    cboMBMMaritalStatus.AddItem "S"
    cboMBMMaritalStatus.AddItem "M"
    
    cboCovAdultChild.AddItem "A"
    cboCovAdultChild.AddItem "C"
    
    cboMBMSalutation.AddItem "MR"
    cboMBMSalutation.AddItem "MS"
    cboMBMSalutation.AddItem "MRS"
    
    cboCovSalutation.AddItem "MR"
    cboCovSalutation.AddItem "MS"
    cboCovSalutation.AddItem "MRS"
    
    
    cboSuppStaffDisc.AddItem "N - No"
    cboSuppStaffDisc.AddItem "Y - Yes"
        
    cboEndorseStatus.AddItem "A" & " - " & "Amendment"
    cboEndorseStatus.AddItem "C" & " - " & "Cancellation"
    cboEndorseStatus.AddItem "S" & " - " & "Add Supplementary"
    cboEndorseStatus.AddItem "K" & " - " & "Cancel Supplementary"
    cboEndorseStatus.AddItem "T" & " - " & "Delete Supplementary"
    cboEndorseStatus.AddItem "X" & " - " & "Suspend Membership"
'    cboEndorseStatus.AddItem "Y" & " - " & "Re-Activate Membership"
'    cboEndorseStatus.AddItem "I" & " - " & "Insert New Member"
    cboEndorseStatus.AddItem "D" & " - " & "Delete Policy"
    cboEndorseStatus.AddItem "P" & " - " & "Plan Change"
    
    cboStatus.AddItem "A" & " - " & "Active"
    cboStatus.AddItem "C" & " - " & "Cancel"
    cboStatus.AddItem "D" & " - " & "Deleted"
    cboStatus.AddItem "S" & " - " & "Suspended"
    cboStatus.AddItem "P" & " - " & "Plan Change"
                
    cboExpStatus.AddItem "Y" & " - " & "Yes"
    cboExpStatus.AddItem "N" & " - " & "No"
                
    cboDataStatus.AddItem "A" & " - " & "Active"
    cboDataStatus.AddItem "D" & " - " & "Deleted"
            
    cboMBMTakeOver.AddItem "N - None"
    cboMBMTakeOver.AddItem "Y - Take Over"
    cboMBMTakeOver.AddItem "C - Conversion"
        
    cboMBMRenewal.AddItem "R - Renew"
    cboMBMRenewal.AddItem "N - New"
    cboMBMRenewal.AddItem "T - Timely"
    cboMBMRenewal.AddItem "P - Partial"
      
    CboInsMode.AddItem "A - Annually"
    CboInsMode.AddItem "Q - Quarterly"
    CboInsMode.AddItem "S - Semi Annual"
    CboInsMode.AddItem "M - Monthly"
    
    CboGuaRenewal.Clear
    CboGuaRenewal.AddItem "  "
    CboGuaRenewal.AddItem "1 - No Life Time Limit"
    CboGuaRenewal.AddItem "3 - With Life Time Limit"
    
End Sub

Private Sub ADJListing()
    Dim ADJ As New ADODB.Recordset
    Dim ADJList As ListItem
    Dim sql As String
    
    If Len(Trim(txtMBMNumber)) Then
         sql = " select MBMAmendID , MBMNumber,  MBMAmendUser , " & _
           " MBMNAme, MBMpolicyNo, MBMICBCPP, MBMAmendEdtType , MBMAmendDate , MBMAENDtEffDate ,   " & _
           " MBMAEndtBordxDate, MBMAEndtBatchNo  " & _
           " from MBMAmendHistory where MBMnumber ='" & Trim(txtMBMNumber) & "'"
         ADJ.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
         
         Do Until ADJ.EOF
         If Len(ADJ!MBMAENDtBordxDate) Then
             Set ADJList = lstAdjustmentHis.ListItems.Add(, , ADJ!MBMAENDtBordxDate)
        Else
        Set ADJList = lstAdjustmentHis.ListItems.Add(, , " ")
        End If
                 If Len(Trim(ADJ!MBMAEndtBatchNo)) Then
                    ADJList.SubItems(1) = ADJ!MBMAEndtBatchNo
                 End If
                 If Len(Trim(ADJ!MBMAmendDate)) Then
                    ADJList.SubItems(2) = ADJ!MBMAmendDate
                 End If

                If Len(ADJ!MBMAmendEdtType) Then
                    ADJList.SubItems(3) = ADJ!MBMAmendEdtType
                End If
                
                If Len(ADJ!MBMAEndtEffDate) Then
                    ADJList.SubItems(4) = ADJ!MBMAEndtEffDate
                End If

                 If Len(Trim(ADJ!MBMAmendUser)) Then
                    ADJList.SubItems(5) = ADJ!MBMAmendUser
                 End If
                
                 ADJ.MoveNext
              '   End If
         Loop
         ADJ.Close
         Set ADJ = Nothing
    End If
End Sub

' ####################### Calculation's Function ###############
Private Sub MemberCallAge()
Dim SelectAge As String
Dim tmpPlnCode As String
Dim tmpINSType As String

    tmpPlnCode = Trim(cboPLNCode)
    tmpINSType = Trim(cboINSCode)
    If Trim(tmpPlnCode) = "" Then
        tmpPlnCode = Trim(txtClnPlanCode)
        tmpINSType = Trim(txtClnInsType)
    End If
    If Trim(txtMBMPayorEffDate) <> "__/__/____" _
        And Trim(txtMBMBirthDate) <> "__/__/____" _
        And Len(Trim(cboHLTCode)) <> 0 Then
        If IsDate(txtMBMBirthDate) Then
            SelectAge = GetAge(CDate(txtMBMPayorEffDate), CDate(txtMBMBirthDate))
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            txtAGECode = GetAgeBand(Trim(Mid(cboHLTCode, 1, 1)), SelectAge, tmpPlnCode, tmpINSType)
          '  medixmasterconn.Close
          '  Set medixmasterconn = Nothing
            If Len(Trim(txtAGECode)) = 0 Then
                MsgBox "Please Enter A Valid Date OR Age Out of Acceptable Range", vbOKOnly + vbCritical, DialogBoxTitle
                txtMBMPayorEffDate.SetFocus
            Else
                txtMBMAdultChild = "A"
                If Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)) = "01" Then
                    txtMBMAdultChild = "C"
                End If
            End If
        End If
    End If
End Sub

Public Sub MemberCallFees()
Dim TempPremium, tempMCOFee As Single
Dim Confirm1, Confirm2
Dim PLNRec As New ADODB.Recordset
Dim SqlPln As String
Dim tmpPremSexStatus As String
Dim tmpPrmInd As String

    If Trim(txtMBMPayorEffDate) <> "__/__/____" _
        And Trim(txtMBMPayorExpDate) <> "__/__/____" _
        And Trim(cboPLNCode) <> "" And Trim(cboINSCode) <> "" _
        And Trim(txtMBMBirthDate) <> "__/__/____" Then

      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
      '  medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        SqlPln = "SELECT PLNRoundUpStatus, PLNProRate,PLNPremGenderStatus,PRMInd from PLN where PlnCode='" & Trim(Mid(cboPLNCode, 1, 4)) & _
        "' and HLTCode = '" & Trim(cboHLTCode) & _
        "' AND PLNEffDate <='" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "'"
        If Trim(cboHLTCode) = "N" Then
            SqlPln = SqlPln & " and PAYCode='" & Trim(cboPAYCode) & "'"
        End If

        PLNRec.Open SqlPln, medixmasterconn, adOpenKeyset, adLockOptimistic
        RoundUpStatus = "Y"
        PLNProRate = "Y"
        If PLNRec.recordCount Then
            RoundUpStatus = IIf(Len(Trim(PLNRec!PLNRoundUpStatus)), Trim(PLNRec!PLNRoundUpStatus), "Y")
            PLNProRate = IIf(Len(Trim(PLNRec!PLNProRate)), Trim(PLNRec!PLNProRate), "Y")
            tmpPremSexStatus = IIf(Len(Trim(PLNRec!PLNPremGenderStatus)), Trim(PLNRec!PLNPremGenderStatus), "N")
            tmpPrmInd = IIf(Len(Trim(PLNRec!PRMInd)), Trim(PLNRec!PRMInd), "")
        End If
        PLNRec.Close
        Set PLNRec = Nothing
        
        If Trim(txtMBMAdultChild) <> "" Then
            tempMCOFee = Format(GetMCO(Trim(Mid(cboINSCode, 1, 1)), Trim(Mid(cboHLTCode, 1, 1)), txtMBMAdultChild, txtMBMPayorEffDate, Trim(Mid(cboPLNCode, 1, 4))), "#,##0.00")
            If PLNProRate <> "N" Then
                tempMCOFee = InsertMCOFees(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(tempMCOFee, "#,##0.00"))
            End If
            tempMCOFee = Format(tempMCOFee, "#,##0.00")
        End If

        If RoundUpStatus <> "N" Then
            tempMCOFee = Round(tempMCOFee)
        End If
        
        If tmpPrmInd = "Y" Then
            TempPremium = Format(GetPremium(Trim(Mid(cboINSCode, 1, 1)), Trim(Mid(cboPAYCode, 1, 2)), Trim(Mid(txtAGECode, 1, 2)), Trim(Mid(cboPLNCode, 1, 4)), Trim(Mid(cboHLTCode, 1, 1)), txtMBMPayorEffDate, Mid(cboMBMSex, 1, 1)), "#,##0.00")
        Else
            TempPremium = 0
        End If
        txtMBMBasicPremium = IIf(Len(TempPremium), TempPremium, 0)
'        PolDics = 0
'        RenDisc = 0
        strsqlPLN = "select PolicyDisc, RenewalDisc from DiscPlan where PLNCode = '" & Trim(Mid(cboPLNCode, 1, 4)) & "'"
        PLN.Open strsqlPLN, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        If PLN.recordCount Then
            PolDics = IIf(IsNumeric(PLN!PolicyDisc), PLN!PolicyDisc, 0)
            RenDisc = IIf(IsNumeric(PLN!RenewalDisc), PLN!RenewalDisc, 0)
        End If
        PLN.Close
        Set PLN = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
       ' medixmasterconnMaint.Close
       ' Set medixmasterconnMaint = Nothing
        If PolDics > 0 Then
            TempPremium = TempPremium * (100 - PolDics) / 100
        End If
        If RenDisc > 0 Then
            TempPremium = TempPremium * (100 - RenDisc) / 100
        End If
        TempPremium = Format(InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(TempPremium, "#,##0.00")), "#,##0.00")
        txtTempPrem = TempPremium
        txtTempMCOFees = tempMCOFee
        txtMBMPremium = TempPremium
        If Format(txtMCOFees, "#,##0.00") <> tempMCOFee Then
            Confirm2 = MsgBox(" Do you want to change member's MCO Fee?", vbCritical + vbYesNo)
            If Confirm2 = vbYes Then
                txtMCOFees = tempMCOFee
            End If
        End If
    End If
End Sub


' ######################### END CALCULATIOn###########
' ######################### START CHECKING ###########
Private Sub CmdNext_Click()
    Dim MBM As New ADODB.Recordset
    Dim sql As String
    Dim MBMNumber As String
    Dim status As Boolean
        
        lblTopupStatus.Visible = False
        
        status = False
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        sql = " select MBMNumber from MBM where MBMnumber > '" & RemStr(txtMBMNumber) & "' " & _
            " Order by MBMnumber "
        MBM.MaxRecords = 1
        MBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBM.recordCount Then
            MBMNumber = MBM!MBMNumber
            status = True
        End If
        MBM.Close
        Set MBM = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        
        If status = True Then
            txtMBMNumber = MBMNumber
            cmdShow_Click
        Else
            MsgBox "Last Record Reached! " & vbNewLine & "Not a valid Membership Number Format! ", vbCritical + vbOKOnly
        End If

End Sub

Private Sub cmdPrev_Click()
    Dim MBM As New ADODB.Recordset
    Dim sql As String
    Dim MBMNumber As String
    Dim status As Boolean
        
        lblTopupStatus.Visible = False
        
        status = False
        
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        sql = " select MBMNumber from MBM where MBMnumber < '" & RemStr(txtMBMNumber) & "' " & _
            " Order by MBMnumber DESC "
        MBM.MaxRecords = 1
        MBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBM.recordCount Then
            MBMNumber = MBM!MBMNumber
            status = True
        End If
        MBM.Close
        Set MBM = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
        
        If status = True Then
            txtMBMNumber = MBMNumber
            cmdShow_Click
        Else
            MsgBox "First Record Reached! " & vbNewLine & "Not a valid Membership Number Format! ", vbCritical + vbOKOnly
        End If
End Sub

Private Sub cboINSCode_LostFocus()
    Clipboard.Clear
End Sub

Private Sub cboINSCode_Click()
    If Len(Trim(cboINSCode)) Then
        If Trim(cboINSCode) = "F" Or Trim(cboINSCode) = "H" Then
            If Trim(cboINSCode) <> Trim(cboINSCode) Then
                Call EnableSupplementary(True)
            Else
                Call EnableSupplementary(False)
            End If
            cboMBMMaritalStatus = "M"
        ElseIf Trim(cboINSCode) = "I" Or Trim(cboINSCode) = "G" Then
            Call EnableSupplementary(False)
            cboMBMMaritalStatus = "S"
        End If
    End If
End Sub

Private Sub cboCovSalutation_Click()
    If ChkStr(cboCovSalutation) = "MR" Then
        cboCovSex = "M"
    Else
        cboCovSex = "F"
    End If
End Sub

Private Sub cboCovSalutation_Change()
    If ChkStr(cboCovSalutation) = "MR" Then
        cboCovSex = "M"
    Else
        cboCovSex = "F"
    End If
End Sub
Private Sub lstAdjustmentHis_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstAdjustmentHis.SortKey = ColumnHeader.Index - 1
    lstAdjustmentHis.Sorted = True
    If lstAdjustmentHis.SortOrder = lvwAscending Then
        lstAdjustmentHis.SortOrder = lvwDescending
    Else
        lstAdjustmentHis.SortOrder = lvwAscending
    End If
End Sub

Private Sub txtCovAllergic_Click()
    KeyPreview = False
End Sub

Private Sub txtCovAllergic_GotFocus()
    KeyPreview = False
End Sub

Private Sub txtCovAllergic_LostFocus()
    KeyPreview = True
    txtCovAllergic = ChkStr(txtCovAllergic)
End Sub

Private Sub txtCovAnnualLimit_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtCovBirthdate_LostFocus()
    Dim MBMValidDate As New ADODB.Recordset
    Dim medixmasterconnCOVDOB As New ADODB.Connection
    Dim strsql As String
    
    'medixmasterconnCOVDOB.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If txtCovBirthdate <> "__/__/____" Then
        If IsDate(txtCovBirthdate) Then
           strsql = "Select ValidID ,MBMDOBStart , MBMDOBEnd  From MBMValidDateRange "
           MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If MBMValidDate.recordCount Then
                If DateDiff("d", MBMValidDate!MBMDOBStart, txtCovBirthdate) < 0 Or DateDiff("d", MBMValidDate!MBMDOBEnd, txtCovBirthdate) > 0 Then
                    MsgBox "Birthdate is out of acceptable range! ", vbCritical + vbOKOnly
                    txtCovBirthdate.SetFocus
                End If
            End If
            MBMValidDate.Close
            Set MBMValidDate = Nothing
        Else
            MsgBox "Invalid date format! ", vbCritical + vbOKCancel
            txtCovBirthdate.SetFocus
        End If
    End If
    'medixmasterconnCOVDOB.Close
    'Set medixmasterconnCOVDOB = Nothing
End Sub

Private Sub txtCovExclusion_Click()
    KeyPreview = False
End Sub

Private Sub txtCovExclusion_LostFocus()
    KeyPreview = True
    txtCovExclusion = ChkStr(txtCovExclusion)
End Sub

Private Sub txtCovIcBcPp_LostFocus()
    txtCovIcBcPp = ChkStr(txtCovIcBcPp)
End Sub

Private Sub txtCovName_LostFocus()
    txtCovName = ChkStr(txtCovName)
End Sub

Private Sub txtCovUndExcess_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtDate_LostFocus()
    If IsDate(txtDate) Then
    Else
    If txtDate <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtDate.SetFocus
    End If
    End If
End Sub


Private Sub txtEndtBordxDate_GotFocus()
        If txtDate <> "__/__/____" Then
            cmdSave.Enabled = True
        End If
End Sub

Private Sub txtendtbordxdate_LostFocus()
 Dim BAT As New ADODB.Recordset
    Dim strsql As String
    
   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If IsDate(txtEndtBordxDate) Then
        strsql = "Select  BATSequence  , bordxType From BAT " & _
                " where BATBordxDate = '" & Format(txtEndtBordxDate, "mm/dd/yyyy") & "'" & _
                " and BordxType = 'E'  "
        
        If Len(Trim(cboPAYCode)) Then
            strsql = strsql + " AND PAYCode ='" & Trim(Mid(cboPAYCode, 1, 2)) & "'"
        End If
        strsql = strsql + " order by BATSequence dESC"
        BAT.MaxRecords = 1
        BAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If Not BAT.EOF Then
            txtLastBatch = BAT!BATSequence
        Else
            txtLastBatch = 1
        End If
        BAT.Close
        Set BAT = Nothing
    Else
        If txtEndtBordxDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
        End If
    End If
    'medixmasterconn.Close
    'Set medixmasterconn = Nothing
End Sub

Private Sub txtLoadingPrem_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub


Private Sub txtMBMBirthdate_LostFocus()
Dim MBMValidDate As New ADODB.Recordset
Dim strsql As String
Dim DateExit
Dim DobStart As String
Dim DOBEnd As String
Dim DateSql As String
    
    DobStart = ""
    DOBEnd = ""
    If IsDate(txtMBMBirthDate) = False Then
        MsgBox "Invalid date format! ", vbCritical + vbOKCancel
        txtMBMBirthDate.SetFocus
    Else
        DateSql = " MBMDOBStart as DateStart, MBMDOBEnd as DateEnd "
        Call GetValidDate(DobStart, DOBEnd, DateSql)
        If Trim(DobStart) <> "" And Trim(DOBEnd) <> "" Then
            
            If DateDiff("d", DobStart, txtMBMBirthDate) >= 0 And DateDiff("d", DOBEnd, txtMBMBirthDate) <= 0 Then
            Else
                DateExit = MsgBox("Uncertain Birthdate, please select OK to continue " & vbNewLine & _
                            " Otherwise press Cancel to make changes! ", vbExclamation + vbOKCancel + vbDefaultButton2)
                If DateExit = vbCancel Then
                    txtMBMBirthDate.SetFocus
                    Exit Sub
                End If
            End If
            Call MemberCallAge
        End If
    End If
End Sub

Private Sub txtMBMBirthdate_gotfocus()
    txtMBMBirthDate.SelStart = 0
End Sub

Private Sub txtMBMBordxDate_LostFocus()
    If IsDate(txtMBMBordxDate) Then
    Else
    If txtMBMBordxDate <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtMBMBordxDate.SetFocus
    End If
    End If
End Sub

Private Sub txtMBMDateJoin_LostFocus()
    If IsDate(txtMBMDateJoin) Then
    Else
    If txtMBMDateJoin <> "__/__/____" Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtMBMDateJoin.SetFocus
    End If
    End If
End Sub

Private Sub txtMBMExclusion_Click()
    KeyPreview = False
End Sub

Private Sub txtMBMExclusion_GotFocus()
    KeyPreview = False
End Sub

Private Sub txtMBMExclusion_LostFocus()
    KeyPreview = True
End Sub


Private Sub txtMBMName_LostFocus()
    ChkStr (txtMBMName)
End Sub

Private Sub txtMBMNumber_GotFocus()
    cmdShow.Default = True
    cmdSave.Enabled = False
    cmdEditMBM.Enabled = False
End Sub

Private Sub txtMBMNumber_LostFocus()
    txtMBMNumber = ChkStr(txtMBMNumber)
    cmdShow.Default = False
End Sub

Private Sub txtMBMPayorEffDate_gotfocus()
    txtMBMPayorEffDate.SelStart = 0
End Sub

Private Sub txtMBMPayorEffDate_LostFocus()
Dim CheckDays As Integer
Dim TempDate As Date
Dim EffStart As String
Dim EffEnd As String
Dim DateSql As String

    EffStart = ""
    EffEnd = ""
    If IsDate(txtMBMPayorEffDate) = False Then
        MsgBox "Invalid Date Format!", vbCritical + vbOKCancel
        txtMBMPayorEffDate.SetFocus
    Else
        DateSql = " MBMEffStart as DateStart, MBMEffEnd as DateEnd "
        Call GetValidDate(EffStart, EffEnd, DateSql)
        If EffStart <> "" And EffEnd <> "" Then
            If DateDiff("d", EffStart, txtMBMPayorEffDate) >= 0 And DateDiff("d", EffEnd, txtMBMPayorEffDate) <= 0 Then
            Else
                MsgBox "Policy Effective Date is out of acceptable range! ", vbCritical + vbOKOnly
                txtMBMPayorEffDate.SetFocus
            End If
        End If
        TempDate = Format(DateAdd("yyyy", 1, txtMBMPayorEffDate), "dd/mm/yyyy")
        'txtMBMPayorExpDate = Format(DateAdd("d", -1, TempDate), "dd/mm/yyyy")
        If Mid(cboEndorseStatus, 1, 1) <> "I" Then
            Call MemberCallAge
            If Len(Trim(txtAGECode)) <> 0 Then
                Call MemberCallFees
            End If
        End If
    End If
End Sub

' ######################### END CHECKING ###########
' ######################### Start Tab Index ###########
Private Sub SSTab1_Click(PreviousTab As Integer)
    
Screen.MousePointer = vbHourglass
    cboCovSalutation = "MR"
    txtCovName = ""
    txtCovIcBcPp = ""
    txtCovBirthdate.mask = "__/__/____"
    txtCovBirthdate.mask = "##/##/####"
    
    cboCovRelationship = "W - Wife"
    cboCovSex = "M"
    cboCovAdultChild = "A"
    txtCovAnnualLimit = ""
    txtAvailableLimit = ""
    txtCovAllergic = ""
    txtCovExclusion = ""
    txtSuppGenWP = ""
    txtSuppPREWP = ""
    TxtSuppSpeWP = ""
    cboSuppStaffDisc = ""
    txtSuppGENWPEnd.mask = ""
    txtSuppGENWPEnd = ""
    txtSuppGENWPEnd.mask = "##/##/####"
    txtSuppPREWPEnd.mask = ""
    txtSuppPREWPEnd = ""
    txtSuppPREWPEnd.mask = "##/##/####"
    txtSuppSPEWPEnd.mask = ""
    txtSuppSPEWPEnd = ""
    txtSuppSPEWPEnd.mask = "##/##/####"
    TxtSuppClmNonPayFr.mask = ""
    TxtSuppClmNonPayFr = ""
    TxtSuppClmNonPayFr.mask = "##/##/####"
    TxtSuppClmNonPayTo.mask = ""
    TxtSuppClmNonPayTo = ""
    TxtSuppClmNonPayTo.mask = "##/##/####"
    
    If SSTab1.Tab = 1 Then
        If Trim(Mid(cboINSCode, 1, 1)) = "F" Or Trim(Mid(cboINSCode, 1, 1)) = "H" Or _
            Trim(txtClnInsType) = "F" Or Trim(txtClnInsType) = "H" Then
            SSTab1.Tab = 1
            cboCovSalutation.Enabled = True
        Else
            MsgBox "This Plan does not have supplementary", vbOKOnly + vbCritical, DialogBoxTitle
            SSTab1.Tab = 0
        End If
    ElseIf SSTab1.Tab = 2 Then
        lstAdjustmentHis.ListItems.Clear
        If lstAdjustmentHis.ListItems.Count = 0 Then
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call ADJListing
            'medixmasterconn.Close
            'Set medixmasterconn = Nothing
        End If
        lstAdjustmentHis.SetFocus
    ElseIf SSTab1.Tab = 4 Then
        If lstMBMHis.ListItems.Count = 0 Then
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call MBMHISListing
           ' medixmasterconn.Close
           ' Set medixmasterconn = Nothing
        End If
        
        
    'ElseIf SSTab1.Tab = 5 Then
    '
    '        If Trim(txtMgmtNotes) = "" And txtMBMNumber <> "" Then
    '            medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    '            medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"'

                'If Len(Trim(txtMBMIcBcPP)) > 4 Then
                '    'If Mid(cboPatientList, 1, 2) = "00" Then
                '        Call MBMHistory("Princ", True)
                '    'Else
                '    '    Call MBMHistory("Supp", True)
                '    'End If
                'Else
                    'If Mid(cboPatientList, 1, 2) = "00" Then
                '        Call MBMHistory("Princ", False)
                '    'Else
                '    '    Call MBMHistory("Supp", False)
                '    'End If
                'End If
                
               ' medixmasterconnMaint.Close
               ' Set medixmasterconnMaint = Nothing
               ' medixmasterconn.Close
               ' Set medixmasterconn = Nothing
                
               ' txtMgmtNotes = "'" & tmpPolHistory & "'"
                'txtMgmtNotes = tmpRemarks
            'End If
        
    End If
Screen.MousePointer = Default
End Sub

Private Sub MBMHISListing()
Dim MBMHIS As New ADODB.Recordset
Dim HISList As ListItem
Dim Cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim tmpName As String
Dim strAppend As String
Dim tmpDOb As String

    If lstMBMHis.ListItems.Count = 0 Then
        tmpName = Trim(Replace(txtMBMName, "'", "''"))
        tmpDOb = Format(txtMBMBirthDate, "yyyy/mm/dd")
        If Len(Trim(txtMBMIcBcPP)) > 4 Then
            'strAppend = " And HISMaintenance.dbo.MBMCrossReference.MBMIcBcPp = '" & Trim(txtMBMIcBcPP) & "'"
            strAppend = " And Replace(HISMaintenance.dbo.MBMCrossReference.MBMIcBcPp,'-','') = Replace('" & Trim(txtMBMIcBcPP) & "','-','')"
        Else
            strAppend = " And HISMaintenance.dbo.MBMCrossReference.MBMName = '" & tmpName
            strAppend = strAppend & "' And HISMaintenance.dbo.MBMCrossReference.MBMDOB='" & tmpDOb & "'"
        End If
        Set Cmd.ActiveConnection = medixmasterconn
        Cmd.CommandText = "SearchMBMHistory06"
        Cmd.CommandType = adCmdStoredProc
        Cmd.CommandTimeout = 300
        Set Prm1 = Cmd.CreateParameter("MBMType", adVarChar, adParamInput, 10, "Princ")
        Cmd.Parameters.Append Prm1
        Set Prm2 = Cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        Cmd.Parameters.Append Prm2
    
        MBMHIS.CursorLocation = adUseClient
        MBMHIS.CursorType = adOpenForwardOnly
        MBMHIS.CacheSize = 100
        Set MBMHIS = Cmd.Execute
        Do While Not MBMHIS.EOF
            Set HISList = lstMBMHis.ListItems.Add(, , MBMHIS!MBMNumber)
                HISList.SubItems(1) = IIf(Len(MBMHIS!MBMTopupInd), MBMHIS!MBMTopupInd, "")
                HISList.SubItems(2) = IIf(Len(MBMHIS!MBMTopupNumber), MBMHIS!MBMTopupNumber, "")
                HISList.SubItems(3) = IIf(Len(MBMHIS!MBMPolicyNo), MBMHIS!MBMPolicyNo, "")
                HISList.SubItems(4) = IIf(MBMHIS!MBMPayorEffDate <> "", MBMHIS!MBMPayorEffDate, "")
                HISList.SubItems(5) = IIf(MBMHIS!MBMPayorEXPDate <> "", MBMHIS!MBMPayorEXPDate, "")
                If MBMHIS!MBMStatus = "C" Then
                    HISList.SubItems(6) = IIf(Len(MBMHIS!MBMENDtEffDate), MBMHIS!MBMENDtEffDate, "")
                End If
                HISList.SubItems(7) = IIf(Len(MBMHIS!MBMDataStatus), MBMHIS!MBMDataStatus, "")
                HISList.SubItems(8) = IIf(Len(MBMHIS!MBMStatus), MBMHIS!MBMStatus, "")
                HISList.SubItems(9) = IIf(Len(MBMHIS!MBMBordxDate), MBMHIS!MBMBordxDate, "")
                HISList.SubItems(10) = IIf(Len(MBMHIS!MBMEndtBordxDate), MBMHIS!MBMEndtBordxDate, "")
                HISList.SubItems(11) = IIf(Len(MBMHIS!MBMValueDate), MBMHIS!MBMValueDate, "")
            MBMHIS.MoveNext
        Loop
        MBMHIS.Close
        Set MBMHIS = Nothing
    End If
End Sub
Private Sub SSTab1_LostFocus()
    If SSTab1.Tab = 0 Then
        If txtMBMAdd1.Enabled = True Then
            txtMBMAdd1.SetFocus
        End If
    ElseIf SSTab1.Tab = 1 Then
        'If cboMBMOccupation.Enabled = True Then
            'cboMBMOccupation.SetFocus
        'End If
    ElseIf SSTab1.Tab = 2 Then
        If cboCovSalutation.Enabled = True Then
            cboCovSalutation.SetFocus
        End If
    Else
        If lstAdjustmentHis.Enabled = True Then
            lstAdjustmentHis.SetFocus
        End If
    End If
End Sub

'
Private Sub SSTab1_GotFocus()
    If SSTab1.Tab = 0 Then
        'txtMBMAdd1.SetFocus
    End If
End Sub

Private Sub cmdExit_LostFocus()
    If Me.ActiveControl.TabIndex = 44 Then
        txtMBMNumber.SetFocus
    End If
End Sub
' ######################### START BUTTON CLICKING ###########


Public Sub cmdShow_Click()
Dim MBMNumber As String
Dim tmpREcCnt As Integer

Call ClearForm(frmNewMBMOthers)
'frmNewMBMOthers.

    SSTab1.Tab = 0
    tmpUpdateAnnLmt = False
    cmdSave.Enabled = False
    txtDate.mask = ""
    txtDate.mask = "__/__/____"
    txtDate.mask = "##/##/####"
    txtEndtBordxDate.mask = ""
    txtEndtBordxDate.mask = "__/__/____"
    txtEndtBordxDate.mask = "##/##/####"
    'getRecord = showMBM
    Screen.MousePointer = vbHourglass
    MBMNumber = txtMBMNumber
    Call ClearEntries
    txtMBMNumber = MBMNumber
    ' ================ Add by CK =============================
    
    If txtMBMNumber = "" Then
        MsgBox "Please Enter Membership No", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMNumber.SetFocus
    Else
        ScanCode = Trim(txtMBMNumber)
        If Mid(ScanCode, 1, 2) = "LO" Then
            txtMBMNumber = ""
            tmpREcCnt = 2
            Call SearchForm(ScanCode)
        ElseIf Len(txtMBMNumber) = 8 Or Len(txtMBMNumber) = 9 Or (Len(Trim(txtMBMNumber)) = 11 And Mid(Trim(txtMBMNumber), 9, 1) <> "*") Then
            tmpREcCnt = 1
            frmSearchMBM05.Source = 2
            frmSearchMBM05.Show
            tmpREcCnt = frmSearchMBM05.SearchRecCnt
            If frmSearchMBM05.SearchRecCnt < 1 Then
                Unload frmSearchMBM05
            ElseIf frmSearchMBM05.SearchRecCnt = 1 Then
                txtMBMNumber = frmSearchMBM05.tmpMBMNo
                Unload frmSearchMBM05
            End If
        End If
        If tmpREcCnt < 2 Then
            'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
            Call FindMBM
           ' medixmasterconn.Close
           ' Set medixmasterconn = Nothing
        End If

    End If
    ' ================ Add by CK =============================
    
    Screen.MousePointer = Default
End Sub

Private Function FindMBM() As Integer
Dim strCount As String
Dim sql As String
Dim rst2 As New ADODB.Recordset
Dim getRecord As Integer
Dim MBMClnRec As New ADODB.Recordset
Dim MBMClnSql As String
Dim TmpCovID As String
    
    
    MBMClnSql = "Select PLNType, INSType, MBMPolEffDate, MBMPolExpDate from MBMCLNOthers where MBMNumber='" & Trim(txtMBMNumber) & "'"
    MBMClnRec.Open MBMClnSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If MBMClnRec.recordCount Then
        txtClnInsType = Trim(MBMClnRec!InsType)
        txtClnPlanCode = Trim(MBMClnRec!PLNType)
        CisEffDate = MBMClnRec!MBMPolEffDate
        CisExpDate = MBMClnRec!MBMPolExpDate
    End If
    
    MBMClnRec.Close
    Set MBMClnRec = Nothing
    
    sql = "Exec SearchMBMCovPersonsProc '" & Trim(txtMBMNumber) & "'"
    Set ADOrstMBM = New ADODB.Recordset
    ADOrstMBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    TmpCovID = ""
    If Not ADOrstMBM.EOF Then
        cmdEditMBM.Enabled = True
        TmpCovID = IIf(Len(ADOrstMBM!MBMCCoverID), Trim(ADOrstMBM!MBMCCoverID), "")
        getRecord = showMBM
        If Trim(ADOrstMBM!INSCode) = "F" Or Trim(ADOrstMBM!INSCode) = "H" Or _
        Trim(txtClnInsType) = "F" Or Trim(txtClnInsType) = "H" Then
            If Len(TmpCovID) And TmpCovID <> "00" Then
                Call showMBMCov
            End If
        End If
    Call ShowReinstatement

    Else
        MsgBox "No record found!", vbCritical
    End If
    ADOrstMBM.Close
    Set ADOrstMBM = Nothing
    HisCisEffDate = txtMBMPayorEffDate
    HISCISExpDate = txtMBMPayorExpDate
    If Trim(cboPLNCode.Text) = "" Then
        HisCisEffDate = CisEffDate
        HISCISExpDate = CisExpDate
    End If


    cmdEditMBM.Enabled = True
    If Mid(cboStatus, 1, 1) = "D" Then
        MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdEditMBM.Enabled = False
        txtMBMNumber.SetFocus
    'ElseIf Mid(cboStatus, 1, 1) = "C" Then
     '   MsgBox "Membership has been cancelled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
      '  cmdEditMBM.Enabled = False
       ' txtMBMNumber.SetFocus
    ElseIf Mid(cboStatus, 1, 1) = "S" Then
        MsgBox "Membership has been Suspended. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
        cmdEditMBM.Enabled = False
        txtMBMNumber.SetFocus
    End If
    Call EnabledFrame(False)
End Function

Private Function showMBM() As Integer
Dim MBM As New ADODB.Recordset
Dim sqlstqMBMRef As String
Dim MBMRef As New ADODB.Recordset
Dim MBMNumber As String
Dim AnnLmt As New ADODB.Recordset
Dim StrSqlAnnLmt, strsql As String
Dim AnnLmtType As New ADODB.Recordset
Dim tmpLmt As String
Dim tmpAvBalLmt As String
Dim TmpEffDate As String
Dim TmpExpDate As String
Dim tmpBPrem As String
Dim tmpPrem As String
Dim tmpBMco As String
Dim tmpMCO As String
Dim tmpInvNo As String
Dim TmpInvDate As String
Dim tmpStatus As String
    tmpPerDisabilitAmt = 0
    cboEndorseStatus.ListIndex = 0
    showMBM = 1
    With ADOrstMBM
        PrevRenewKey1 = ""
        If InStr(1, txtMBMNumber, "*") <> 0 Then
            If Trim(!PLNCode) <> "" And Trim(!INSCode) <> "" Then
                PrevRenewKey1 = Trim(!PAYCode) & "*" & Trim(!INSCode) & "*" & Format(!MBMDOB, "YYYYMMDD") & "*" & Trim(!MBMName)
            End If
        End If
        txtMBMName = Trim(!MBMName)
        txtBatchNo = IIf(Len(!MBMBatchNo), !MBMBatchNo, "")
        cboHLTCode = !HLTCode
        LblHnS = IIf(Len(!HSCLNInd), !HSCLNInd, "H")
        txtMBMIcBcPP = IIf(Len(!MBMIcBcPp), !MBMIcBcPp, "")
        txtMBMPolNo = IIf(Len(!MBMPolicyNo), !MBMPolicyNo, "")
        cboINSCode = IIf(Len(!INSCode), !INSCode, "")
        cboPLNCode = IIf(Len(!PLNCode), !PLNCode, "")
        
        If (!MBMEndEffInd) Then
            ChkEffInd.Value = 1
        Else
            ChkEffInd.Value = 0
        End If
        
        If Trim(!MBMTopupInd) = "Y" Then
            lblTopupStatus.Visible = True
        End If
        
        SptLmtInd = "A"
        tmpLmt = 0
        StrSqlAnnLmt = "Select AnnLimit, DisabilityLmt, SuppLimitStatus from AnnualLimit where PLNCode = '" & Trim(!PLNCode) & _
                   "' AND INSCode='" & Trim(!INSCode) & _
                   "' And HLTCode = '" & Trim(!HLTCode) & _
                   "' AND AnnualEffDate <= '" & Format(!MBMPayorEffDate, "mm/dd/yyyy") & "' order by AnnualEffDate  desc"
        AnnLmt.Open StrSqlAnnLmt, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        If AnnLmt.EOF = False Then
            SptLmtInd = IIf(Len(Trim(AnnLmt!SuppLimitStatus)), Trim(AnnLmt!SuppLimitStatus), "A")
            tmpPerDisabilitAmt = IIf(IsNumeric(AnnLmt!DisabilityLmt), Trim(AnnLmt!DisabilityLmt), 0)
            tmpLmt = IIf(IsNumeric(AnnLmt!AnnLimit), Format(AnnLmt!AnnLimit, "#,##0.00"), 0)
        End If
        AnnLmt.Close
        Set AnnLmt = Nothing
        lblLimitType = ""
        If Trim(cboINSCode) <> "" Then
            If Mid(cboINSCode, 1, 1) = "F" Or Mid(cboINSCode, 1, 1) = "H" Then
                lblLimitType = SptLmtInd
                strsql = "Select LimitDesc from AnnualLimitType where LimitCode='" & Trim(SptLmtInd) & " '"
                AnnLmtType.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                If AnnLmtType.EOF = False Then
                    lblLimitType = SptLmtInd & " - " & AnnLmtType!LimitDesc
                End If
            End If
        End If
        txtMBMAdd1 = IIf(Len(!MBMAddress1), !MBMAddress1, "")
        txtMBMAdd2 = IIf(Len(!MBMAddress2), !MBMAddress2, "")
        txtMBMAdd3 = IIf(Len(!MBMAddress3), !MBMAddress3, "")
        txtMBMBatchNo = IIf(Len(!MBMBatchNo), !MBMBatchNo, "")
        txtMBMPostCode = IIf(Len(!MBMPostCode), !MBMPostCode, "")
        cboCTYCode = IIf(Len(!CTYCode), !CTYCode, "")
        cboMBMRenewal = IIf(Len(!MBMRenewFlag), !MBMRenewFlag, "")
        cboMBMSex = IIf(Len(!MBMSex), !MBMSex, "")
        cboBRNCode = IIf(Len(!MBMBranch), !MBMBranch, "")
        cboRACCode = IIf(Len(!RACCode), !RACCode, "")
        cboSTACode = IIf(Len(!STACode), !STACode, "")
        cboHLTCode = IIf(Len(!HLTCode), !HLTCode, "")
        cboPAYCode = IIf(Len(!PAYCode), !PAYCode, "")
        txtAGECode = IIf(Len(!AGECode), !AGECode, "")

        If Trim(!MBMDOB) <> "" Then
            txtMBMBirthDate = Format(!MBMDOB, "dd/mm/yyyy")
        End If
        If Trim(!MBMTakeover) = "C" Then
            cboMBMTakeOver.ListIndex = 2
        ElseIf Trim(!MBMTakeover) = "Y" Then
            cboMBMTakeOver.ListIndex = 1
        Else
            cboMBMTakeOver.ListIndex = 0
        End If
        If Trim(!MBMBordxDate) <> "" Then
            txtMBMBordxDate = Format(!MBMBordxDate, "dd/mm/yyyy")
        End If
        If Trim(!MBMPayorEffDate) <> "" Then
            txtMBMPayorEffDate = Format(!MBMPayorEffDate, "dd/mm/yyyy")
        End If
        If Trim(!MBMPayorEXPDate) <> "" Then
            txtMBMPayorExpDate = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
        End If
        
        tmpPAYGRPCode = IIf(Len(!PAYGRPCode), !PAYGRPCode, "")
        
        'If !PAYGRPCode = "ETS" Then
        '    Label16(1).Caption = "Next Due Date"
            If Trim(!MBMNextDueDate) <> "" Then
                txtNxDueDate = Format(!MBMNextDueDate, "dd/mm/yyyy")
            End If
        'Else
        '    Label16(1).Caption = "Final Exp Date"
            If Trim(!MBMFinalExpiryDate) <> "" Then
                txtFinalExpDate = Format(!MBMFinalExpiryDate, "dd/mm/yyyy")
            Else
                If Trim(!MBMPayorEXPDate) <> "" Then
                    txtFinalExpDate = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                End If
            End If
        'End If
        
        If Len(!MBMPolicyYear) Then
            txtMBMPolicyYear = Trim(!MBMPolicyYear)
        End If
        cboDataStatus = IIf(Len(!MBMDataStatus), Trim(!MBMDataStatus), "")
        tmpAvBalLmt = IIf(IsNumeric(!MBMAvailableLimit), Format(Trim(!MBMAvailableLimit), "#,##0.00"), 0)
        txtOriAnnLmt = IIf(IsNumeric(!MBMAvailableLimit), Format(Trim(!MBMAvailableLimit), "#,##0.00"), 0)
        txtMBMAdultChild = IIf(Len(!MBMAdultChild), Trim(!MBMAdultChild), "")
        If (!MBMValueDate) <> "" Then
            txtDOR = Format(!MBMValueDate, "dd/mm/yyyy")
        End If
        If Trim(!MBMStatus) = "A" Then
            cboStatus.ListIndex = 0
        ElseIf Trim(!MBMStatus) = "C" Then
            cboStatus.ListIndex = 1
        ElseIf Trim(!MBMStatus) = "D" Then
            cboStatus.ListIndex = 2
        ElseIf Trim(!MBMStatus) = "S" Then
            cboStatus.ListIndex = 3
        ElseIf Trim(!MBMStatus) = "P" Then
            cboStatus.ListIndex = 4
        End If
        tmpStatus = !MBMStatus
        cboExpStatus.ListIndex = IIf(Trim(!MBMExpiredStatus) = "N", 0, 1)
        tmpPrem = IIf(IsNumeric(!MBMPremium), !MBMPremium, 0)
        tmpBPrem = IIf(IsNumeric(!MBMBasicPrem), !MBMBasicPrem, 0)
        tmpMCO = IIf(IsNumeric(!MBMMCOFee), !MBMMCOFee, 0)
        tmpBMco = IIf(IsNumeric(!MBMBasicMCO), !MBMBasicMCO, 0)
        tmpInvNo = IIf(Len(!MBMInvoiceNo), !MBMInvoiceNo, "")
        TmpInvDate = IIf(Len(!MBMInvoiceDate), !MBMInvoiceDate, "")
        tmpCRNote = IIf(Len(!MBMCRNote), !MBMCRNote, "")
        tmpCRDate = IIf(Len(!MBMCRNoteDate), !MBMCRNoteDate, "")
        TmpEffDate = IIf(Len(!MBMPayorEffDate), !MBMPayorEffDate, "")
        TmpExpDate = IIf(Len(!MBMPayorEXPDate), !MBMPayorEXPDate, "")
        tmpCancelDate = IIf(Len(!MBMCancelDate), !MBMCancelDate, "")
        Call FeesListing(txtMBMNumber, txtMBMName, "P", "00", tmpStatus, cboPLNCode, _
        TmpEffDate, TmpExpDate, tmpCancelDate, _
        tmpLmt, tmpAvBalLmt, tmpBPrem, tmpPrem, tmpBMco, tmpMCO, _
        tmpInvNo, TmpInvDate, tmpCRNote, tmpCRDate)


    End With
    cmdEditMBM.Enabled = True
    Call showMBMTwo
    Call showMBMOthers
    Call showMBMOthersTwo
End Function

Private Sub ShowReinstatement()
Dim ListCount As Integer
Dim ForLoop As Integer
Dim FeesList As ListItem
Dim tmpName, tmpRELCode, TmpCovID, tmpStatus, tmpPlnCode, TmpEffDate, TmpExpDate, tmpCancelDate As String
Dim tmpAnnLmt, tmpAvLmt, tmpPerDisabilitAmt As Double
Dim strsqlRein As String
Dim rein As New ADODB.Recordset
Dim tmpAdd As Integer

    ForLoop = 0
    ListCount = 0
    tmpAdd = 0
    
    ListCount = lstFeesView.ListItems.Count
    For ForLoop = 1 To ListCount
        tmpName = lstFeesView.ListItems(ForLoop).SubItems(1)
        tmpRELCode = lstFeesView.ListItems(ForLoop).SubItems(2)
        TmpCovID = lstFeesView.ListItems(ForLoop).SubItems(3)
        tmpStatus = lstFeesView.ListItems(ForLoop).SubItems(4)
        tmpPlnCode = lstFeesView.ListItems(ForLoop).SubItems(5)
        TmpEffDate = lstFeesView.ListItems(ForLoop).SubItems(7)
        TmpExpDate = lstFeesView.ListItems(ForLoop).SubItems(8)
        tmpCancelDate = lstFeesView.ListItems(ForLoop).SubItems(9)
        tmpAnnLmt = 0
        tmpAvLmt = 0
        tmpPerDisabilitAmt = 0
            
        strsqlRein = ""
        strsqlRein = "Select MBMMCOFee,MBMPremium,MBMInvoiceNo,MBMInvoiceDate from MBMReinstatement where " & _
                    " MBMNumber = '" & txtMBMNumber & "' AND MBMCoverID= '" & TmpCovID & "' "
        rein.Open strsqlRein, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            Do While Not rein.EOF
                Set FeesList = lstFeesView.ListItems.Add(, , "")
                FeesList.SubItems(1) = tmpName
                FeesList.SubItems(2) = tmpRELCode
                FeesList.SubItems(3) = TmpCovID
                FeesList.SubItems(4) = tmpStatus
                FeesList.SubItems(5) = IIf(Len(tmpPlnCode), tmpPlnCode, "")
                FeesList.SubItems(7) = IIf(Len(TmpEffDate), Format(TmpEffDate, "dd/mm/yyyy"), "")
                FeesList.SubItems(8) = IIf(Len(TmpExpDate), Format(TmpExpDate, "dd/mm/yyyy"), "")
                'FeesList.SubItems(9) = IIf(Len(tmpCancelDate), Format(tmpCancelDate, "dd/mm/yyyy"), "")
                'FeesList.SubItems(10) = tmpAnnLmt
                'FeesList.SubItems(11) = IIf(IsNumeric(tmpAvLmt), tmpAvLmt, 0)
                If SptLmtInd = "A" And TmpCovID = "00" Or SptLmtInd = "B" And TmpCovID = "01" Or SptLmtInd = "C" Then
                    FeesList.SubItems(12) = IIf(IsNumeric(tmpPerDisabilitAmt), tmpPerDisabilitAmt, 0)
                End If
                FeesList.SubItems(13) = 0
                FeesList.SubItems(14) = 0
                FeesList.SubItems(15) = 0
                FeesList.SubItems(16) = 0
                FeesList.SubItems(17) = IIf(Len(rein!MBMInvoiceNo), rein!MBMInvoiceNo, "")
                FeesList.SubItems(18) = IIf(Len(rein!MBMInvoiceDate), rein!MBMInvoiceDate, "")
                FeesList.SubItems(19) = ""
                FeesList.SubItems(20) = ""
                FeesList.SubItems(25) = IIf(IsNumeric(rein!MBMMCOFee), Format(Round(rein!MBMMCOFee, 0)), "#,##0.00")
                FeesList.SubItems(26) = Format(0, "#,##0.00")
                FeesList.SubItems(27) = IIf(IsNumeric(rein!MBMPremium), Format(Round(rein!MBMPremium, 0)), "#,##0.00")
                FeesList.SubItems(28) = Format(0, "#,##0.00")
            rein.MoveNext
            Loop
        rein.Close
        Set rein = Nothing
    Next
    

End Sub

Private Sub showMBMCov()
Dim MBMCov As New ADODB.Recordset
Dim sqlstr As String
Dim MBMCovTwo As New ADODB.Recordset
Dim sqlstrTwo As String
Dim CovTwo As Boolean
Dim tmpLmt As String
Dim tmpAvLmt As String
Dim TmpEffDate As String
Dim TmpExpDate As String
Dim tmpCancelDate As String
Dim tmpBPrem As String
Dim tmpBMco As String
Dim tmpPrem As String
Dim tmpMCO As String
Dim tmpInvNo As String
Dim TmpInvDate As String
Dim MBMCovList As ListItem
Dim tmpCStatus As String
Dim tmpPln As String
Dim tmpName As String
Dim tmpRELCode As String
Dim TmpCovID As String
Dim ListCount As Integer
Dim ForLoop As Integer
Dim strsqlRein As String
Dim rein As New ADODB.Recordset
Dim FeesList As ListItem
    lstCovAdd.ListItems.Clear
        Do While Not ADOrstMBM.EOF
        With ADOrstMBM
            tmpName = IIf(Trim(!MBMCName) <> "", !MBMCName, "")
            Set MBMCovList = lstCovAdd.ListItems.Add(, , IIf(Trim(!MBMCName) <> "", !MBMCName, ""))
            If IsNull(!MBMRetireDate) Then
                txtRtDate = "__/__/____"
            ElseIf Len(!MBMRetireDate) Then
                txtRtDate = !MBMRetireDate
            Else
                txtRtDate = "__/__/____"
            End If
                'txtRtDate = IIf(Len(!MBMRetireDate), !MBMRetireDate, "")
                tmpCStatus = IIf(Len(!MBMCStatus), Trim(!MBMCStatus), "")
                tmpRELCode = IIf(Len(!MBMCRELCode), Trim(!MBMCRELCode), "")
                tmpPln = IIf(Len(!MBMCPLNCode), Trim(!MBMCPLNCode), cboPLNCode)
                tmpLmt = Format(IIf(IsNumeric(!MBMCAnnualLimit), !MBMCAnnualLimit, 0), "#,##0.00")
                TmpEffDate = IIf(Len(!MBMCEffDate), !MBMCEffDate, txtMBMPayorEffDate)
                TmpExpDate = IIf(Len(!MBMCExpDate), !MBMCExpDate, txtMBMPayorExpDate)
                'If Trim(!MBMStatus) = "C" Then
                If !MBMCCancelDate <> "" Then
                    tmpCancelDate = !MBMCCancelDate
                Else
                     tmpCancelDate = ""
                End If
                tmpBPrem = IIf(IsNumeric(!MBMCBasicPremium), !MBMCBasicPremium, 0)
                tmpPrem = IIf(IsNumeric(!MBMCPremium), !MBMCPremium, 0)
                tmpBMco = IIf(IsNumeric(!MBMCBasicMCO), Trim(!MBMCBasicMCO), 0)
                tmpMCO = IIf(IsNumeric(!MBMCMCO), Trim(!MBMCMCO), 0)
                tmpInvNo = IIf(Len(!MBMCInvoiceNo), !MBMCInvoiceNo, "")
                TmpInvDate = IIf(Len(!MBMCInvoiceDate), !MBMCInvoiceDate, "")
                tmpAvLmt = Format(IIf(IsNumeric(!MBMCAvailableLimit), !MBMCAvailableLimit, 0), "#,##0.00")
                TmpCovID = IIf(Len(!MBMCCoverID), Trim(!MBMCCoverID), "")
                MBMCovList.SubItems(1) = IIf(Len(!MBMCSalutation), Trim(!MBMCSalutation), "")
                MBMCovList.SubItems(2) = IIf(Len(!MBMCICBCPPNo), Trim(!MBMCICBCPPNo), "")
                MBMCovList.SubItems(3) = IIf(Len(!MBMCDOB), Format(!MBMCDOB, "dd/mm/yyyy"), "")
                MBMCovList.SubItems(4) = IIf(Len(!MBMCSex), Trim(!MBMCSex), "")
                MBMCovList.SubItems(5) = IIf(Len(!MBMCAdultChild), Trim(!MBMCAdultChild), "")
                MBMCovList.SubItems(6) = tmpRELCode
                MBMCovList.SubItems(7) = tmpPln
                MBMCovList.SubItems(8) = tmpCStatus 'Dependent Status change to Editable - By CHEng 3 Dec 2016
                MBMCovList.SubItems(9) = IIf(Len(!MBMCExclusion), Trim(!MBMCExclusion), "")
                MBMCovList.SubItems(10) = TmpCovID
                MBMCovList.SubItems(12) = tmpLmt
                MBMCovList.SubItems(13) = TmpEffDate
                MBMCovList.SubItems(14) = TmpExpDate
                MBMCovList.SubItems(15) = IIf(IsNumeric(!MBMCUndExcess), Trim(!MBMCUndExcess), 0)
                MBMCovList.SubItems(16) = tmpMCO
                MBMCovList.SubItems(17) = tmpAvLmt
                MBMCovList.SubItems(18) = tmpPrem
                MBMCovList.SubItems(19) = IIf(Len(!MBMCAllergic), Trim(!MBMCAllergic), "")
                MBMCovList.SubItems(20) = IIf(Len(!MBMCPrevPLNCode), Trim(!MBMCPrevPLNCode), "")
                MBMCovList.SubItems(21) = IIf(IsNumeric(!MBMCRBAmt), !MBMCRBAmt, 0)
                MBMCovList.SubItems(22) = IIf(Len(!MBMCPAYRemarks), Trim(!MBMCPAYRemarks), "")
                MBMCovList.SubItems(23) = IIf(Len(!MBMCPrevInsCom), Trim(!MBMCPrevInsCom), "")

                If Len(!MBMCGenWP) Then
                    MBMCovList.SubItems(24) = !MBMCGenWP
                End If
                If Len(!MBMCGenEndWP) Then
                    MBMCovList.SubItems(25) = !MBMCGenEndWP
                End If
                If Len(!MBMCPreExWP) Then
                    MBMCovList.SubItems(26) = !MBMCPreExWP
                End If
                If Len(!MBMCPreExEndWP) Then
                    MBMCovList.SubItems(27) = !MBMCPreExEndWP
                End If
                If Len(!MBMCSpeIllWP) Then
                    MBMCovList.SubItems(28) = !MBMCSpeIllWP
                End If
                If Len(!MBMSpeIllEndWP) Then
                    MBMCovList.SubItems(29) = !MBMSpeIllEndWP
                End If
                If Len(!MBMCStaffDec) Then
                   MBMCovList.SubItems(30) = !MBMCStaffDec
                End If
                If Len(!MBMCCLMNonPayFr) Then
                    MBMCovList.SubItems(31) = !MBMCCLMNonPayFr
                End If
                If Len(!MBMCCLMNonPayTo) Then
                    MBMCovList.SubItems(32) = !MBMCCLMNonPayTo
                End If
                MBMCovList.SubItems(34) = IIf(IsNumeric(!MBMCPayorPremGross), !MBMCPayorPremGross, 0)
                MBMCovList.SubItems(35) = IIf(IsNumeric(!MBMCPayorPremNet), !MBMCPayorPremNet, 0)
                MBMCovList.SubItems(36) = IIf(IsNumeric(!MBMCPayorMCOGross), !MBMCPayorMCOGross, 0)
                MBMCovList.SubItems(37) = IIf(IsNumeric(!MBMCPayorMCONet), !MBMCPayorMCONet, 0)
                MBMCovList.SubItems(38) = IIf(IsNumeric(!MBMCRenewalDisc), !MBMCRenewalDisc, 0)
                MBMCovList.SubItems(39) = IIf(IsNumeric(!MBMCPolicyDisc), !MBMCPolicyDisc, 0)
                MBMCovList.SubItems(40) = IIf(IsNumeric(!MBMCFamDiscAmt), !MBMCFamDiscAmt, 0)
                MBMCovList.SubItems(41) = IIf(IsNumeric(!MBMCRenewDiscAmt), !MBMCRenewDiscAmt, 0)
                MBMCovList.SubItems(42) = IIf(IsNumeric(!MBMCLoadingPrem), !MBMCLoadingPrem, 0)
                MBMCovList.SubItems(43) = IIf(Len(!MBMCPayNo), Trim(!MBMCPayNo), "")
                MBMCovList.SubItems(44) = IIf(Len(!MBMCPaySeqNo), Trim(!MBMCPaySeqNo), "")
                MBMCovList.SubItems(45) = IIf(Len(!MBMCClientNo), Trim(!MBMCClientNo), "")
                MBMCovList.SubItems(46) = IIf(Len(!MBMCClientID), Trim(!MBMCClientID), "")
                MBMCovList.SubItems(47) = IIf(Len(!MBMCIcBcPp2nd), Trim(!MBMCIcBcPp2nd), "")
                'Add Student/OKU Indicator to cover > 21yrs old - By CHEng 1 Dec 2016'
                MBMCovList.SubItems(48) = IIf(Len(!MBMHRAppInd), Trim(!MBMHRAppInd), "")
                            
                    Call FeesListing(txtMBMNumber, tmpName, tmpRELCode, TmpCovID, tmpCStatus, tmpPln, _
                    TmpEffDate, TmpExpDate, tmpCancelDate, _
                    tmpLmt, tmpAvLmt, tmpBPrem, tmpPrem, tmpBMco, tmpMCO, _
                    tmpInvNo, TmpInvDate, tmpCRNote, tmpCRDate)


            End With
            ADOrstMBM.MoveNext
        Loop
End Sub

Private Sub cmdEditMBM_Click()
    EnableEntries (True)
    cboINSCode.Enabled = False
    'cboPLNCode.Enabled = False
    Frame10.Enabled = True
    SSTab1.Enabled = True
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    cmdRemove.Enabled = False
    cboEndorseStatus.Enabled = True
    cboEndorseStatus.SetFocus
    'Dependent Expired Date change to Editable - by CHEng 2 Dec 2016
    txtSuppExpDate.Enabled = True
End Sub

Private Sub CmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
    If frmNewMBMOthers.Visible = False Then
        Load frmNewMBMOthers
    End If
    
    Call ClearForm(frmNewMBMOthers)
    Unload Me
'    End If
End Sub

Private Sub cmdSearchMBM_Click()
    lblTopupStatus.Visible = False
    
    Call ClearEntries
    frmSearchMBM.Source = 2
    frmSearchMBM.Show
    
End Sub

Private Sub cmdAdd_Click()
    Dim MemberConveredPerson As ListItem
    Dim RecordSaved

    If txtCovName = "" Then
        MsgBox "Please Enter Covered Person Name!", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf txtCovIcBcPp = "" Then
        MsgBox "Please Enter Covered Person Identification!", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboCovRelationship = "" Then
        MsgBox "Please Enter Covered Person Relationship!", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf IsDate(txtCovBirthdate) = False Then
        MsgBox "Please Enter Covered Person Birth Date!", vbOKOnly + vbCritical, DialogBoxTitle
    'ElseIf txtCovBirthdate > Date Then
        'MsgBox "Birth Date can not greater than System Date!", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboCovSex = "" Then
        MsgBox "Please Enter Covered Person Sex!", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf Trim(txtEndtBordxDate) = "__/__/____" Then
        MsgBox "Please Enter Endt Bordx Date !", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            Set MemberConveredPerson = lstCovAdd.ListItems.Add(, , txtCovName)
            MemberConveredPerson.SubItems(1) = cboCovSalutation
            MemberConveredPerson.SubItems(2) = txtCovIcBcPp
            MemberConveredPerson.SubItems(3) = txtCovBirthdate
            MemberConveredPerson.SubItems(4) = cboCovSex
            MemberConveredPerson.SubItems(5) = cboCovAdultChild
            MemberConveredPerson.SubItems(6) = cboCovRelationship
            MemberConveredPerson.SubItems(19) = txtCovAllergic
            MemberConveredPerson.SubItems(7) = IIf(Len(txtSuppPLNCode), Trim(txtSuppPLNCode), cboPLNCode)
            MemberConveredPerson.SubItems(9) = txtCovExclusion
            MemberConveredPerson.SubItems(11) = "N"
            MemberConveredPerson.SubItems(12) = txtCovAnnualLimit
            MemberConveredPerson.SubItems(13) = IIf(IsDate(txtSuppEffDate), txtSuppEffDate, txtMBMPayorEffDate)
            MemberConveredPerson.SubItems(14) = IIf(IsDate(txtSuppExpDate), txtSuppExpDate, txtMBMPayorExpDate)
            MemberConveredPerson.SubItems(15) = IIf(IsNumeric(txtCovUndExcess), txtCovUndExcess, 0)
            MemberConveredPerson.SubItems(20) = IIf(Len(txtPrevPLNCode), Trim(txtPrevPLNCode), "")
            MemberConveredPerson.SubItems(21) = IIf(IsNumeric(txtRBAmt), txtRBAmt, 0)
            MemberConveredPerson.SubItems(22) = IIf(Len(txtCOVPayRemarks), Trim(txtCOVPayRemarks), "")
            MemberConveredPerson.SubItems(23) = IIf(Len(txtPrevINSCom), Trim(txtPrevINSCom), "")
            If TxtSuppClmNonPayFr <> "__/__/____" Then
                MemberConveredPerson.SubItems(31) = TxtSuppClmNonPayFr
            End If
            If TxtSuppClmNonPayTo <> "__/__/____" Then
                MemberConveredPerson.SubItems(32) = TxtSuppClmNonPayTo
            End If
            If txtEndtBordxDate <> "__/__/____" Then
                MemberConveredPerson.SubItems(33) = txtEndtBordxDate
            End If

            MemberConveredPerson.SubItems(34) = IIf(IsNumeric(txtSuppPayPremGross), txtSuppPayPremGross, 0)
            MemberConveredPerson.SubItems(35) = IIf(IsNumeric(txtSuppPayPremNet), txtSuppPayPremNet, 0)
            MemberConveredPerson.SubItems(36) = IIf(IsNumeric(txtSuppPayMCOGross), txtSuppPayMCOGross, 0)
            MemberConveredPerson.SubItems(37) = IIf(IsNumeric(txtSuppPayMCONet), txtSuppPayMCONet, 0)
            MemberConveredPerson.SubItems(38) = IIf(IsNumeric(txtRenewalDisc), txtRenewalDisc, 0)
            MemberConveredPerson.SubItems(39) = IIf(IsNumeric(txtPolicyDisc), txtPolicyDisc, 0)
            MemberConveredPerson.SubItems(40) = IIf(IsNumeric(txtPolicyDiscAmt), txtPolicyDiscAmt, 0)
            MemberConveredPerson.SubItems(41) = IIf(IsNumeric(txtRenewalDiscAmt), txtRenewalDiscAmt, 0)
            MemberConveredPerson.SubItems(42) = IIf(IsNumeric(txtLoadingPrem), txtLoadingPrem, 0)
            MemberConveredPerson.SubItems(43) = IIf(Len(txtPayMBMNo), Trim(txtPayMBMNo), "")
            MemberConveredPerson.SubItems(44) = IIf(Len(txtPaySeqID), Trim(txtPaySeqID), "")
            MemberConveredPerson.SubItems(45) = IIf(Len(txtClientNo), Trim(txtClientNo), "")
            MemberConveredPerson.SubItems(46) = IIf(Len(txtClientID), Trim(txtClientID), "")
            Call ClearCoveredEntry
            cboCovSalutation.SetFocus
        End If
    End If
End Sub

Private Sub CmdRemove_Click()
    Dim ForLoop As Integer
    Dim ListCount As Integer
    Dim strsql As String
    Dim CAS1 As New ADODB.Recordset
    Dim CAS2 As New ADODB.Recordset
    Dim tmpEndtStatus As String
    Dim RecordSaved
    Dim HISCASFound As Boolean
    Dim CISCASFound As Boolean
    
    tmpEndtStatus = Mid(cboEndorseStatus, 1, 1)
    ListCount = 0
    If txtCovName = "" Then
        MsgBox "Please Choose a Record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            ListCount = lstCovAdd.ListItems.Count
            For ForLoop = 1 To ListCount
            
                If lstCovAdd.ListItems(ForLoop).Selected = True Then
                    If Trim(lstCovAdd.SelectedItem.SubItems(10)) <> "" Then
                        HISCASFound = False
                        CISCASFound = False
                        strsql = "Select CASNumber from CASCrossReference Where MBMNumber ='" & txtMBMNumber & _
                        "' AND MBMPatientCovID ='" & Trim(lstCovAdd.SelectedItem.SubItems(10)) & "' And CASStatus = 'O'"
                       ' medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

                        CAS1.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
                        If CAS1.recordCount > 0 Then
                            HISCASFound = True
                        Else
                          '  MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"

                            strsql = "Select CASNumber from CISCas Where MBMNumber ='" & txtMBMNumber & _
                            "' AND CoverID ='" & Trim(lstCovAdd.SelectedItem.SubItems(10)) & "' And CASStatus = 'O'"
                            CAS2.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic
                            If CAS2.recordCount > 0 Then
                                CISCASFound = True
                            Else
                                lstCovAdd.SelectedItem.SubItems(11) = "R"
                            End If
                            CAS2.Close
                            Set CAS2 = Nothing
                            'MediConnCISDB.Close
                            'Set MediConnCISDB = Nothing
                        End If
                        CAS1.Close
                        Set CAS1 = Nothing
                      '  medixmasterconnMaint.Close
                      '  Set medixmasterconnMaint = Nothing
                        If HISCASFound = True Then
                            MsgBox "There is at least an outstanding HIS case under this Supplementary!", vbOKOnly + vbCritical, DialogBoxTitle
                        End If
                        If CISCASFound = True Then
                            MsgBox "There is at least an outstanding Clinical case under this Supplementary!", vbOKOnly + vbCritical, DialogBoxTitle
                        End If
                    Else
                        lstCovAdd.ListItems.Remove (ForLoop)
                        Exit For
                    End If
                End If
            Next ForLoop
        End If
    End If

End Sub

Private Sub cmdUpdate_Click()
Dim ForLoop As Integer

Dim RecordSaved
    If txtCovName = "" Then
       MsgBox "Please Enter Covered Person Name", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf txtCovIcBcPp = "" Then
       MsgBox "Please Enter Covered Person Identification", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboCovRelationship = "" Then
       MsgBox "Please Enter Covered Person Relationship", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            For ForLoop = 1 To lstCovAdd.ListItems.Count
                If lstCovAdd.ListItems(ForLoop).Selected = True Then
                        lstCovAdd.SelectedItem.Text = txtCovName
                        lstCovAdd.SelectedItem.SubItems(1) = Trim(cboCovSalutation)
                        lstCovAdd.SelectedItem.SubItems(2) = txtCovIcBcPp
                        lstCovAdd.SelectedItem.SubItems(3) = txtCovBirthdate
                        lstCovAdd.SelectedItem.SubItems(4) = cboCovSex
                        lstCovAdd.SelectedItem.SubItems(5) = cboCovAdultChild
                        lstCovAdd.SelectedItem.SubItems(6) = cboCovRelationship
                        lstCovAdd.SelectedItem.SubItems(7) = txtSuppPLNCode
                        'Dependent Status change to Editable - By CHEng 3 Dec 2016
                        lstCovAdd.SelectedItem.SubItems(8) = Cmb_Cstatus.Text
                        lstCovAdd.SelectedItem.SubItems(19) = txtCovAllergic
                        lstCovAdd.SelectedItem.SubItems(9) = txtCovExclusion
                        lstCovAdd.SelectedItem.SubItems(11) = "E"
                        lstCovAdd.SelectedItem.SubItems(12) = txtCovAnnualLimit
                        lstCovAdd.SelectedItem.SubItems(13) = txtSuppEffDate
                        lstCovAdd.SelectedItem.SubItems(14) = txtSuppExpDate
                        lstCovAdd.SelectedItem.SubItems(15) = txtCovUndExcess
                        lstCovAdd.SelectedItem.SubItems(20) = txtPrevPLNCode
                        lstCovAdd.SelectedItem.SubItems(21) = txtRBAmt
                        lstCovAdd.SelectedItem.SubItems(22) = txtCOVPayRemarks
                        lstCovAdd.SelectedItem.SubItems(23) = txtPrevINSCom
                        lstCovAdd.SelectedItem.SubItems(24) = txtSuppGenWP
                        lstCovAdd.SelectedItem.SubItems(25) = txtSuppGENWPEnd
                        lstCovAdd.SelectedItem.SubItems(26) = txtSuppPREWP
                        lstCovAdd.SelectedItem.SubItems(27) = txtSuppPREWPEnd
                        lstCovAdd.SelectedItem.SubItems(28) = TxtSuppSpeWP
                        lstCovAdd.SelectedItem.SubItems(29) = txtSuppSPEWPEnd
                        lstCovAdd.SelectedItem.SubItems(30) = Mid(cboSuppStaffDisc, 1, 1)
                        lstCovAdd.SelectedItem.SubItems(31) = IIf(IsDate(TxtSuppClmNonPayFr), TxtSuppClmNonPayFr, "")
                        lstCovAdd.SelectedItem.SubItems(32) = IIf(IsDate(TxtSuppClmNonPayTo), TxtSuppClmNonPayTo, "")
                        lstCovAdd.SelectedItem.SubItems(34) = IIf(IsNumeric(txtSuppPayPremGross), txtSuppPayPremGross, 0)
                        lstCovAdd.SelectedItem.SubItems(35) = IIf(IsNumeric(txtSuppPayPremNet), txtSuppPayPremNet, 0)
                        lstCovAdd.SelectedItem.SubItems(36) = IIf(IsNumeric(txtSuppPayMCOGross), txtSuppPayMCOGross, 0)
                        lstCovAdd.SelectedItem.SubItems(37) = IIf(IsNumeric(txtSuppPayMCONet), txtSuppPayMCONet, 0)
                        lstCovAdd.SelectedItem.SubItems(38) = IIf(IsNumeric(txtRenewalDisc), txtRenewalDisc, 0)
                        lstCovAdd.SelectedItem.SubItems(39) = IIf(IsNumeric(txtPolicyDisc), txtPolicyDisc, 0)
                        lstCovAdd.SelectedItem.SubItems(40) = IIf(IsNumeric(txtPolicyDiscAmt), txtPolicyDiscAmt, 0)
                        lstCovAdd.SelectedItem.SubItems(41) = IIf(IsNumeric(txtRenewalDiscAmt), txtRenewalDiscAmt, 0)
                        lstCovAdd.SelectedItem.SubItems(42) = IIf(IsNumeric(txtLoadingPrem), txtLoadingPrem, 0)
                        lstCovAdd.SelectedItem.SubItems(43) = IIf(Len(txtPayMBMNo), Trim(txtPayMBMNo), "")
                        lstCovAdd.SelectedItem.SubItems(44) = IIf(Len(txtPaySeqID), Trim(txtPaySeqID), "")
                        lstCovAdd.SelectedItem.SubItems(45) = IIf(Len(txtClientNo), Trim(txtClientNo), "")
                        lstCovAdd.SelectedItem.SubItems(46) = IIf(Len(txtClientID), Trim(txtClientID), "")
                        lstCovAdd.SelectedItem.SubItems(48) = Cmd_HRInd.Text
                        
                    Exit For
                End If
            Next ForLoop
            cboCovSalutation.SetFocus
        End If
       Call ClearCoveredEntry
    End If
End Sub
Private Sub lstCovAdd_Click()
    
    If lstCovAdd.ListItems.Count Then
        txtCovName = lstCovAdd.SelectedItem.Text
        cboCovSalutation.Text = lstCovAdd.SelectedItem.SubItems(1)
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate.mask = ""
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        txtSuppPLNCode = lstCovAdd.SelectedItem.SubItems(7)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(19)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(10)
        txtCovAnnualLimit = Format(lstCovAdd.SelectedItem.SubItems(12), "###,##0.00")
        txtSuppEffDate = IIf(Len(lstCovAdd.SelectedItem.SubItems(13)), lstCovAdd.SelectedItem.SubItems(13), txtMBMPayorEffDate)
        txtSuppExpDate = IIf(Len(lstCovAdd.SelectedItem.SubItems(14)), lstCovAdd.SelectedItem.SubItems(14), txtMBMPayorExpDate)
        txtCovUndExcess = lstCovAdd.SelectedItem.SubItems(15)
        txtAvailableLimit = Format(lstCovAdd.SelectedItem.SubItems(17), "#,##0.00")
        txtPrevPLNCode = lstCovAdd.SelectedItem.SubItems(20)
        txtRBAmt = lstCovAdd.SelectedItem.SubItems(21)
        txtCOVPayRemarks = lstCovAdd.SelectedItem.SubItems(22)
        txtPrevINSCom = lstCovAdd.SelectedItem.SubItems(23)
        
        txtSuppGenWP = lstCovAdd.SelectedItem.SubItems(24)
        If lstCovAdd.SelectedItem.SubItems(25) <> "" Then
            txtSuppGENWPEnd = lstCovAdd.SelectedItem.SubItems(25)
        End If
        txtSuppPREWP = lstCovAdd.SelectedItem.SubItems(26)
        If lstCovAdd.SelectedItem.SubItems(27) <> "" Then
            txtSuppPREWPEnd = lstCovAdd.SelectedItem.SubItems(27)
        End If
        
        TxtSuppSpeWP = lstCovAdd.SelectedItem.SubItems(28)
        If lstCovAdd.SelectedItem.SubItems(29) <> "" Then
            txtSuppSPEWPEnd = lstCovAdd.SelectedItem.SubItems(29)
        End If
        
        If lstCovAdd.SelectedItem.SubItems(30) <> "" Then
            cboSuppStaffDisc = lstCovAdd.SelectedItem.SubItems(30)
        End If
        TxtSuppClmNonPayFr.mask = ""
        TxtSuppClmNonPayFr.Text = ""
        TxtSuppClmNonPayFr.mask = "##/##/####"
        TxtSuppClmNonPayTo.mask = ""
        TxtSuppClmNonPayTo.Text = ""
        TxtSuppClmNonPayTo.mask = "##/##/####"
        If Len(lstCovAdd.SelectedItem.SubItems(31)) Then
            TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(31)
        End If
        If Len(lstCovAdd.SelectedItem.SubItems(32)) Then
            TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(32)
        End If
        txtSuppPayPremGross = lstCovAdd.SelectedItem.SubItems(34)
        txtSuppPayPremNet = lstCovAdd.SelectedItem.SubItems(35)
        txtSuppPayMCOGross = lstCovAdd.SelectedItem.SubItems(36)
        txtSuppPayMCONet = lstCovAdd.SelectedItem.SubItems(37)
        txtRenewalDisc = lstCovAdd.SelectedItem.SubItems(38)
        txtPolicyDisc = lstCovAdd.SelectedItem.SubItems(39)
        txtPolicyDiscAmt = lstCovAdd.SelectedItem.SubItems(40)
        txtRenewalDiscAmt = lstCovAdd.SelectedItem.SubItems(41)
        txtLoadingPrem = lstCovAdd.SelectedItem.SubItems(42)
        txtPayMBMNo = IIf(Len(lstCovAdd.SelectedItem.SubItems(43)), Trim(lstCovAdd.SelectedItem.SubItems(43)), "")
        txtPaySeqID = IIf(Len(lstCovAdd.SelectedItem.SubItems(44)), Trim(lstCovAdd.SelectedItem.SubItems(44)), "")
        txtClientNo = IIf(Len(lstCovAdd.SelectedItem.SubItems(45)), Trim(lstCovAdd.SelectedItem.SubItems(45)), "")
        txtClientID = IIf(Len(lstCovAdd.SelectedItem.SubItems(46)), Trim(lstCovAdd.SelectedItem.SubItems(46)), "")
        'Add Student/OKU Indicator to cover > 21yrs old - By CHEng 1 Dec 2016'
        Cmd_HRInd.Text = IIf(Len(lstCovAdd.SelectedItem.SubItems(48)), Trim(lstCovAdd.SelectedItem.SubItems(48)), "")
        'Dependent Status change to Editable - By CHEng 3 Dec 2016
        Cmb_Cstatus.Text = IIf(Len(lstCovAdd.SelectedItem.SubItems(8)), Trim(lstCovAdd.SelectedItem.SubItems(8)), "")

    End If
End Sub

Private Sub lstCovAdd_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstCovAdd.SortKey = ColumnHeader.Index - 1
    lstCovAdd.Sorted = True
    If lstCovAdd.SortOrder = lvwAscending Then
        lstCovAdd.SortOrder = lvwDescending
    Else
        lstCovAdd.SortOrder = lvwAscending
    End If
End Sub

Private Sub cmdClear_Click()
    Call ClearCoveredEntry
End Sub

Public Sub ClearCoveredEntry()
    cboCovSalutation = "MR"
    txtCovName = ""
    txtCovIcBcPp = ""
    txtCovBirthdate.mask = ""
    txtCovBirthdate = ""
    txtCovBirthdate.mask = "##/##/####"
    cboCovSex = "M"
    cboCovAdultChild = "A"
    cboCovRelationship = "S - Son"
    txtCovAllergic = ""
    txtCovExclusion = ""
    txtCovAnnualLimit = ""
    txtPrevPLNCode = ""
    txtRBAmt = ""
    txtPrevINSCom = ""
    txtCOVPayRemarks = ""
    txtCovAnnualLimit = ""
    txtAvailableLimit = ""
    txtSuppGenWP = ""
    txtSuppPREWP = ""
    TxtSuppSpeWP = ""
    cboSuppStaffDisc = ""
    txtSuppGENWPEnd.mask = ""
    txtSuppGENWPEnd = ""
    txtSuppGENWPEnd.mask = "##/##/####"
    txtSuppPREWPEnd.mask = ""
    txtSuppPREWPEnd = ""
    txtSuppPREWPEnd.mask = "##/##/####"
    txtSuppSPEWPEnd.mask = ""
    txtSuppSPEWPEnd = ""
    txtSuppSPEWPEnd.mask = "##/##/####"
    TxtSuppClmNonPayFr.mask = ""
    TxtSuppClmNonPayFr.Text = ""
    TxtSuppClmNonPayFr.mask = "##/##/####"
    TxtSuppClmNonPayTo.mask = ""
    TxtSuppClmNonPayTo.Text = ""
    TxtSuppClmNonPayTo.mask = "##/##/####"
    txtSuppPayPremGross = ""
    txtSuppPayPremNet = ""
    txtSuppPayMCOGross = ""
    txtSuppPayMCONet = ""
    txtRenewalDisc = ""
    txtPolicyDisc = ""
    txtPolicyDiscAmt = ""
    txtRenewalDiscAmt = ""
    txtLoadingPrem = ""
    txtSuppPLNCode = cboPLNCode.Text
    txtCovUndExcess = ""
    txtSuppEffDate.mask = ""
    txtSuppEffDate.Text = ""
    txtSuppEffDate.mask = "##/##/####"
    txtSuppEffDate = txtMBMPayorEffDate
    txtSuppExpDate.mask = ""
    txtSuppExpDate.Text = ""
    txtSuppExpDate.mask = "##/##/####"
    txtSuppExpDate = txtMBMPayorExpDate
    txtPayMBMNo = ""
    txtPaySeqID = ""
    txtClientNo = ""
    txtClientID = ""
End Sub


Private Sub EnableSupplementary(status As Boolean)
    lstCovAdd.Enabled = status
    cboCovSalutation.Enabled = status
    txtCovName.Enabled = status
    txtCovIcBcPp.Enabled = status
    txtCovBirthdate.Enabled = status
    cboCovSex.Enabled = status
    cboCovAdultChild.Enabled = status
    cboCovRelationship.Enabled = status
    txtCovAllergic.Enabled = status
    txtCovExclusion.Enabled = status
    cmdRemove.Enabled = status
    cmdAdd.Enabled = status
End Sub
'  ################# END Suppplementary ######################

' ################################ Enable Tab / button   ######

Private Sub TabEnabled(status As Boolean)
    SSTab1.TabEnabled(0) = status
    SSTab1.TabEnabled(1) = status
    SSTab1.TabEnabled(2) = status
    SSTab1.TabEnabled(3) = status
End Sub

Private Sub EnableEntries(status As Boolean)
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is CommandButton Then
            ctl.Enabled = status
         ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.Enabled = status
        End If
    Next ctl
    cmdSave.Enabled = False
    cboEndorseStatus.Enabled = False
    cboStatus.Enabled = False
    txtCovAnnualLimit.Enabled = False
    txtAvailableLimit.Enabled = False
    cboExpStatus.Enabled = False
    cboPAYCode.Enabled = False
    txtSuppGENWPEnd.Enabled = False
    txtSuppPREWPEnd.Enabled = False
    txtSuppSPEWPEnd.Enabled = False
    txtSuppPLNCode.Enabled = False
    txtSuppEffDate.Enabled = False
    txtSuppExpDate.Enabled = False
End Sub

'  ############### SAVE DATA INTO DATABSE #####################
Private Sub CmdSave_Click()

Dim RecordSaved, confirmCancel
Dim strsql As String
Dim CAS As New ADODB.Recordset
Dim bSucess  As Boolean
Dim tmpEndtStatus As String
Dim CasFound As Boolean
Dim tmpSEffDate As String
Dim tmpSExpDate As String
Dim TmpSCovID As String
Dim tmpSDOB As String
Dim tmpSPlnCode As String
Dim tmpSPrem As Double
Dim tmpSMco As Double
Dim tmpSName As String
Dim tmpSIC As String
Dim AA As Integer
Dim bEffInd As Integer
If ChkEffInd.Value = 1 Then
    bEffInd = 1
Else
    bEffInd = 0
End If

    tmpEndtStatus = Mid(cboEndorseStatus, 1, 1)
    bSucess = False
    If Trim(tmpEndtStatus) = "" Then
        MsgBox "Please Choose An Endorsement Type", vbOKOnly + vbCritical, DialogBoxTitle
        cboEndorseStatus.SetFocus
    ElseIf Trim(txtMBMNumber) = "" And tmpEndtStatus <> "I" Then
        MsgBox "Please select a record to edit!", vbCritical
        txtMBMNumber.SetFocus
    ElseIf Trim(cboPLNCode) <> "" And Trim(txtMBMPolNo) = "" Then
        MsgBox "Please Enter Policy No!", vbCritical
        txtMBMPolNo.SetFocus
    ElseIf txtEndtBordxDate = "__/__/____" Then
        MsgBox "Please Enter Endt. Bordx Date!", vbCritical
        txtEndtBordxDate.SetFocus
    ElseIf txtDate = "__/__/____" Then
        MsgBox "Please Enter Endt. Effective Date!", vbCritical
        txtDate.SetFocus
    ElseIf txtMBMBirthDate = "__/__/____" Then
        MsgBox "Please Enter Birth Date!", vbCritical
        txtMBMBirthDate.SetFocus
    ElseIf Trim(cboPLNCode) <> "" And txtMBMPayorEffDate = "__/__/____" Then
        MsgBox "Please Enter Payor Eff. Date!", vbCritical
        txtMBMPayorEffDate.SetFocus
    ElseIf Trim(cboPLNCode) <> "" And txtMBMPayorExpDate = "__/__/____" Then
        MsgBox "Please Enter Payor Exp. Date!", vbCritical
        txtMBMPayorExpDate.SetFocus
    ElseIf txtMBMBordxDate = "__/__/____" Then
        MsgBox "Please Enter Bordx Date!", vbCritical
        txtMBMBordxDate.SetFocus
    ElseIf txtMBMAdd1 = "" Then
        MsgBox "Please Enter Address!", vbCritical
    ElseIf txtMBMPostCode = "" Then
        MsgBox "Please Enter Post Code!", vbCritical
    ElseIf cboSTACode = "" Then
        MsgBox "Please Enter State!", vbCritical
    ElseIf txtMBMIcBcPP = "" Then
        MsgBox "Please Enter IC Number!", vbCritical
        txtMBMIcBcPP.SetFocus
    ElseIf cboMBMSex = "" Then
        MsgBox "Please Enter Sex!", vbCritical
    ElseIf frmNewMBMOthers.txtPLNCode <> "" And frmNewMBMOthers.txtNewPolEffDate = "__/__/____" Then
        MsgBox "Please Enter New Policy Eff Date", vbOKOnly + vbCritical, DialogBoxTitle
        frmNewMBMOthers.txtNewPolEffDate = "__/__/____"
        Call CmdMBMOthers_Click
        frmNewMBMOthers.txtNewPolEffDate.SetFocus
    ElseIf frmNewMBMOthers.txtPLNCode = "" And frmNewMBMOthers.txtNewPolEffDate <> "__/__/____" Then
        MsgBox "Please Enter New Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
        frmNewMBMOthers.txtPLNCode.SetFocus
    ElseIf Trim(cboINSCode) = "" And Trim(cboPLNCode) = "" And _
    (tmpEndtStatus = "C" Or tmpEndtStatus = "X" Or tmpEndtStatus = "D" Or tmpEndtStatus = "Y") And Mid(cboStatus, 1, 1) = "A" Then
            MsgBox "This is Clinical Member, pls use Clinical Adjustment !", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
        On Error GoTo ErrorSave
        
            Call UpdateMBMRetire
            
            Screen.MousePointer = vbHourglass
            strsqlPLN = ""
            strsqlPLN = "Update MBM  set MBMEndEffInd='" & CStr(bEffInd) & "' where MBMNumber = '" & Trim(txtMBMNumber) & "'"
            medixmasterconn.Execute strsqlPLN
            strsqlPLN = ""
            medixmasterconn.beginTrans
            medixmasterconnMaint.beginTrans
            bSucess = True
            tmpServerDate = GetServerTime
            strsqlPLN = "select * from DiscPlan where PLNCode = '" & Trim(Mid(cboPLNCode, 1, 4)) & "'"
            PLN.Open strsqlPLN, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            Disc = False
            If PLN.recordCount Then
                Disc = True
            End If
            PLN.Close
            Set PLN = Nothing
            
            CasFound = False
            If tmpEndtStatus = "C" Then ' Cancellation
                Call Write2MBMnHistory("C", 1, Trim(txtMBMNumber), Trim(txtMBMPolNo), Trim(txtMBMIcBcPP), _
                Trim(txtMBMName), txtMBMPayorExpDate, Trim(cboPLNCode))
                Call WriteAdjStatus2Cov("C", tmpEndtStatus, "01")
                Call MBMMCORefundCal
            ElseIf tmpEndtStatus = "T" Then   ' Delete Supplementary
                Call Write2MBMnHistory(Mid(cboStatus, 1, 1), cboStatus.ListIndex, Trim(txtMBMNumber), Trim(txtMBMPolNo), Trim(txtMBMIcBcPP), _
                    Trim(txtMBMName), txtMBMPayorExpDate, Trim(cboPLNCode))
                Call DeleteSupplementary
            ElseIf tmpEndtStatus = "S" Then   ' Add  Supplementary
                Call Write2MBMnHistory(Mid(cboStatus, 1, 1), cboStatus.ListIndex, Trim(txtMBMNumber), Trim(txtMBMPolNo), Trim(txtMBMIcBcPP), _
                    Trim(txtMBMName), txtMBMPayorExpDate, Trim(cboPLNCode))
                Call AddSupplementaryAdjustment
            ElseIf tmpEndtStatus = "K" Then 'Cancel Supplementary
                If lstCovAdd.ListItems.Count > 0 Then
                    For AA = 1 To lstCovAdd.ListItems.Count
                        If Trim(lstCovAdd.ListItems(AA).SubItems(11)) = "C" Then
                            tmpSEffDate = IIf(Len(lstCovAdd.ListItems(AA).SubItems(13)), lstCovAdd.ListItems(AA).SubItems(13), txtMBMPayorEffDate)
                            tmpSExpDate = IIf(Len(lstCovAdd.ListItems(AA).SubItems(14)), lstCovAdd.ListItems(AA).SubItems(14), txtMBMPayorExpDate)
                            TmpSCovID = Format(lstCovAdd.ListItems(AA).SubItems(10), "00")
                            tmpSDOB = IIf(Len(lstCovAdd.ListItems(AA).SubItems(3)), lstCovAdd.ListItems(AA).SubItems(3), txtMBMBirthDate)
                            tmpSPlnCode = IIf(Len(lstCovAdd.ListItems(AA).SubItems(7)), lstCovAdd.ListItems(AA).SubItems(7), cboPLNCode)
                            tmpSPrem = IIf(IsNumeric(lstCovAdd.ListItems(AA).SubItems(18)), lstCovAdd.ListItems(AA).SubItems(18), 0)
                            tmpSMco = IIf(IsNumeric(lstCovAdd.ListItems(AA).SubItems(16)), lstCovAdd.ListItems(AA).SubItems(16), 0)
                            tmpSName = IIf(Len(lstCovAdd.ListItems(AA).Text), lstCovAdd.ListItems(AA).Text, "")
                            tmpSIC = IIf(Len(lstCovAdd.ListItems(AA).SubItems(2)), lstCovAdd.ListItems(AA).SubItems(2), "")
                            Call Write2MBMnHistory(tmpEndtStatus, cboStatus.ListIndex, Trim(txtMBMNumber), Trim(txtMBMPolNo), Trim(tmpSIC), _
                                Trim(tmpSName), tmpSExpDate, Trim(tmpSPlnCode))
                            Call MBMSuppMCORefundCal(TmpSCovID, tmpSEffDate, tmpSExpDate, tmpSDOB, tmpSPlnCode, tmpEndtStatus, tmpSPrem, tmpSMco)
                            Call WriteAdjStatus2Cov("C", tmpEndtStatus, TmpSCovID)
                        End If
                    Next AA
                End If
            ElseIf tmpEndtStatus = "X" Then   ' Suspend MEMbership
                Call Write2MBMnHistory("S", 3, Trim(txtMBMNumber), Trim(txtMBMPolNo), Trim(txtMBMIcBcPP), _
                    Trim(txtMBMName), txtMBMPayorExpDate, Trim(cboPLNCode))
                Call WriteAdjStatus2Cov("S", tmpEndtStatus, "01")
            ElseIf tmpEndtStatus = "A" Or tmpEndtStatus = "P" Then ' Amendment n Plan Change
                Call Write2MBMnHistory(tmpEndtStatus, 0, Trim(txtMBMNumber), Trim(txtMBMPolNo), Trim(txtMBMIcBcPP), _
                    Trim(txtMBMName), txtMBMPayorExpDate, Trim(cboPLNCode))
                Call WriteAdjStatus2Cov(tmpEndtStatus, tmpEndtStatus, "01")
                Call SavePrincipalAdjustment
                Call SaveSupplementaryAdjustment
                Call UpdateMCOPending
                If Mid(cboEndorseStatus, 1, 1) = "P" Then
                    Call MBMMCORefundCal  'ChangePlanRefund
                End If
            ElseIf tmpEndtStatus = "D" Then  ' Delete Policy
                Call Write2MBMnHistory("D", 2, Trim(txtMBMNumber), Trim(txtMBMPolNo), Trim(txtMBMIcBcPP), _
                    Trim(txtMBMName), txtMBMPayorExpDate, Trim(cboPLNCode))
                Call WriteAdjStatus2Cov("D", tmpEndtStatus, "01")
            End If
            medixmasterconn.commitTrans
            medixmasterconnMaint.commitTrans
            bSucess = False
            If tmpEndtStatus = "X" Then
                MsgBox "Record is Suspended", vbOKOnly + vbInformation, DialogBoxTitle
            ElseIf tmpEndtStatus = "D" Then
                MsgBox "Policy Deleted Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
            ElseIf tmpEndtStatus = "Y" Then
                MsgBox "Record is Activated", vbOKOnly + vbInformation, DialogBoxTitle
            ElseIf tmpEndtStatus = "A" Or tmpEndtStatus = "P" Then
                MsgBox "Updated Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
            ElseIf tmpEndtStatus = "C" Then
                MsgBox "This Membership has been cancelled successfully", vbOKOnly + vbInformation, DialogBoxTitle
            ElseIf tmpEndtStatus = "T" Then
                MsgBox "Supplementary Deleted Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
            ElseIf tmpEndtStatus = "K" Then
                MsgBox "Supplementary cancelled Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
            ElseIf tmpEndtStatus = "S" Then
                MsgBox "Supplementary Added Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
            End If
        End If
        Screen.MousePointer = vbDefault
        txtDate.mask = ""
        txtDate.mask = "__/__/____"
        txtDate.mask = "##/##/####"
        txtEndtBordxDate.mask = ""
        txtEndtBordxDate.mask = "__/__/____"
        txtEndtBordxDate.mask = "##/##/####"
        tmpUpdateAnnLmt = False
    End If
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If bSucess = True Then
        medixmasterconn.RollbackTrans
        medixmasterconnMaint.RollbackTrans
    End If
    If CasFound = False Then
        MsgBox Err.Description
    End If
End Sub

Private Sub MBMMCORefundCal()
    Dim strsql As String
    Dim PayEffDate As String
    Dim PAYExpDate As String
    Dim MBMDOB As String
    Dim MBMCancelStatus As String
    Dim PayRec As New ADODB.Recordset
    Dim tmpPlnCode As String
    Dim tmpINSType As String
    Dim AA As Integer
    AA = InStr(cboPLNCode, "-")
    tmpPlnCode = Trim(Mid(cboPLNCode, 1, IIf(AA > 0, AA - 1, Len(cboPLNCode))))
    tmpINSType = Mid(cboINSCode, 1, 1)
    MBMDOB = txtMBMBirthDate
    PayEffDate = txtMBMPayorEffDate
    PAYExpDate = txtMBMPayorExpDate
    MBMCancelStatus = "N"
    strsql = "Select MBMCancelStatus From PAY where PAYCode ='" & Mid(cboPAYCode, 1, 2) & "'"
    PayRec.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If PayRec.recordCount Then
       MBMCancelStatus = IIf(Len(Trim(PayRec!MBMCancelStatus)), Trim(PayRec!MBMCancelStatus), "N")
    End If
    PayRec.Close
    Set PayRec = Nothing
    Call SaveMBMRefund(CDate(txtMBMPayorEffDate), CDate(txtMBMPayorExpDate), "00", CInt(txtMBMPremium), CInt(txtMCOFees), MBMCancelStatus)
    If Trim(tmpINSType) = "H" Or Trim(tmpINSType) = "F" Then
        Call MBMSuppMCORefundCal("00", txtMBMPayorEffDate, txtMBMPayorExpDate, MBMDOB, tmpPlnCode, "C", CDbl(txtMBMPremium), CDbl(txtMCOFees))
    End If
End Sub
Public Sub RecalculateMCO()
'    Dim MBMRefund As New ADODB.Recordset
'    Dim StrSqlRefund As String
'    Dim MBMOthers As New ADODB.Recordset
'    Dim StrSqlOhers As String
'    Dim PremiumRefundValue As Double
'    Dim MCORefundValue As Double
'
'    StrSqlRefund = "Select MBMNumber,MBMPremReduction,MBMMCORefund from MBMRefund Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
'    MBMRefund.Open StrSqlRefund, medixmasterconn, adOpenKeyset, adLockOptimistic
'
'        If MBMRefund.recordCount Then
'            PremiumRefundValue = MBMRefund!MBMPremReduction
'            MCORefundValue = MBMRefund!MBMMCORefund
'        Else
'            PremiumRefundValue = 0
'            MCORefundValue = 0
'        End If
'    MBMRefund.Close
'    Set MBMRefund = Nothing

'    StrSqlOhers = "Select MBMNumber,MBMMCOFee,MBMPremium from MBMOthers Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
'    MBMOthers.Open StrSqlOhers, medixmasterconn, adOpenKeyset, adLockOptimistic
'
'        If MBMOthers.recordCount Then
'            If MCORefundValue <> "0" Then
'                MBMOthers!MBMMCOFee = MCORefundValue
'            End If
'            If PremiumRefundValue <> "0" Then
'                MBMOthers!MBMPremium = PremiumRefundValue
'            End If
'        MBMOthers.Update
'
'            txtMBMPremium = Format(Round(MBMOthers!MBMPremium, 0), "#,##0.00")
'            txtMCOFees = Format(Round(MBMOthers!MBMMCOFee, 0), "#,##0.00")
'        End If
'    MBMOthers.Close
'    Set MBMOthers = Nothing
End Sub

Private Sub SavePrincipalAdjustment()
Dim MBM As New ADODB.Recordset
Dim MBMOthers As New ADODB.Recordset
Dim strsql As String
Dim strsqlOthers As String
Dim RecordSaved
Dim strsqlTwo As String
Dim MBMTwo As New ADODB.Recordset
Dim strsqlOthersTwo As String
Dim MBMOthersTwo As New ADODB.Recordset
Dim MBMCRef As New ADODB.Recordset
Dim StrSqlCRef As String
Dim TempMBMName As String
Dim NewRenewKey1 As String
Dim NewRenewKey2 As String
Dim BB As Integer
Dim CC As Integer
Dim DD As Integer
Dim RenewSql1 As String
Dim NewMBMSql As String
Dim tmpUMBMNo As String
Dim UpdMBMSql, UPbMBMXref As String
Dim MBMEndEffInd As Boolean



    
    BB = InStr(1, PrevRenewKey1, "'")
    CC = InStr(1, txtMBMName, "'")
    DD = InStr(1, txtMBMNumber, "*")
    If DD > 0 Then
        tmpUMBMNo = Mid(txtMBMNumber, 1, DD - 1)
    End If
    If BB > 0 Then
        PrevRenewKey1 = Replace(PrevRenewKey1, "'", "''")
    End If
    NewRenewKey1 = ""
    If Trim(PrevRenewKey1) <> "" Then
        If Trim(cboPLNCode) <> "" And Trim(cboINSCode) <> "" Then
            NewRenewKey1 = Trim(cboPAYCode) & "*" & Trim(cboINSCode) & "*" & Format(txtMBMBirthDate, "yyyymmdd") & "*" & Trim(txtMBMName)
            If CC = 0 Then
                TempMBMName = Trim(txtMBMName)
            Else
                TempMBMName = Replace(txtMBMName, "'", "''")
            End If
            NewRenewKey2 = Trim(cboPAYCode) & "*" & Trim(cboINSCode) & "*" & Format(txtMBMBirthDate, "yyyymmdd") & "*" & Trim(TempMBMName)
        End If
    End If
    If Trim(PrevRenewKey1) <> Trim(NewRenewKey1) Then
        RenewSql1 = "Update MBMSEQRenew05 set RewMBMSEQKey ='" & Trim(NewRenewKey2) & "' where RewMBMSEQKey = '" & Trim(PrevRenewKey1) & "'"
        medixmasterconnMaint.Execute RenewSql1
        UpdMBMSql = "Update MBM set MBMName ='" & Trim(TempMBMName) & "',MBMDOB ='" & Format(txtMBMBirthDate, "yyyy/mm/dd") & "'  where MBMNumber like '%" & tmpUMBMNo & "%'"
        medixmasterconn.Execute UpdMBMSql
        UPbMBMXref = "Update MBMCrossReference set MBMName ='" & Trim(TempMBMName) & "',MBMDOB ='" & Format(txtMBMBirthDate, "yyyy/mm/dd") & "' where MBMNumber like '%" & tmpUMBMNo & "%'"
        medixmasterconnMaint.Execute UPbMBMXref
    End If
    
    strsql = " Select MBMNumber, MBMName, MBMPolicyNo, MBMSex, MBMIcBcPp, MBMAddress1," & _
            " MBMAddress2, MBMAddress3, CTYCode, STACode, MBMPostCode,PLNCode," & _
            " MBMIcBcPp, MBMDOB, RACCode, MBMPayorEffDate, MBMPayorEXPDate, " & _
            " BRNCode,  MBMBordxDate, AGECode, PLNCode,MBMAvailableLimit," & _
            " MBMTakeOver, MBMRenewFlag, MBMMemberType, MBMDataStatus, MBMStatus "

    strsql = strsql & " From MBM Where MBMNumber ='" & Trim(txtMBMNumber) & "'"
    
    MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
    StrSqlCRef = " SELECT MBMName, MBMPolicyNo,  MBMDOB, MBMIcBcPp,MBMStatus," & _
            " MBMPayorEffDate, MBMPayorExpDate , MBMBordxDate, MBMDataStatus, PLNCode, MBMBatchNo "
    StrSqlCRef = StrSqlCRef & " From MBMCrossReference Where MBMNumber ='" & Trim(txtMBMNumber) & "'"
    
    MBMCRef.Open StrSqlCRef, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
    strsqlTwo = " SELECT MBMSalutation, MBMEmail, LGNCode, NATCode,MBMTelnoH,MBMLastEndorseStatus," & _
            " MBMTelNoM, MBMTelNoO , MBMLastAdjustmentDate, MBMLastUpdateUser,MBMMmgtNotes "
    strsqlTwo = strsqlTwo & " From MBMTwo Where MBMNumber ='" & Trim(txtMBMNumber) & "'"
    
    MBMTwo.Open strsqlTwo, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    With MBM
        !MBMName = ChkStr(txtMBMName)
        !MBMPolicyNo = Trim(ChkStr(txtMBMPolNo))
        !MBMSex = ChkStr(cboMBMSex)
        !MBMAddress1 = Trim(ChkStr(txtMBMAdd1))
        !MBMAddress2 = Trim(ChkStr(txtMBMAdd2))
        !MBMAddress3 = Trim(ChkStr(txtMBMAdd3))
        !CTYCode = Trim(ChkStr(cboCTYCode))
        !STACode = Trim(ChkStr(cboSTACode))
        !MBMPostCode = txtMBMPostCode
        !MBMIcBcPp = ChkStr(txtMBMIcBcPP)
        !MBMDOB = txtMBMBirthDate
        !RACCode = Trim(Mid(cboRACCode, 1, 2))
        !AGECode = Mid(txtAGECode, 1, 2)
        If IsDate(txtMBMPayorEffDate) Then
            !MBMPayorEffDate = txtMBMPayorEffDate
        End If
        If IsDate(txtMBMPayorExpDate) Then
            !MBMPayorEXPDate = txtMBMPayorExpDate
        End If
        !MBMBordxDate = txtMBMBordxDate
        !MBMTakeover = Mid(cboMBMTakeOver, 1, 1)
        !MBMRenewFlag = Mid(cboMBMRenewal, 1, 1)
        !MBMDataStatus = Mid(cboDataStatus, 1, 1)
        
        If Mid(cboEndorseStatus, 1, 1) = "P" Then
            !MBMStatus = "P"
        Else
            !MBMStatus = Mid(cboStatus, 1, 1)
        End If
        !PLNCode = cboPLNCode
            If tmpUpdateAnnLmt = True Then
                !MBMAvailableLimit = txtAnnLmt
            End If
        MBM.Update
    End With
    
    With MBMTwo
        If MBMTwo.recordCount = 0 Then
            MBMTwo.AddNew
            !MBMNumber = ChkStr(MBM!MBMNumber)
        End If
        
        !MBMSalutation = ChkStr(cboMBMSalutation)
        '!MBMEmail = ChkStr(frmNewMBMOthers.txtMBMEmail)
        !MBMEmail = ChkStr(tmpEmail)
        !LGNCode = Trim(Mid(frmNewMBMOthers.cboLGNCode, 1, 1))
        !MBMLastEndorseStatus = Trim(Mid(cboEndorseStatus, 1, InStr(1, cboEndorseStatus, "-") - 1))
        If InStr(1, cboNATCode, "-") > 0 Or Len(cboNATCode) = 2 Then
            If InStr(1, cboNATCode, "-") > 0 Then
                !NATCode = Trim(Mid(cboNATCode, 1, InStr(1, cboNATCode, "-") - 1))
            Else
                !NATCode = Trim(Mid(cboNATCode, 1, 2))
            End If
        End If
        !MBMTelnoH = frmNewMBMOthers.txtMBMTelNoH
        !MBMTelNoM = frmNewMBMOthers.txtMBMTelNoM
        !MBMTelNoO = frmNewMBMOthers.txtMBMTelNoO
        !MBMLastAdjustmentDate = tmpServerDate
        !MBMLastUpdateUser = Logon
        !MBMMmgtNotes = IIf(Len(txtMgmtNotes), txtMgmtNotes, "")
        MBMTwo.Update
    End With
    
    With MBMCRef
        !MBMName = ChkStr(txtMBMName)
        !MBMPolicyNo = Trim(ChkStr(txtMBMPolNo))
        !MBMIcBcPp = ChkStr(txtMBMIcBcPP)
        !MBMDOB = txtMBMBirthDate
        If IsDate(txtMBMPayorEffDate) Then
            !MBMPayorEffDate = txtMBMPayorEffDate
        End If
        If IsDate(txtMBMPayorExpDate) Then
            !MBMPayorEXPDate = txtMBMPayorExpDate
        End If
        !MBMBordxDate = txtMBMBordxDate
        !MBMDataStatus = Mid(cboDataStatus, 1, 1)
        If Mid(cboEndorseStatus, 1, 1) = "P" Then
            !MBMStatus = "P"
        Else
            !MBMStatus = Mid(cboStatus, 1, 1)
        End If
        '!MBMBatchNo = txtMBMBatchNo
        !PLNCode = cboPLNCode
        MBMCRef.Update
    End With
    
    
    If MBMPayorEXPDate = 1 Then
        RecordSaved = MsgBox("Do you want to edit Expiry Date?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            Call InsertMCO
            Call InsertPrem
         End If
     End If
    strsqlOthers = " Select Top 1 MBMNumber, MBMDepartment, MBMBranch, MBMNextDueDate, " & _
                    " MBMGroupCompany, MBMPolSubNo1, LabelStatus, MBMEmployeeNo," & _
                    " MBMAgentCode, RELCODE, MBMInstallment, MBMMCOFee, MBMPremium, MBMBasicPrem," & _
                    " MBMAccessFee, MBMFirstJoinedDate,  MBMCLMNonPayFr, MBMUndExcess,MBMENDtRefNo, " & _
                    " MBMCLMNonPayTo, MBMNewPlanCode, MBMNewPolEffDate, MBMPolicyDisc, MBMRenewalDisc, " & _
                    " MBMPolCondition, MBMCostCenter, MBMPosition,MBMGroupCategory,MBMDivision,MBMSubDivision,Subsidiary "
                
    strsqlOthers = strsqlOthers & " From MBMOthers Where MBMNumber ='" & Trim(txtMBMNumber) & "'"
    
    MBMOthers.Open strsqlOthers, medixmasterconn, adOpenKeyset, adLockOptimistic

        If MBMOthers.recordCount = 0 Then
            MBMOthers.AddNew
            MBMOthers!MBMNumber = txtMBMNumber
        End If
            With MBMOthers
                 !MBMDepartment = ChkStr(txtDeptName)
                 !MBMGroupCompany = ChkStr(txtMBMGroupCompany)
                 !MBMPolSubNo1 = frmNewMBMOthers.TxtPOLSubNo
                 If frmNewMBMOthers.Check1.Value = 1 Then
                    !LabelStatus = "Y"
                 Else
                    !LabelStatus = Null
                 End If
                 !MBMEmployeeNo = ChkStr(txtMBMEmployeeNo)
                 !MBMAgentCode = ChkStr(txtMBMAgentCode)
                 !RELCode = "P"
                 !MBMInstallment = IIf(Len(Mid(CboInsMode, 1, 1)), Mid(CboInsMode, 1, 1), "")
                If Len(txtMCOFees) Then
                    !MBMMCOFee = txtMCOFees
                End If
                !MBMBranch = IIf(Len(cboBRNCode), Trim(cboBRNCode), "")
                If Len(txtMBMBasicPremium) Then
                    !MBMBasicPrem = txtMBMBasicPremium
                End If
                
                If Len(txtMBMPremium) Then
                    !MBMPremium = txtMBMPremium
                End If
                
                If txtMBMDateJoin <> "__/__/____" Then
                    !MBMFirstJoinedDate = txtMBMDateJoin
                End If
                
                If frmNewMBMOthers.TxtClmNonPayFr <> "__/__/____" And frmNewMBMOthers.TxtClmNonPayFr <> "" Then
                   !MBMCLMNonPayFr = frmNewMBMOthers.TxtClmNonPayFr
                Else
                    !MBMCLMNonPayFr = Null
                End If
                If frmNewMBMOthers.TxtClmNonPayTo <> "__/__/____" And frmNewMBMOthers.TxtClmNonPayTo <> "" Then
                   !MBMCLMNonPayTo = frmNewMBMOthers.TxtClmNonPayTo
                Else
                    !MBMCLMNonPayTo = Null
                End If
             '=================== Add By CK ================

                !MBMNewPlanCode = IIf(Len(frmNewMBMOthers.txtPLNCode), frmNewMBMOthers.txtPLNCode, "")
                If frmNewMBMOthers.txtNewPolEffDate <> "__/__/____" And frmNewMBMOthers.txtNewPolEffDate <> "" Then
                    !MBMNewPolEffDate = frmNewMBMOthers.txtNewPolEffDate
                Else
                    !MBMNewPolEffDate = Null
                End If
                !MBMPolicyDisc = IIf(IsNumeric(frmNewMBMOthers.txtPolicyDisc), frmNewMBMOthers.txtPolicyDisc, 0)
                !MBMRenewalDisc = IIf(IsNumeric(frmNewMBMOthers.txtRenewalDisc), frmNewMBMOthers.txtRenewalDisc, 0)
                !MBMPolCondition = IIf(Len(frmNewMBMOthers.cboPOLCondition), frmNewMBMOthers.cboPOLCondition, "")
                !MBMUndExcess = IIf(IsNumeric(frmNewMBMOthers.txtUndExcess), frmNewMBMOthers.txtUndExcess, 0)
                !MBMENDtRefNo = IIf(Len(frmNewMBMOthers.txtEndtRefNo), frmNewMBMOthers.txtEndtRefNo, "")
                !MBMCostCenter = IIf(Len(txtCostCenter), txtCostCenter, "")
                !MBMPosition = IIf(Len(txtMBMPosition), txtMBMPosition, "")
                !MBMGroupCategory = IIf(Len(txtGrpCat), txtGrpCat, "")
                !MBMDivision = IIf(Len(txtDivision), txtDivision, "")
                !MBMSubDivision = IIf(Len(txtSubDivision), txtSubDivision, "")
                !Subsidiary = IIf(Len(txtSubsidiary), txtSubsidiary, "")
                
                'If tmpPAYGRPCode = "ETS" Then
                    If txtNxDueDate <> "__/__/____" Then
                        !MBMNextDueDate = txtNxDueDate
                    End If
                'End If
         End With
         MBMOthers.Update
        
    strsqlOthersTwo = " Select MBMNumber, MBMWeight, MBMHeight, MBMWorkNature, MBMBlood, MBMAllergic, MBMRemarks," & _
                    " MBMSmoking, MBMAlcohol, MBMMaritalStatus, MBMPrevPLNCode," & _
                    " MBMRBAmt, MBMPAYRemarks, MBMPrevInsCom, MBMOccupation, MBMCoPayRB, " & _
                    " MBMGenWP, MBMGenEndWP, MBMPreExWP, MBMPreExEndWP, MBMSpeIllWP, " & _
                    " MBMSpeIllEndWP, MBMStaffDec, MBMBTBG, MBMPayorPremGross, MBMPayorPremNet, " & _
                    " MBMPayorMCOGross, MBMPayorMCONet, MBMFamDiscAmt, MBMRenewDiscAmt, MBMLoadingPrem,MBMExclusion, MBMBankACNo, MBMTopupJoinDate"
                
    strsqlOthersTwo = strsqlOthersTwo & " From MBMOthersTwo Where MBMNumber ='" & Trim(txtMBMNumber) & "'"
    
    MBMOthersTwo.Open strsqlOthersTwo, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBMOthersTwo.recordCount = 0 Then
            MBMOthersTwo.AddNew
            MBMOthersTwo!MBMNumber = txtMBMNumber
        End If
            With MBMOthersTwo
                 !MBMAllergic = ChkStr(frmNewMBMOthers.txtMBMAllergic)
                 !MBMRemarks = IIf(Len(txtRemarks), Trim(txtRemarks), "")
                 !MBMWeight = ChkStr(frmNewMBMOthers.txtMBMWeight)
                 !MBMHeight = ChkStr(frmNewMBMOthers.txtMBMHeight)
                 If InStr(1, frmNewMBMOthers.txtMBMWorkNature, "-") > 0 Then
                    !MBMWorkNature = Trim(Mid(frmNewMBMOthers.txtMBMWorkNature, 1, InStr(1, frmNewMBMOthers.txtMBMWorkNature, "-") - 1))
                 Else
                   !MBMWorkNature = Trim(ChkStr(frmNewMBMOthers.txtMBMWorkNature))
                 End If
                 !MBMBlood = Trim(Mid(frmNewMBMOthers.cboMBMBlood, 1, 1))
                 !MBMSmoking = Trim(Mid(frmNewMBMOthers.cboMBMSmoking, 1, 1))
                 !MBMAlcohol = Mid(frmNewMBMOthers.cboMBMAlcohol, 1, 1)
                 If Len(Trim(cboMBMMaritalStatus)) Then
                    !MBMMaritalStatus = Mid(cboMBMMaritalStatus, 1, 1)
                 End If
                    !MBMPrevPLNCode = IIf(Len(frmNewMBMOthers.txtMBMPrevPlan), ChkStr(frmNewMBMOthers.txtMBMPrevPlan), "")
                    !MBMRBAmt = IIf(IsNumeric(frmNewMBMOthers.txtMBMRBAmt), frmNewMBMOthers.txtMBMRBAmt, 0)
                    '!MBMPAYRemarks = IIf(Len(txtMBMPAYRemarks), ChkStr(txtMBMPAYRemarks), "")
                    !MBMPrevInsCom = IIf(Len(frmNewMBMOthers.txtMBMPrevINSCom), ChkStr(frmNewMBMOthers.txtMBMPrevINSCom), "")
                If InStr(1, frmNewMBMOthers.cboMBMOccupation, "-") > 0 Then
                    !MBMOccupation = Trim(Mid(frmNewMBMOthers.cboMBMOccupation, 1, InStr(1, frmNewMBMOthers.cboMBMOccupation, "-") - 1))
                Else
                    !MBMOccupation = Trim(ChkStr(frmNewMBMOthers.cboMBMOccupation))
                End If
                If Len(frmNewMBMOthers.cboCOPayRB) Then
                    !MBMCoPayRB = frmNewMBMOthers.cboCOPayRB
                Else
                    !MBMCoPayRB = ""
                End If
                             '=================== Add By CK ================
                If Len(frmNewMBMOthers.TxtGenWP) Then
                    !MBMGenWP = frmNewMBMOthers.TxtGenWP
                Else
                    !MBMGenWP = Null
                End If
                If frmNewMBMOthers.txtGENWPEnd <> "__/__/____" And frmNewMBMOthers.txtGENWPEnd <> "" Then
                   !MBMGenEndWP = frmNewMBMOthers.txtGENWPEnd
                Else
                    !MBMGenEndWP = Null
                End If
                If Len(frmNewMBMOthers.txtPREWP) Then
                    !MBMPreExWP = frmNewMBMOthers.txtPREWP
                Else
                    !MBMPreExWP = Null
                End If
                If frmNewMBMOthers.txtPREWPEnd <> "__/__/____" And frmNewMBMOthers.txtPREWPEnd <> "" Then
                   !MBMPreExEndWP = frmNewMBMOthers.txtPREWPEnd
                Else
                   !MBMPreExEndWP = Null
                End If
                If Len(frmNewMBMOthers.TxtSpeWP) Then
                    !MBMSpeIllWP = Trim(frmNewMBMOthers.TxtSpeWP)
                Else
                    !MBMSpeIllWP = Null
                End If
                If frmNewMBMOthers.txtSPEWPEnd <> "__/__/____" And frmNewMBMOthers.txtSPEWPEnd <> "" Then
                   !MBMSpeIllEndWP = frmNewMBMOthers.txtSPEWPEnd
                Else
                    !MBMSpeIllEndWP = Null
                End If
                If Len(frmNewMBMOthers.cboStaffDisc) Then
                    !MBMStaffDec = Mid(frmNewMBMOthers.cboStaffDisc, 1, 1)
                Else
                    !MBMStaffDec = ""
                End If
          
                If Len(frmNewMBMOthers.cboBTBG) Then
                    !MBMBTBG = Mid(frmNewMBMOthers.cboBTBG, 1, 1)
                Else
                    !MBMBTBG = Null
                End If
                
                If Len(frmNewMBMOthers.txtPayPrem) Then
                    !MBMPayorPremGross = frmNewMBMOthers.txtPayPrem
                End If
               
            !MBMPayorPremNet = IIf(IsNumeric(frmNewMBMOthers.txtPayPremNett), frmNewMBMOthers.txtPayPremNett, 0)
            !MBMPayorMCOGross = IIf(IsNumeric(frmNewMBMOthers.txtPayMCOGross), frmNewMBMOthers.txtPayMCOGross, 0)
            !MBMPayorMCONet = IIf(IsNumeric(frmNewMBMOthers.txtPayMCONett), frmNewMBMOthers.txtPayMCONett, 0)
            !MBMFamDiscAmt = IIf(IsNumeric(frmNewMBMOthers.txtPolicyDiscAmt), frmNewMBMOthers.txtPolicyDiscAmt, 0)
            !MBMRenewDiscAmt = IIf(IsNumeric(frmNewMBMOthers.txtRenewalDiscAmt), frmNewMBMOthers.txtRenewalDiscAmt, 0)
            !MBMLoadingPrem = IIf(IsNumeric(frmNewMBMOthers.txtLoadingPrem), frmNewMBMOthers.txtLoadingPrem, 0)
            !MBMExclusion = IIf(Len(txtMBMExclusion), txtMBMExclusion, "")
            !MBMBankACNo = IIf(Len(frmNewMBMOthers.txtBankACNo), frmNewMBMOthers.txtBankACNo, "")
            '!MBMTopupJoinDate = IIf(txtMBMTopUpJoinDate <> "__/__/____", Format(txtMBMTopUpJoinDate, "dd/mm/yyyy"), "")
            
         End With
         MBMOthersTwo.Update
              
        MBM.Close
        Set MBM = Nothing
        MBMTwo.Close
        Set MBMTwo = Nothing
        MBMCRef.Close
        Set MBMCRef = Nothing
        MBMOthers.Close
        Set MBMOthers = Nothing
        MBMOthersTwo.Close
        Set MBMOthersTwo = Nothing
End Sub
Public Function GetMainLastCoveredID(MembershipNo As String) As Integer
On Error Resume Next
    Dim MBMCovered As New ADODB.Recordset
    Dim strsql As String
    Dim tmpCover As Integer
    strsql = ""
    strsql = "SELECT MBMCNumber, MBMCMemberType, MBMCCoverID AS CoveredID "
    strsql = strsql & "From HISDB.dbo.MBMCoveredPersons  "
    strsql = strsql & "where (MBMCNumber = '" & MembershipNo & "') ORDER BY  MBMCCoverID DESC"
    MBMCovered.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'Do While Not MBMCovered.EOF
   '
   ' MBMCovered.MoveNext
   ' Loop
    
    If MBMCovered.EOF = False Then
        tmpCover = IIf(IsNumeric(MBMCovered!CoveredID), MBMCovered!CoveredID, 0)
        GetMainLastCoveredID = tmpCover + 1
    Else
        GetMainLastCoveredID = tmpCover + 1
    End If
    
    MBMCovered.Close
    Set MBMCovered = Nothing
End Function

Private Sub SaveSupplementaryAdjustment()
'On Error Resume Next
Dim Arraycnt, arrayCurrRow As Integer
Dim strsql As String
Dim ListCount As Integer
Dim strsql2 As String
Dim MBMCov As New ADODB.Recordset
Dim LastCoveredID As Integer
Dim listloop As Integer
Dim strsqlTwo As String
Dim MBMCovTwo As New ADODB.Recordset
Dim TempMBMCName  As String

    strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    strsqlTwo = "Select * from MBMCoveredPersonsTwo Where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
    MBMCovTwo.Open strsqlTwo, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'LastCoveredID = GetLastCoveredID(txtMBMNumber)
    LastCoveredID = GetMainLastCoveredID(txtMBMNumber)
    ListCount = lstCovAdd.ListItems.Count
    For listloop = 1 To ListCount
        TempMBMCName = Replace(lstCovAdd.ListItems(listloop).Text, "'", "''")
        With MBMCov
            MBMCov.Filter = ("MBMCCoverID = '" & Format(lstCovAdd.ListItems(listloop).SubItems(10), "00") & "'")
            !MBMCName = ChkStr(TempMBMCName)
            !MBMCICBCPPNo = ChkStr(lstCovAdd.ListItems(listloop).SubItems(2))
            !MBMCDOB = lstCovAdd.ListItems(listloop).SubItems(3)
            !MBMCSex = ChkStr(lstCovAdd.ListItems(listloop).SubItems(4))
            !MBMCAdultChild = ChkStr(lstCovAdd.ListItems(listloop).SubItems(5))
            !MBMCRELCode = Mid(lstCovAdd.ListItems(listloop).SubItems(6), 1, 1)
            If Len(lstCovAdd.ListItems(listloop).SubItems(13)) Then
                If IsDate(lstCovAdd.ListItems(listloop).SubItems(13)) Then
                    !MBMCEffDate = lstCovAdd.ListItems(listloop).SubItems(13)
                End If
            End If
            If Len(lstCovAdd.ListItems(listloop).SubItems(14)) Then
                If IsDate(lstCovAdd.ListItems(listloop).SubItems(14)) Then
                     !MBMCExpDate = lstCovAdd.ListItems(listloop).SubItems(14)
                End If
            End If
            If Len(lstCovAdd.ListItems(listloop).SubItems(18)) Then
                 !MBMCPremium = lstCovAdd.ListItems(listloop).SubItems(18)
            End If
            
            If Len(lstCovAdd.ListItems(listloop).SubItems(7)) Then
                !MBMCPLNCode = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(7)), lstCovAdd.ListItems(listloop).SubItems(7), cboPLNCode)
            End If
                
'            SuppPlanCode = IIf(Len(lstCovAdd.listitems(listloop).SubItems(7)), lstCovAdd.listitems(listloop).SubItems(7), cboPLNCode)

'            TempSuppExpDate = GetSuppExpDate(lstCovAdd.listitems(listloop).SubItems(3), SuppPlanCode, txtMBMPayorEffDate, txtMBMPayorExpDate)
            
'            !MBMCExpDate = TempSuppExpDate
            
            If Mid(cboEndorseStatus, 1, 1) = "P" Then
                !MBMCStatus = "P"
            End If
            'Add Student/OKU Indicator to cover > 21yrs old - By CHEng 1 Dec 2016'
            If Len(lstCovAdd.ListItems(listloop).SubItems(48)) Then
                !MBMHRAppInd = lstCovAdd.ListItems(listloop).SubItems(48)
            End If
            'Dependent Status change to Editable - By CHEng 3 Dec 2016
            If Len(lstCovAdd.ListItems(listloop).SubItems(8)) Then
                !MBMCStatus = lstCovAdd.ListItems(listloop).SubItems(8)
            End If
            
            MBMCov.Update
        End With
    
        With MBMCovTwo
            MBMCovTwo.Filter = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(10)) = 1, ("MBMCCoverID = '" & 0 & lstCovAdd.ListItems(listloop).SubItems(10) & "'"), ("MBMCCoverID = '" & lstCovAdd.ListItems(listloop).SubItems(10) & "'"))
            MBMCovTwo!MBMCSalutation = lstCovAdd.ListItems(listloop).SubItems(1)
            !MBMCAllergic = lstCovAdd.ListItems(listloop).SubItems(19)
            !MBMCExclusion = lstCovAdd.ListItems(listloop).SubItems(9)
'            If Len(lstCovAdd.listitems(listloop).SubItems(13)) Then
'              !MBMCOccupation = lstCovAdd.listitems(listloop).SubItems(13)
'            End If
'            If Len(lstCovAdd.listitems(listloop).SubItems(14)) Then
'              !MBMCWorkNature = lstCovAdd.listitems(listloop).SubItems(14)
'            End If
'            If Len(lstCovAdd.listitems(listloop).SubItems(15)) Then
'              !MBMCSmoking = lstCovAdd.listitems(listloop).SubItems(15)
'            End If
'            If Len(lstCovAdd.listitems(listloop).SubItems(16)) Then
'              !MBMCAlcohol = lstCovAdd.listitems(listloop).SubItems(16)
'            End If
            If lstCovAdd.ListItems(listloop).SubItems(31) <> "" And lstCovAdd.ListItems(listloop).SubItems(31) <> "__/__/____" Then
                !MBMCCLMNonPayFr = ChkStr(lstCovAdd.ListItems(listloop).SubItems(31))
            Else
                !MBMCCLMNonPayFr = Null
            End If
            If lstCovAdd.ListItems(listloop).SubItems(32) <> "" And lstCovAdd.ListItems(listloop).SubItems(32) <> "__/__/____" Then
                !MBMCCLMNonPayTo = ChkStr(lstCovAdd.ListItems(listloop).SubItems(32))
            Else
                !MBMCCLMNonPayTo = Null
            End If
            
            !MBMCPolicyDisc = IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(39)), lstCovAdd.ListItems(listloop).SubItems(39), 0)
            !MBMCRenewalDisc = IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(38)), lstCovAdd.ListItems(listloop).SubItems(38), 0)

            !MBMCPrevPLNCode = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(20)), lstCovAdd.ListItems(listloop).SubItems(20), "")
            !MBMCRBAmt = IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(21)), lstCovAdd.ListItems(listloop).SubItems(21), 0)
            !MBMCPAYRemarks = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(22)), ChkStr(lstCovAdd.ListItems(listloop).SubItems(22)), "")
            !MBMCPrevInsCom = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(23)), ChkStr(lstCovAdd.ListItems(listloop).SubItems(23)), "")
            If Len(lstCovAdd.ListItems(listloop).SubItems(24)) Then
               !MBMCGenWP = ChkStr(lstCovAdd.ListItems(listloop).SubItems(24))
            Else
               !MBMCGenWP = Null
            End If
            If lstCovAdd.ListItems(listloop).SubItems(25) <> "" And lstCovAdd.ListItems(listloop).SubItems(25) <> "__/__/____" Then
                !MBMCGenEndWP = ChkStr(lstCovAdd.ListItems(listloop).SubItems(25))
            Else
                !MBMCGenEndWP = Null
            End If
            If Len(lstCovAdd.ListItems(listloop).SubItems(26)) Then
                !MBMCPreExWP = ChkStr(lstCovAdd.ListItems(listloop).SubItems(26))
            Else
                !MBMCPreExWP = Null
            End If
            If lstCovAdd.ListItems(listloop).SubItems(27) <> "" And lstCovAdd.ListItems(listloop).SubItems(27) <> "__/__/____" Then
                !MBMCPreExEndWP = ChkStr(lstCovAdd.ListItems(listloop).SubItems(27))
            Else
                !MBMCPreExEndWP = Null
            End If
            If Len(lstCovAdd.ListItems(listloop).SubItems(28)) Then
                !MBMCSpeIllWP = ChkStr(lstCovAdd.ListItems(listloop).SubItems(28))
            Else
                !MBMCSpeIllWP = Null
            End If
            If lstCovAdd.ListItems(listloop).SubItems(29) <> "" And lstCovAdd.ListItems(listloop).SubItems(29) <> "__/__/____" Then
                !MBMSpeIllEndWP = ChkStr(lstCovAdd.ListItems(listloop).SubItems(29))
            Else
                !MBMSpeIllEndWP = Null
            End If
            If Len(lstCovAdd.ListItems(listloop).SubItems(30)) Then
                !MBMCStaffDec = ChkStr(lstCovAdd.ListItems(listloop).SubItems(30))
            End If
            If lstCovAdd.ListItems(listloop).SubItems(34) <> "" Then
                !MBMCPayorPremGross = lstCovAdd.ListItems(listloop).SubItems(34)
            Else
                !MBMCPayorPremGross = Null
            End If
            
            !MBMCPayorPremNet = IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(35)), lstCovAdd.ListItems(listloop).SubItems(35), 0)
            !MBMCPayorMCOGross = IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(36)), lstCovAdd.ListItems(listloop).SubItems(36), 0)
            !MBMCPayorMCONet = IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(37)), lstCovAdd.ListItems(listloop).SubItems(37), 0)
            !MBMCFamDiscAmt = IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(40)), lstCovAdd.ListItems(listloop).SubItems(40), 0)
            !MBMCRenewDiscAmt = IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(41)), lstCovAdd.ListItems(listloop).SubItems(41), 0)
            !MBMCLoadingPrem = IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(42)), lstCovAdd.ListItems(listloop).SubItems(42), 0)
            !MBMCUndExcess = IIf(IsNumeric(lstCovAdd.ListItems(listloop).SubItems(15)), lstCovAdd.ListItems(listloop).SubItems(15), 0)
            !MBMCPayNo = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(43)), lstCovAdd.ListItems(listloop).SubItems(43), "")
            !MBMCPaySeqNo = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(44)), lstCovAdd.ListItems(listloop).SubItems(44), "")
            !MBMCClientNo = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(45)), lstCovAdd.ListItems(listloop).SubItems(45), "")
            !MBMCClientID = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(46)), lstCovAdd.ListItems(listloop).SubItems(46), "")
            MBMCovTwo.Update
        End With
        
        lstCovAdd.ListItems(listloop).SubItems(11) = ""
    Next listloop
    MBMCov.Close
    Set MBMCov = Nothing
    MBMCovTwo.Close
    Set MBMCovTwo = Nothing

End Sub

Private Sub DeleteSupplementary()
Dim strsql As String
Dim listloop As Integer
Dim strdelcov As String
Dim MBMCovPersons As New ADODB.Recordset
Dim StrSql1 As String
Dim MBMLife As New ADODB.Recordset
Dim TmpTopupNo As String

    For listloop = 1 To lstCovAdd.ListItems.Count
        If lstCovAdd.ListItems(listloop).SubItems(11) = "R" Then
            strsql = " Where MBMCNumber ='" & txtMBMNumber & "' And MBMCCoverID ='" & lstCovAdd.ListItems(listloop).SubItems(10) & "'"
            
    ' ----- MBMCoveredPersons
            strdelcov = "Delete from MBMCoveredPersons " & strsql
            'MBMCovPersons.Open strdelcov, medixmasterconn, adOpenKeyset, adLockOptimistic
            medixmasterconn.Execute strdelcov

    ' ----- MBMCoveredPersonsTwo
            strdelcov = ""
            strdelcov = "Delete From MBMCoveredPersonsTwo " & strsql
            'MBMCovPersons.Open strdelcov, medixmasterconn, adOpenKeyset, adLockOptimistic
            medixmasterconn.Execute strdelcov

    ' ----- MBMCoveredPersonsCLN
            strdelcov = ""
            strdelcov = "Delete From MBMCoveredPersonsCLN " & strsql
'            MBMCovPersons.Open strdelcov, medixmasterconn, adOpenKeyset, adLockOptimistic
            medixmasterconn.Execute strdelcov

    ' ----- MBMLifeTime
            If Mid(CboGuaRenewal, 1, 1) = "3" Then
                StrSql1 = "Select MBMTopupNumber from MBM Where MBMNumber ='" & txtMBMNumber & "'"
                MBMLife.Open StrSql1, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                Do While Not MBMLife.EOF
                    TmpTopupNo = IIf(Len(MBMLife!MBMTopupNumber), Trim(MBMLife!MBMTopupNumber), "")
                    MBMLife.MoveNext
                Loop
                MBMLife.Close
                Set MBMLife = Nothing
                If Trim(TmpTopupNo) <> "" Then
                    strsql = " Where MBMNumber ='" & TmpTopupNo & "' And MBMCovID ='" & lstCovAdd.ListItems(listloop).SubItems(10) & "'"
                    strdelcov = ""
                    strdelcov = "Delete From MBMLifeTime " & strsql
    '                MBMCovPersons.Open strdelcov, medixmasterconn, adOpenKeyset, adLockOptimistic
                    medixmasterconn.Execute strdelcov
                End If
            End If
    ' ----- MBMBnfLmt
            strsql = " Where MBMNumber ='" & txtMBMNumber & "' And MBMCoveredID ='" & lstCovAdd.ListItems(listloop).SubItems(10) & "'"
            strdelcov = ""
            strdelcov = "Delete From MBMBnfLmt " & strsql
'            MBMCovPersons.Open strdelcov, medixmasterconn, adOpenKeyset, adLockOptimistic
            medixmasterconn.Execute strdelcov
        End If
    Next listloop
    
    Call DelSuppLoop
End Sub

Private Sub AddSupplementaryAdjustment()
Dim listloop As Integer
Dim LastCoveredID As String
Dim SelectAge As String
Dim AgeValue As String
Dim StrSqlAnnLmt As String
Dim AnnLmt As New ADODB.Recordset
Dim strsqlPrem  As String
Dim Prem As New ADODB.Recordset
Dim TempMBMCName As String
Dim tmpUndExcess As String
Dim MBMCovSql, MBMCovTwoSql As String
Dim TmpBordxDate, tmpDOb As String
Dim tmpAnnLmt As Double
Dim tmpPremium, TmpBPremium, tmpREwDisc, tmpPolDisc As String
Dim tmpSex, tmpRELCode, tmpICNo As String
Dim TmpPayPremNet, TmpPayMCOGross, TmpPayMcoNet, TmpLoadingPrem, tmpRewAmt, TmpPolAmt As String
Dim tmpAllergies, TmpExclusion, TmpMemberType As String
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset
Dim TmpSql2 As String
Dim tmpMBM As New ADODB.Recordset
Dim tmpPlnCode As String
Dim tmpINSType As String
Dim tmpKidneyAnnLmt As Double
Dim tmpCancerAnnLmt As Double
Dim tmpHomeNursAnnLmt As Double
Dim tmpSuppBNFStatus  As String
Dim strSptMCOSql As String
Dim SptMCOInd As String
Dim McoCovIDChrg As String
Dim SptPlnMCO As New ADODB.Recordset
Dim TmpBasicMCO, tmpMCO As String
Dim tmpAdultChild As String
Dim tmpSuppEffDate As String
Dim tmpClmNonPayFr As String
Dim tmpClmNonPayTo As String
Dim SuppPlanCode As String
Dim TempSuppExpDate As String
Dim tempSPremium As String
Dim tmpCNextLmt As Double

    tmpPlnCode = Trim(cboPLNCode)
    tmpINSType = Trim(cboINSCode)
    If Trim(tmpPlnCode) = "" Then
        tmpPlnCode = Trim(txtClnPlanCode)
        tmpINSType = Trim(txtClnInsType)
    End If

    'LastCoveredID = Format(GetLastCoveredID(txtMBMNumber), "00")
    LastCoveredID = Format(GetMainLastCoveredID(txtMBMNumber), "00")
    For listloop = 1 To lstCovAdd.ListItems.Count
    If lstCovAdd.ListItems(listloop).SubItems(11) = "N" Then
        SelectAge = GetAge(CDate(HisCisEffDate), lstCovAdd.ListItems(listloop).SubItems(3))
        AgeValue = Mid(GetAgeBand(Trim(cboHLTCode), SelectAge, tmpPlnCode, tmpINSType), 1, 2)
        tmpAdultChild = "C"
        If AgeValue <> "01" Then
            tmpAdultChild = "A"
        End If
        SuppPlanCode = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(7)), lstCovAdd.ListItems(listloop).SubItems(7), cboPLNCode)
        tmpAnnLmt = 0
        TmpBPremium = 0
        tmpPremium = 0
        tmpSuppEffDate = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(13)), lstCovAdd.ListItems(listloop).SubItems(13), txtMBMPayorEffDate)
        TempSuppExpDate = IIf(Len(lstCovAdd.ListItems(listloop).SubItems(14)), lstCovAdd.ListItems(listloop).SubItems(14), txtMBMPayorExpDate)
        If Trim(SuppPlanCode) <> "" And Trim(cboINSCode) <> "" Then
            If IsDate(txtMBMPayorEffDate) And IsDate(txtMBMPayorExpDate) Then
                TempSuppExpDate = GetSuppExpDate(lstCovAdd.ListItems(listloop).SubItems(3), SuppPlanCode, CDate(tmpSuppEffDate), CDate(TempSuppExpDate))
            End If
            
            TempLifeTimeStatus = ""
            PrevLifeMBMNo = ""
            SptLmtInd = "A"
            AnnLmt.MaxRecords = 1
            StrSqlAnnLmt = "Select SuppLimitStatus, PLNCode, SuppLimit from AnnualLimit where PLNCode = '" & Trim(SuppPlanCode) & _
                           "' ORDER BY AnnualEffDate DESC"
            AnnLmt.Open StrSqlAnnLmt, medixmasterconn, adOpenForwardOnly, adLockReadOnly
            Do While Not AnnLmt.EOF
                SptLmtInd = IIf(Len(Trim(AnnLmt!SuppLimitStatus)), Trim(AnnLmt!SuppLimitStatus), "A")
                If Trim(SptLmtInd) = "B" And Format(LastCoveredID, "00") = "01" Or Trim(SptLmtInd) = "C" Then
                    tmpAnnLmt = IIf(IsNumeric(AnnLmt!SuppLimit), AnnLmt!SuppLimit, 0)
                End If
                AnnLmt.MoveNext
            Loop
            AnnLmt.Close
            Set AnnLmt = Nothing
'---------- Life Time Limit
            strsqlPLN = "select PLNLifeTimeStatus from PLN where PLNCode = '" & Trim(SuppPlanCode) & "'"
            PLN.Open strsqlPLN, medixmasterconn, adOpenKeyset, adLockOptimistic
            If PLN.recordCount Then
                TempLifeTimeStatus = IIf(Len(PLN!PLNLifeTimeStatus), PLN!PLNLifeTimeStatus, "")
            End If
            PLN.Close
            Set PLN = Nothing
            If Trim(CboGuaRenewal) <> "" Then
                If Trim(TempLifeTimeStatus) = "Y" Then
                    tmpKidneyAnnLmt = 0
                    tmpCancerAnnLmt = 0
                    tmpHomeNursAnnLmt = 0
                    Call GetBenefitLimit(Trim(cboINSCode), Trim(cboPLNCode), Trim(cboHLTCode), CDate(txtMBMPayorEffDate), tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt, tmpSuppBNFStatus)
                    Call PrincipalLifeTimeChecking
                    If (tmpSuppBNFStatus = "B" And LastCoveredID = "01") Or tmpSuppBNFStatus = "C" Then
                        Call SaveMBMBnfAnnLmt(Trim(txtMBMNumber), LastCoveredID, tmpSuppBNFStatus, tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt)
                    End If
                    If Trim(SptLmtInd) = "B" And Format(LastCoveredID, "00") = "01" Or Trim(SptLmtInd) = "C" Then
                        tmpCNextLmt = tmpAnnLmt
                        TmpSql2 = "Select top 1 MBMCNumber,MBMCAvailableLimit From MBMCoveredPersons where MBMCNumber='" & Trim(PrevLifeMBMNo) & _
                                 "' AND MBMCCoverID='" & Trim(LastCoveredID) & "'"
                        tmpMBM.Open TmpSql2, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                        If tmpMBM.EOF = False Then
                            If Mid(CboGuaRenewal, 1, 1) <> "3" Then
                                tmpCNextLmt = IIf(IsNumeric(tmpMBM!MBMCAvailableLimit), tmpMBM!MBMCAvailableLimit, 0)
                            End If
                        End If
                        tmpMBM.Close
                        Set tmpMBM = Nothing
                    
                        If WLifeStatus Then
                            Call SaveMBMLifeTimeLmt(PrevLifeMBMNo, LastCoveredID, SptLmtInd, tmpAnnLmt, tmpCNextLmt)
                        End If
                        
                    End If
                End If
            End If
'---- End of Life Time Limit
            SptMCOInd = "A"
            McoCovIDChrg = ""
            strSptMCOSql = "Select SuppMCOStatus, CovIDChrg from MCO where PLNCode = '" & Trim(SuppPlanCode) & _
                   "' AND HLTCode='" & Trim(cboHLTCode) & "' AND INSCode='" & Trim(cboINSCode) & _
                   "' AND MCOEffdate <= '" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"

            SptPlnMCO.Open strSptMCOSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
            If SptPlnMCO.recordCount > 0 Then
            Else
                SptPlnMCO.Close
                strSptMCOSql = "Select SuppMCOStatus, CovIDChrg from MCO where PLNCode is null " & _
                               " AND HLTCode='" & Trim(cboHLTCode) & "' AND INSCode='" & Trim(cboINSCode) & _
                               "' AND MCOEffdate <= '" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
                SptPlnMCO.Open strSptMCOSql, medixmasterconn, adOpenKeyset, adLockOptimistic
            End If
            
            If Trim(SptPlnMCO!SuppMCOStatus) <> "" Then
                SptMCOInd = Trim(SptPlnMCO!SuppMCOStatus)
            End If
            If Trim(SptPlnMCO!CovIDChrg) <> "" Then
                McoCovIDChrg = Trim(SptPlnMCO!CovIDChrg)
            End If
            
            SptPlnMCO.Close
            Set SptPlnMCO = Nothing

            If (Trim(SptMCOInd) = "B" And Trim(LastCoveredID) = "01") _
            Or Trim(SptMCOInd) = "C" _
            Or (Trim(SptMCOInd) = "D" And Trim(LastCoveredID) >= Trim(McoCovIDChrg)) Then
                TmpBasicMCO = Format(GetSuppMCO(Trim(cboINSCode), Trim(cboHLTCode), tmpAdultChild, CDate(tmpSuppEffDate), Trim(SuppPlanCode)), "#,##0.00")
                tmpMCO = InsertMCOFees(CDate(TempSuppExpDate), CDate(tmpSuppEffDate), Format(TmpBasicMCO, "#,##0.00"))
                tmpMCO = Round(tmpMCO)
            End If

            strsqlPrem = "Select SuppPrmStatus, PLNCode, SuppPrmAmt from PRM where PLNCode = '" & Trim(SuppPlanCode) & "' And AgeCode = '" & Trim(AgeValue) & "' And INSCode = '" & Mid(cboINSCode, 1, 1) & "'"
            Prem.Open strsqlPrem, medixmasterconn, adOpenKeyset, adLockOptimistic

            If Prem.recordCount Then
                If Trim(Prem!SuppPrmStatus) = "B" And Format(LastCoveredID, "00") = "01" Or Trim(Prem!SuppPrmStatus) = "C" Then

                     tempSPremium = Prem!SuppPrmAmt
                     TmpBPremium = IIf(IsNumeric(Prem!SuppPrmAmt), Prem!SuppPrmAmt, 0)
                     tempSPremium = Format(GetSPremium(Mid(cboINSCode, 1, 1), Mid(cboPAYCode, 1, 2), AgeValue, Trim(Mid(cboPLNCode, 1, 4)), Trim(cboHLTCode), CDate(tmpSuppEffDate)), "#,##0.00")
                     tempSPremium = InsertPremium(CDate(TempSuppExpDate), CDate(tmpSuppEffDate), Format(tempSPremium, "#,##0.00"))
                     
                     If Disc = True Then
                         If lstCovAdd.ListItems(listloop).SubItems(38) <> "" Then
                             tempSPremium = tempSPremium * (100 - CDbl(lstCovAdd.ListItems(listloop).SubItems(38))) / 100
                         End If
                         If lstCovAdd.ListItems(listloop).SubItems(39) <> "" Then
                             tempSPremium = tempSPremium * (100 - CDbl(lstCovAdd.ListItems(listloop).SubItems(39))) / 100
                         End If
                     End If
                     tmpPremium = tempSPremium
                End If
            End If
            Prem.Close
            Set Prem = Nothing
        End If
        tmpAnnLmt = CDbl(IIf(IsNumeric(tmpAnnLmt) = True, tmpAnnLmt, 0))
        TmpBPremium = CDbl(IIf(IsNumeric(TmpBPremium) = True, TmpBPremium, 0))
        tmpPremium = CDbl(IIf(IsNumeric(tmpPremium) = True, tmpPremium, 0))
        tmpREwDisc = lstCovAdd.ListItems(listloop).SubItems(38)
        tmpREwDisc = CDbl(IIf(IsNumeric(tmpREwDisc) = True, tmpREwDisc, 0))
        tmpPolDisc = lstCovAdd.ListItems(listloop).SubItems(39)
        tmpPolDisc = CDbl(IIf(IsNumeric(tmpPolDisc) = True, tmpPolDisc, 0))
        TmpMemberType = "C"
        TempMBMCName = Replace(lstCovAdd.ListItems(listloop).Text, "'", "''")
        tmpICNo = Trim(lstCovAdd.ListItems(listloop).SubItems(2))
        TmpBordxDate = Format(lstCovAdd.ListItems(listloop).SubItems(33), "yyyy/mm/dd")
        tmpDOb = Format(lstCovAdd.ListItems(listloop).SubItems(3), "yyyy/mm/dd")
        tmpSex = lstCovAdd.ListItems(listloop).SubItems(4)
        TmpBasicMCO = CDbl(IIf(IsNumeric(TmpBasicMCO) = True, TmpBasicMCO, 0))
        tmpMCO = CDbl(IIf(IsNumeric(tmpMCO) = True, tmpMCO, 0))
        tmpRELCode = Mid(lstCovAdd.ListItems(listloop).SubItems(6), 1, 1)
        LastCoveredID = Format(LastCoveredID, "00")
        lstCovAdd.ListItems(listloop).SubItems(11) = ""
        lstCovAdd.ListItems(listloop).SubItems(10) = LastCoveredID
        TempSuppExpDate = IIf(IsDate(TempSuppExpDate), Format(TempSuppExpDate, "yyyy/mm/dd"), "")
        tmpSuppEffDate = IIf(IsDate(tmpSuppEffDate), Format(tmpSuppEffDate, "yyyy/mm/dd"), "")
        
        Dim AgeLmt As String
        AgeLmt = ""
        AgeLmt = Mid(cmbAgeLmt, 1, 1)
        
        If Trim(cboPLNCode) <> "" Then
            MBMCovSql = "Insert into MBMCoveredPersons(MBMCNumber, MBMCCoverID, MBMCAgeband, MBMCStatus, MBMCBordxDate, MBMCBatchNo," & _
                        "MBMCName, MBMCICBCPPNo, MBMCDOB, MBMCSex, MBMCAdultChild, MBMCRELCode, " & _
                        "MBMCMemberType, MBMCRegDate,  MBMCPLNCode, MBMCEffDate, MBMCExpDate, " & _
                        "MBMCAnnualLimit, MBMCAvailableLimit, MBMCBasicPremium, MBMCPremium, " & _
                        "MBMCBasicMCO, MBMCMCO,MBMHRAppInd) " & _
                        "Values('" & Trim(txtMBMNumber) & "','" & LastCoveredID & "','" & AgeValue & "','" & Mid(cboStatus, 1, 1) & "','" & TmpBordxDate & "','" & Trim(txtBatchNo) & "'," & _
                        "'" & Trim(TempMBMCName) & "','" & tmpICNo & "','" & tmpDOb & "','" & tmpSex & "','" & tmpAdultChild & "','" & tmpRELCode & "'," & _
                        "'" & TmpMemberType & "','" & Format(tmpServerDate, "yyyy/mm/dd") & "','" & SuppPlanCode & "','" & tmpSuppEffDate & "','" & TempSuppExpDate & "'," & _
                        "" & tmpAnnLmt & "," & tmpAnnLmt & "," & TmpBPremium & "," & tmpPremium & "," & _
                        "" & TmpBasicMCO & "," & tmpMCO & ",'" & AgeLmt & "')"
        Else
            MBMCovSql = "Insert into MBMCoveredPersons(MBMCNumber, MBMCCoverID, MBMCAgeband, MBMCStatus, MBMCBordxDate, MBMCBatchNo," & _
                        "MBMCName, MBMCICBCPPNo, MBMCDOB, MBMCSex, MBMCAdultChild, MBMCRELCode, " & _
                        "MBMCMemberType, MBMCRegDate,  MBMCPLNCode, " & _
                        "MBMCAnnualLimit, MBMCAvailableLimit, MBMCBasicPremium, MBMCPremium, " & _
                        "MBMCBasicMCO, MBMCMCO,MBMHRAppInd) " & _
                        "Values('" & Trim(txtMBMNumber) & "','" & LastCoveredID & "','" & AgeValue & "','" & Mid(cboStatus, 1, 1) & "','" & TmpBordxDate & "','" & Trim(txtBatchNo) & "'," & _
                        "'" & Trim(TempMBMCName) & "','" & tmpICNo & "','" & tmpDOb & "','" & tmpSex & "','" & tmpAdultChild & "','" & tmpRELCode & "'," & _
                        "'" & TmpMemberType & "','" & Format(tmpServerDate, "yyyy/mm/dd") & "','" & SuppPlanCode & "'," & _
                        "" & tmpAnnLmt & "," & tmpAnnLmt & "," & TmpBPremium & "," & tmpPremium & "," & _
                        "" & TmpBasicMCO & "," & tmpMCO & ",'" & AgeLmt & "')"
        End If
        medixmasterconn.Execute MBMCovSql
        
        TmpPayPremNet = lstCovAdd.ListItems(listloop).SubItems(35)
        TmpPayPremNet = CDbl(IIf(IsNumeric(TmpPayPremNet) = True, TmpPayPremNet, 0))
        TmpPayMCOGross = lstCovAdd.ListItems(listloop).SubItems(36)
        TmpPayMCOGross = CDbl(IIf(IsNumeric(TmpPayMCOGross) = True, TmpPayMCOGross, 0))
        TmpPayMcoNet = lstCovAdd.ListItems(listloop).SubItems(37)
        TmpPayMcoNet = CDbl(IIf(IsNumeric(TmpPayMcoNet) = True, TmpPayMcoNet, 0))
        TmpPolAmt = lstCovAdd.ListItems(listloop).SubItems(40)
        TmpPolAmt = CDbl(IIf(IsNumeric(TmpPolAmt) = True, TmpPolAmt, 0))
        tmpRewAmt = lstCovAdd.ListItems(listloop).SubItems(41)
        tmpRewAmt = CDbl(IIf(IsNumeric(tmpRewAmt) = True, tmpRewAmt, 0))
        TmpLoadingPrem = lstCovAdd.ListItems(listloop).SubItems(42)
        TmpLoadingPrem = CDbl(IIf(IsNumeric(TmpLoadingPrem) = True, TmpLoadingPrem, 0))
        tmpAllergies = Replace(lstCovAdd.ListItems(listloop).SubItems(19), "'", "''")
        TmpExclusion = Replace(lstCovAdd.ListItems(listloop).SubItems(9), "'", "''")
        tmpUndExcess = lstCovAdd.ListItems(listloop).SubItems(15)
        tmpUndExcess = CDbl(IIf(IsNumeric(tmpUndExcess) = True, tmpUndExcess, 0))
        tmpClmNonPayFr = lstCovAdd.ListItems(listloop).SubItems(31)
        tmpClmNonPayTo = lstCovAdd.ListItems(listloop).SubItems(32)
        If Len(tmpClmNonPayFr) And Len(tmpClmNonPayTo) Then
            MBMCovTwoSql = "Insert into MBMCoveredPersonsTwo(MBMCNumber, MBMCCoverID, MBMCSalutation, " & _
                           "MBMCAllergic, MBMCExclusion, " & _
                           "MBMCPayorPremNet, MBMCPayorMCOGross, MBMCPayorMCONet, MBMCLoadingPrem, MBMCFamDiscAmt, " & _
                           "MBMCRenewDiscAmt, MBMCUndExcess, MBMCRenewalDisc, MBMCPolicyDisc, MBMCClmNonPayFr, MBMCClmNonPayTo)" & _
                           "Values('" & Trim(txtMBMNumber) & "','" & LastCoveredID & "','" & lstCovAdd.ListItems(listloop).SubItems(1) & "'," & _
                           "'" & tmpAllergies & "','" & TmpExclusion & "'," & _
                           "" & TmpPayPremNet & "," & TmpPayMCOGross & "," & TmpPayMcoNet & "," & TmpLoadingPrem & "," & TmpPolAmt & "," & _
                           "" & tmpRewAmt & "," & tmpUndExcess & "," & tmpREwDisc & "," & tmpPolDisc & "," & _
                           "'" & Format(tmpClmNonPayFr, "yyyy/mm/dd") & "','" & Format(tmpClmNonPayTo, "yyyy/mm/dd") & "')"
        Else
            MBMCovTwoSql = "Insert into MBMCoveredPersonsTwo(MBMCNumber, MBMCCoverID, MBMCSalutation, " & _
                           "MBMCAllergic, MBMCExclusion, " & _
                           "MBMCPayorPremNet, MBMCPayorMCOGross, MBMCPayorMCONet, MBMCLoadingPrem, MBMCFamDiscAmt, " & _
                           "MBMCRenewDiscAmt, MBMCUndExcess, MBMCRenewalDisc, MBMCPolicyDisc)" & _
                           "Values('" & Trim(txtMBMNumber) & "','" & LastCoveredID & "','" & lstCovAdd.ListItems(listloop).SubItems(1) & "'," & _
                           "'" & tmpAllergies & "','" & TmpExclusion & "'," & _
                           "" & TmpPayPremNet & "," & TmpPayMCOGross & "," & TmpPayMcoNet & "," & TmpLoadingPrem & "," & TmpPolAmt & "," & _
                           "" & tmpRewAmt & "," & tmpUndExcess & "," & tmpREwDisc & "," & tmpPolDisc & ")"
        End If
        medixmasterconn.Execute MBMCovTwoSql
        
        If Trim(txtClnInsType) <> "" And Trim(txtClnPlanCode) <> "" Then
            Call SaveCISCovPerson(LastCoveredID, lstCovAdd.ListItems(listloop).SubItems(3), lstCovAdd.ListItems(listloop).Text)
        End If

        LastCoveredID = LastCoveredID + 1
    End If
    Next listloop
End Sub

Private Sub cboCovRelationship_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboCovRelationship, KeyAscii)
End Sub

Private Sub cboCovSalutation_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboCovSalutation, KeyAscii)
End Sub

Private Sub cboCovSex_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboCovSex, KeyAscii)
End Sub

Private Sub cboINSCode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboINSCode, KeyAscii)
End Sub

Private Sub cboPLNCode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboPLNCode, KeyAscii)
End Sub

Private Sub cboRACCode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboRACCode, KeyAscii)
End Sub

Private Sub cboSRVCode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboSRVCode, KeyAscii)
End Sub

Private Sub cboMBMSalutation_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboMBMSalutation, KeyAscii)
End Sub

Private Sub cboMBMSex_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboMBMSex, KeyAscii)
End Sub

Private Sub cboNATCode_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboNATCode, KeyAscii)
End Sub

Private Sub cboMBMMaritalStatus_KeyPress(KeyAscii As Integer)
    Call ComboKeyPress(cboMBMMaritalStatus, KeyAscii)
End Sub

Private Sub txtMBMPayorExpDate_LostFocus()
Dim MBMValidDate As New ADODB.Recordset
Dim strsql As String
Dim CheckDays As Integer
Dim Confirm1
Dim Confirm2
Dim ExpStart As String
Dim ExpEnd As String
Dim DateSql As String
    
    ExpStart = ""
    ExpEnd = ""
    CheckDays = 0
    If IsDate(txtMBMPayorExpDate) = False Then
        MsgBox "Invalid Date Format! ", vbCritical
        txtMBMPayorExpDate.SetFocus
    Else
        'If Mid(cboEndorseStatus, 1, 1) <> "I" Then
        DateSql = " MBMExpStart as DateStart, MBMExpEnd as DateEnd "
        Call GetValidDate(ExpStart, ExpEnd, DateSql)
        If Trim(ExpStart) <> "" And Trim(ExpEnd) <> "" Then
            If DateDiff("d", ExpStart, txtMBMPayorExpDate) >= 0 And DateDiff("d", ExpEnd, txtMBMPayorExpDate) <= 0 Then
                CheckDays = DateDiff("d", txtMBMPayorEffDate, txtMBMPayorExpDate)
                If CheckDays > 365 Then
                    Confirm1 = MsgBox("Policy more than 365 days. Do you want to continue?", vbCritical + vbYesNo)
                    If Confirm1 = vbNo Then
                        txtMBMPayorExpDate.SetFocus
                    End If
                ElseIf CheckDays <= 30 Then
                    Confirm2 = MsgBox("Policy less than 30 days. Do you want to continue?", vbCritical + vbYesNo)
                    If Confirm2 = vbNo Then
                        txtMBMPayorExpDate.SetFocus
                    End If
                End If
                Call MemberCallAge
                If Len(Trim(txtAGECode)) <> 0 Then
                    Call MemberCallFees
                End If
            Else
                MsgBox "Policy Expiry Date is out of acceptable range! ", vbCritical + vbOKOnly
                txtMBMPayorExpDate.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtMBMPolNo_LostFocus()
    Dim CheckRecord
    Dim MBMPOLNO As New ADODB.Recordset
    Dim strsql As String
    Dim msgstr As String
    Dim RecordConfirm
    Dim CheckFlag As Boolean
    
    'medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    
    If Len(Trim(txtMBMPolNo)) Then
        Screen.MousePointer = vbHourglass
        If CheckFlag = False Then
            strsql = "Select MBMNumber From MBMCrossReference Where MBMPolicyNo ='" & Trim(txtMBMPolNo) & "'"
            MBMPOLNO.Open strsql, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            
           If MBMPOLNO.recordCount >= 1 Then
                If Trim(MBMPOLNO!MBMNumber) <> Trim(txtMBMNumber) Then
                    msgstr = "Duplicate Policy No In The Membership" & Chr(13) & "Membership No: " & MBMPOLNO!MBMNumber & vbNewLine & _
                            " Do you want to accept it? "
                    RecordConfirm = MsgBox(msgstr, vbCritical + vbYesNoCancel, DialogBoxTitle)
                    If RecordConfirm = vbNo Then
                        txtMBMPolNo.SetFocus
                        txtMBMPolNo.SelStart = 0
                        CheckFlag = False
                    ElseIf RecordConfirm = vbCancel Then
                        txtMBMPolNo = ""
                        txtMBMPolNo.SetFocus
                        CheckFlag = False
                    Else
                        CheckFlag = True
                    End If
                End If
            End If
            MBMPOLNO.Close
            Set MBMPOLNO = Nothing
        End If
        Screen.MousePointer = vbDefault
    End If
   ' medixmasterconnMaint.Close
   ' Set medixmasterconnMaint = Nothing
End Sub

Private Sub InsertMCO()
Dim MBMInsertMCOFees As Double
    If txtTempMCOFees <> 0 Then
        MBMInsertMCOFees = Format(txtTempMCOFees, "#,##0.00")
        If PLNProRate <> "N" Then
            MBMInsertMCOFees = InsertMCOFees(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(txtTempMCOFees, "#,##0.00"))
        End If
        If RoundUpStatus <> "N" Then
            MBMInsertMCOFees = Round(MBMInsertMCOFees)
        End If
        txtMCOFees = ""
        txtMCOFees = Format(Round(MBMInsertMCOFees, 0), "#,##0.00")
    End If
End Sub

Private Sub InsertPrem()
Dim MBMInsertPrem As Double
Dim MBMPrem As Double
    
    If IsNumeric(PolDics) = False Then
        PolDics = 0
    End If
    
    If IsNumeric(RenDisc) = False Then
        RenDisc = 0
    End If
    
    If Trim(cboPLNCode) <> "" And Trim(txtMBMPremium) <> "" Then
        MBMInsertPrem = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(txtTempPrem, "#,##0.00"))
        txtMBMPremium = ""
        txtMBMPremium = Format(Round(MBMInsertPrem, 0), "#,##0.00")
            strsqlPLN = "select * from DiscPlan where PLNCode = '" & Trim(Mid(cboPLNCode, 1, 4)) & "'"
            PLN.Open strsqlPLN, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            If PLN.recordCount Then
                If PolDics > 0 Then
                    txtMBMPremium = txtMBMPremium * (100 - PolDics) / 100
                End If
                If RenDisc > 0 Then
                    txtMBMPremium = txtMBMPremium * (100 - RenDisc) / 100
                End If
            End If
            PLN.Close
            Set PLN = Nothing
    End If
End Sub

Private Sub ClearEntries()
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        ElseIf TypeOf ctl Is ComboBox Then
            'ctl. = ""
            cboHLTCode.Text = ""
            cboCTYCode.Text = ""
            cboSTACode.Text = ""
            cboMBMSex.Text = ""
            cboMBMMaritalStatus.Text = ""
            cboRACCode.Text = ""
            cboNATCode.Text = ""
            cboMBMRenewal.Text = ""
            cboINSCode.Text = ""
            cboPLNCode.Text = ""
            CboInsMode.Text = ""
            CboGuaRenewal.Text = ""
        ElseIf TypeOf ctl Is ListView Then
            ctl.ListItems.Clear
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.mask = ""
            ctl.Text = ""
            ctl.mask = "##/##/####"
        ElseIf TypeOf ctl Is Label Then
            txtMBMPremium = ""
            txtMCOFees = ""
        End If
    Next ctl
    '  Call ComboListing
      'cmdSaveKIV.Enabled = True
      'cmdGenerate.Enabled = True
End Sub

Private Sub DelSuppLoop()
Dim listloop As Integer
Dim ListCount As Integer
    
    ListCount = lstCovAdd.ListItems.Count
    For listloop = 1 To lstCovAdd.ListItems.Count
        If listloop > ListCount Then
            Exit Sub
        ElseIf listloop <= ListCount Then
            If lstCovAdd.ListItems(listloop).SubItems(11) = "R" Then
                lstCovAdd.ListItems.Remove (listloop)
                ListCount = ListCount - 1
            End If
        End If
    Next listloop
End Sub

Private Sub cboCTYCode_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyCharacter(KeyAscii)
  
End Sub

Private Sub cboCTYCode_LostFocus()
    If Len(cboCTYCode) > 50 Then
        MsgBox "City/Town Exceeded 50 Characters!", vbCritical
        cboCTYCode.SetFocus
    End If
End Sub

Private Sub cboSTACode_LostFocus()
    If Len(cboSTACode) > 50 Then
        MsgBox "State Exceeded 50 Characters!", vbCritical
        cboSTACode.SetFocus
    End If
End Sub


Private Sub cboSTACode_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyCharacter(KeyAscii)
End Sub

Private Sub cboBRNCode_LostFocus()
    If Len(cboBRNCode) > 50 Then
        MsgBox "Branch Exceeded 50 characters!", vbCritical
        cboBRNCode.SetFocus
    End If
    cboBRNCode = UCase(cboBRNCode)
End Sub

Private Sub cboBRNCode_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyCharacter(KeyAscii)
End Sub

Private Sub txtMgmtNotes_Click()
    KeyPreview = False
End Sub

Private Sub txtMgmtNotes_LostFocus()
    KeyPreview = True
End Sub

Private Sub txtPolicyDisc_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtPolicyDiscAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtRBAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtRemarks_Click()
    KeyPreview = False
End Sub

Private Sub txtRemarks_LostFocus()
    KeyPreview = True
    txtRemarks = ChkStr(txtRemarks)
End Sub

Private Sub txtRenewalDisc_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtRenewalDiscAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSuppGenWP_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSuppGenWP_LostFocus()
    If txtSuppGenWP <> "" Then
        If txtSuppGenWP > 0 Then
            If txtMBMPayorEffDate <> "__/__/____" Then
                txtSuppGENWPEnd = Format(DateAdd("d", Trim(txtSuppGenWP), txtMBMPayorEffDate), "dd/mm/yyyy")
            End If
        Else
            txtSuppGENWPEnd.mask = ""
            txtSuppGENWPEnd = ""
            txtSuppGENWPEnd.mask = "##/##/####"
        End If
    Else
        txtSuppGENWPEnd.mask = ""
        txtSuppGENWPEnd = ""
        txtSuppGENWPEnd.mask = "##/##/####"
    End If
End Sub

Private Sub txtSuppPayMCOGross_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSuppPayMCONet_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSuppPayPremGross_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSuppPayPremNet_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSuppPLNCode_LostFocus()
Dim TempSuppExpDate As String
    txtSuppPLNCode = IIf(Trim(txtSuppPLNCode) <> "", txtSuppPLNCode, cboPLNCode)
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    If Trim(txtSuppPLNCode) <> "" And IsDate(txtCovBirthdate) = True And IsDate(txtSuppEffDate) And IsDate(txtSuppExpDate) Then
        TempSuppExpDate = GetSuppExpDate(txtCovBirthdate, txtSuppPLNCode, txtSuppEffDate, txtSuppExpDate)
        txtSuppExpDate = TempSuppExpDate
    End If
'    medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    If OverAge = True Then
        MsgBox "Supplementary age is above " & PlnAgeLmt
    End If
End Sub

Private Sub txtSuppPREWP_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtSuppPREWP_LostFocus()
    If txtSuppPREWP <> "" Then
        If Trim(txtSuppPREWP) > 0 Then
            If txtMBMPayorEffDate <> "__/__/____" Then
                txtSuppPREWPEnd = Format(DateAdd("d", Trim(txtSuppPREWP), txtMBMPayorEffDate), "dd/mm/yyyy")
            End If
        Else
            txtSuppPREWPEnd.mask = ""
            txtSuppPREWPEnd = ""
            txtSuppPREWPEnd.mask = "##/##/####"
        End If
    Else
        txtSuppPREWPEnd.mask = ""
        txtSuppPREWPEnd = ""
        txtSuppPREWPEnd.mask = "##/##/####"
    End If
End Sub

Private Sub TxtSuppSpeWP_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub TxtSuppSpeWP_LostFocus()
    If TxtSuppSpeWP <> "" Then
        If Trim(TxtSuppSpeWP) > 0 Then
            If txtMBMPayorEffDate <> "__/__/____" Then
                txtSuppSPEWPEnd = Format(DateAdd("d", Trim(TxtSuppSpeWP), txtMBMPayorEffDate), "dd/mm/yyyy")
            End If
        Else
            txtSuppSPEWPEnd.mask = ""
            txtSuppSPEWPEnd = ""
            txtSuppSPEWPEnd.mask = "##/##/####"
        End If
    Else
        txtSuppSPEWPEnd.mask = ""
        txtSuppSPEWPEnd = ""
        txtSuppSPEWPEnd.mask = "##/##/####"
    End If
End Sub

Private Sub TxtSuppClmNonPayFr_LostFocus()
    If TxtSuppClmNonPayFr <> "__/__/____" Then
        If IsDate(TxtSuppClmNonPayFr) = False Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtSuppClmNonPayFr.SetFocus
        End If
    End If
End Sub

Private Sub TxtSuppClmNonPayTo_LostFocus()
    If TxtSuppClmNonPayTo <> "__/__/____" Then
        If IsDate(TxtSuppClmNonPayTo) = False Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtSuppClmNonPayTo.SetFocus
        End If
    End If
End Sub

Public Function showMBMOthers() As Integer
Dim sqlstr As String
Dim MBM As New ADODB.Recordset
Dim strCount As String
Dim rstCount As New ADODB.Recordset
Dim MBMNumber As String
Dim TempMBMOthers As New ADODB.Recordset
Dim strMBMOthers As String
    
    With ADOrstMBM
        
        frmNewMBMOthers.Label4 = Trim(MBMNumber)
        If Trim(!MBMCLMNonPayFr) <> "" Then
            frmNewMBMOthers.TxtClmNonPayFr = Format(!MBMCLMNonPayFr, "dd/mm/yyyy")
        End If
        If Trim(!MBMCLMNonPayTo) <> "" Then
            frmNewMBMOthers.TxtClmNonPayTo = Format(!MBMCLMNonPayTo, "dd/mm/yyyy")
        End If
        
        If Trim(!LabelStatus) = "Y" Then
            frmNewMBMOthers.Check1.Value = 1
        End If
        txtMBMPremium = 0
        If Len(Trim(cboPAYCode)) <> 0 And Len(cboINSCode) <> 0 And txtAGECode <> "" And Len(cboHLTCode) <> 0 And txtMBMPayorEffDate <> "__/__/____" Then
            txtMBMPremium = IIf(IsNumeric(!MBMPremium), Format(!MBMPremium, "#,##0.00"), 0)
        End If
        txtMCOFees = 0
        If Len(Trim(cboPAYCode)) <> 0 And Len(cboINSCode) <> 0 And txtAGECode <> "" And Len(cboHLTCode) <> 0 And txtMBMPayorEffDate <> "__/__/____" Then
            txtMCOFees = Format(IIf(IsNull(!MBMMCOFee) = True, Format(GetMCO(Trim(cboINSCode), Trim(cboHLTCode), txtMBMAdultChild, txtMBMPayorEffDate, Mid(cboPLNCode, 1, 4)), "#,##0.00"), !MBMMCOFee), "#,##0.00")
        End If
        
        txtTempMCOFees = txtMCOFees
        txtTempPrem = txtMBMPremium
        txtMBMAgentCode = IIf(Len(!MBMAgentCode), Trim(!MBMAgentCode), "")
        txtDeptName = IIf(Len(!MBMDepartment), Trim(!MBMDepartment), "")
        txtMBMGroupCompany = IIf(Len(!MBMGroupCompany), Trim(!MBMGroupCompany), "")
        txtSubsidiary = IIf(Len(!Subsidiary), Trim(!Subsidiary), "")
        txtMBMEmployeeNo = IIf(Len(!MBMEmployeeNo), Trim(!MBMEmployeeNo), "")
        CboInsMode = IIf(Len(!MBMInstallment), Trim(!MBMInstallment), "")
        If Trim(!MBMFirstJoinedDate) <> "" Then
            txtMBMDateJoin = Format(!MBMFirstJoinedDate, "dd/mm/yyyy")
        End If
        If Len(!MBMGuaRenewal) Then
            If Trim(!MBMGuaRenewal) = "1" Then
                CboGuaRenewal.ListIndex = 1
            ElseIf Trim(!MBMGuaRenewal) = "3" Then
                CboGuaRenewal.ListIndex = 2
            End If
        End If
        frmNewMBMOthers.TxtPOLSubNo = IIf(Len(!MBMPolSubNo1), Trim(!MBMPolSubNo1), "")
        frmNewMBMOthers.txtPLNCode = IIf(Len(!MBMNewPlanCode), Trim(!MBMNewPlanCode), "")
        If Len(!MBMNewPolEffDate) Then
            frmNewMBMOthers.txtNewPolEffDate = Format(!MBMNewPolEffDate, "dd/mm/yyyy")
        End If
        frmNewMBMOthers.txtPolicyDisc = IIf(IsNumeric(!MBMPolicyDisc), !MBMPolicyDisc, 0)
        frmNewMBMOthers.txtRenewalDisc = IIf(IsNumeric(!MBMRenewalDisc), !MBMRenewalDisc, 0)
        PolDics = IIf(IsNumeric(!MBMPolicyDisc), !MBMPolicyDisc, 0)
        RenDisc = IIf(IsNumeric(!MBMRenewalDisc), !MBMRenewalDisc, 0)
        frmNewMBMOthers.cboPOLCondition = IIf(Len(!MBMPolCondition), Trim(!MBMPolCondition), "")
        frmNewMBMOthers.txtUndExcess = IIf(IsNumeric(!MBMUndExcess), !MBMUndExcess, 0)
        frmNewMBMOthers.txtEndtRefNo = IIf(Len(!MBMENDtRefNo), !MBMENDtRefNo, "")
        txtCostCenter = IIf(Len(!MBMCostCenter), !MBMCostCenter, "")
        txtMBMPosition = IIf(Len(!MBMPosition), !MBMPosition, "")
        txtGrpCat = IIf(Len(!MBMGroupCategory), !MBMGroupCategory, "")
        txtDivision = IIf(Len(!MBMDivision), !MBMDivision, "")
        txtSubDivision = IIf(Len(!MBMSubDivision), !MBMSubDivision, "")

    End With
End Function

Public Function showMBMOthersTwo() As Integer
Dim sqlstr As String
Dim MBM As New ADODB.Recordset
Dim strCount As String
Dim rstCount As New ADODB.Recordset
Dim MBMNumber As String
Dim TempMBMOthers As New ADODB.Recordset
Dim strMBMOthers As String
    

    If frmNewMBMOthers.Visible = False Then
        Load frmNewMBMOthers
    End If
    MBMNumber = txtMBMNumber
        
        With ADOrstMBM
            
            frmNewMBMOthers.Label4 = Trim(MBMNumber)
            If Len(!MBMRemarks) Then
                txtRemarks = Trim(!MBMRemarks)
            End If
            If Trim(!MBMAllergic) <> "" Then
                frmNewMBMOthers.txtMBMAllergic = Trim(!MBMAllergic)
            Else
                frmNewMBMOthers.txtMBMAllergic = ""
            End If

            If Len(!MBMPrevPLNCode) Then
                frmNewMBMOthers.txtMBMPrevPlan = !MBMPrevPLNCode
            Else
                frmNewMBMOthers.txtMBMPrevPlan = ""
            End If
            
            If Len(!MBMRBAmt) Then
                frmNewMBMOthers.txtMBMRBAmt = !MBMRBAmt
            Else
            
                frmNewMBMOthers.txtMBMRBAmt = ""
            End If
            
            'If Len(!MBMPAYRemarks) Then
            '    frmNewMBMOthers.txtMBMPAYRemarks = !MBMPAYRemarks
            'End If
                        
            If Len(!MBMPrevInsCom) Then
                frmNewMBMOthers.txtMBMPrevINSCom = !MBMPrevInsCom
            Else
                frmNewMBMOthers.txtMBMPrevINSCom = ""
            End If
        
            If Len(!MBMCoPayRB) Then
                frmNewMBMOthers.cboCOPayRB = !MBMCoPayRB
            Else
                frmNewMBMOthers.cboCOPayRB = ""
            End If
            
            If Len(!MBMGenWP) Then
                frmNewMBMOthers.TxtGenWP = !MBMGenWP
            Else
                frmNewMBMOthers.TxtGenWP = ""
            End If
            
            If Len(!MBMPreExWP) Then
                frmNewMBMOthers.txtPREWP = !MBMPreExWP
            Else
                frmNewMBMOthers.txtPREWP = ""
            End If
            
            If Len(!MBMSpeIllWP) Then
                frmNewMBMOthers.TxtSpeWP = !MBMSpeIllWP
            Else
                frmNewMBMOthers.TxtSpeWP = ""
            End If
            
            If Len(!MBMStaffDec) Then
                frmNewMBMOthers.cboStaffDisc = !MBMStaffDec
            Else
                frmNewMBMOthers.cboStaffDisc = ""
            End If
            
            If Len(!MBMGenEndWP) Then
                frmNewMBMOthers.txtGENWPEnd = !MBMGenEndWP
            Else
                frmNewMBMOthers.txtGENWPEnd = "__/__/____"
            End If
            
            If Len(!MBMPreExEndWP) Then
                frmNewMBMOthers.txtPREWPEnd = !MBMPreExEndWP
            Else
                frmNewMBMOthers.txtPREWPEnd = "__/__/____"
            End If
            
            If Len(!MBMSpeIllEndWP) Then
                frmNewMBMOthers.txtSPEWPEnd = !MBMSpeIllEndWP
            Else
                frmNewMBMOthers.txtSPEWPEnd = "__/__/____"
            End If
                            
            If Len(!MBMBTBG) Then
                frmNewMBMOthers.cboBTBG = !MBMBTBG
            Else
                frmNewMBMOthers.cboBTBG = ""
            End If
            
            If Len(!MBMOccupation) Then
                frmNewMBMOthers.cboMBMOccupation = !MBMOccupation
            Else
                frmNewMBMOthers.cboMBMOccupation = ""
            End If
            
            If Len(!MBMWorkNature) Then
                frmNewMBMOthers.txtMBMWorkNature = !MBMWorkNature
            Else
                frmNewMBMOthers.txtMBMWorkNature = ""
            End If
            
            If Len(!MBMWeight) Then
                frmNewMBMOthers.txtMBMWeight = !MBMWeight
            Else
                frmNewMBMOthers.txtMBMWeight = ""
            End If
            
            If Len(!MBMHeight) Then
                frmNewMBMOthers.txtMBMHeight = !MBMHeight
            Else
                frmNewMBMOthers.txtMBMHeight = ""
            End If
                        
            If Len(!MBMBlood) Then
                frmNewMBMOthers.cboMBMBlood = !MBMBlood
            Else
                frmNewMBMOthers.cboMBMBlood = ""
            End If
                         
            If Len(!MBMSmoking) Then
                frmNewMBMOthers.cboMBMSmoking = !MBMSmoking
            Else
                frmNewMBMOthers.cboMBMSmoking = ""
            End If
            
            If Len(!MBMAlcohol) Then
                frmNewMBMOthers.cboMBMAlcohol = !MBMAlcohol
            Else
                frmNewMBMOthers.cboMBMAlcohol = ""
            End If
            
            If Len(!MBMPayorPremGross) Then
                frmNewMBMOthers.txtPayPrem = !MBMPayorPremGross
            Else
                frmNewMBMOthers.txtPayPrem = ""
            End If
            
            If Len(!MBMPayorPremNet) Then
                frmNewMBMOthers.txtPayPremNett = !MBMPayorPremNet
            Else
                frmNewMBMOthers.txtPayPremNett = ""
            End If
            
            If Len(!MBMPayorMCOGross) Then
                frmNewMBMOthers.txtPayMCOGross = !MBMPayorMCOGross
            Else
                frmNewMBMOthers.txtPayMCOGross = ""
            End If
            
            If Len(!MBMPayorMCONet) Then
                frmNewMBMOthers.txtPayMCONett = !MBMPayorMCONet
            Else
                frmNewMBMOthers.txtPayMCONett = ""
            End If
            
            If Len(!MBMFamDiscAmt) Then
                frmNewMBMOthers.txtPolicyDiscAmt = !MBMFamDiscAmt
            Else
                frmNewMBMOthers.txtPolicyDiscAmt = ""
            End If
    
            If Len(!MBMRenewDiscAmt) Then
                frmNewMBMOthers.txtRenewalDiscAmt = !MBMRenewDiscAmt
            Else
                frmNewMBMOthers.txtRenewalDiscAmt = ""
            End If
    
            If Len(!MBMLoadingPrem) Then
                frmNewMBMOthers.txtLoadingPrem = !MBMLoadingPrem
            Else
                frmNewMBMOthers.txtLoadingPrem = ""
            End If
            
            If Len(!MBMExclusion) Then
                txtMBMExclusion = !MBMExclusion
            Else
                txtMBMExclusion = ""
            End If
            
            If Len(!MBMBankACNo) Then
                frmNewMBMOthers.txtBankACNo = !MBMBankACNo
            Else
                frmNewMBMOthers.txtBankACNo = ""
            End If
            
            If Len(!MBMTopupJoinDate) Then
                txtMBMTopUpJoinDate = !MBMTopupJoinDate
            Else
                txtMBMTopUpJoinDate = "__/__/____"
            End If
            
        End With
End Function

Private Function showMBMTwo() As Integer
Dim sqlstrTwo As String
Dim MBMTwo As New ADODB.Recordset
    
'    If frmNewMBMOthers.Visible = False Then
'        Load frmNewMBMOthers
'    End If
'    Call ClearForm(frmNewMBMOthers)
    
        With ADOrstMBM
            If Trim(!MBMSalutation) <> "" Then
                cboMBMSalutation = Trim(!MBMSalutation)
            End If
            If Trim(!SRVCodeList) <> "" Then
                cboSRVCode = Trim(!SRVCodeList)
            End If
            If Trim(!NATCode) <> "" Then
                cboNATCode = Trim(!NATCode)
            End If
            If Trim(!MBMTelnoH) <> "" Then
                frmNewMBMOthers.txtMBMTelNoH = Trim(!MBMTelnoH)
            End If
            
            If Trim(!MBMTelNoM) <> "" Then
                frmNewMBMOthers.txtMBMTelNoM = Trim(!MBMTelNoM)
            End If
            
            If Trim(!MBMTelNoO) <> "" Then
                frmNewMBMOthers.txtMBMTelNoO = Trim(!MBMTelNoO)
            End If
            
            If Len(!MBMEmail) Then
                tmpEmail = !MBMEmail
                frmNewMBMOthers.txtMBMEmail = !MBMEmail
            End If
            
            If Len(!LGNCode) Then
                frmNewMBMOthers.cboLGNCode = !LGNCode
            End If
            'If Len(!MBMSendingMode) Then
                frmNewMBMOthers.txtSendingMode = IIf(Len(!MBMSendingMode), !MBMSendingMode, "")
            'End If
            
            'If Len(!MBMCardDate) Then
                frmNewMBMOthers.txtCardSentDate = IIf(Len(!MBMCardDate), !MBMCardDate, "__/__/____")
            'End If
            txtMgmtNotes = IIf(Len(!MBMMmgtNotes), !MBMMmgtNotes, "")
        End With
End Function

Private Function GetSuppExpDate(DateOfBirth, tmpPlanCode, TmpEffDate, TmpExpDate As String)
Dim CheckYr As Long
Dim PLNRec As New ADODB.Recordset
Dim PLNSql As String
Dim SuppDobEffDate As String
Dim SuppDobExpDate As String

    PlnAgeLmt = 0
    OverAge = False
    SuppDobEffDate = Format(DateOfBirth, "dd/mm") & "/" & Year(TmpEffDate)
    SuppDobExpDate = Format(DateOfBirth, "dd/mm") & "/" & Year(TmpExpDate)
    CheckYr = Year(TmpExpDate) - Year(DateOfBirth)
    
    PLNSql = "select PlnAgeLimit from PLN where PLNCode = '" & Trim(tmpPlanCode) & "'"
    PLNRec.Open PLNSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If PLNRec.recordCount > 0 Then
        If Trim(PLNRec!PlnAgeLimit) <> "" Then
             PlnAgeLmt = IIf(IsNumeric(PLNRec!PlnAgeLimit), PLNRec!PlnAgeLimit, 0)
        End If
    End If
    PLNRec.Close
    Set PLNRec = Nothing
    
    GetSuppExpDate = TmpExpDate
    If Trim(PlnAgeLmt) <> 0 Then
        If CInt(CheckYr) = CInt(PlnAgeLmt) Then
            If CDate(SuppDobExpDate) <= CDate(TmpExpDate) Then
                GetSuppExpDate = CDate(SuppDobExpDate)
            End If
            If CDate(SuppDobEffDate) <= CDate(TmpEffDate) Then
                OverAge = True
            End If
        ElseIf CheckYr > PlnAgeLmt Then
            OverAge = True
        End If
    End If

End Function

Private Sub UpdateMCOPending()
Dim strsqlMBM As String
Dim MBM As New ADODB.Recordset

    strsqlMBM = " update MCOPending set " & _
          " MBMBordxDate  = '" & Format(txtMBMBordxDate, "YYYY/MM/DD") & "'" & _
          " Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    medixmasterconnMaint.Execute strsqlMBM
End Sub

Private Sub UpdateMBMRetire()
Dim strsqlMBM As String
Dim MBM As New ADODB.Recordset
If txtRtDate <> "__/__/____" Then
    strsqlMBM = " update MBMOthers   set " & _
          " MBMRetireDate  = '" & Format(txtRtDate, "YYYY/MM/DD") & "'" & _
          " Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    medixmasterconn.Execute strsqlMBM
End If
End Sub


Private Sub EnabledFrame(EStatus)
    Frame1.Enabled = EStatus
    CboGuaRenewal.Enabled = False
    Frame6.Enabled = EStatus
    Frame4.Enabled = EStatus
    Frame10.Enabled = EStatus
End Sub

Private Sub SaveCISCovPerson(CoveredIDCount As String, MemberDOB As String, SuppName As String)
Dim FindAge As String
Dim AgeValue As String
Dim CISSuppAnnLmt, CISSuppVisLmt As String
Dim CISSuppBPrem, CISSuppPrem As String
Dim CISSuppBMCO, CISSuppMCO As String
Dim CisDentalLmt As Double
Dim CisSpecLmt As Double
Dim CisMepLmt As Double
Dim CisDentalExtLmt As Double
Dim Cis1stSpecLmt As Double
Dim CisAnnLmtType As String
Dim CisSuppSql As String
Dim TmpCovID As String
Dim SysDate As String
Dim tmpCisExpDate As String
Dim tmpSuppExpDate As String
Dim CISDCPayor As String
Dim CISstrSptLmt As String
Dim CISSptLmtInd As String
Dim CISSptPlnLmt As New ADODB.Recordset
Dim CISstrSptPrem As String
Dim CISSptPremInd As String
Dim CISSptPlnPrem As New ADODB.Recordset
Dim CISstrSptMCOSql As String
Dim CISSptMCOInd As String
Dim CISSptPlnMCO As New ADODB.Recordset
Dim tmpCISPln As String
Dim tmpCisInsType As String
Dim tmpCisEffDate As String
Dim CLNPlnSql  As String
Dim CLNPln As New ADODB.Recordset
Dim ClnAnnLmtSql As String
Dim CLnAnnLmt As New ADODB.Recordset
Dim CisPlnCode, CisInsType As String
Dim CISCovSuppLMtStat As String
Dim CISDCStatus As String
Dim CISPDentStatus As String
Dim CISSDentStatus As String
Dim CISPSpecStatus As String
Dim CISSSpecStatus As String
Dim CISPMepStatus As String
Dim CISSMepStatus As String
Dim tmpCardType As String
Dim tmpWarr, CLNStatus, strCLN  As String
Dim CLNOther As New ADODB.Recordset
Dim tmpPlnCode As String
Dim tmpINSType As String
Dim CisRoundUp As String
Dim CisProRate As String
Dim tCMCOProcfeeStatus As String
Dim FFSInvNo As String
Dim FFSInvBy As String
Dim FFSInvDate As String
Dim FFSInved As String

    tmpPlnCode = Trim(cboPLNCode)
    tmpINSType = Trim(cboINSCode)
    If Trim(tmpPlnCode) = "" Then
        tmpPlnCode = Trim(txtClnPlanCode)
        tmpINSType = Trim(txtClnInsType)
    End If

    CisPlnCode = Trim(txtClnPlanCode)
    CisInsType = Trim(txtClnInsType)
    
    CLNPlnSql = "SELECT AnnLmtIND, VisLmtIND, DCStatus, PAYCode, PDentalStatus, SDentalStatus, " & _
                "PSpecStatus, SSpecStatus, PMEPStatus, SMEPStatus, PLNRoundUpStatus, PLNProRate, " & _
                "MCOProcfeeStatus from PLN where PLNCode='" & Trim(CisPlnCode) & _
                "' AND HLTCode='" & Trim(cboHLTCode) & "'"
    CLNPln.Open CLNPlnSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    CISCovSuppLMtStat = "A"
    CISDCStatus = ""
    CISPDentStatus = ""
    CISSDentStatus = ""
    CISPSpecStatus = ""
    CISSSpecStatus = ""
    CISPMepStatus = ""
    CISSMepStatus = ""
    CISDCPayor = Trim(cboPAYCode)
    CisRoundUp = "Y"
    CisProRate = "Y"
    
    If CLNPln.recordCount > 0 Then
        CISDCStatus = IIf(Len(Trim(CLNPln!DCStatus)), Trim(CLNPln!DCStatus), "")
        CISDCPayor = IIf(Len(Trim(CLNPln!PAYCode)), Trim(CLNPln!PAYCode), "")
        CISPDentStatus = IIf(Len(Trim(CLNPln!PDentalStatus)), Trim(CLNPln!PDentalStatus), "")
        CISSDentStatus = IIf(Len(Trim(CLNPln!SDentalStatus)), Trim(CLNPln!SDentalStatus), "")
        CISPSpecStatus = IIf(Len(Trim(CLNPln!PSpecStatus)), Trim(CLNPln!PSpecStatus), "")
        CISSSpecStatus = IIf(Len(Trim(CLNPln!SSpecStatus)), Trim(CLNPln!SSpecStatus), "")
        CISPMepStatus = IIf(Len(Trim(CLNPln!PMepStatus)), Trim(CLNPln!PMepStatus), "")
        CISSMepStatus = IIf(Len(Trim(CLNPln!SMepStatus)), Trim(CLNPln!SMepStatus), "")
        CisRoundUp = IIf(Len(Trim(CLNPln!PLNRoundUpStatus)), Trim(CLNPln!PLNRoundUpStatus), "Y")
        CisProRate = IIf(Len(Trim(CLNPln!PLNProRate)), Trim(CLNPln!PLNProRate), "Y")
        tCMCOProcfeeStatus = IIf(Len(Trim(CLNPln!MCOProcfeeStatus)), Trim(CLNPln!MCOProcfeeStatus), "N")
        If Trim(CLNPln!AnnLmtInd) = "Y" Then
            ClnAnnLmtSql = "Select SuppLimitStatus from AnnualLimit where PLNCode = '" & Trim(CisPlnCode) & _
                           "' AND HLTCode='" & Trim(cboHLTCode) & "' AND INSCode='" & Trim(CisInsType) & "'"
            CLnAnnLmt.Open ClnAnnLmtSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            If CLnAnnLmt.recordCount > 0 Then
                CISCovSuppLMtStat = CLnAnnLmt!SuppLimitStatus
            End If
            CLnAnnLmt.Close
            Set CLnAnnLmt = Nothing
        End If
    End If
    CLNPln.Close
    Set CLNPln = Nothing
    CoveredIDCount = Format(CoveredIDCount, "00")
    CISDCPayor = IIf(Trim(CISDCStatus) = "Y", CISDCPayor, cboPAYCode)
    tmpCisEffDate = CisEffDate
    tmpCisExpDate = CisExpDate
    tmpCISPln = CisPlnCode
    tmpCisInsType = CisInsType
    
    TmpCovID = Format(CoveredIDCount, "00")
    FindAge = GetAge(CDate(tmpCisEffDate), CDate(MemberDOB))
    AgeValue = Mid(GetAgeBand(cboHLTCode, FindAge, tmpPlnCode, tmpINSType), 1, 2)
    
    If Trim(tmpCISPln) <> "" Then
        tmpSuppExpDate = GetSuppExpDate(MemberDOB, tmpCISPln, tmpCisEffDate, tmpCisExpDate)
        tmpSuppExpDate = Format(tmpSuppExpDate, "yyyy/mm/dd")
    End If

' ******* Clinical Supp Limit
        CISstrSptLmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(tmpCISPln) & _
                       "' AND HLTCode='" & cboHLTCode & "' AND INSCode='" & Trim(tmpCisInsType) & _
                       "' AND AnnEffDate <= '" & Format(tmpCisEffDate, "mm/dd/yyyy") & "' order by AnnEffDate desc"
        CISSptPlnLmt.Open CISstrSptLmt, MediConnCISDB, adOpenKeyset, adLockOptimistic
        CISSptLmtInd = "A"
        If CISSptPlnLmt.recordCount > 0 Then
            If Trim(CISSptPlnLmt!SuppLimitStatus) <> "" Then
                CISSptLmtInd = Trim(CISSptPlnLmt!SuppLimitStatus)
            End If
        End If
        CISSptPlnLmt.Close
        Set CISSptPlnLmt = Nothing
        If Trim(cboHLTCode) = "S" Then
            CISstrSptPrem = "Select SuppPrmStatus, PLNCode from PRM where PLNCode = '" & Trim(tmpCISPln) & _
                             "' AND HLTCode='" & cboHLTCode & "' AND INSCode='" & Trim(tmpCisInsType) & _
                             "' AND PRMEffDate <= '" & Format(tmpCisEffDate, "mm/dd/yyyy") & "' order by PRMEffDate desc"
        Else
            CISstrSptPrem = "Select SuppPrmStatus, PLNCode from PRM where PLNCode = '" & Trim(tmpCISPln) & _
                             "' AND HLTCode='" & cboHLTCode & "' AND INSCode='" & Trim(tmpCisInsType) & _
                             "' AND PAYCode='" & Trim(CISDCPayor) & _
                             "' AND PRMEffDate <= '" & Format(tmpCisEffDate, "mm/dd/yyyy") & "' order by PRMEffDate desc"
        End If
        CISSptPlnPrem.Open CISstrSptPrem, MediConnCISDB, adOpenKeyset, adLockOptimistic
        CISSptPremInd = "A"
        If CISSptPlnPrem.recordCount > 0 Then
            If Trim(CISSptPlnPrem!SuppPrmStatus) <> "" Then
                CISSptPremInd = Trim(CISSptPlnPrem!SuppPrmStatus)
            End If
        End If
        CISSptPlnPrem.Close
        Set CISSptPlnPrem = Nothing
        
        CISstrSptMCOSql = "Select SuppMCOStatus from MCO where PLNCode = '" & Trim(tmpCISPln) & _
                         "' AND HLTCode='" & cboHLTCode & "' AND INSCode='" & Trim(tmpCisInsType) & _
                         "' AND MCOEffdate <= '" & Format(tmpCisEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"

        CISSptPlnMCO.Open CISstrSptMCOSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        CISSptMCOInd = "A"
        If CISSptPlnMCO.recordCount > 0 Then
            If Trim(CISSptPlnMCO!SuppMCOStatus) <> "" Then
                CISSptMCOInd = Trim(CISSptPlnMCO!SuppMCOStatus)
            End If
        Else
            CISSptPlnMCO.Close
            CISstrSptMCOSql = "Select SuppMCOStatus from MCO where PLNCode is null " & _
                              " AND HLTCode='" & cboHLTCode & "' AND INSCode='" & Trim(tmpCisInsType) & _
                              "' AND MCOEffdate <= '" & Format(tmpCisEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"

            CISSptPlnMCO.Open CISstrSptMCOSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
            If CISSptPlnMCO.recordCount > 0 Then
                If Trim(CISSptPlnMCO!SuppMCOStatus) <> "" Then
                    CISSptMCOInd = Trim(CISSptPlnMCO!SuppMCOStatus)
                End If
            End If
        End If
        CISSptPlnMCO.Close
        Set CISSptPlnMCO = Nothing
        
'----
    If Trim(CISSptPremInd) = "B" And Format(CoveredIDCount, "00") = "01" Or Trim(CISSptPremInd) = "C" Then
        CISSuppPrem = Format(GetCLNSPremium(tmpCisInsType, CISDCPayor, AgeValue, tmpCISPln, Trim(cboHLTCode), CDate(tmpCisEffDate)), "#,##0.00")
        CISSuppBPrem = CISSuppPrem
        CISSuppPrem = InsertPremium(CDate(tmpSuppExpDate), CDate(tmpCisEffDate), Format(CISSuppPrem, "#,##0.00"))
        CISSuppPrem = Format(Round(CISSuppPrem, 0), "#,##0.00")
    End If
    If Trim(CISSptMCOInd) = "B" And Format(CoveredIDCount, "00") = "01" Or Trim(CISSptMCOInd) = "C" Then
        If AgeValue = "01" Then
            CISSuppMCO = Format(GetCLNSMCO(tmpCisInsType, cboHLTCode, "C", CDate(tmpCisEffDate), tmpCISPln), "#,##0.00")
        Else
            CISSuppMCO = Format(GetCLNSMCO(tmpCisInsType, cboHLTCode, "A", CDate(tmpCisEffDate), tmpCISPln), "#,##0.00")
        End If
        CISSuppBMCO = CISSuppMCO
        If CisProRate <> "N" Then
            CISSuppMCO = InsertMCOFees(CDate(tmpSuppExpDate), CDate(tmpCisEffDate), Format(CISSuppMCO, "#,##0.00"))
        End If
        CISSuppMCO = Format(CISSuppMCO, "#,##0.00")
        If CisRoundUp <> "N" Then
            CISSuppMCO = Round(CISSuppMCO)
        End If
    End If
    If Trim(CISSptLmtInd) = "B" And Format(CoveredIDCount, "00") = "01" Or Trim(CISSptLmtInd) = "C" Then
        CISSuppAnnLmt = Format(GetCLNSuppLimit(tmpCisInsType, tmpCISPln, cboHLTCode, CDate(tmpCisEffDate)), "#,##0.00")
        CISSuppVisLmt = GetCLNSVisitLimit(tmpCisInsType, tmpCISPln, cboHLTCode, CDate(tmpCisEffDate))
        If Trim(CISSDentStatus) = "S" Then 'stand alone dental limit
            CisAnnLmtType = "SDentalAnnLimit"
            CisDentalLmt = GetCISLimit(tmpCisInsType, tmpCISPln, cboHLTCode, CDate(tmpCisEffDate), CisAnnLmtType)
        End If
        If Trim(CISSSpecStatus) = "S" Then 'stand alone Spec limit
            CisAnnLmtType = "SSpecAnnLimit"
            CisSpecLmt = GetCISLimit(tmpCisInsType, tmpCISPln, cboHLTCode, CDate(tmpCisEffDate), CisAnnLmtType)
        End If
        If Trim(CISSMepStatus) = "S" Then 'stand alone MEP limit
            CisAnnLmtType = "SMepAnnLimit"
            CisMepLmt = GetCISLimit(tmpCisInsType, tmpCISPln, cboHLTCode, CDate(tmpCisEffDate), CisAnnLmtType)
        End If
    End If
    
' ****** End of Clinical Supp Limit
    CISSuppBPrem = CDbl(IIf(IsNumeric(CISSuppBPrem) = True, CISSuppBPrem, 0))
    CISSuppPrem = CDbl(IIf(IsNumeric(CISSuppPrem) = True, CISSuppPrem, 0))
    CISSuppBMCO = CDbl(IIf(IsNumeric(CISSuppBMCO) = True, CISSuppBMCO, 0))
    CISSuppMCO = CDbl(IIf(IsNumeric(CISSuppMCO) = True, CISSuppMCO, 0))
    CISSuppAnnLmt = CDbl(IIf(IsNumeric(CISSuppAnnLmt) = True, CISSuppAnnLmt, 0))
    CISSuppVisLmt = CDbl(IIf(IsNumeric(CISSuppVisLmt) = True, CISSuppVisLmt, 0))
    SysDate = Format(tmpServerDate, "yyyy/mm/dd")
    tmpCardType = ""
    CLNStatus = "A"
    FFSInved = ""
    FFSInvNo = ""
    FFSInvBy = ""
    FFSInvDate = ""
    If Trim(tCMCOProcfeeStatus) = "Y" Then
        FFSInved = "Y"
        FFSInvNo = "FFS"
        FFSInvBy = Logon
        FFSInvDate = Format(tmpServerDate, "YYYY/MM/DD")
    End If

    CisSuppSql = "EXEC Upload_MBMCovPersonsCis '" & Trim(txtMBMNumber) & "','" & TmpCovID & "', " & _
                "" & CISSuppMCO & "," & CISSuppBMCO & "," & CISSuppPrem & "," & CISSuppBPrem & ", " & _
                "" & CISSuppAnnLmt & "," & CISSuppVisLmt & "," & CisDentalLmt & "," & CisSpecLmt & ", " & _
                "" & CisMepLmt & ",'" & SysDate & "','" & Logon & "'," & _
                "'" & SysDate & "','" & Logon & "','" & tmpSuppExpDate & "','" & tmpCardType & "'," & _
                "'" & tmpWarr & "','" & CLNStatus & "','" & Trim(CisPlnCode) & "'," & _
                    "'" & FFSInvNo & "','" & FFSInvDate & "','" & FFSInvBy & "'"
    medixmasterconn.Execute CisSuppSql
End Sub

Private Sub Write2MBMnHistory(WriteStatus As String, cboIndex As Integer, HtmpMBMNo As String, HtmpPolicy As String, _
HtmpIcNo As String, HtmpName As String, HtmpExpDate As String, HtmpPlnCode As String)

Dim MBM As New ADODB.Recordset
Dim MBMCov As New ADODB.Recordset
Dim MBMTwo As New ADODB.Recordset
Dim StrSqlCRef, StrMBMTwo As String
Dim strsql, StrCovSql, StrMBMSql As String
Dim StrAmendSql As String
Dim ListCount As Integer
Dim listloop As Integer
Dim tmpName As String
    If WriteStatus = "C" Then
        StrSqlCRef = "Update MBMCrossReference set MBMStatus='" & WriteStatus & "', MBMFinalExpiryDate = '" & Format(txtDate, "yyyy/mm/dd") & "' Where MBMNumber = '" & txtMBMNumber & "'"
        medixmasterconnMaint.Execute StrSqlCRef
        StrMBMSql = "Update MBM set MBMStatus ='" & WriteStatus & "' , MBMFinalExpiryDate = '" & Format(txtDate, "yyyy/mm/dd") & "' Where MBMNumber = '" & txtMBMNumber & "'"
        medixmasterconn.Execute StrMBMSql
    ElseIf WriteStatus <> "K" Then
        StrSqlCRef = "Update MBMCrossReference set MBMStatus='" & WriteStatus & "' Where MBMNumber = '" & txtMBMNumber & "'"
        medixmasterconnMaint.Execute StrSqlCRef
        StrMBMSql = "Update MBM set MBMStatus ='" & WriteStatus & "'  Where MBMNumber = '" & txtMBMNumber & "'"
        medixmasterconn.Execute StrMBMSql
    
    End If
    If WriteStatus = "C" Then ' New field in MBMTwo - MBMCancelDate
        StrMBMTwo = "Update MBMTwo Set MBMLastEndorseStatus ='" & Mid(cboEndorseStatus, 1, 1) & "'," & _
                    "MBMENDtEffDate='" & Format(txtDate, "yyyy/mm/dd") & "'," & _
                    "MBMENDtBordxDate ='" & Format(txtEndtBordxDate, "yyyy/mm/dd") & "'," & _
                    "MBMCancelDate ='" & Format(txtDate, "yyyy/mm/dd") & _
                    "' Where MBMNumber = '" & txtMBMNumber & "'"
    Else
        StrMBMTwo = "Update MBMTwo Set MBMLastEndorseStatus ='" & Mid(cboEndorseStatus, 1, 1) & "'," & _
                    "MBMENDtEffDate='" & Format(txtDate, "yyyy/mm/dd") & "'," & _
                    "MBMENDtBordxDate ='" & Format(txtEndtBordxDate, "yyyy/mm/dd") & _
                    "' Where MBMNumber = '" & txtMBMNumber & "'"
    End If
    medixmasterconn.Execute StrMBMTwo
    tmpName = Replace(HtmpName, "'", "''")
    If IsDate(HtmpExpDate) Then
        StrAmendSql = "Insert into MBMAmendHistory(MBMNumber, MBMPolicyNo, MBMICBcPp, MBMName, " & _
                  "MBMExpiryDate, MBMPlanCode, MBMAmendEdtType, " & _
                  "MBMAEndtBordxDate, MBMAEndtEffDate, " & _
                  "MBMAmendUser, MBMAmendDate, MBMAEndtBatchNo) " & _
                  "Values ('" & HtmpMBMNo & "','" & HtmpPolicy & "','" & HtmpIcNo & "','" & tmpName & "'," & _
                  "'" & Format(HtmpExpDate, "yyyy/mm/dd") & "','" & HtmpPlnCode & "','" & Mid(cboEndorseStatus, 1, 1) & "'," & _
                  "'" & Format(txtEndtBordxDate, "YYYY/MM/DD") & "','" & Format(txtDate, "YYYY/MM/DD") & "'," & _
                  "'" & Logon & "','" & Format(tmpServerDate, "YYYY/MM/DD") & "','" & txtBatchNo & "')"
    Else
        StrAmendSql = "Insert into MBMAmendHistory(MBMNumber, MBMPolicyNo, MBMICBcPp, MBMName, " & _
                  "MBMAmendEdtType, " & _
                  "MBMAEndtBordxDate, MBMAEndtEffDate, " & _
                  "MBMAmendUser, MBMAmendDate, MBMAEndtBatchNo) " & _
                  "Values ('" & HtmpMBMNo & "','" & HtmpPolicy & "','" & HtmpIcNo & "','" & tmpName & "'," & _
                  "'" & Mid(cboEndorseStatus, 1, 1) & "'," & _
                  "'" & Format(txtEndtBordxDate, "YYYY/MM/DD") & "','" & Format(txtDate, "YYYY/MM/DD") & "'," & _
                  "'" & Logon & "','" & Format(tmpServerDate, "YYYY/MM/DD") & "','" & txtBatchNo & "')"

    End If
    medixmasterconn.Execute StrAmendSql
    
    cboStatus.ListIndex = cboIndex
    If Mid(cboEndorseStatus, 1, 1) = "C" Or Mid(cboEndorseStatus, 1, 1) = "X" _
        Or Mid(cboEndorseStatus, 1, 1) = "D" Then
        strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & txtMBMNumber & "'"
        MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        ListCount = lstCovAdd.ListItems.Count
        For listloop = 1 To ListCount
    
            With MBMCov
                MBMCov.Filter = ("MBMCCoverID = '" & lstCovAdd.ListItems(listloop).SubItems(10) & "'")
                !MBMCStatus = WriteStatus
                MBMCov.Update
            End With
                
        Next listloop
        MBMCov.Close
        Set MBMCov = Nothing
    End If
End Sub

Private Sub WriteAdjStatus2Cov(AdjStatus As String, AmendStatus As String, TmpCovID As String)
Dim strsqlMBMCov As String
Dim MBMCov As New ADODB.Recordset

    strsqlMBMCov = "SELECT MBMCNumber, MBMCCancelDate, MBMCStatus From MBMCoveredPersons  Where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
    If Trim(AmendStatus) = "K" Then
        strsqlMBMCov = strsqlMBMCov & " AND MBMCCoverID='" & Trim(TmpCovID) & "'"
    End If
    MBMCov.Open strsqlMBMCov, medixmasterconn, adOpenKeyset, adLockOptimistic
    If MBMCov.recordCount Then
        Do Until MBMCov.EOF
            With MBMCov
                If Trim(!MBMCStatus) <> "K" And Trim(!MBMCStatus) <> "C" Then
                    !MBMCStatus = AdjStatus
                End If
                If Trim(AmendStatus) = "K" Or Trim(AmendStatus) = "C" Then
                    !MBMCCancelDate = Format(txtDate, "yyyy/mm/dd")
                End If
                                
            End With
            MBMCov.Update
            MBMCov.MoveNext
        Loop
    End If
    MBMCov.Close
    Set MBMCov = Nothing

End Sub

Private Sub SaveMBMLifeTimeLmt(LMBMNumber, LMBMCovID, LmtType As String, ActualLmt As Double, tmpCNextLmt As Double)
Dim tmpSql As String
Dim TmpLifeSql As String
Dim TmpLimit As Double
Dim TmpLmtVar As String
Dim TmpSql2 As String
Dim TmpLMBM As New ADODB.Recordset

    TmpSql2 = "Select MBMNumber, MBMLifeTimeLmt From MBMLifeTime where MBMNumber='" & Trim(LMBMNumber) & _
              "' AND MBMCovID='" & LMBMCovID & "'"
    TmpLMBM.Open TmpSql2, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    If TmpLMBM.EOF Then
        If Mid(CboGuaRenewal, 1, 1) = "3" Then
            WLifeStatus = True
        End If
    Else
        tmpCNextLmt = ActualLmt
        If CDbl(ActualLmt) > CDbl(TmpLMBM!MBMLifeTimeLmt) Then
            ActualLmt = CDbl(TmpLMBM!MBMLifeTimeLmt)
        End If
    End If
    TmpLMBM.Close
    Set TmpLMBM = Nothing
    
    If WLifeStatus = True Then
        TmpLimit = 0
        TmpLmtVar = "SuppLifeTimeLimit"
        If LmtType = "P" Then
            TmpLmtVar = "LifeTimeLimit"
        End If
        TmpLimit = GetLifeTimeLimit(TmpLmtVar, Trim(cboINSCode), Trim(cboPLNCode), Trim(cboHLTCode), CDate(txtMBMPayorEffDate))
        TmpLimit = CDbl(TmpLimit) - (ActualLmt - tmpCNextLmt)
        TmpLifeSql = "Insert Into MBMLifeTime(MBMNumber, MBMCOVID, MBMLifeTimeLmt) " & _
              "Values('" & LMBMNumber & "','" & LMBMCovID & "'," & TmpLimit & ")"
        medixmasterconn.Execute TmpLifeSql
    End If
End Sub

Private Sub PrincipalLifeTimeChecking()
Dim tmpSuppBNFStatus  As String
Dim TmpSql2 As String
Dim TmpLMBM As New ADODB.Recordset
    WLifeStatus = False
    If Trim(TempLifeTimeStatus) = "Y" Then
        TmpSql2 = "Select top 1 MBMNumber, MBMTopupNumber, MBMAvailableLimit From MBM where MBMPolicyNo='" & Trim(txtMBMPolNo) & _
                  "' And MBMStatus = 'A' AND PayCode ='" & Trim(cboPAYCode) & "' AND INSCode='" & Trim(cboINSCode) & _
                  "' ORDER BY MBMPayorEffDate DESC "
        TmpLMBM.Open TmpSql2, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        PrevLifeMBMNo = Trim(txtMBMNumber)

        If TmpLMBM.EOF = False Then
            PrevLifeMBMNo = TmpLMBM!MBMTopupNumber
        End If
        TmpLMBM.Close
        Set TmpLMBM = Nothing
        
        TmpSql2 = "Select MBMNumber, MBMLifeTimeLmt From MBMLifeTime where MBMNumber='" & Trim(PrevLifeMBMNo) & _
                  "' AND MBMCovID='00'"
        TmpLMBM.Open TmpSql2, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        If TmpLMBM.EOF = False Then
            If Mid(CboGuaRenewal, 1, 1) = "3" Then
                WLifeStatus = True
            End If
        End If
        TmpLMBM.Close
        Set TmpLMBM = Nothing
    End If
End Sub

Private Sub SaveMBMBnfAnnLmt(LMBMNumber, LMBMCovID, LmtType, tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt)
Dim TmpBnfLmtSql As String
    LMBMCovID = Format(LMBMCovID, "00")
    TmpBnfLmtSql = "Insert Into MBMBnfLmt(MBMNumber, MBMCoveredID, MBMKidneyAnnLmt, MBMKidneyBalLmt, " & _
                 "MBMCancerAnnLmt, MBMCancerBalLmt, MBMHomeNCAnnLmt, MBMHomeNCBalLmt) " & _
                 "Values('" & LMBMNumber & "','" & LMBMCovID & "'," & tmpKidneyAnnLmt & "," & tmpKidneyAnnLmt & _
                 "," & tmpCancerAnnLmt & "," & tmpCancerAnnLmt & "," & tmpHomeNursAnnLmt & "," & tmpHomeNursAnnLmt & ")"
    medixmasterconn.Execute TmpBnfLmtSql

End Sub

Private Sub ComboKeyPress(tmpComboBox As ComboBox, KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(tmpComboBox.Text, tmpComboBox.SelStart) + Chr(KeyAscii) + Mid(tmpComboBox.Text, tmpComboBox.SelStart + tmpComboBox.SelLength + 1))
    ValidCount = 0
    ValidValue = ""
    For i = 0 To tmpComboBox.ListCount - 1
        If NewText = LCase(Left(tmpComboBox.List(i), Len(NewText))) Then
            ValidCount = ValidCount + 1
            ValidValue = tmpComboBox.List(i)
        End If
    Next
    If ValidCount <= 1 Then KeyAscii = 0
         If ValidCount = 1 Then
           tmpComboBox.Text = ValidValue
           tmpComboBox.SelStart = 0
           tmpComboBox.SelLength = Len(ValidValue)
         End If
    End If
End Sub

Private Sub GetValidDate(EffStart As String, EffEnd As String, DateSql As String)

Dim strsql As String
Dim MBMValidDate As New ADODB.Recordset

   ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    strsql = "Select " & DateSql & " From MBMValidDateRange "
    MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If MBMValidDate.recordCount Then
        EffStart = MBMValidDate!DateStart
        EffEnd = MBMValidDate!DateEnd
    End If
    MBMValidDate.Close
    Set MBMValidDate = Nothing
  '  medixmasterconn.Close
 '   Set medixmasterconn = Nothing
End Sub

Private Sub SaveMBMRefund(PayEffDate As Date, PAYExpDate As Date, TmpCovID As String, tmpPremium As Double, tmpMCO As Double, MBMCancelStatus As String)
Dim MBMMCOFees As Double
Dim HLTCode As String
Dim RefundSql As String
Dim MCORefundValue As Double
Dim PremiumRefundValue As Double
Dim PremiumEffective As Double
Dim MCOEffective As Double
    
    HLTCode = Mid(cboHLTCode, 1, 1)
    MCORefundValue = Format(MCORefund(CDate(PAYExpDate), CDate(txtDate), CDate(PayEffDate), tmpMCO, MBMCancelStatus), "#0.00")
    If Trim(RoundUpStatus) = "Y" Then
        MCORefundValue = Round(MCORefundValue)
    End If
    
    PremiumRefundValue = Format(MCOPremiumReduction(CDate(PAYExpDate), CDate(txtDate), CDate(PayEffDate), tmpPremium), "#0.00")
    If RoundUpStatus = "Y" Then
        PremiumRefundValue = Round(PremiumRefundValue)
    End If
    
    PremiumEffective = Format(CDbl(tmpPremium) - CDbl(PremiumRefundValue), "#0.00")
    
    If RoundUpStatus = "Y" Then
        PremiumEffective = Round(PremiumEffective)
    End If
    
    MCOEffective = Format(CDbl(tmpMCO) - CDbl(MCORefundValue), "#0.00")
    If RoundUpStatus = "Y" Then
        MCOEffective = Round(MCOEffective)
    End If
    
    RefundSql = "EXEC Upload_MBMMcoRefund '" & Trim(txtMBMNumber) & "','" & Trim(TmpCovID) & _
              "'," & PremiumEffective & "," & PremiumRefundValue & "," & MCOEffective & _
              "," & MCORefundValue & ",'" & Trim(cboPAYCode) & "'," & _
              "'" & Format(txtMBMBordxDate, "yyyy/mm/dd") & "','" & txtBatchNo & "','" & HLTCode & "'"
    medixmasterconn.Execute RefundSql
End Sub

Private Sub CmdCancelSupp_Click()

Dim ForLoop As Integer
Dim ListCount As Integer
Dim RecordSaved

    ListCount = 0

    If txtCovName = "" Then
        MsgBox "Please Choose a Record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            ListCount = lstCovAdd.ListItems.Count
            For ForLoop = 1 To ListCount
                If lstCovAdd.ListItems(ForLoop).Selected = True Then
                    lstCovAdd.SelectedItem.SubItems(11) = "C"
                End If
            Next ForLoop
        End If
    End If
End Sub

Private Sub MBMSuppMCORefundCal(TmpCovID As String, TmpEffDate As String, TmpExpDate As String, tmpDOb As String, tmpPln As String, EndtStatus As String, tmpPrem As Double, tmpMCO As Double)
Dim strsql As String
Dim MCORefundValue As String
Dim PayEffDate As String
Dim PAYExpDate As String
Dim MBMDOB As String
Dim MBMCancelStatus As String
Dim PayRec As New ADODB.Recordset
Dim tmpPlnCode As String
Dim tmpINSType As String
Dim MCORec As New ADODB.Recordset
Dim sqlMCO, SuppMCOStatus, tmpCovIDChrg As String
Dim AA As Integer
Dim tmpCovMcoFee As Double
Dim tEffDate As String
Dim tExpDate As String
Dim tDOB As String
Dim tPrem As Double
Dim tMCO As Double

    AA = InStr(cboPLNCode, "-")
    MBMCancelStatus = "N"
    strsql = "Select CLNStatus, MBMCancelStatus From PAY where PAYCode ='" & Mid(cboPAYCode, 1, 2) & "'"
    PayRec.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If PayRec.recordCount Then
       MBMCancelStatus = IIf(Len(Trim(PayRec!MBMCancelStatus)), Trim(PayRec!MBMCancelStatus), "N")
    End If
    PayRec.Close
    Set PayRec = Nothing

    SuppMCOStatus = "A"
    tmpCovIDChrg = ""
    sqlMCO = "Select SuppMCOStatus,CovIDChrg from MCO where HLTCode = '" & Mid(cboHLTCode, 1, 1) & _
            "' and PlnCode='" & Trim(tmpPln) & "' And INSCode='" & Trim(cboINSCode) & _
            "' AND MCOEffdate <= '" & Format(TmpEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
    MCORec.Open sqlMCO, medixmasterconn, adOpenKeyset, adLockOptimistic
    If MCORec.recordCount Then
        SuppMCOStatus = IIf(Len(MCORec!SuppMCOStatus), Trim(MCORec!SuppMCOStatus), "A")
        tmpCovIDChrg = IIf(Len(MCORec!CovIDChrg), Trim(MCORec!CovIDChrg), "")
    Else
        MCORec.Close
        Set MCORec = Nothing
        sqlMCO = "Select SuppMCOStatus, CovIDChrg from MCO where HLTCode = '" & Mid(cboHLTCode, 1, 1) & _
                "' and PLNCode is Null And INSCode='" & Trim(cboINSCode) & _
                "' AND MCOEffdate <= '" & Format(TmpEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
        MCORec.Open sqlMCO, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MCORec.recordCount Then
            SuppMCOStatus = IIf(Len(MCORec!SuppMCOStatus), Trim(MCORec!SuppMCOStatus), "A")
            tmpCovIDChrg = IIf(Len(MCORec!CovIDChrg), Trim(MCORec!CovIDChrg), "")
        End If
    End If
    MCORec.Close
    Set MCORec = Nothing
    If EndtStatus = "C" Then
        For AA = 1 To lstCovAdd.ListItems.Count
            tmpCovMcoFee = lstCovAdd.ListItems(AA).SubItems(16)
            tEffDate = IIf(Len(lstCovAdd.ListItems(AA).SubItems(13)), lstCovAdd.ListItems(AA).SubItems(13), txtMBMPayorEffDate)
            tExpDate = IIf(Len(lstCovAdd.ListItems(AA).SubItems(14)), lstCovAdd.ListItems(AA).SubItems(14), txtMBMPayorExpDate)
            TmpCovID = Format(lstCovAdd.ListItems(AA).SubItems(10), "00")
            tPrem = IIf(Len(lstCovAdd.ListItems(AA).SubItems(18)), lstCovAdd.ListItems(AA).SubItems(18), 0)
            If SuppMCOStatus = "B" Then
                If Format(lstCovAdd.ListItems(AA).SubItems(10), "00") = "01" Then
                    Call SaveMBMRefund(CDate(tEffDate), CDate(tExpDate), TmpCovID, tPrem, CInt(tmpCovMcoFee), MBMCancelStatus)
                Else
                    Exit For
                End If
            End If
            If SuppMCOStatus = "C" Or _
                SuppMCOStatus = "D" And TmpCovID >= Format(tmpCovIDChrg, "00") Then
                Call SaveMBMRefund(CDate(tEffDate), CDate(tExpDate), TmpCovID, tPrem, CInt(tmpCovMcoFee), MBMCancelStatus)
            End If
        Next AA
    ElseIf EndtStatus = "K" Then
        If SuppMCOStatus = "B" And TmpCovID = "01" Then
            Call SaveMBMRefund(CDate(TmpEffDate), CDate(TmpExpDate), TmpCovID, tmpPrem, tmpMCO, MBMCancelStatus)
        End If
        If SuppMCOStatus = "C" Or SuppMCOStatus = "D" Then
            Call SaveMBMRefund(CDate(TmpEffDate), CDate(TmpExpDate), TmpCovID, tmpPrem, tmpMCO, MBMCancelStatus)
        End If
    End If
End Sub

Private Sub FeesListing(tmpMBMNo As String, tmpName As String, tmpRELCode As String, TmpCovID As String, _
tmpStatus As String, tmpPlnCode As String, TmpEffDate As String, TmpExpDate As String, tmpCancelDate As String, _
tmpAnnLmt As String, tmpAvLmt As String, tmpBPrem As String, tmpPrem As String, tmpBMco As String, tmpMCO As String, _
tmpInvNo As String, TmpInvDate As String, tmpCRNote As String, tmpCRDate As String)

Dim FeesList As ListItem
Dim strRedSql As String
Dim RefundRec As New ADODB.Recordset
Dim strsqlRein As String
Dim rein As New ADODB.Recordset

    Set FeesList = lstFeesView.ListItems.Add(, , "")
    FeesList.SubItems(1) = tmpName
    FeesList.SubItems(2) = tmpRELCode
    FeesList.SubItems(3) = TmpCovID
    FeesList.SubItems(4) = tmpStatus
    FeesList.SubItems(5) = IIf(Len(tmpPlnCode), tmpPlnCode, "")
    FeesList.SubItems(7) = IIf(Len(TmpEffDate), Format(TmpEffDate, "dd/mm/yyyy"), "")
    FeesList.SubItems(8) = IIf(Len(TmpExpDate), Format(TmpExpDate, "dd/mm/yyyy"), "")
    FeesList.SubItems(9) = IIf(Len(tmpCancelDate), Format(tmpCancelDate, "dd/mm/yyyy"), "")
    FeesList.SubItems(10) = tmpAnnLmt
    FeesList.SubItems(11) = IIf(IsNumeric(tmpAvLmt), tmpAvLmt, 0)
    If SptLmtInd = "A" And TmpCovID = "00" Or SptLmtInd = "B" And TmpCovID = "01" Or SptLmtInd = "C" Then
        FeesList.SubItems(12) = IIf(IsNumeric(tmpPerDisabilitAmt), tmpPerDisabilitAmt, 0)
    End If
    FeesList.SubItems(13) = IIf(Len(tmpBPrem), tmpBPrem, "")
    FeesList.SubItems(14) = IIf(IsNumeric(tmpPrem), tmpPrem, 0)
    FeesList.SubItems(15) = IIf(IsNumeric(tmpBMco), tmpBMco, 0)
    FeesList.SubItems(16) = IIf(IsNumeric(tmpMCO), tmpMCO, 0)
    FeesList.SubItems(17) = IIf(Len(tmpInvNo), tmpInvNo, "")
    FeesList.SubItems(18) = IIf(Len(TmpInvDate), TmpInvDate, "")
    FeesList.SubItems(19) = IIf(Len(tmpCRNote), tmpCRNote, "")
    FeesList.SubItems(20) = IIf(Len(tmpCRDate), tmpCRDate, "")
    '21,22 reserve
    strRedSql = "Select MBMRefInvoiceNo, MBMRefInvoiceDate, MBMPremReduction, MBMMCORefund, " & _
                "MBMMCOEffective, MBMPREMEffective from MBMRefund where " & _
                " MBMNumber = '" & tmpMBMNo & "' AND MBMCovID='" & TmpCovID & "'"
    RefundRec.Open strRedSql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    Do While Not RefundRec.EOF
        With RefundRec
            'MBMRefInvoiceNo & MBMRefInvoiceDate this 2 fields not using at all
'            FeesList.SubItems(23) = IIf(Len(!MBMRefInvoiceNo), !MBMRefInvoiceNo, "")
'            FeesList.SubItems(24) = IIf(Len(!MBMRefInvoiceDate), !MBMRefInvoiceDate, "")
            FeesList.SubItems(25) = IIf(IsNumeric(!MBMMCOEffective), Format(Round(!MBMMCOEffective, 0)), "#,##0.00")
            FeesList.SubItems(26) = IIf(IsNumeric(!MBMMCORefund), Format(Round(!MBMMCORefund, 0)), "#,##0.00")
            FeesList.SubItems(27) = IIf(IsNumeric(!MBMPREMEffective), Format(Round(!MBMPREMEffective, 0)), "#,##0.00")
            FeesList.SubItems(28) = IIf(IsNumeric(!MBMPremReduction), Format(Round(!MBMPremReduction, 0)), "#,##0.00")
        End With
        RefundRec.MoveNext
    Loop
    RefundRec.Close
    Set RefundRec = Nothing
    
    'lstFeesView.SortOrder = lvwAscending
    
    'strsqlRein = ""
    'strsqlRein = "Select MBMMCOFee,MBMPremium,MBMInvoiceNo,MBMInvoiceDate from MBMReinstatement where " & _
    '            " MBMNumber = '" & tmpMBMNo & "' AND MBMCoverID='" & TmpCovID & "'"
    'Rein.Open strsqlRein, medixmasterconn, adOpenForwardOnly, adLockReadOnly
    '    Do While Not Rein.EOF
    '        Set FeesList = lstFeesView.listitems.Add(, , "")
    '        FeesList.SubItems(1) = tmpName
    '        FeesList.SubItems(2) = tmpRELCode
    '        FeesList.SubItems(3) = TmpCovID
    '        FeesList.SubItems(4) = tmpStatus
    '        FeesList.SubItems(5) = IIf(Len(tmpPlnCode), tmpPlnCode, "")
    '        FeesList.SubItems(7) = IIf(Len(TmpEffDate), Format(TmpEffDate, "dd/mm/yyyy"), "")
    '        FeesList.SubItems(8) = IIf(Len(TmpExpDate), Format(TmpExpDate, "dd/mm/yyyy"), "")
    '        FeesList.SubItems(9) = IIf(Len(tmpCancelDate), Format(tmpCancelDate, "dd/mm/yyyy"), "")
    '        FeesList.SubItems(10) = tmpAnnLmt
    '        FeesList.SubItems(11) = IIf(IsNumeric(tmpAvLmt), tmpAvLmt, 0)
    '        If SptLmtInd = "A" And TmpCovID = "00" Or SptLmtInd = "B" And TmpCovID = "01" Or SptLmtInd = "C" Then
    '            FeesList.SubItems(12) = IIf(IsNumeric(tmpPerDisabilitAmt), tmpPerDisabilitAmt, 0)
    '        End If
    '        FeesList.SubItems(13) = 0
    '        FeesList.SubItems(14) = 0
    '        FeesList.SubItems(15) = 0
    '        FeesList.SubItems(16) = 0
    '        FeesList.SubItems(17) = IIf(Len(Rein!MBMInvoiceNo), Rein!MBMInvoiceNo, "")
    '        FeesList.SubItems(18) = IIf(Len(Rein!MBMInvoiceDate), Rein!MBMInvoiceDate, "")
    '        FeesList.SubItems(19) = ""
    '        FeesList.SubItems(20) = ""
    '        '21,22 reserve
    '            'MBMRefInvoiceNo & MBMRefInvoiceDate this 2 fields not using at all
    ''            FeesList.SubItems(23) = IIf(Len(!MBMRefInvoiceNo), !MBMRefInvoiceNo, "")
    ''            FeesList.SubItems(24) = IIf(Len(!MBMRefInvoiceDate), !MBMRefInvoiceDate, "")
    '        FeesList.SubItems(25) = IIf(IsNumeric(Rein!MBMMCOFee), Format(Round(Rein!MBMMCOFee, 0)), "#,##0.00")
    '        FeesList.SubItems(26) = Format(0, "#,##0.00")
    '        FeesList.SubItems(27) = IIf(IsNumeric(Rein!MBMPremium), Format(Round(Rein!MBMPremium, 0)), "#,##0.00")
    '        FeesList.SubItems(28) = Format(0, "#,##0.00")
    '    Rein.MoveNext
    '    Loop
    'Rein.Close
    'Set Rein = Nothing
        
End Sub

Private Sub MBMHistory(TmpMBMType, SearchByICNo As Boolean)
Dim TmpREc1 As New ADODB.Recordset
Dim HISList As ListItem
Dim Cmd As New ADODB.Command
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim tmpName As String
Dim strAppend As String
Dim tmpDOb As String
Dim Cnt As Integer
Dim tmpPolHistory1 As String
    tmpPolHistory = ""
    'tmpPolHistory1 = ""
    tmpName = Trim(Replace(txtMBMName, "'", "''"))
    'tmpDOb = Format(txtMBMBirthDate, "yyyy/mm/dd")
    Cnt = 0
    'If TmpMBMType = "Princ" Then
        If SearchByICNo = True Then
            strAppend = " And HISMaintenance.dbo.MBMCrossReference.MBMIcBcPp = '" & Trim(txtMBMIcBcPP) & "' And HISMaintenance.dbo.MBMCrossReference.MBMStatus <> 'D'"
        Else
            strAppend = " And HISMaintenance.dbo.MBMCrossReference.MBMName = '" & tmpName & "' And HISMaintenance.dbo.MBMCrossReference.MBMStatus <> 'D'"
            'strAppend = strAppend & "' And HISMaintenance.dbo.MBMCrossReference.MBMDOB='" & tmpDOb & "'"
        End If
    'Else
    '    If SearchByICNo = True Then
    '        strAppend = " And dbo.MBMCoveredPersons.MBMCICBCPPNo = '" & Trim(txtMBMIcBcPP) & "'"
    '    Else
    '        strAppend = " And dbo.MBMCoveredPersons.MBMCName = '" & tmpName
    '        strAppend = strAppend & "' And dbo.MBMCoveredPersons.MBMCDOB='" & tmpDOb & "'"
    '    End If
    'End If
    Set Cmd.ActiveConnection = medixmasterconn
    Cmd.CommandText = "SearchMBMHistory06"
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandTimeout = 300
    Set Prm1 = Cmd.CreateParameter("MBMType", adVarChar, adParamInput, 10, TmpMBMType)
    Cmd.Parameters.Append Prm1
    Set Prm2 = Cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
    Cmd.Parameters.Append Prm2

    TmpREc1.CursorLocation = adUseClient
    TmpREc1.CursorType = adOpenForwardOnly
    TmpREc1.CacheSize = 100
    Set TmpREc1 = Cmd.Execute
    Do While Not TmpREc1.EOF
        Cnt = Cnt + 1
        If Cnt = 1 Then
                tmpPolHistory = "Policy History" & _
                vbNewLine & "" & TmpREc1!MBMNumber & " - " & TmpREc1!MBMPolicyNo & " - " & TmpREc1!MBMPayorEffDate & " - " & TmpREc1!MBMPayorEXPDate & """"
        Else
            tmpPolHistory1 = ""
            
            tmpPolHistory1 = "" & _
                vbNewLine & "" & TmpREc1!MBMNumber & " - " & TmpREc1!MBMPolicyNo & " - " & TmpREc1!MBMPayorEffDate & " - " & TmpREc1!MBMPayorEXPDate & ""
             tmpPolHistory = tmpPolHistory + tmpPolHistory1
        End If
        'Set HISList = lstMBMHis.listitems.Add(, , TmpREc1!MBMNumber)
        'HISList.SubItems(1) = IIf(Len(TmpREc1!MBMCOVID), TmpREc1!MBMCOVID, "")
        'HISList.SubItems(2) = IIf(Len(TmpREc1!MBMTopupInd), TmpREc1!MBMTopupInd, "")
        'HISList.SubItems(3) = IIf(Len(TmpREc1!MBMTopupNumber), TmpREc1!MBMTopupNumber, "")
        'HISList.SubItems(4) = IIf(Len(TmpREc1!MBMPolicyNo), TmpREc1!MBMPolicyNo, "")
        'HISList.SubItems(5) = IIf(TmpREc1!MBMPayorEffDate <> "", TmpREc1!MBMPayorEffDate, "")
        'HISList.SubItems(6) = IIf(TmpREc1!MBMPayorEXPDate <> "", TmpREc1!MBMPayorEXPDate, "")
        'HISList.SubItems(7) = IIf(Len(TmpREc1!MBMDataStatus), TmpREc1!MBMDataStatus, "")
        'HISList.SubItems(8) = IIf(Len(TmpREc1!MBMStatus), TmpREc1!MBMStatus, "")
        'HISList.SubItems(9) = IIf(Len(TmpREc1!MBMBordxDate), TmpREc1!MBMBordxDate, "")
        TmpREc1.MoveNext
    Loop
    TmpREc1.Close
    Set TmpREc1 = Nothing

End Sub



