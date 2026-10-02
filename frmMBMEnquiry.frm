VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMBMEnquiry 
   Caption         =   "Member Enquiry"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   74
      Top             =   480
      Width           =   11655
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Go"
         Height          =   285
         Left            =   3240
         TabIndex        =   78
         Top             =   225
         Width           =   765
      End
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdSearchMBM 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   77
         Top             =   225
         Width           =   675
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&Pre"
         Height          =   285
         Left            =   3960
         TabIndex        =   76
         Top             =   225
         Width           =   765
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
         Height          =   285
         Left            =   4680
         TabIndex        =   75
         Top             =   225
         Width           =   690
      End
      Begin VB.Label txtMBMStatus 
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7040
         TabIndex        =   79
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblEXPStatus 
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   9260
         TabIndex        =   99
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label103 
         Caption         =   "Exp. Status"
         Height          =   330
         Left            =   8360
         TabIndex        =   100
         Top             =   240
         Width           =   915
      End
      Begin VB.Label txtDataStatus 
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   10580
         TabIndex        =   97
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label90 
         Caption         =   "Data Status"
         Height          =   330
         Left            =   9660
         TabIndex        =   96
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label76 
         Caption         =   "Mem Num."
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   210
         Width           =   1275
      End
      Begin VB.Label Label25 
         Caption         =   "Pol Status"
         Height          =   330
         Left            =   6240
         TabIndex        =   80
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   12015
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   15
         Left            =   240
         TabIndex        =   66
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Member Enquiry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   240
         TabIndex        =   67
         Top             =   120
         Width           =   6615
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5445
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9604
      _Version        =   393216
      Tabs            =   11
      TabsPerRow      =   11
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
      TabCaption(0)   =   "Principal &Detail"
      TabPicture(0)   =   "frmMBMEnquiry.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Supp&lementary"
      TabPicture(1)   =   "frmMBMEnquiry.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstCovAdd"
      Tab(1).Control(1)=   "Frame6"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Adjustment History"
      TabPicture(2)   =   "frmMBMEnquiry.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstAdjustmentHis"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "HIS &Benefits"
      TabPicture(3)   =   "frmMBMEnquiry.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstBenefitDetails"
      Tab(3).Control(1)=   "lstNewBenefitDetails"
      Tab(3).Control(2)=   "Label37"
      Tab(3).Control(3)=   "Label105"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "&Current Case"
      TabPicture(4)   =   "frmMBMEnquiry.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label7"
      Tab(4).Control(1)=   "lstCaseList"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Case History"
      TabPicture(5)   =   "frmMBMEnquiry.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lstCASHistory"
      Tab(5).Control(1)=   "Label13"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Mem &History"
      TabPicture(6)   =   "frmMBMEnquiry.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lstMBMHis"
      Tab(6).Control(1)=   "Label26"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Account"
      TabPicture(7)   =   "frmMBMEnquiry.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "lstMBMAcc"
      Tab(7).Control(1)=   "Frame9"
      Tab(7).Control(2)=   "Frame13"
      Tab(7).Control(3)=   "Frame12"
      Tab(7).Control(4)=   "Frame11"
      Tab(7).Control(5)=   "cmdTotal"
      Tab(7).ControlCount=   6
      TabCaption(8)   =   "Limitation"
      TabPicture(8)   =   "frmMBMEnquiry.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lstLimitation"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Surgical Proc"
      TabPicture(9)   =   "frmMBMEnquiry.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "lstSurgicalProc"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Fees"
      TabPicture(10)  =   "frmMBMEnquiry.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame2"
      Tab(10).ControlCount=   1
      Begin VB.Frame Frame6 
         Height          =   3075
         Left            =   -74880
         TabIndex        =   182
         Top             =   360
         Width           =   11460
         Begin VB.TextBox txtCovAllergic 
            Height          =   915
            Left            =   4680
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   200
            Top             =   2040
            Width           =   2415
         End
         Begin VB.TextBox txtCovExclusion 
            Height          =   915
            Left            =   7200
            MaxLength       =   2500
            MultiLine       =   -1  'True
            TabIndex        =   199
            Top             =   2040
            Width           =   2775
         End
         Begin VB.ComboBox cboCovSalutation 
            Height          =   315
            Left            =   1320
            TabIndex        =   198
            Text            =   "MS"
            Top             =   195
            Width           =   975
         End
         Begin VB.TextBox txtCovName 
            Height          =   330
            Left            =   3480
            MaxLength       =   60
            TabIndex        =   197
            Top             =   180
            Width           =   3660
         End
         Begin VB.TextBox txtCovIcBcPp 
            Height          =   285
            Left            =   7800
            MaxLength       =   30
            TabIndex        =   196
            Top             =   180
            Width           =   2055
         End
         Begin VB.ComboBox cboCovRelationship 
            Height          =   315
            Left            =   1320
            TabIndex        =   195
            Text            =   "W -Wife"
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox cboCovAdultChild 
            Height          =   315
            Left            =   1320
            TabIndex        =   194
            Text            =   "A"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtPrevPLNCode 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5760
            TabIndex        =   193
            Top             =   480
            Width           =   1380
         End
         Begin VB.TextBox txtAvailableLimit 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   192
            Top             =   1010
            Width           =   975
         End
         Begin VB.TextBox txtPrevINSCom 
            Height          =   330
            Left            =   7800
            MaxLength       =   60
            TabIndex        =   191
            Top             =   440
            Width           =   2055
         End
         Begin VB.TextBox txtRBAmt 
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   190
            Top             =   1290
            Width           =   975
         End
         Begin VB.TextBox txtCOVPayRemarks 
            Height          =   285
            Left            =   5760
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   189
            Top             =   720
            Width           =   4095
         End
         Begin VB.TextBox txtSuppGenWP 
            Height          =   285
            Left            =   5760
            MaxLength       =   9
            TabIndex        =   188
            Top             =   960
            Width           =   945
         End
         Begin VB.TextBox txtSuppPREWP 
            Height          =   285
            Left            =   5760
            MaxLength       =   9
            TabIndex        =   187
            Top             =   1200
            Width           =   945
         End
         Begin VB.TextBox TxtSuppSpeWP 
            Height          =   300
            Left            =   5760
            MaxLength       =   9
            TabIndex        =   186
            Top             =   1455
            Width           =   945
         End
         Begin MSMask.MaskEdBox txtCovBirthdate 
            Height          =   285
            Left            =   3480
            TabIndex        =   201
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
            Left            =   8880
            TabIndex        =   202
            Top             =   1200
            Width           =   975
            _ExtentX        =   1720
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
            Left            =   8880
            TabIndex        =   203
            Top             =   1440
            Width           =   975
            _ExtentX        =   1720
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
            Left            =   7320
            TabIndex        =   204
            Top             =   960
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
            Left            =   7320
            TabIndex        =   205
            Top             =   1190
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
            Left            =   7320
            TabIndex        =   206
            Top             =   1430
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
            TabIndex        =   185
            Text            =   "M"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtCovAnnualLimit 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   184
            Top             =   1005
            Width           =   1095
         End
         Begin VB.ComboBox cboSuppStaffDisc 
            Height          =   315
            Left            =   3480
            Sorted          =   -1  'True
            TabIndex        =   183
            Top             =   1260
            Width           =   1095
         End
         Begin VB.TextBox txtSuppPayPremGross 
            Height          =   285
            Left            =   1320
            TabIndex        =   238
            Top             =   1550
            Width           =   975
         End
         Begin VB.TextBox txtSuppPayMCOGross 
            Height          =   285
            Left            =   1320
            TabIndex        =   239
            Top             =   1805
            Width           =   975
         End
         Begin VB.TextBox txtRenewalDisc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   240
            Top             =   2060
            Width           =   975
         End
         Begin VB.TextBox txtSuppPayPremNet 
            Height          =   285
            Left            =   3480
            TabIndex        =   241
            Top             =   1550
            Width           =   1095
         End
         Begin VB.TextBox txtSuppPayMCONet 
            Height          =   285
            Left            =   3480
            TabIndex        =   242
            Top             =   1805
            Width           =   1095
         End
         Begin VB.TextBox txtPolicyDisc 
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   243
            Top             =   2060
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Policy Disc."
            Height          =   255
            Left            =   2475
            TabIndex        =   237
            Top             =   2040
            Width           =   1365
         End
         Begin VB.Label Label6 
            Caption         =   "Renewal Disc."
            Height          =   255
            Left            =   120
            TabIndex        =   236
            Top             =   2055
            Width           =   1485
         End
         Begin VB.Label Label33 
            Caption         =   "Pay MCO Net"
            Height          =   255
            Left            =   2475
            TabIndex        =   235
            Top             =   1815
            Width           =   1365
         End
         Begin VB.Label Label34 
            Caption         =   "Pay MCO Gross"
            Height          =   255
            Left            =   120
            TabIndex        =   234
            Top             =   1815
            Width           =   1365
         End
         Begin VB.Label Label35 
            Caption         =   "Pay Prem Net"
            Height          =   255
            Left            =   2480
            TabIndex        =   233
            Top             =   1560
            Width           =   1365
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            Caption         =   "To"
            Height          =   255
            Left            =   8400
            TabIndex        =   232
            Top             =   1440
            Width           =   285
         End
         Begin VB.Label Label41 
            Alignment       =   2  'Center
            Caption         =   "From"
            Height          =   255
            Left            =   8400
            TabIndex        =   231
            Top             =   1245
            Width           =   405
         End
         Begin VB.Label Label111 
            Caption         =   "Pay Prem Gross"
            Height          =   255
            Left            =   120
            TabIndex        =   230
            Top             =   1560
            Width           =   1365
         End
         Begin VB.Label Label42 
            Caption         =   "Clm Non-Pay Date"
            Height          =   375
            Left            =   8450
            TabIndex        =   229
            Top             =   1000
            Width           =   1365
         End
         Begin VB.Label Label53 
            Caption         =   "EOP"
            Height          =   255
            Left            =   6840
            TabIndex        =   228
            Top             =   1245
            Width           =   1005
         End
         Begin VB.Label Label104 
            Caption         =   "Spe. Ill. W. P."
            Height          =   255
            Left            =   4680
            TabIndex        =   227
            Top             =   1530
            Width           =   1065
         End
         Begin VB.Label Label43 
            Caption         =   "Pre W. P."
            Height          =   255
            Left            =   4680
            TabIndex        =   226
            Top             =   1245
            Width           =   780
         End
         Begin VB.Label Label44 
            Caption         =   "Gen W. P."
            Height          =   255
            Left            =   4680
            TabIndex        =   225
            Top             =   990
            Width           =   900
         End
         Begin VB.Label Label46 
            Caption         =   "EOP"
            Height          =   255
            Left            =   6840
            TabIndex        =   224
            Top             =   1000
            Width           =   1005
         End
         Begin VB.Label Label47 
            Caption         =   "EOP"
            Height          =   255
            Left            =   6840
            TabIndex        =   223
            Top             =   1500
            Width           =   1005
         End
         Begin VB.Label Label48 
            Caption         =   "Disc Req"
            Height          =   255
            Left            =   2480
            TabIndex        =   222
            Top             =   1320
            Width           =   1065
         End
         Begin VB.Label Label49 
            Caption         =   "Pay Remarks"
            Height          =   255
            Left            =   4680
            TabIndex        =   221
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label91 
            Caption         =   "T/O Ins "
            Height          =   255
            Left            =   7200
            TabIndex        =   220
            Top             =   480
            Width           =   1140
         End
         Begin VB.Label Label50 
            Caption         =   "R&&B Amt"
            Height          =   255
            Left            =   120
            TabIndex        =   219
            Top             =   1320
            Width           =   1140
         End
         Begin VB.Label Label51 
            Caption         =   "P. Pln Code"
            Height          =   255
            Left            =   4680
            TabIndex        =   218
            Top             =   525
            Width           =   1125
         End
         Begin VB.Label Label52 
            Caption         =   "Name"
            Height          =   255
            Left            =   2480
            TabIndex        =   217
            Top             =   195
            Width           =   1140
         End
         Begin VB.Label Label62 
            Caption         =   "Adult/Child"
            Height          =   255
            Left            =   120
            TabIndex        =   216
            Top             =   795
            Width           =   825
         End
         Begin VB.Label Label65 
            Caption         =   "Sex "
            Height          =   255
            Left            =   2480
            TabIndex        =   215
            Top             =   735
            Width           =   825
         End
         Begin VB.Label Label66 
            Caption         =   "Allergic"
            Height          =   255
            Left            =   4680
            TabIndex        =   214
            Top             =   1800
            Width           =   600
         End
         Begin VB.Label Label93 
            Caption         =   "Salutation"
            Height          =   255
            Left            =   120
            TabIndex        =   213
            Top             =   225
            Width           =   1140
         End
         Begin VB.Label Label94 
            Caption         =   "IC Num"
            Height          =   255
            Left            =   7200
            TabIndex        =   212
            Top             =   165
            Width           =   1125
         End
         Begin VB.Label Label95 
            Caption         =   "DOB"
            Height          =   255
            Left            =   2480
            TabIndex        =   211
            Top             =   495
            Width           =   1020
         End
         Begin VB.Label Label96 
            Caption         =   "Supp Lmt"
            Height          =   255
            Left            =   2480
            TabIndex        =   210
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label97 
            Caption         =   "R'Ship"
            Height          =   255
            Left            =   120
            TabIndex        =   209
            Top             =   525
            Width           =   1005
         End
         Begin VB.Label Label98 
            Caption         =   "Supplementary Warranty "
            Height          =   255
            Left            =   7200
            TabIndex        =   208
            Top             =   1800
            Width           =   2100
         End
         Begin VB.Label Label99 
            Caption         =   "Bal. Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   207
            Top             =   1020
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdTotal 
         Caption         =   "Total"
         Height          =   375
         Left            =   -64440
         TabIndex        =   177
         Top             =   2640
         Width           =   855
      End
      Begin VB.Frame Frame11 
         Caption         =   "Others"
         Height          =   1215
         Left            =   -69360
         TabIndex        =   167
         Top             =   1800
         Width           =   4815
         Begin VB.TextBox txtReimbursement 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   170
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtExcess 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   169
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtExcessPaid 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   168
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label92 
            Caption         =   "Excess Paid"
            Height          =   255
            Left            =   120
            TabIndex        =   173
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label78 
            Caption         =   "Total Excess"
            Height          =   255
            Left            =   120
            TabIndex        =   172
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label80 
            Caption         =   "Total Reimburse"
            Height          =   255
            Left            =   120
            TabIndex        =   171
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Payable to Member"
         Height          =   1335
         Left            =   -69360
         TabIndex        =   158
         Top             =   480
         Width           =   4815
         Begin VB.TextBox txtPayableMem 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   162
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtPaid2MBM 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   161
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtPaidByMem 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   160
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtMemOS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   159
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label81 
            Caption         =   "Paid To Mem"
            Height          =   255
            Left            =   120
            TabIndex        =   166
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label82 
            Caption         =   "O/S"
            Height          =   255
            Left            =   120
            TabIndex        =   165
            Top             =   960
            Width           =   1425
         End
         Begin VB.Label Label84 
            Caption         =   "Received From Mem"
            Height          =   255
            Left            =   120
            TabIndex        =   164
            Top             =   720
            Width           =   1785
         End
         Begin VB.Label txtMBMPayable 
            Caption         =   "Payable to/(by) Mem"
            Height          =   255
            Left            =   120
            TabIndex        =   163
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Worksheet Summary"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   149
         Top             =   480
         Width           =   5055
         Begin VB.TextBox txtAnnualLmt 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   153
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtAvailableAnnualLmt 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   152
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtApprovedAmt 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   151
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtBordxAmt 
            Height          =   285
            Left            =   1320
            TabIndex        =   150
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label89 
            Caption         =   "Bordx Amt"
            Height          =   255
            Left            =   120
            TabIndex        =   157
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label Label86 
            Caption         =   "Annual Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   156
            Top             =   255
            Width           =   1065
         End
         Begin VB.Label Label87 
            Caption         =   "Approved Amt"
            Height          =   255
            Left            =   120
            TabIndex        =   155
            Top             =   735
            Width           =   1065
         End
         Begin VB.Label Label88 
            Caption         =   "Balance Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   154
            Top             =   495
            Width           =   945
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Payable To Hospital"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   142
         Top             =   1800
         Width           =   5055
         Begin VB.TextBox txtPayable2Hos 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   145
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtPaid2Hos 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   144
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtHosOS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   143
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label59 
            Caption         =   "Bill Amt"
            Height          =   255
            Left            =   240
            TabIndex        =   148
            Top             =   255
            Width           =   1065
         End
         Begin VB.Label Label67 
            Caption         =   "O/S "
            Height          =   255
            Left            =   240
            TabIndex        =   147
            Top             =   735
            Width           =   825
         End
         Begin VB.Label Label68 
            Caption         =   "Paid Amount"
            Height          =   255
            Left            =   240
            TabIndex        =   146
            Top             =   495
            Width           =   945
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1275
         Left            =   -74760
         TabIndex        =   121
         Top             =   600
         Width           =   8625
         Begin VB.Label Label11 
            Caption         =   "Access Fee"
            Height          =   255
            Left            =   5880
            TabIndex        =   139
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label lblMCOFee 
            Caption         =   "MCO Fees"
            Height          =   255
            Left            =   2955
            TabIndex        =   138
            Top             =   195
            Width           =   870
         End
         Begin VB.Label Label85 
            Caption         =   "Effective MCO Fee "
            Height          =   255
            Left            =   2955
            TabIndex        =   137
            Top             =   915
            Width           =   1455
         End
         Begin VB.Label Label83 
            Caption         =   "Effective Prem "
            Height          =   255
            Left            =   120
            TabIndex        =   136
            Top             =   915
            Width           =   1140
         End
         Begin VB.Label txtMCOFees 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            TabIndex        =   135
            Top             =   180
            Width           =   945
         End
         Begin VB.Label txtMBMPremiumEff 
            Alignment       =   1  'Right Justify
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
            Left            =   1320
            TabIndex        =   134
            Top             =   915
            Width           =   1035
         End
         Begin VB.Label txtMCOFeesEff 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            TabIndex        =   133
            Top             =   915
            Width           =   945
         End
         Begin VB.Label txtMBMProviderCharges 
            Alignment       =   1  'Right Justify
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
            Left            =   7560
            TabIndex        =   132
            Top             =   180
            Width           =   945
         End
         Begin VB.Label lblPremium 
            Caption         =   "Premium"
            Height          =   255
            Left            =   120
            TabIndex        =   131
            Top             =   195
            Width           =   960
         End
         Begin VB.Label txtMBMPremium 
            Alignment       =   1  'Right Justify
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
            Left            =   1320
            TabIndex        =   130
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label Label58 
            Caption         =   "Prem Refund"
            Height          =   255
            Left            =   120
            TabIndex        =   129
            Top             =   560
            Width           =   1020
         End
         Begin VB.Label txtMBMPremiumRed 
            Alignment       =   1  'Right Justify
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
            Left            =   1320
            TabIndex        =   128
            Top             =   560
            Width           =   1035
         End
         Begin VB.Label txtMBMMCORed 
            Alignment       =   1  'Right Justify
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
            Left            =   4440
            TabIndex        =   127
            Top             =   555
            Width           =   945
         End
         Begin VB.Label Label61 
            Caption         =   "MCO Refund"
            Height          =   255
            Left            =   2955
            TabIndex        =   126
            Top             =   555
            Width           =   975
         End
         Begin VB.Label Label79 
            Caption         =   "Acces Fee Refund"
            Height          =   255
            Left            =   5880
            TabIndex        =   125
            Top             =   555
            Width           =   1380
         End
         Begin VB.Label Label77 
            Caption         =   "Effective Access Fee"
            Height          =   255
            Left            =   5880
            TabIndex        =   124
            Top             =   915
            Width           =   1620
         End
         Begin VB.Label txtMBMProviderEffCharges 
            Alignment       =   1  'Right Justify
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
            Left            =   7560
            TabIndex        =   123
            Top             =   915
            Width           =   945
         End
         Begin VB.Label txtMBMProviderRedCharges 
            Alignment       =   1  'Right Justify
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
            Left            =   7560
            TabIndex        =   122
            Top             =   555
            Width           =   945
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   8520
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   8520
            Y1              =   840
            Y2              =   840
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5010
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   11445
         Begin VB.TextBox txtRemarks 
            Height          =   1185
            Left            =   5640
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   3600
            Width           =   5295
         End
         Begin VB.TextBox txtMBMExclusion 
            Height          =   1185
            Left            =   120
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   3600
            Width           =   4860
         End
         Begin VB.TextBox txtMBMAdd1 
            Height          =   285
            Left            =   1080
            MaxLength       =   40
            TabIndex        =   9
            Top             =   195
            Width           =   4065
         End
         Begin VB.TextBox cboINSCode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6840
            TabIndex        =   4
            Top             =   210
            Width           =   1875
         End
         Begin VB.TextBox cboPLNCode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   9960
            TabIndex        =   3
            Top             =   210
            Width           =   1200
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
            Height          =   315
            Left            =   6840
            TabIndex        =   5
            Top             =   480
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMBMBordxDate 
            Height          =   300
            Left            =   6840
            TabIndex        =   6
            Top             =   760
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
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
            Left            =   9960
            TabIndex        =   8
            Top             =   480
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtMBMAdd2 
            Height          =   285
            Left            =   1080
            MaxLength       =   40
            TabIndex        =   10
            Top             =   450
            Width           =   4065
         End
         Begin VB.TextBox txtMBMAdd3 
            Height          =   285
            Left            =   1080
            MaxLength       =   40
            TabIndex        =   11
            Top             =   710
            Width           =   4065
         End
         Begin VB.TextBox txtMBMPostCode 
            Height          =   285
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   12
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox cboCTYCode 
            Height          =   315
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   960
            Width           =   1545
         End
         Begin VB.ComboBox cboSTACode 
            Height          =   315
            Left            =   1080
            TabIndex        =   14
            Top             =   1220
            Width           =   4065
         End
         Begin VB.ComboBox cboMBMSex 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   1505
            Width           =   1335
         End
         Begin VB.ComboBox cboMBMTakeOver 
            Height          =   315
            Left            =   3600
            TabIndex        =   16
            Top             =   1505
            Width           =   1545
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
            Left            =   1080
            TabIndex        =   17
            Top             =   1770
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboMBMMaritalStatus 
            Height          =   315
            Left            =   3600
            TabIndex        =   18
            Top             =   1770
            Width           =   1545
         End
         Begin VB.ComboBox cboRACCode 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   19
            Top             =   2030
            Width           =   1335
         End
         Begin VB.ComboBox cboNATCode 
            Height          =   315
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   20
            Top             =   2030
            Width           =   1545
         End
         Begin VB.TextBox cboBRNCode 
            Height          =   315
            Left            =   1080
            TabIndex        =   21
            Top             =   2300
            Width           =   4065
         End
         Begin VB.ComboBox cboMBMRenewal 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   23
            Top             =   2580
            Width           =   1290
         End
         Begin VB.TextBox txtMBMAgentCode 
            Height          =   315
            Left            =   3600
            MaxLength       =   15
            TabIndex        =   22
            Top             =   2580
            Width           =   1545
         End
         Begin VB.TextBox txtMBMDateJoin 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6840
            MaxLength       =   25
            TabIndex        =   24
            Top             =   1020
            Width           =   1875
         End
         Begin VB.TextBox txtMBMPolicyYear 
            Enabled         =   0   'False
            Height          =   315
            Left            =   9960
            MaxLength       =   25
            TabIndex        =   7
            Top             =   1010
            Width           =   1200
         End
         Begin VB.TextBox txtMBMEmployeeNo 
            Height          =   315
            Left            =   6840
            MaxLength       =   15
            TabIndex        =   25
            Top             =   1290
            Width           =   4335
         End
         Begin VB.TextBox txtDeptName 
            Height          =   315
            Left            =   6840
            MaxLength       =   150
            TabIndex        =   26
            Top             =   1580
            Width           =   4335
         End
         Begin VB.TextBox txtMBMGroupCompany 
            Height          =   315
            Left            =   6840
            MaxLength       =   150
            TabIndex        =   27
            Top             =   1860
            Width           =   4335
         End
         Begin VB.TextBox txtAGECode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6840
            TabIndex        =   28
            Top             =   2140
            Width           =   4335
         End
         Begin VB.ComboBox cboSRVCode 
            Height          =   315
            Left            =   6840
            Sorted          =   -1  'True
            TabIndex        =   29
            Text            =   "EA - Europe Assistance"
            Top             =   2410
            Width           =   4335
         End
         Begin VB.ComboBox CboInsMode 
            Height          =   315
            Left            =   9840
            TabIndex        =   32
            Top             =   2700
            Width           =   1305
         End
         Begin MSMask.MaskEdBox txtNewPolEffDate 
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
            Left            =   9840
            TabIndex        =   36
            Top             =   2955
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboMBMRelCode 
            Height          =   315
            Left            =   6840
            TabIndex        =   30
            Top             =   2700
            Width           =   1815
         End
         Begin VB.TextBox txtPLNCode 
            Height          =   285
            Left            =   6840
            MaxLength       =   4
            TabIndex        =   33
            Top             =   2950
            Width           =   1815
         End
         Begin VB.Label Label23 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   5640
            TabIndex        =   181
            Top             =   3360
            Width           =   1590
         End
         Begin VB.Label Label40 
            Caption         =   "Principal Warranty "
            Height          =   255
            Left            =   120
            TabIndex        =   180
            Top             =   3360
            Width           =   1590
         End
         Begin VB.Label Label39 
            Caption         =   "New PLN Code"
            Height          =   255
            Left            =   5640
            TabIndex        =   102
            Top             =   3045
            Width           =   1125
         End
         Begin VB.Label Label38 
            Caption         =   "New Eff Date"
            Height          =   255
            Left            =   8760
            TabIndex        =   101
            Top             =   3000
            Width           =   1005
         End
         Begin VB.Label Label30 
            Caption         =   "Supp Type"
            Height          =   255
            Left            =   5640
            TabIndex        =   64
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label27 
            Caption         =   "Marital Status"
            Height          =   255
            Left            =   2520
            TabIndex        =   63
            Top             =   1860
            Width           =   975
         End
         Begin VB.Label Label71 
            Caption         =   "Nationality"
            Height          =   275
            Left            =   2520
            TabIndex        =   62
            Top             =   2080
            Width           =   960
         End
         Begin VB.Label Label70 
            Caption         =   "Race"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   2140
            Width           =   420
         End
         Begin VB.Label Label12 
            Caption         =   "Branch"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   2400
            Width           =   960
         End
         Begin VB.Label Label63 
            Caption         =   "TakeOver"
            Height          =   255
            Left            =   2520
            TabIndex        =   59
            Top             =   1620
            Width           =   825
         End
         Begin VB.Label Label15 
            Caption         =   "Eff Date"
            Height          =   255
            Left            =   5640
            TabIndex        =   58
            Top             =   510
            Width           =   1275
         End
         Begin VB.Label Label16 
            Caption         =   "Exp Date"
            Height          =   255
            Left            =   8880
            TabIndex        =   57
            Top             =   450
            Width           =   795
         End
         Begin VB.Label Label9 
            Caption         =   "Pol Year"
            Height          =   255
            Left            =   8880
            TabIndex        =   56
            Top             =   1080
            Width           =   1050
         End
         Begin VB.Label Label56 
            Caption         =   "Sex "
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   1560
            Width           =   825
         End
         Begin VB.Label Label54 
            Caption         =   "Address"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label57 
            Caption         =   "City/Town"
            Height          =   255
            Left            =   2520
            TabIndex        =   53
            Top             =   1020
            Width           =   1020
         End
         Begin VB.Label Label60 
            Caption         =   "State"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1320
            Width           =   465
         End
         Begin VB.Label Label69 
            Caption         =   "Birthdate"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label Label72 
            Caption         =   "PostCode"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1015
            Width           =   1140
         End
         Begin VB.Label Label29 
            Caption         =   "Age Band"
            Height          =   255
            Left            =   5640
            TabIndex        =   49
            Top             =   2170
            Width           =   735
         End
         Begin VB.Label Label21 
            Caption         =   "Plan Code"
            Height          =   255
            Left            =   8880
            TabIndex        =   48
            Top             =   210
            Width           =   825
         End
         Begin VB.Label Label20 
            Caption         =   "Ins Type"
            Height          =   255
            Left            =   5640
            TabIndex        =   47
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label14 
            Caption         =   "Agent Code"
            Height          =   255
            Left            =   2520
            TabIndex        =   46
            Top             =   2640
            Width           =   1140
         End
         Begin VB.Label Label17 
            Caption         =   "Bordx. Date"
            Height          =   255
            Left            =   5640
            TabIndex        =   45
            Top             =   780
            Width           =   1275
         End
         Begin VB.Label Label8 
            Caption         =   "Batch Num"
            Height          =   255
            Left            =   8880
            TabIndex        =   44
            Top             =   760
            Width           =   885
         End
         Begin VB.Label Label10 
            Caption         =   "Date Join"
            Height          =   255
            Left            =   5640
            TabIndex        =   43
            Top             =   1080
            Width           =   780
         End
         Begin VB.Label Label19 
            Caption         =   "Employee Num"
            Height          =   255
            Left            =   5640
            TabIndex        =   42
            Top             =   1400
            Width           =   1545
         End
         Begin VB.Label Label18 
            Caption         =   "Group Company"
            Height          =   255
            Left            =   5640
            TabIndex        =   41
            Top             =   1935
            Width           =   1545
         End
         Begin VB.Label Label22 
            Caption         =   "Service Provider"
            Height          =   255
            Left            =   5640
            TabIndex        =   40
            Top             =   2480
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Dept. Name"
            Height          =   255
            Left            =   5640
            TabIndex        =   39
            Top             =   1680
            Width           =   1545
         End
         Begin VB.Label Label45 
            Caption         =   "Renewal "
            Height          =   240
            Left            =   120
            TabIndex        =   38
            Top             =   2640
            Width           =   1065
         End
         Begin VB.Label txtMBMBatchNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   9960
            TabIndex        =   37
            Top             =   780
            Width           =   1200
         End
         Begin VB.Label Label55 
            Caption         =   "Ins mode"
            Height          =   255
            Left            =   8760
            TabIndex        =   31
            Top             =   2760
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView lstCovAdd 
         Height          =   1845
         Left            =   -74880
         TabIndex        =   109
         Top             =   3480
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   3254
         SortKey         =   9
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
         NumItems        =   31
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
            Object.Width           =   2540
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
            Text            =   "Blood"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Allergic"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Exclusion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Status"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "AnnualLimit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Available Limit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Occupation"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Work Nature"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Smoking"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Alcohol"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Premium"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "Prev Plan"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "RB Amt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Payor Remarks"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "Prev Insurance Com"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Clm Non Pay Fr"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Clm Non Pay To"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   25
            Text            =   "Payor Prem"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "Payor Prem Net"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "Payor MCO Gross"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "Payor MCO Net"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Text            =   "RenewalDisc"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Text            =   "Pol Disc"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstAdjustmentHis 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   110
         Top             =   360
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   8705
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Bordx Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Batch No"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Adjust. Date"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Endt. Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Endt. Eff Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstBenefitDetails 
         Height          =   4755
         Left            =   -74880
         TabIndex        =   111
         Top             =   600
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   8387
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "BNF Desc"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Days"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "RM PDay"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Per Dis"
            Object.Width           =   1587
         EndProperty
      End
      Begin MSComctlLib.ListView lstNewBenefitDetails 
         Height          =   4755
         Left            =   -69105
         TabIndex        =   112
         Top             =   600
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   8387
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "BNF Desc"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Days"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "RM PDay"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Per Dis"
            Object.Width           =   1587
         EndProperty
      End
      Begin MSComctlLib.ListView lstCaseList 
         Height          =   4755
         Left            =   -74880
         TabIndex        =   115
         Top             =   660
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   8387
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Case No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Mem No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Patient"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Rel"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DOA"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "DOD"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Claimability"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "FinalDiag"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Case Status"
            Object.Width           =   1905
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Hosp"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Payor"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstSurgicalProc 
         Height          =   4875
         Left            =   -74880
         TabIndex        =   140
         Top             =   360
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   8599
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sur Proc Cat"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Sur Proc Cat Desc"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Sur Proc Code"
            Object.Width           =   2294
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Sur Proc Code Desc"
            Object.Width           =   9879
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Percentage"
            Object.Width           =   1941
         EndProperty
      End
      Begin MSComctlLib.ListView lstLimitation 
         Height          =   4905
         Left            =   -74880
         TabIndex        =   141
         Top             =   360
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   8652
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Relation"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Gen WP"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "End of Gen WP"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Pre WP"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "End of Pre WP"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Spec Illness WP"
            Object.Width           =   2910
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "End of Spec Illness WP"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Staff Disclosure Req"
            Object.Width           =   2822
         EndProperty
      End
      Begin MSComctlLib.ListView lstMBMAcc 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   174
         Top             =   3120
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   3836
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MBMNumber"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cov ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Approved Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Pay2HOS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Paid2HOS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "PAY2MBM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Paid2MBM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "PaidByMBM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Reimburse"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Excess"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Bordx Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Excess Paid"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstMBMHis 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   175
         Top             =   600
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8070
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
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "MBM Reg Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstCASHistory 
         Height          =   4755
         Left            =   -74880
         TabIndex        =   178
         Top             =   600
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   8387
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Case No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Mem No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Patient"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Rel"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DOA"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "DOD"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Claimability"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "FinalDiag"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Case Status"
            Object.Width           =   1905
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Hosp"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Payor"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label13 
         Caption         =   "Double click to open the case"
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
         Left            =   -74880
         TabIndex        =   179
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label26 
         Caption         =   "Double Click to Open Old Membership data (Members With Same IC are Listed Below)"
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
         Left            =   -74880
         TabIndex        =   176
         Top             =   360
         Width           =   8055
      End
      Begin VB.Label Label7 
         Caption         =   "Double click to open the case"
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
         Left            =   -74880
         TabIndex        =   116
         Top             =   420
         Width           =   3135
      End
      Begin VB.Label Label37 
         Caption         =   "Existing Bnf Listing"
         Height          =   255
         Left            =   -74820
         TabIndex        =   114
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label105 
         Caption         =   "New Bnf Listing"
         Height          =   255
         Left            =   -69060
         TabIndex        =   113
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame7 
      Height          =   525
      Left            =   120
      TabIndex        =   68
      Top             =   7440
      Width           =   11655
      Begin VB.CommandButton CmdMBMOthers 
         Caption         =   "&Others Details"
         Enabled         =   0   'False
         Height          =   315
         Left            =   8040
         TabIndex        =   107
         Top             =   150
         Width           =   1200
      End
      Begin VB.CommandButton CmdClinical 
         Caption         =   "&Clinical"
         Enabled         =   0   'False
         Height          =   315
         Left            =   8280
         TabIndex        =   108
         Top             =   150
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox txtcount 
         Height          =   285
         Left            =   360
         TabIndex        =   98
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   9240
         TabIndex        =   95
         Top             =   150
         Width           =   960
      End
      Begin VB.TextBox txtLastEdtDate 
         Height          =   285
         Left            =   2760
         TabIndex        =   73
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtLastEndStatus 
         Height          =   285
         Left            =   1680
         TabIndex        =   72
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtMBMAdultChild 
         Height          =   285
         Left            =   840
         TabIndex        =   71
         Tag             =   "T"
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   10200
         TabIndex        =   70
         Top             =   150
         Width           =   960
      End
      Begin VB.TextBox txtFirstMeMNo 
         Height          =   285
         Left            =   4080
         TabIndex        =   69
         Text            =   "FirstMeMNo"
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame10 
      Height          =   975
      Left            =   120
      TabIndex        =   82
      Top             =   1080
      Width           =   11655
      Begin VB.TextBox txtMBMName 
         Height          =   330
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   85
         Top             =   187
         Width           =   3075
      End
      Begin VB.ComboBox cboMBMSalutation 
         Height          =   315
         Left            =   1080
         TabIndex        =   84
         Top             =   195
         Width           =   795
      End
      Begin VB.TextBox txtMBMIcBcPP 
         Height          =   285
         Left            =   6840
         MaxLength       =   30
         TabIndex        =   83
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtMBMPolNo 
         Height          =   330
         Left            =   3000
         MaxLength       =   25
         TabIndex        =   87
         Top             =   480
         Width           =   3075
      End
      Begin VB.ComboBox cboPAYCode 
         Height          =   315
         Left            =   1080
         TabIndex        =   86
         Top             =   480
         Width           =   795
      End
      Begin VB.Label Label32 
         Caption         =   "H&&C"
         Height          =   255
         Left            =   10560
         TabIndex        =   120
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label28 
         Caption         =   "Clinical"
         Height          =   255
         Left            =   10560
         TabIndex        =   119
         Top             =   495
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label LblClinical 
         Alignment       =   1  'Right Justify
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
         Left            =   11040
         TabIndex        =   118
         Top             =   495
         Visible         =   0   'False
         Width           =   480
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
         Left            =   11040
         TabIndex        =   117
         Top             =   210
         Width           =   480
      End
      Begin VB.Label txtAvailableAnnualLimit 
         Alignment       =   1  'Right Justify
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
         Left            =   9480
         TabIndex        =   106
         Top             =   495
         Width           =   960
      End
      Begin VB.Label cboHLTCode 
         BackColor       =   &H80000009&
         Height          =   285
         Left            =   6840
         TabIndex        =   88
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label txtAnnualLimit 
         Alignment       =   1  'Right Justify
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
         Left            =   9480
         TabIndex        =   105
         Top             =   210
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "Bal Lmt"
         Height          =   255
         Left            =   8760
         TabIndex        =   104
         Top             =   480
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Ann Lmt"
         Height          =   255
         Left            =   8760
         TabIndex        =   103
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label75 
         Caption         =   "Hlt Type"
         Height          =   255
         Left            =   6120
         TabIndex        =   90
         Top             =   165
         Width           =   795
      End
      Begin VB.Label Label74 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   94
         Top             =   520
         Width           =   915
      End
      Begin VB.Label Label73 
         Caption         =   "Policy Num"
         Height          =   255
         Left            =   2040
         TabIndex        =   93
         Top             =   525
         Width           =   900
      End
      Begin VB.Label Label31 
         Caption         =   "Mem Name"
         Height          =   255
         Left            =   2040
         TabIndex        =   92
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label64 
         Caption         =   "IC Num"
         Height          =   330
         Left            =   6120
         TabIndex        =   91
         Top             =   480
         Width           =   795
      End
      Begin VB.Label Label24 
         Caption         =   "Salutation"
         Height          =   255
         Left            =   240
         TabIndex        =   89
         Top             =   240
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport crptSummary 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmMBMEnquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bExit As Boolean
Public bDataStatus As Boolean
Public ReportCriteria As String
Dim ScanCode As String
Private m_blnDirection(1 To 9) As Boolean

Private Sub CmdClinical_Click()
    frmMBMClinical.Source = 3
    Call frmMBMClinical.showMBMClinical
    frmMBMClinical.show
    frmMBMClinical.Frame1.Enabled = False
End Sub

Private Sub CmdMBMOthers_Click()
    FrmMBMOthers.Source = 2
    Call FrmMBMOthers.showMBMOthers
    FrmMBMOthers.show
    'FrmMBMOthers.cmdClose_Click
    FrmMBMOthers.txtMBMTelNoH.Enabled = False
    FrmMBMOthers.txtMBMTelNoM.Enabled = False
    FrmMBMOthers.txtMBMTelNoO.Enabled = False
    FrmMBMOthers.TxtPOLSubNo.Enabled = False
    FrmMBMOthers.txtMBMPrevPlan.Enabled = False
    FrmMBMOthers.txtMBMRBAmt.Enabled = False
    FrmMBMOthers.txtMBMPrevINSCom.Enabled = False
    FrmMBMOthers.txtMBMPAYRemarks.Enabled = False
    FrmMBMOthers.cboCOPayRB.Enabled = False
    FrmMBMOthers.Check1.Enabled = False
    FrmMBMOthers.Frame11.Enabled = False
    FrmMBMOthers.Frame12.Enabled = False
    FrmMBMOthers.Frame14.Enabled = False
    FrmMBMOthers.cboBTBG.Enabled = False
    FrmMBMOthers.txtPayPrem.Enabled = False
    FrmMBMOthers.txtMBMAllergic.Enabled = True
    'FrmMBMOthers.txtRemarks.Enabled = True
    'FrmMBMOthers.txtMBMExclusion.Enabled = True
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

Private Sub CmdPrint_Click()
    'RptMBMInquiry.show
    'Call PrintTxtBoxes(frmMBMEnquiry)
    Screen.MousePointer = vbHourglass
        
    If txtMBMNumber <> "" Then
        'RptCaseInquiry.show
        
        crptSummary.ReportFileName = crdPath & "\Report\RptMbmEnquiry.rpt"
        crptSummary.Destination = crptToWindow
        crptSummary.WindowState = crptMaximized
        crptSummary.SelectionFormula = "{VIEW_MBM_ENQUIRY.MBMNumber} = '" & Trim(txtMBMNumber) & "'"
        crptSummary.ParameterFields(0) = "MemberNo;" & Trim(txtMBMNumber) & ";true"
        crptSummary.Action = 1
        
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
              
    Screen.MousePointer = Default
End Sub

Private Sub CmdTotal_Click()
Dim Acc As New ADODB.Recordset
Dim strsql As String
Dim masterlist As ListItem
Dim strsqlMBM As String
Dim MBM As New ADODB.Recordset
Dim strsqlMBMArc As String
Dim MBMArc As New ADODB.Recordset
Dim totalApprove As Variant
Dim totalPayable2Hos As Variant
Dim totalPaid2Hos As Variant
Dim totalPayable2Mbm As Variant
Dim totalPaidByMbm As Variant
Dim totalPaid2MBM As Variant
Dim TotalExcess As Variant
Dim TotalReimburse As Variant
Dim MBMTotalAnnualLmt As Variant
Dim MBMTotalAvailLmt As Variant

    If lstMBMAcc.listitems.count <> 0 Then
        strsqlMBM = " Select MBMNumber, MBMMemberType, MBMCAnnualLimit, " & _
                    " MBMCAvailableLimit " & _
                    " From View_MBM_MemberType where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
        
    If MBM.recordCount Then
        If MBM!MBMMemberType = "P" Then
            txtAnnualLmt = txtAnnualLimit
            txtAvailableAnnualLmt = txtAvailableAnnualLimit
        Else
            Do Until MBM.EOF
                MBMTotalAnnualLmt = Format(MBMTotalAnnualLmt + MBM!MBMCAnnualLimit, "#,##0.00")
                MBMTotalAvailLmt = Format(MBMTotalAvailLmt + MBM!MBMCAvailableLimit, "#,##0.00")
            MBM.MoveNext
            Loop
                txtAnnualLmt = Format(MBMTotalAnnualLmt + txtAnnualLimit, "#,##0.00")
                txtAvailableAnnualLmt = Format(MBMTotalAvailLmt + txtAvailableAnnualLimit, "#,##0.00")
        End If
        MBM.Close
        Set MBM = Nothing
    Else
    
        strsqlMBMArc = " Select MBMNumber, MBMMemberType, MBMCAnnualLimit, " & _
                       " MBMCAvailableLimit, From View_MBM_MemTypeArc where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        MBMArc.Open strsqlMBMArc, medixmasterconn, adOpenKeyset, adLockOptimistic

        If MBMArc.recordCount Then
            If MBMArc!MBMMemberType = "P" Then
                txtAnnualLmt = Format(txtAnnualLimit, "#,##0.00")
                txtAvailableAnnualLmt = Format(txtAvailableAnnualLimit, "#,##0.00")
            Else
                Do Until MBM.EOF
                    MBMTotalAnnualLmt = Format(MBMTotalAnnualLmt + MBMArc!MBMCAnnualLimit, "#,##0.00")
                    MBMTotalAvailLmt = Format(MBMTotalAvailLmt + MBMArc!MBMCAvailableLimit, "#,##0.00")
                MBMArc.MoveNext
                Loop
                    txtAnnualLmt = Format(MBMTotalAnnualLmt + txtAnnualLimit, "#,##0.00")
                    txtAvailableAnnualLmt = Format(MBMTotalAvailLmt + txtAvailableAnnualLimit, "#,##0.00")
            End If
        End If
        MBMArc.Close
        Set MBMArc = Nothing
    End If

        strsql = " select totalApprove, totalPayable2Hos, " & _
                " totalPaid2Hos, totalPayable2Mbm, totalPaidByMbm, MBMPatientCOVID, " & _
                " totalPaid2MBM , totalExcess , totalReimburse from view_MBMTotal_summary where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        Acc.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If Not Acc.EOF Then
            Do Until Acc.EOF
            With Acc
                totalApprove = Format(totalApprove + !totalApprove, "#,##0.00")
                totalPayable2Hos = Format(totalPayable2Hos + !totalPayable2Hos, "#,##0.00")
                If Trim(!totalPaid2Hos) <> "" Then
                    totalPaid2Hos = Format(totalPaid2Hos + !totalPaid2Hos, "#,##0.00")
                End If
                totalPayable2Mbm = Format(totalPayable2Mbm + !totalPayable2Mbm, "#,##0.00")
                If Trim(!totalPaid2MBM) <> "" Then
                    totalPaid2MBM = Format(totalPaid2MBM + totalPaid2MBM, "#,##0.00")
                End If
                If Trim(!totalPaidByMbm) <> "" Then
                    totalPaidByMbm = Format(totalPaidByMbm + totalPaidByMbm, "#,##0.00")
                End If
                'txtMemOS = Format(CDbl(txtPayableMem) - CDbl(txtPaid2MBM) + CDbl(txtPaidByMem), "#,#.00;(#,#.00)")
                TotalReimburse = Format(TotalReimburse + TotalReimburse, "#,##0.00")
                TotalExcess = Format(TotalExcess + !TotalExcess, "#,##0.00")
            End With
            Acc.MoveNext
            Loop
            
            txtApprovedAmt = Format(totalApprove, "#,##0.00")
            
            txtPayable2Hos = totalPayable2Hos
            If IsNumeric(txtPayable2Hos) = False Then
                txtPayable2Hos = "0.00"
            End If
            
            txtPaid2Hos = totalPaid2Hos
            If IsNumeric(txtPaid2Hos) = False Then
                txtPaid2Hos = "0.00"
            End If
            ' Out standing hos
            txtHosOS = Format(CDbl(txtPayable2Hos) - CDbl(txtPaid2Hos), "#,##0.00;(#,##0.00)")
            
            txtPayableMem = Format(totalPayable2Mbm, "#,##0.00;(#,##0.00)")
            
            txtPaid2MBM = Format(totalPaid2MBM, "#,##0.00")
            If IsNumeric(txtPaid2MBM) = False Then
                txtPaid2MBM = "0.00"
            End If
                
            txtPaidByMem = Format(totalPaidByMbm, "#,##0.00")
            If IsNumeric(txtPaidByMem) = False Then
                txtPaidByMem = "0.00"
            End If
            ' Out standing MBM
            txtMemOS = Format(CDbl(txtPayableMem) - CDbl(txtPaid2MBM) + CDbl(txtPaidByMem), "#,#.00;(#,#.00)")
            txtReimbursement = Format(TotalReimburse, "#,##0.00;(#,##0.00)")
            txtExcess = Format(TotalExcess, "#,##0.00;(#,##0.00)")
        End If
        Acc.Close
        Set Acc = Nothing
    End If
    'End If
    
End Sub

Private Sub Form_Load()
    Call ComboListing
    'Call EnableEntries(False)
    bDataStatus = True
    'counter = 0
   
    cmdShow.Enabled = True
    txtMBMNumber.Enabled = True
    'txtMBMNumber.SetFocus
    cmdSearchMBM.Enabled = True
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Dim RecordExit
'    If bExit = False Then
'        RecordExit = MsgBox("Are you sure?", vbYesNo + vbExclamation)
'        If RecordExit = vbYes Then
            Unload Me
'        Else
'            Cancel = True
'        End If
'    End If
End Sub


Private Sub ClearEntries()
    On Error Resume Next
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            'If ctl.Enabled = True Then
            If ctl.Name <> "txtMBMNumber" Then
               ctl.Enabled = True
               ctl.Text = ""
               'ctl.Enabled = False
            End If
            'End If
        ElseIf TypeOf ctl Is ComboBox Then
            'If ctl.Enabled = True Then
                ctl.Enabled = True
                ctl.Text = ""
                'ctl.Enabled = False
            'End If
        ElseIf TypeOf ctl Is CommandButton Then
            ctl.Enabled = True
            
        ElseIf TypeOf ctl Is MaskEdBox Then
            'If ctl.Enabled = True Then
                ctl.Enabled = True
                ctl.mask = ""
                ctl.Text = ""
                ctl.mask = "##/##/####"
                'ctl.Enabled = False
            'End If
        End If
    Next
    'Check1.value = 0
End Sub


Private Sub ComboListing()
    cboMBMTakeOver.AddItem "N - None"
    cboMBMTakeOver.AddItem "Y - Take Over"
    cboMBMTakeOver.AddItem "C - Conversion"
    
    CboInsMode.AddItem "A - Annually"
    CboInsMode.AddItem "Q - Quarterly"
    CboInsMode.AddItem "S - Semi Annual"
    CboInsMode.AddItem "M - Monthly"
    
End Sub

Private Sub BNFListing()
    Dim rst2 As New ADODB.Recordset
    Dim BNF As ListItem
    Dim sql As String
    If lstBenefitDetails.listitems.count = 0 And Len(Trim(cboHLTCode)) <> 0 _
     And Len(Trim(cboINSCode)) <> 0 And Len(Trim(cboPLNCode)) <> 0 Then
        sql = " select BNFCode, BNFNoOFDays , " & _
          " BNFclaimableAmount, BNFdescription, BNFRMPerDIs " & _
          " from View_BNF where PLNCOde ='" & Trim(cboPLNCode) & "'" & _
          " AND HLTCode ='" & Trim(cboHLTCode) & "'" & _
          " AND INSCODe ='" & Trim(cboINSCode) & "'"
        rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        Do Until rst2.EOF
            Set BNF = lstBenefitDetails.listitems.Add(, , rst2!BNFCode)
                If Trim(rst2!BNFdescription) <> "" Then
                    BNF.SubItems(1) = rst2!BNFdescription
                End If
                
                BNF.SubItems(2) = IIf(IsNull(rst2!BNFNoOfDays), 0, rst2!BNFNoOfDays)
                If Trim(rst2!BNFClaimAbleAmount) <> "" Then
                    BNF.SubItems(3) = Format(Round(rst2!BNFClaimAbleAmount, 2), "###,###,###.00")
                End If
                If Trim(rst2!BNFRMperDis) <> "" Then
                    BNF.SubItems(4) = rst2!BNFRMperDis
                End If
                
                rst2.MoveNext
        Loop
       
        rst2.Close
        Set rst2 = Nothing
    End If
End Sub

Private Sub NewBNFListing()
    Dim rst2 As New ADODB.Recordset
    Dim BNF As ListItem
    Dim sql As String
    If lstNewBenefitDetails.listitems.count = 0 And Len(Trim(cboHLTCode)) <> 0 _
     And Len(Trim(cboINSCode)) <> 0 And Len(Trim(txtPLNCode)) <> 0 Then
        sql = " select BNFCode, BNFNoOFDays , " & _
          " BNFclaimableAmount, BNFdescription, BNFRMPerDIs " & _
          " from View_BNF where PLNCOde ='" & Trim(txtPLNCode) & "'" & _
          " AND HLTCode ='" & Trim(cboHLTCode) & "'" & _
          " AND INSCODe ='" & Trim(cboINSCode) & "'"
        rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        Do Until rst2.EOF
            Set BNF = lstNewBenefitDetails.listitems.Add(, , rst2!BNFCode)
                If Trim(rst2!BNFdescription) <> "" Then
                    BNF.SubItems(1) = rst2!BNFdescription
                End If
                
                BNF.SubItems(2) = IIf(IsNull(rst2!BNFNoOfDays), 0, rst2!BNFNoOfDays)
                If Trim(rst2!BNFClaimAbleAmount) <> "" Then
                    BNF.SubItems(3) = Format(Round(rst2!BNFClaimAbleAmount, 2), "###,###,###.00")
                End If
                If Trim(rst2!BNFRMperDis) <> "" Then
                    BNF.SubItems(4) = rst2!BNFRMperDis
                End If
                
                rst2.MoveNext
        Loop
       
        rst2.Close
        Set rst2 = Nothing
    End If
End Sub

Private Sub MBMHISListing()
Dim MBMHIS As New ADODB.Recordset
Dim SUPPHIS As New ADODB.Recordset
Dim HISList As ListItem
Dim sql As String
Dim TopupInd As String
Dim TopupNumber  As String
Dim PolicyNo As String
Dim PayorEffDate As String
Dim PayorEXPDate As String
Dim ENDtEffDate As String
Dim DataStatus As String
Dim status As String
Dim BordxDate  As String
Dim ENDtBordxDate As String
Dim ValueDate As String
Dim INSHisCode As String
INSHisCode = ""
    If Len(Trim(txtMBMIcBcPP)) <> 0 Then
        If lstMBMHis.listitems.count = 0 Then
            sql = " select MBMNumber, MBMPolicyNo, " & _
                " MBMPayorEffDate, MBMPayorExpDate,MBMENDtBordxDate, MBMTopupInd,MBMTopupNumber," & _
                " INSCode, PLNCode, MBMStatus, MBMENDTEffDate, MBMDataStatus, MBMBordxDate, MBMValueDate from View_MBM_MBMOthers " & _
                " where MBMIcBcPp ='" & Trim(txtMBMIcBcPP) & "'" '& _
                '" and MBMFirstJoinedDate ='" & Format(txtMBMDateJoin, "mm/dd/yyyy") & "'"
                '" where MBMFirstMemNo = '" & RemStr(txtFirstMeMNo) & "'"
                'View_MBM_MBMOthers
            MBMHIS.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
            Do Until MBMHIS.EOF
                Set HISList = lstMBMHis.listitems.Add(, , MBMHIS!MBMNumber)
                    'With MBMHIS
                        If Len(MBMHIS!MBMTopupInd) Then
                            HISList.SubItems(1) = MBMHIS!MBMTopupInd
                            TopupInd = MBMHIS!MBMTopupInd
                        Else
                            TopupInd = ""
                        End If
                        
                        
                        If Len(MBMHIS!MBMTopupNumber) Then
                            HISList.SubItems(2) = MBMHIS!MBMTopupNumber
                            HISList.SubItems(2) = TopupNumber
                        Else
                            TopupNumber = ""
                        End If

                        
                        If Len(MBMHIS!MBMPolicyNo) Then
                            HISList.SubItems(3) = MBMHIS!MBMPolicyNo
                            PolicyNo = MBMHIS!MBMPolicyNo
                        Else
                            PolicyNo = ""
                        End If
                        If Trim(MBMHIS!MBMPayorEffDate) <> "" Then
                            HISList.SubItems(4) = MBMHIS!MBMPayorEffDate
                            PayorEffDate = MBMHIS!MBMPayorEffDate
                        Else
                            PayorEffDate = ""
                        End If
                        If Trim(MBMHIS!MBMPayorEXPDate) <> "" Then
                            HISList.SubItems(5) = MBMHIS!MBMPayorEXPDate
                            PayorEXPDate = MBMHIS!MBMPayorEXPDate
                        Else
                            PayorEXPDate = ""
                        End If
                        If MBMHIS!MBMStatus = "C" Then
                            If Trim(MBMHIS!MBMENDtEffDate) <> "" Then
                                HISList.SubItems(6) = MBMHIS!MBMENDtEffDate
                                ENDtEffDate = MBMHIS!MBMENDtEffDate
                            Else
                                ENDtEffDate = ""
                            End If
                        Else
                            HISList.SubItems(6) = ""
                            ENDtEffDate = ""
                        End If
                        If Len(MBMHIS!MBMDataStatus) Then
                            HISList.SubItems(7) = MBMHIS!MBMDataStatus
                            DataStatus = MBMHIS!MBMDataStatus
                        Else
                            DataStatus = ""
                        End If
                        If Len(MBMHIS!MBMStatus) Then
                            HISList.SubItems(8) = MBMHIS!MBMStatus
                            status = MBMHIS!MBMStatus
                        Else
                            status = ""
                        End If
                        If Len(MBMHIS!MBMBordxDate) Then
                            HISList.SubItems(9) = MBMHIS!MBMBordxDate
                            BordxDate = MBMHIS!MBMBordxDate
                        Else
                            BordxDate = ""
                        End If
                        If Len(MBMHIS!MBMENDtBordxDate) Then
                            HISList.SubItems(10) = MBMHIS!MBMENDtBordxDate
                            ENDtBordxDate = MBMHIS!MBMENDtBordxDate
                        Else
                            ENDtBordxDate = ""
                        End If
                        If Len(MBMHIS!MBMValueDate) Then
                            HISList.SubItems(11) = MBMHIS!MBMValueDate
                            ValueDate = MBMHIS!MBMValueDate
                        Else
                            ValueDate = ""
                        End If
                        INSHisCode = MBMHIS!INSCode
                        MBMHIS.MoveNext
                    'End With
                        
            Loop
           
            'If Trim(INSHisCode) = "I" Or Trim(INSHisCode) = "G" Then
                'sql = " select MBMCNumber from MBMCoveredPersons " & _
                    " Where MBMCICBCPPNo ='" & Trim(txtMBMIcBcPP) & "'"
                    
                'SUPPHIS.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
                'If SUPPHIS.recordCount Then
                    'Do Until SUPPHIS.EOF
                        'Set HISList = lstMBMHis.listitems.Add(, , SUPPHIS!MBMCNumber)
                        
                       'if MBMHIS.recordCount Then
                        
                            
                                'If Len(TopupInd) Then
                                    'HISList.SubItems(1) = TopupInd
                                'End If
                                'If Len(TopupNumber) Then
                                    'HISList.SubItems(2) = TopupNumber
                                'End If
        
                                'HISList.SubItems(3) = PolicyNo
                                'If Trim(PayorEffDate) <> "" Then
                                    'HISList.SubItems(4) = PayorEffDate
                                'End If
                                'If Trim(PayorEXPDate) <> "" Then
                                    'HISList.SubItems(5) = PayorEXPDate
                                'End If
                                'If status = "C" Then
                                    'If Trim(ENDtEffDate) <> "" Then
                                        'HISList.SubItems(6) = ENDtEffDate
                                    'End If
                                'Else
                                    'HISList.SubItems(6) = ""
                                'End If
                                'If Len(DataStatus) Then
                                    'HISList.SubItems(7) = DataStatus
                                'End If
                                'If Len(status) Then
                                    'HISList.SubItems(8) = status
                                'End If
                                'If Len(BordxDate) Then
                                    'HISList.SubItems(9) = BordxDate
                                'End If
                                'If Len(ENDtBordxDate) Then
                                    'HISList.SubItems(10) = ENDtBordxDate
                                'End If
                                'If Len(ValueDate) Then
                                    'HISList.SubItems(11) = ValueDate
                                'End If
                        
                                'SUPPHIS.MoveNext
                    
                            
                        'End If
                    'Loop
                
                'End If
                'SUPPHIS.Close
                'Set SUPPHIS = Nothing
        'End If
                MBMHIS.Close
                Set MBMHIS = Nothing
            
        End If
    End If
End Sub


Private Sub ADJListing()
    Dim ADJ As New ADODB.Recordset
    Dim ADJList As ListItem
    Dim sql As String
    
    If Len(Trim(txtMBMNumber)) Then
         sql = " select MBMAmendID, MBMNumber, " & _
           " MBMAmendEdtType, MBMAmendDate, MBMAENDtEffDate, " & _
           " MBMAEndtBordxDate, MBMAEndtBatchNo " & _
           " From MBMAmendHistory where MBMnumber = '" & Trim(txtMBMNumber) & "'"
         ADJ.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
         
         Do Until ADJ.EOF
         
             Set ADJList = lstAdjustmentHis.listitems.Add(, , IIf(Len(Trim(ADJ!MBMAENDtBordxDate)) <> 0, ADJ!MBMAENDtBordxDate, "-"))
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

               
                 ADJ.MoveNext
         Loop
         ADJ.Close
         Set ADJ = Nothing
    End If
End Sub

Private Sub MBMLimitationListing()
Dim MBMLimit As New ADODB.Recordset
Dim MBMCovLimit As New ADODB.Recordset
Dim LimitList As ListItem
Dim sql As String
Dim sql2 As String
    
    lstLimitation.listitems.Clear
    
    If Len(Trim(txtMBMNumber)) Then
         sql = " select MBMNumber, MBMName, MBMGenWP, MBMPreExWP, " & _
           " MBMSpeIllWP, MBMStaffDec, MBMGenEndWP, " & _
           " MBMPreExEndWP, MBMSpeIllEndWP, MBMCOVID " & _
           " From VIEW_MBM_Limitation where MBMNumber = '" & Trim(txtMBMNumber) & "'"
         MBMLimit.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
         
         Do Until MBMLimit.EOF
         
             Set LimitList = lstLimitation.listitems.Add(, , Trim(MBMLimit!MBMName))
                 If Len(Trim(MBMLimit!MBMCOVID)) Then
                    LimitList.SubItems(1) = MBMLimit!MBMCOVID
                 End If
                 If Len(Trim(MBMLimit!MBMGenWP)) Then
                    LimitList.SubItems(2) = MBMLimit!MBMGenWP
                 End If

                If Len(MBMLimit!MBMGenEndWP) Then
                    LimitList.SubItems(3) = MBMLimit!MBMGenEndWP
                End If
                
                If Len(MBMLimit!MBMPreExWP) Then
                    LimitList.SubItems(4) = MBMLimit!MBMPreExWP
                End If
                
                If Len(MBMLimit!MBMPreExEndWP) Then
                    LimitList.SubItems(5) = MBMLimit!MBMPreExEndWP
                End If
                
                If Len(MBMLimit!MBMSpeIllWP) Then
                    LimitList.SubItems(6) = MBMLimit!MBMSpeIllWP
                End If
                
                If Len(MBMLimit!MBMSpeIllEndWP) Then
                    LimitList.SubItems(7) = MBMLimit!MBMSpeIllEndWP
                End If
                
                If Len(MBMLimit!MBMStaffDec) Then
                    LimitList.SubItems(8) = MBMLimit!MBMStaffDec
                End If

               
                 MBMLimit.MoveNext
         Loop
         MBMLimit.Close
         Set MBMLimit = Nothing
         
         sql2 = " select MBMCNumber, MBMCName, MBMCGenWP, MBMCPreExWP, " & _
                " MBMCSpeIllWP, MBMCStaffDec, MBMCGenEndWP, " & _
                " MBMCPreExEndWP, MBMSpeIllEndWP, MBMCCoverID " & _
                " From MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
         MBMCovLimit.Open sql2, medixmasterconn, adOpenKeyset, adLockOptimistic
         
         Do Until MBMCovLimit.EOF
         
             Set LimitList = lstLimitation.listitems.Add(, , Trim(MBMCovLimit!MBMCName))
                 If Len(Trim(MBMCovLimit!MBMCCoverID)) Then
                    LimitList.SubItems(1) = MBMCovLimit!MBMCCoverID
                 End If
                 If Len(Trim(MBMCovLimit!MBMCGenWP)) Then
                    LimitList.SubItems(2) = MBMCovLimit!MBMCGenWP
                 End If

                If Len(MBMCovLimit!MBMCGenEndWP) Then
                    LimitList.SubItems(3) = MBMCovLimit!MBMCGenEndWP
                End If
                
                If Len(MBMCovLimit!MBMCPreExWP) Then
                    LimitList.SubItems(4) = MBMCovLimit!MBMCPreExWP
                End If
                
                If Len(MBMCovLimit!MBMCPreExEndWP) Then
                    LimitList.SubItems(5) = MBMCovLimit!MBMCPreExEndWP
                End If
                
                If Len(MBMCovLimit!MBMCSpeIllWP) Then
                    LimitList.SubItems(6) = MBMCovLimit!MBMCSpeIllWP
                End If
                
                If Len(MBMCovLimit!MBMSpeIllEndWP) Then
                    LimitList.SubItems(7) = MBMCovLimit!MBMSpeIllEndWP
                End If
                
                If Len(MBMCovLimit!MBMCStaffDec) Then
                    LimitList.SubItems(8) = MBMCovLimit!MBMCStaffDec
                End If
                MBMCovLimit.MoveNext
         Loop
         MBMCovLimit.Close
         Set MBMCovLimit = Nothing
         
    End If
End Sub

' ######## List out all the cases which have been discharged
Private Sub ListCases()
    Dim CASList As ListItem
    Dim strsql As String
    Dim CAS As New ADODB.Recordset

    strsql = " Select CASNumber, MBMCName, " & _
            " DIAPrelimDiag1, MBMPatientCovID, " & _
            " CASDischargeDate, CASDateAdmitted, MBMCRELCode, HOSName, PAYCode, " & _
            " CASStatus, MBMName, MBMNumber, DIAFinalDiag1, CASClaimability  "
    strsql = strsql & " From VIEW_MBMInquiry_CASe where 0=0"
    strsql = strsql & " AND MBMNumber = '" & RemStr(txtMBMNumber) & "' "
            
    CAS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If CAS.recordCount <> 0 Then
        Screen.MousePointer = vbHourglass
        lstCaseList.listitems.Clear
        Do Until CAS.EOF
            Set CASList = lstCaseList.listitems.Add(, , CAS!CASNumber)
                CASList.SubItems(1) = CAS!MBMNumber
                'If CAS!MBMCName = "" Then
                    'CASList.SubItems(2) = "Self"
                    'CASList.SubItems(3) = "Self"
                If CAS!mbmpatientcovid = "00" Then
                    If Len(CAS!MBMName) Then
                        CASList.SubItems(2) = Trim(CAS!MBMName)
                    End If
                    CASList.SubItems(3) = "P"
                Else
                    If Len(CAS!MBMCName) Then
                        CASList.SubItems(2) = Trim(CAS!MBMCName)
                    End If
                    If Len(CAS!MBMCRELCode) Then
                        CASList.SubItems(3) = CAS!MBMCRELCode
                    End If
                End If
                
                'CASList.SubItems(4) = CAS!CASStatus
                'CASList.SubItems(5) = CAS!MBMPatientCovID
                'CASList.SubItems(6) = CAS!MBMPolicyNo
                If Len(Trim(CAS!CASDateAdmitted)) Then
                    CASList.SubItems(4) = CAS!CASDateAdmitted
                End If
                If Len(Trim(CAS!CASDischargeDate)) Then
                    CASList.SubItems(5) = CAS!CASDischargeDate
                End If
                If Len(Trim(CAS!CASClaimability)) Then
                    CASList.SubItems(6) = CAS!CASClaimability
                End If
                If Trim(CAS!DIAFinalDiag1) <> "" Then
                    CASList.SubItems(7) = CAS!DIAFinalDiag1
                End If
                If Len(CAS!CASStatus) Then
                    CASList.SubItems(8) = CAS!CASStatus
                End If
                If Trim(CAS!HOSName) <> "" Then
                    CASList.SubItems(9) = CAS!HOSName
                End If
                If Trim(CAS!PAYCode) <> "" Then
                    CASList.SubItems(10) = CAS!PAYCode
                End If
                
                
                'CASList.SubItems(12) = CAS!MBMPayorEffDate
                'CASList.SubItems(13) = CAS!MBMPayorExpDate
                'If Len(Trim(CAS!CASHOSCode)) Then
                    'CASList.SubItems(14) = CAS!CASHOSCode
                'End If
                
'                CASList.SubItems(15) = CAS!HLTCode
 '               CASList.SubItems(16) = CAS!MBMAvailableLimit
  '              CASList.SubItems(17) = CAS!CASReserveAmount
                CAS.MoveNext
        Loop
        Screen.MousePointer = vbDefault
    End If
    CAS.Close
    Set CAS = Nothing
End Sub

Private Sub ListALLCases()
    Dim CASList As ListItem
    Dim strsql As String
    Dim CAS As New ADODB.Recordset

    strsql = " Select CASNumber, MBMCName, " & _
            " DIAPrelimDiag1, MBMPatientCovID, " & _
            " CASDischargeDate, CASDateAdmitted, MBMCRELCode, HOSName, PAYCode, " & _
            " CASStatus, MBMName, MBMNumber, DIAFinalDiag1, CASClaimability  "
    strsql = strsql & " From VIEW_MBMInquiry_CASe where 0=0"
    strsql = strsql & " AND MBMIcBcPp = '" & RemStr(txtMBMIcBcPP) & "' "
            
    CAS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If CAS.recordCount <> 0 Then
        Screen.MousePointer = vbHourglass
        lstCASHistory.listitems.Clear
        Do Until CAS.EOF
            Set CASList = lstCASHistory.listitems.Add(, , CAS!CASNumber)
                CASList.SubItems(1) = CAS!MBMNumber
                'If CAS!MBMCName = "" Then
                    'CASList.SubItems(2) = "Self"
                    'CASList.SubItems(3) = "Self"
                If CAS!mbmpatientcovid = "00" Then
                    If Len(CAS!MBMName) Then
                        CASList.SubItems(2) = Trim(CAS!MBMName)
                    End If
                    CASList.SubItems(3) = "P"
                Else
                    If Len(CAS!MBMCName) Then
                        CASList.SubItems(2) = Trim(CAS!MBMCName)
                    End If
                    If Len(CAS!MBMCRELCode) Then
                        CASList.SubItems(3) = CAS!MBMCRELCode
                    End If
                End If
                
                'CASList.SubItems(4) = CAS!CASStatus
                'CASList.SubItems(5) = CAS!MBMPatientCovID
                'CASList.SubItems(6) = CAS!MBMPolicyNo
                If Len(Trim(CAS!CASDateAdmitted)) Then
                    CASList.SubItems(4) = CAS!CASDateAdmitted
                End If
                If Len(Trim(CAS!CASDischargeDate)) Then
                    CASList.SubItems(5) = CAS!CASDischargeDate
                End If
                If Len(Trim(CAS!CASClaimability)) Then
                    CASList.SubItems(6) = CAS!CASClaimability
                End If
                If Trim(CAS!DIAFinalDiag1) <> "" Then
                    CASList.SubItems(7) = CAS!DIAFinalDiag1
                End If
                If Len(CAS!CASStatus) Then
                    CASList.SubItems(8) = CAS!CASStatus
                End If
                If Trim(CAS!HOSName) <> "" Then
                    CASList.SubItems(9) = CAS!HOSName
                End If
                If Trim(CAS!PAYCode) <> "" Then
                    CASList.SubItems(10) = CAS!PAYCode
                End If
                
                'CASList.SubItems(12) = CAS!MBMPayorEffDate
                'CASList.SubItems(13) = CAS!MBMPayorExpDate
                'If Len(Trim(CAS!CASHOSCode)) Then
                    'CASList.SubItems(14) = CAS!CASHOSCode
                'End If
                
'                CASList.SubItems(15) = CAS!HLTCode
 '               CASList.SubItems(16) = CAS!MBMAvailableLimit
  '              CASList.SubItems(17) = CAS!CASReserveAmount
                CAS.MoveNext
        Loop
        Screen.MousePointer = vbDefault
    End If
    CAS.Close
    Set CAS = Nothing
End Sub
' ####################### Calculation's Function ###############
Private Sub MemberCallAge()
'On Error Resume Next
    Dim SelectAge As String
    If Trim(txtMBMPayorEffDate) <> "__/__/____" _
        And Len(Trim(txtMBMBirthDate)) <> 0 _
        And Trim(txtMBMBirthDate) <> "__/__/____" _
        And Len(Trim(cboHLTCode)) <> 0 Then
        If IsDate(txtMBMBirthDate) Then
            SelectAge = GetAge(txtMBMPayorEffDate, txtMBMBirthDate)
            txtAGECode = GetAgeBand(Trim(Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1)), SelectAge)
            If Len(Trim(txtAGECode)) = 0 Then
                MsgBox "Please Enter A Valid Date OR Age Out of Acceptable Range", vbOKOnly + vbCritical, DialogBoxTitle
                txtMBMPayorEffDate.SetFocus
            Else
                If Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)) = "01" Then
                    txtMBMAdultChild = "C"
                Else
                    txtMBMAdultChild = "A"
                End If
            End If
        End If
    End If
End Sub

Public Sub MemberCallFees()
    txtMBMPremium = ""
    txtMCOFees = ""
    If InStr(1, cboPLNCode, "-") > 1 Then
        txtMBMPremium = Format(GetPremium(Trim(Mid(cboINSCode, 1, 1)), Trim(Mid(cboPAYCode, 1, InStr(1, cboPAYCode, "-") - 1)), Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)), Trim(Mid(cboPLNCode, 1, InStr(1, cboPLNCode, "-") - 1)), Trim(Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1)), CDate(txtMBMPayorEffDate)), "#,##0.00")
        txtMCOFees = Format(GetMCO(Trim(Mid(cboINSCode, 1, 1)), Trim(Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1)), txtMBMAdultChild, CDate(txtMBMPayorEffDate), Trim(cboPLNCode)), "#,##0.00")
    End If
End Sub

Public Sub ProviderCallFees()
'On Error Resume Next
    Dim GetProString As String
    
    If Len(Trim(cboINSCode)) <> 0 And Len(Trim(cboSRVCode)) <> 0 And Len(Trim(cboHLTCode)) <> 0 Then
        txtMBMProviderCharges = ""
       ' txtMBMCaseCharges = ""
        GetProString = GetProviderCharges(Trim(cboSRVCode), Trim(cboINSCode), Trim(cboHLTCode), txtMBMAdultChild)
        If Len(Trim(GetProString)) Then
            txtMBMProviderCharges = Format(GetProString, "###,##0.00")
        End If
     '   txtMBMCaseCharges = Format(Mid(GetProString, InStr(1, GetProString, "-") + 1), "#,##0.00")
    Else
        txtMBMProviderCharges = 0
  '      txtMBMCaseCharges = 0
    End If
            
End Sub

Public Sub CallLimit()
'On Error Resume Next
    txtAnnualLimit = ""
    If Len(Trim(cboPLNCode)) Then
        txtAnnualLimit = Format(GetAnnualLimit(cboINSCode, cboPLNCode, cboHLTCode, txtMBMPayorEffDate), "#,##0.00")
    End If
End Sub

' ######################### END CALCULATIOn###########
' ######################### START CHECKING ###########

' ######################### END CHECKING ###########
' ######################### START BUTTON CLICKING ###########

Private Sub cmdPrev_Click()
Dim MBM As New ADODB.Recordset
Dim sql As String
    
        
    'If Trim(txtMBMNumber) <> "" Then
        sql = " select MBMNumber from MBM where MBMnumber < '" & RemStr(txtMBMNumber) & "' " & _
            " Order by MBMnumber DESC "
        MBM.MaxRecords = 1
        MBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBM.recordCount Then
            txtMBMNumber.Text = MBM!MBMNumber
            cmdshow_Click
        Else
            MsgBox "First Record Reached! " & vbNewLine & "Not a valid Membership Number Format! ", vbCritical + vbOKOnly
        End If
        MBM.Close
        Set MBM = Nothing
    'End If
End Sub

Private Sub CmdNext_Click()
    Dim MBM As New ADODB.Recordset
    Dim sql As String
    
        
    sql = " select MBMNumber from MBM where MBMnumber > '" & RemStr(txtMBMNumber) & "' " & _
        " Order by MBMnumber "
    MBM.MaxRecords = 1
    MBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If MBM.recordCount And MBM.recordCount <> -1 Then
        txtMBMNumber.Text = MBM!MBMNumber
        cmdshow_Click
    Else
        MsgBox " Last Record Reached! ", vbCritical + vbOKOnly
    End If
    
    MBM.Close
    Set MBM = Nothing
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

Private Sub lstBenefitDetails_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstBenefitDetails.SortKey = ColumnHeader.Index - 1
    lstBenefitDetails.Sorted = True
    If lstBenefitDetails.SortOrder = lvwAscending Then
        lstBenefitDetails.SortOrder = lvwDescending
    Else
        lstBenefitDetails.SortOrder = lvwAscending
    End If
End Sub

Private Sub lstCASHistory_DblClick()
    If lstCaseList.listitems.count Then
        frmCaseInquiry.txtCaseNo = ChkStr(lstCASHistory.SelectedItem.Text)
        Call frmCaseInquiry.cmdshow_Click
        frmCaseInquiry.show
    End If
End Sub

Private Sub lstNewBenefitDetails_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstNewBenefitDetails.SortKey = ColumnHeader.Index - 1
    lstNewBenefitDetails.Sorted = True
    If lstNewBenefitDetails.SortOrder = lvwAscending Then
        lstNewBenefitDetails.SortOrder = lvwDescending
    Else
        lstNewBenefitDetails.SortOrder = lvwAscending
    End If
End Sub

Private Sub lstCaseList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstCovAdd.SortKey = ColumnHeader.Index - 1
    lstCovAdd.Sorted = True
    If lstCovAdd.SortOrder = lvwAscending Then
        lstCovAdd.SortOrder = lvwDescending
    Else
        lstCovAdd.SortOrder = lvwAscending
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

Private Sub lstCovAdd_KeyDown(KeyCode As Integer, Shift As Integer)
    If lstCovAdd.listitems.count Then
        txtCovName = lstCovAdd.SelectedItem.Text
        cboCovSalutation = lstCovAdd.SelectedItem.SubItems(1)
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate.mask = ""
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        'cboCovBlood = lstCovAdd.SelectedItem.SubItems(7)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(10)
        txtCovAnnualLimit = Format(lstCovAdd.SelectedItem.SubItems(12), "###,##0.00")
        'txtAvailLimit = lstCovAdd.SelectedItem.SubItems(13)
        'cboOccupation = lstCovAdd.SelectedItem.SubItems(14)
        'cboWorkNature = lstCovAdd.SelectedItem.SubItems(15)
        'cboSmoking = lstCovAdd.SelectedItem.SubItems(16)
        'cboAlcohol = lstCovAdd.SelectedItem.SubItems(17)
        'txtSuppPrem = Format(lstCovAdd.SelectedItem.SubItems(18), "#,##0.00")
        txtPrevPLNCode = lstCovAdd.SelectedItem.SubItems(19)
        txtRBAmt = lstCovAdd.SelectedItem.SubItems(20)
        txtCOVPayRemarks = lstCovAdd.SelectedItem.SubItems(21)
        txtPrevINSCom = lstCovAdd.SelectedItem.SubItems(22)
        TxtSuppClmNonPayFr.mask = ""
        TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(23)
        TxtSuppClmNonPayTo.mask = ""
        TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(24)
        txtSuppPayPremGross = lstCovAdd.SelectedItem.SubItems(25)
        

    End If
    'Call GetAccountSummary
End Sub

Private Sub lstCovAdd_KeyUp(KeyCode As Integer, Shift As Integer)
    If lstCovAdd.listitems.count Then
        txtCovName = lstCovAdd.SelectedItem.Text
        cboCovSalutation = lstCovAdd.SelectedItem.SubItems(1)
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate.mask = ""
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        'cboCovBlood = lstCovAdd.SelectedItem.SubItems(7)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(10)
        txtCovAnnualLimit = Format(lstCovAdd.SelectedItem.SubItems(12), "###,##0.00")
        'txtAvailLimit = lstCovAdd.SelectedItem.SubItems(13)
        'cboOccupation = lstCovAdd.SelectedItem.SubItems(14)
        'cboWorkNature = lstCovAdd.SelectedItem.SubItems(15)
        'cboSmoking = lstCovAdd.SelectedItem.SubItems(16)
        'cboAlcohol = lstCovAdd.SelectedItem.SubItems(17)
        'txtSuppPrem = Format(lstCovAdd.SelectedItem.SubItems(18), "#,##0.00")
        txtPrevPLNCode = lstCovAdd.SelectedItem.SubItems(19)
        txtRBAmt = lstCovAdd.SelectedItem.SubItems(20)
        txtCOVPayRemarks = lstCovAdd.SelectedItem.SubItems(21)
        txtPrevINSCom = lstCovAdd.SelectedItem.SubItems(22)
        TxtSuppClmNonPayFr.mask = ""
        TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(23)
        TxtSuppClmNonPayTo.mask = ""
        TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(24)
        txtSuppPayPremGross = lstCovAdd.SelectedItem.SubItems(25)

    End If
    'Call GetAccountSummary
End Sub

Private Sub lstMBMAcc_Click()
Dim strsql As String
Dim MBM As New ADODB.Recordset
    
    If lstMBMAcc.listitems.count <> 0 Then
    
        strsql = " Select MBMNumber, MBMMemberType, MBMCAnnualLimit, " & _
                 " MBMCAvailableLimit From View_MBM_MemberType where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBM.recordCount Then
        If MBM!MBMMemberType = "P" Then
            txtAnnualLmt = txtAnnualLimit
            txtAvailableAnnualLmt = txtAvailableAnnualLimit
        ElseIf MBM!MBMMemberType = "Q" And lstMBMAcc.SelectedItem.SubItems(1) <> 0 Then
            txtAnnualLmt = MBM!MBMCAnnualLimit
            txtAvailableAnnualLmt = MBM!MBMCAvailableLimit
        ElseIf MBM!MBMMemberType = "Q" And lstMBMAcc.SelectedItem.SubItems(1) = 0 Then
            txtAnnualLmt = txtAnnualLimit
            txtAvailableAnnualLmt = txtAvailableAnnualLimit
        End If
        
        txtApprovedAmt = Format(lstMBMAcc.SelectedItem.SubItems(2), "#,##0.00")
        txtPayable2Hos = Format(lstMBMAcc.SelectedItem.SubItems(3), "#,##0.00")
        If IsNumeric(txtPayable2Hos) = False Then
            txtPayable2Hos = 0
        End If
        If Trim(lstMBMAcc.SelectedItem.SubItems(4)) <> "" Then
            txtPaid2Hos = Format(lstMBMAcc.SelectedItem.SubItems(4), "#,##0.00")
        End If
        If IsNumeric(txtPaid2Hos) = False Then
            txtPaid2Hos = 0
        End If
        
        ' Out standing hos
        txtHosOS = Format(CDbl(txtPayable2Hos) - CDbl(txtPaid2Hos), "#,##0.00;(#,##0.00)")
        
        txtPayableMem = Format(lstMBMAcc.SelectedItem.SubItems(5), "#,##0.00;(#,##0.00)")
        If IsNumeric(txtPayableMem) = False Then
            txtPayableMem = 0
        End If
        If Trim(lstMBMAcc.SelectedItem.SubItems(6)) <> "" Then
            txtPaid2MBM = Format(lstMBMAcc.SelectedItem.SubItems(6), "#,##0.00")
        End If
        If IsNumeric(txtPaid2MBM) = False Then
            txtPaid2MBM = 0
        End If
        If Trim(lstMBMAcc.SelectedItem.SubItems(7)) <> "" Then
            txtPaidByMem = Format(lstMBMAcc.SelectedItem.SubItems(7), "#,##0.00")
        End If
        If IsNumeric(txtPaidByMem) = False Then
            txtPaidByMem = 0
        End If
        txtMemOS = Format(CDbl(txtPayableMem) - CDbl(txtPaid2MBM) + CDbl(txtPaidByMem), "#,##0.00;(#,##0.00)")
        txtReimbursement = Format(lstMBMAcc.SelectedItem.SubItems(8), "#,##0.00")
        txtExcess = Format(lstMBMAcc.SelectedItem.SubItems(9), "#,##0.00;(#,##0.00)")
        txtBordxAmt = Format(lstMBMAcc.SelectedItem.SubItems(10), "#,##0.00;(#,##0.00)")
        If lstMBMAcc.SelectedItem.SubItems(9) = "" Then
            lstMBMAcc.SelectedItem.SubItems(9) = "0"
            txtExcessPaid = Format(lstMBMAcc.SelectedItem.SubItems(9) - lstMBMAcc.SelectedItem.SubItems(5), "#,#.00;(#,#.00)")
        End If

    End If
    End If
End Sub

Private Sub lstMBMHis_DblClick()
    'Dim f As New frmMBMInquiry
    'f.show
    'f.txtMBMNumber = lstMBMHis.SelectedItem.Text
    'f.cmdshow_Click
    If lstMBMHis.listitems.count Then
        txtMBMNumber = lstMBMHis.SelectedItem.Text
        cmdshow_Click
        SSTab1.Tab = 0
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Screen.MousePointer = vbHourglass
    
    If SSTab1.Tab = 1 Then
        If Mid(cboINSCode, 1, 1) = "I" Or Mid(cboINSCode, 1, 1) = "G" Then
            MsgBox "This Plan does not have supplementary", vbOKOnly + vbCritical, DialogBoxTitle
            SSTab1.Tab = 0
        Else
            SSTab1.Tab = 1
            'If lstCovAdd.listitems.count = 0 Then
            '    Call showMBMCov
            'End If
        End If
        
    End If
    
    If SSTab1.Tab = 2 Then
        lstAdjustmentHis.listitems.Clear
        If lstAdjustmentHis.listitems.count = 0 Then
            Call ADJListing
        End If
    End If
    
    If SSTab1.Tab = 3 Then
        If lstBenefitDetails.listitems.count = 0 Then
            Call BNFListing
        End If
        If lstNewBenefitDetails.listitems.count = 0 Then
            If txtPLNCode <> "" Or txtNewPolEffDate <> "__/__/____" Then
                Call NewBNFListing
            End If
        End If
    End If
    
    If SSTab1.Tab = 4 Then
        If Trim(txtMBMNumber) <> "" And lstCaseList.listitems.count = 0 Then
            Call ListCases
        End If
    End If
    
    If SSTab1.Tab = 5 Then
        If Trim(txtMBMNumber) <> "" And lstCASHistory.listitems.count = 0 Then
            Call ListALLCases
        End If
    End If
    
    If SSTab1.Tab = 6 Then
        If lstMBMHis.listitems.count = 0 Then
            Call MBMHISListing
        End If
    End If
    
    If SSTab1.Tab = 7 Then
        'If lstMBMAcc.listitems.count = 0 Then
        lstMBMAcc.listitems.Clear
            Call GetAccountSummary
        'End If
    End If
    
    'If SSTab1.Tab = 9 Then
    '    Call SGCProcListing
    'End If
Screen.MousePointer = Default
End Sub

Private Sub lstCaseList_DblClick()
    If lstCaseList.listitems.count Then
        frmCaseInquiry.txtCaseNo = ChkStr(lstCaseList.SelectedItem.Text)
        Call frmCaseInquiry.cmdshow_Click
        frmCaseInquiry.show
    End If
End Sub


Public Sub cmdshow_Click()
Dim getRecord As Integer
Dim MBMNumber As String

    
    MBMNumber = txtMBMNumber
    Call ClearEntries
    txtMBMNumber = MBMNumber
    Screen.MousePointer = vbHourglass
        If txtMBMNumber = "" Then
            MsgBox "Please Enter Membership No", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMNumber.SetFocus
        Else
        ScanCode = Trim(txtMBMNumber)
        If Mid(ScanCode, 1, 2) = "LO" Then
            txtMBMNumber = ""
            Call SearchForm(ScanCode)
        Else

            Call FindMBM
        End If
    End If
    
    Screen.MousePointer = Default
    
End Sub

Private Function FindMBM() As Integer
Dim strCount As String
Dim sqlstr2 As String
Dim rst2 As New ADODB.Recordset
Dim getRecord As Integer
   sqlstr2 = " Select MBMNumber from MBM Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
   rst2.Open sqlstr2, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbCritical
    Else
        getRecord = showMBM
        lstAdjustmentHis.listitems.Clear
        lstBenefitDetails.listitems.Clear
        lstNewBenefitDetails.listitems.Clear
        lstCaseList.listitems.Clear
        lstMBMHis.listitems.Clear
        lstSurgicalProc.listitems.Clear
        lstCASHistory.listitems.Clear
        lstCovAdd.listitems.Clear
        Call showMBMCov
        Call CallLimit
        
        Call SGCProcListing
        Call MBMLimitationListing
    End If
    rst2.Close
    Set rst2 = Nothing
End Function

Private Function showMBM() As Integer
Dim sqlstr, sqlPolStatus As String
Dim sqlstqMBMRef As String
Dim MBM As New ADODB.Recordset
Dim MBMRef As New ADODB.Recordset
Dim PolStatus As New ADODB.Recordset
Dim rstCount As New ADODB.Recordset
Dim MBMNumber As String
Dim strCount As String


    txtMBMPremiumRed = ""
    txtMBMMCORed = ""
    txtMBMProviderRedCharges = ""
    txtMBMProviderEffCharges = ""
    txtMCOFeesEff = ""
    txtMBMPremiumEff = ""
    txtMBMPremium = 0
    txtMCOFees = 0
    txtMBMProviderCharges = 0
    txtAnnualLimit = 0
    
    sqlstr = " Select MBMNumber, HLTCode, MBMPolicyNo, MBMName, MBMIcBcPp, " & _
            " MBMMemberType, MBMStatus, INSCode, PLNCode, MBMRenewFlag, " & _
            " MBMAddress1, MBMAddress2, MBMAddress3, MBMPostCode, CTYCode, " & _
            " MBMSalutation, MBMSex, SRVCodeList, " & _
            " MBMWorkNature, MBMGroupCompany, MBMEmployeeNo, " & _
            " MBMMaritalStatus, MBMAccessFee, MBMPremium, MBMMCOfee, " & _
            " MBMFirstJoinedDate, BRNCode, MBMDOB, NATCode, RACCode," & _
            " STACode, MBMFirstMEMNo, MBMExpiredStatus, " & _
            " BRNCode, MBMTakeOver, MBMAgentCode, MBMPolicyYear, MBMDataStatus"
            
    sqlstr = sqlstr & ", PAYCode, MBMBatchNo, " & _
            " MBMBordxDate, AGECode, HSCLNInd, " & _
            " MBMPayorEffDate, MBMPayorExpDate, MBMAdultChild, MBMAvailableLimit, " & _
            " MBMLastEndorseStatus, MBMENDtBordxDate, " & _
            " MBMReinstate, MBMLastAdjustmentDate, MBMInstallment, " & _
            " MBMDepartment, MBMNewPlanCode, MBMNewPolEffDate,MBMExclusion, MBMRemarks "
        
    sqlstr = sqlstr & " From View_MBM_MBMOthers Where MBMNumber ='" & Trim(txtMBMNumber) & "'"
    

    ReportCriteria = sqlstr
           
    If bDataStatus = True Then
        MBM.Open ReportCriteria, medixmasterconn, adOpenKeyset, adLockOptimistic
    Else
        MBM.Open ReportCriteria, medixmasterconnA, adOpenKeyset, adLockOptimistic
    End If
    Call ClearEntries
    
        showMBM = 1
        With MBM
            txtMBMName = IIf(!MBMStatus = "S", Trim(!MBMName) & " (Suspended)", Trim(!MBMName))
            If Len(!MBMIcBcPp) Then
                txtMBMIcBcPP = Trim(!MBMIcBcPp)
            End If
            If Len(!MBMStatus) Then
                txtMBMStatus = Trim(!MBMStatus)
            End If
            
            sqlPolStatus = " select pstatuscode ,pstatusDesc from PolicyStatus where PStatusCode ='" & Trim(txtMBMStatus) & "'"
            PolStatus.Open sqlPolStatus, medixmasterconn, adOpenKeyset, adLockOptimistic
            If PolStatus.recordCount Then
                txtMBMStatus = PolStatus!pstatusDesc
            End If
            
            PolStatus.Close
            Set PolStatus = Nothing
            
            If Len(!MBMExpiredStatus) Then
                lblEXPStatus = !MBMExpiredStatus
            End If
            
            If Len(!MBMPolicyNo) Then
                txtMBMPolNo = Trim(!MBMPolicyNo)
            End If
            If Len(!INSCode) Then
                cboINSCode = Trim(!INSCode)
            End If
            If Len(!PLNCode) Then
                cboPLNCode = Trim(!PLNCode)
            End If
            
            If Len(!MBMAddress1) Then
                txtMBMAdd1 = Trim(!MBMAddress1)
            End If
            
            If Len(!MBMAddress2) Then
                txtMBMAdd2 = Trim(!MBMAddress2)
            End If
            
            If Len(!MBMAddress3) Then
                txtMBMAdd3 = Trim(!MBMAddress3)
            End If
            
            If Len(!MBMPostCode) Then
                txtMBMPostCode = Trim(!MBMPostCode)
            End If
            If Len(!CTYCode) Then
                cboCTYCode = Trim(!CTYCode)
            End If
            If Len(!MBMRenewFlag) Then
                cboMBMRenewal = !MBMRenewFlag
            End If
            
            If Trim(!MBMSalutation) <> "" Then
                cboMBMSalutation = Trim(!MBMSalutation)
            End If
            If Len(!MBMSex) Then
                cboMBMSex = Trim(!MBMSex)
            End If
            If Len(Trim(!SRVCodeList)) Then
                cboSRVCode = Trim(!SRVCodeList)
            End If
            If Trim(!MBMGroupCompany) <> "" Then
                txtMBMGroupCompany = Trim(!MBMGroupCompany)
            End If
            If Trim(!MBMEmployeeNo) <> "" Then
                txtMBMEmployeeNo = Trim(!MBMEmployeeNo)
            End If
            If Trim(!MBMExclusion) <> "" Then
                txtMBMExclusion = Trim(!MBMExclusion)
            End If
            If Trim(!MBMInstallment) <> "" Then
                CboInsMode = Trim(!MBMInstallment)
            End If
            If Trim(!MBMMaritalStatus) <> "" Then
                cboMBMMaritalStatus = Trim(!MBMMaritalStatus)
            End If
            If Trim(!MBMFirstJoinedDate) <> "" Then
                txtMBMDateJoin = Trim(!MBMFirstJoinedDate)
            End If
            If Trim(!BRNCode) <> "" Then
                cboBRNCode = Trim(!BRNCode)
            End If
            If Trim(!NATCode) <> "" Then
                cboNATCode = Trim(!NATCode)
            End If
            If Trim(!RACCode) <> "" Then
                cboRACCode = Trim(!RACCode)
            End If
            If Trim(!STACode) <> "" Then
                cboSTACode = Trim(!STACode)
            End If
            If Trim(!MBMDepartment) <> "" Then
                txtDeptName = Trim(!MBMDepartment)
            End If
            If Trim(!MBMDOB) <> "" Then
                txtMBMBirthDate.mask = ""
                txtMBMBirthDate.Text = Trim(!MBMDOB)
            End If
            If Trim(!MBMFirstMEMNo) <> "" Then
                txtFirstMeMNo = Trim(!MBMFirstMEMNo)
            End If
            If Trim(!BRNCode) <> "" Then
                cboBRNCode = Trim(!BRNCode)
            End If
            If Trim(!MBMTakeOver) = "C" Then
                cboMBMTakeOver.ListIndex = 2
            ElseIf Trim(!MBMTakeOver) = "Y" Then
                cboMBMTakeOver.ListIndex = 1
            Else
                cboMBMTakeOver.ListIndex = 0
            End If
            If Len(!MBMAgentCode) Then
                txtMBMAgentCode = Trim(!MBMAgentCode)
            End If
            If Trim(!MBMMemberType) = "P" Then
                cboMBMRelCode = "C - Combined Limit"
            ElseIf Trim(!MBMMemberType) = "Q" Then
                cboMBMRelCode = "S - Separated Limit"
            Else
                cboMBMRelCode = Trim(!MBMMemberType)
            End If
            If Trim(!MBMDataStatus) = "A" Then
                txtDataStatus = "Active"
            Else
                txtDataStatus = "Deleted"
            End If
            If Len(!PAYCode) Then
                cboPAYCode = Trim(!PAYCode)
            End If
            If Len(!AgeCode) Then
                txtAGECode = Trim(!AgeCode)
            End If
            If Len(!MBMBordxDate) Then
                txtMBMBordxDate.mask = ""
                txtMBMBordxDate = Trim(!MBMBordxDate)
            End If
            If Len(!MBMBatchNo) Then
                txtMBMBatchNo = Trim(!MBMBatchNo)
            End If
            If Len(!HLTCode) Then
                cboHLTCode = Trim(!HLTCode)
            End If

            If Trim(!MBMPayorEffDate) <> "" Then
                txtMBMPayorEffDate = Trim(!MBMPayorEffDate)
            End If
            If Trim(!MBMPayorEXPDate) <> "" Then
                txtMBMPayorExpDate = Trim(!MBMPayorEXPDate)
            End If
            If Len(!MBMPolicyYear) Then
                txtMBMPolicyYear = Trim(!MBMPolicyYear)
            End If
            If Len(!MBMAvailableLimit) Then
                txtAvailableAnnualLimit = Format(Trim(!MBMAvailableLimit), "#,##0.00")
            End If
            If Len(!MBMAdultChild) Then
                txtMBMAdultChild = Trim(!MBMAdultChild)
            End If
            If Len(!MBMRemarks) Then
                txtRemarks = Trim(!MBMRemarks)
            End If
            If Trim(!HSClnInd) = "H" Or Trim(!HSClnInd) = "M" Then
                txtMBMPremium = Format(IIf(IsNull(!MBMPremium) = True, Format(GetPremium(Trim(cboINSCode), Trim(cboPAYCode), Trim(txtAGECode), Trim(cboPLNCode), Trim(cboHLTCode), txtMBMPayorEffDate), "#,##0.00"), !MBMPremium), "####0.00")
                txtMCOFees = Format(IIf(IsNull(!MBMMCOFee) = True, Format(GetMCO(Trim(cboINSCode), Trim(cboHLTCode), txtMBMAdultChild, txtMBMPayorEffDate, cboPLNCode), "#,##0.00"), !MBMMCOFee), "####0.00")
                txtMBMProviderCharges = Format(IIf(IsNull(!MBMAccessFee) = True, Format(GetProviderCharges(Trim(Mid(cboSRVCode, 1, 2)), Trim(cboINSCode), Trim(cboHLTCode), Trim(txtMBMAdultChild)), "####0.00"), !MBMAccessFee), "#,##0.00")
            ElseIf Trim(!HSClnInd) = "C" Then
                txtMBMPremium = "0.00"
                txtMCOFees = "0.00"
                txtMBMProviderCharges = IIf(IsNumeric(!MBMAccessFee) = True, !MBMAccessFee, "0.00")
            Else
                txtMBMPremium = Format(IIf(IsNull(!MBMPremium) = True, Format(GetPremium(Trim(cboINSCode), Trim(cboPAYCode), Trim(txtAGECode), Trim(cboPLNCode), Trim(cboHLTCode), txtMBMPayorEffDate), "#,##0.00"), !MBMPremium), "####0.00")
                txtMCOFees = Format(IIf(IsNull(!MBMMCOFee) = True, Format(GetMCO(Trim(cboINSCode), Trim(cboHLTCode), txtMBMAdultChild, txtMBMPayorEffDate, cboPLNCode), "#,##0.00"), !MBMMCOFee), "####0.00")
                txtMBMProviderCharges = Format(IIf(IsNull(!MBMAccessFee) = True, Format(GetProviderCharges(Trim(Mid(cboSRVCode, 1, 2)), Trim(cboINSCode), Trim(cboHLTCode), Trim(txtMBMAdultChild)), "####0.00"), !MBMAccessFee), "#,##0.00")
            End If
            If Len(!MBMNewPlanCode) Then
                txtPLNCode = !MBMNewPlanCode
            End If
            
            If Len(!MBMNewPolEffDate) Then
                txtNewPolEffDate = Format(!MBMNewPolEffDate, "dd/mm/yyyy")
            End If
            
            If Trim(!HSClnInd) = "H" Then
                LblHnS = !HSClnInd
                CmdClinical.Enabled = False
            Else
                If Len(!HSClnInd) Then
                    LblHnS = !HSClnInd
                    CmdClinical.Enabled = True
                Else
                    LblHnS = ""
                    CmdClinical.Enabled = False
                End If
            End If
            
          End With
        
            If MBM!MBMStatus = "C" Then
            If Len(MBM!MBMNumber) Then
                MBMNumber = Trim(MBM!MBMNumber)
            End If
            sqlstqMBMRef = "Select * from MBMRefund where MBMNumber =  '" & Trim(MBMNumber) & "'"
            MBMRef.Open sqlstqMBMRef, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            If MBMRef.recordCount Then
                With MBMRef
                    txtMBMPremiumRed = Format(Round(!MBMPremReduction, 0), "#,##0.00")
                    txtMBMMCORed = Format(Round(!MBMMCORefund, 0), "#,##0.00")
                    txtMBMProviderRedCharges = Format(Round(!MBMProviderRef, 2), "#,##0.00")
                    txtMBMProviderEffCharges = Format(Round(!MBMProviderEff, 2), "#,##0.00")
                    txtMCOFeesEff = Format(Round(!MBMMCOEffective, 0), "#,##0.00")
                    txtMBMPremiumEff = Format(Round(!MBMPREMEffective, 0), "#,##0.00")
                End With
            
            End If
            
            MBMRef.Close
            Set MBMRef = Nothing
        End If
    MBM.Close
    Set MBM = Nothing
End Function

Private Sub showMBMCov()
    Dim MBMCov As New ADODB.Recordset
    Dim sqlstr As String
    sqlstr = "select MBMCNumber, MBMCCoverID, MBMCName, MBMCSalutation," & _
        " MBMCICBCPPNo , MBMCDOB, MBMCSex, MBMCAdultChild, MBMCAgeband, " & _
        " MBMCRELCode, MBMCBloodType, MBMCAllergic,  " & _
        " MBMCExclusion, MBMCStatus , MBMCAnnualLimit , MBMCAvailableLimit, " & _
        "  MBMCOccupation, MBMCCLMNonPayFr, MBMCCLMNonPayTo, " & _
        " MBMCWorkNature, MBMCSmoking, MBMCAlcohol, MBMCPremium, MBMPayorPrem, " & _
        " MBMCPAYRemarks, MBMCRBAmt, MBMCPrevInsCom, MBMCPrevPLNCode, " & _
        " MBMCPayorPremGross, MBMCPayorPremNet, MBMCPayorMCOGross, MBMCPayorMCONet, MBMCPolicyDisc, MBMCRenewalDisc "

        
    If bDataStatus = True Then
        sqlstr = sqlstr & " from MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
        MBMCov.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    Else
        sqlstr = sqlstr & " from MBMCoveredPersonsArchive where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
        MBMCov.Open sqlstr, medixmasterconnA, adOpenKeyset, adLockOptimistic
    End If

    lstCovAdd.listitems.Clear
    If MBMCov.recordCount Then
        Dim MemberConveredPerson As ListItem
        Do While Not MBMCov.EOF
            Set MemberConveredPerson = lstCovAdd.listitems.Add(, , MBMCov!MBMCName)
            With MBMCov
                If Len(Trim(!MBMCSalutation)) Then
                    MemberConveredPerson.SubItems(1) = ChkStr(!MBMCSalutation)
                End If
                If Len(Trim(!MBMCIcBcPpNo)) Then
                    MemberConveredPerson.SubItems(2) = !MBMCIcBcPpNo
                End If
                If Len(Trim(!MBMCDOB)) Then
                    MemberConveredPerson.SubItems(3) = !MBMCDOB
                End If
                
                If Len(Trim(!MBMCSex)) Then
                    MemberConveredPerson.SubItems(4) = ChkStr(!MBMCSex)
                End If
                
                If Len(Trim(!MBMCAdultChild)) Then
                    MemberConveredPerson.SubItems(5) = ChkStr(!MBMCAdultChild)
                End If
                
                If Len(Trim(!MBMCRELCode)) Then
                    MemberConveredPerson.SubItems(6) = ChkStr(!MBMCRELCode)
                End If
                
                If Len(Trim(!MBMCBloodType)) Then
                    MemberConveredPerson.SubItems(7) = ChkStr(!MBMCBloodType)
                End If
                
                If Len(Trim(!MBMCAllergic)) Then
                    MemberConveredPerson.SubItems(8) = Trim(!MBMCAllergic)
                End If
                
                If Len(Trim(!MBMCExclusion)) Then
                    MemberConveredPerson.SubItems(9) = Trim(!MBMCExclusion)
                End If
                MemberConveredPerson.SubItems(10) = !MBMCCoverID
                
                If Len(Trim(!MBMCAnnualLimit)) Then
                    MemberConveredPerson.SubItems(12) = !MBMCAnnualLimit
                End If
                
                If Len(Trim(!MBMCAvailableLimit)) Then
                    MemberConveredPerson.SubItems(13) = !MBMCAvailableLimit
                End If
                
                If Len(Trim(!MBMCOccupation)) Then
                    MemberConveredPerson.SubItems(14) = !MBMCOccupation
                End If
 
                If Len(Trim(!MBMCWorkNature)) Then
                    MemberConveredPerson.SubItems(15) = !MBMCWorkNature
                End If
                
                If Len(Trim(!MBMCSmoking)) Then
                    MemberConveredPerson.SubItems(16) = !MBMCSmoking
                End If
                
                If Len(Trim(!MBMCAlcohol)) Then
                    MemberConveredPerson.SubItems(17) = !MBMCAlcohol
                End If
                
                If Len(Trim(!MBMCPremium)) Then
                    MemberConveredPerson.SubItems(18) = !MBMCPremium
                End If
                If Len(Trim(!MBMCPrevPLNCode)) Then
                    MemberConveredPerson.SubItems(19) = !MBMCPrevPLNCode
                End If
                If Len(Trim(!MBMCRBAmt)) Then
                    MemberConveredPerson.SubItems(20) = !MBMCRBAmt
                End If
                If Len(Trim(!MBMCPAYRemarks)) Then
                    MemberConveredPerson.SubItems(21) = !MBMCPAYRemarks
                End If
                If Len(Trim(!MBMCPrevInsCom)) Then
                    MemberConveredPerson.SubItems(22) = !MBMCPrevInsCom
                End If
                If Len(Trim(!MBMCCLMNonPayFr)) Then
                    MemberConveredPerson.SubItems(23) = !MBMCCLMNonPayFr
                End If
                If Len(Trim(!MBMCCLMNonPayTo)) Then
                    MemberConveredPerson.SubItems(24) = !MBMCCLMNonPayTo
                End If
                If Len(!MBMCPayorPremGross) Then
                    MemberConveredPerson.SubItems(25) = !MBMCPayorPremGross
                End If
                If Len(!MBMCPayorPremNet) Then
                    MemberConveredPerson.SubItems(26) = !MBMCPayorPremNet
                End If
                
                If Len(!MBMCPayorMCOGross) Then
                    MemberConveredPerson.SubItems(27) = !MBMCPayorMCOGross
                End If
                
                If Len(!MBMCPayorMCONet) Then
                    MemberConveredPerson.SubItems(28) = !MBMCPayorMCONet
                End If
                
                If Len(!MBMCRenewalDisc) Then
                    MemberConveredPerson.SubItems(29) = !MBMCRenewalDisc
                End If
                
                If Len(!MBMCPolicyDisc) Then
                    MemberConveredPerson.SubItems(30) = !MBMCPolicyDisc
                End If
            
            End With
            MBMCov.MoveNext
        Loop
    End If
    MBMCov.Close
    Set MBMCov = Nothing
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdMBMPolicy_Click()
    frmSearchMBM.txtMBMPolicyNo1.value = True
    frmSearchMBM.TxtMBMNo1.value = False
    frmSearchMBM.Source = 2
    frmSearchMBM.show
End Sub

Private Sub cmdSearchMBM_Click()
    bDataStatus = True
    frmSearchMBM.Source = 3
    frmSearchMBM.show vbModal
End Sub


Private Sub lstCovAdd_Click()
    If lstCovAdd.listitems.count Then
        txtCovName = lstCovAdd.SelectedItem.Text
        cboCovSalutation = lstCovAdd.SelectedItem.SubItems(1)
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate.mask = ""
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(10)
        txtCovAnnualLimit = Format(lstCovAdd.SelectedItem.SubItems(12), "###,##0.00")
        txtPrevPLNCode = lstCovAdd.SelectedItem.SubItems(19)
        txtRBAmt = lstCovAdd.SelectedItem.SubItems(20)
        txtCOVPayRemarks = lstCovAdd.SelectedItem.SubItems(21)
        txtPrevINSCom = lstCovAdd.SelectedItem.SubItems(22)
        TxtSuppClmNonPayFr.mask = ""
        TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(23)
        TxtSuppClmNonPayTo.mask = ""
        TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(24)
        txtSuppPayPremGross = lstCovAdd.SelectedItem.SubItems(25)
        txtSuppPayPremNet = lstCovAdd.SelectedItem.SubItems(26)
        txtSuppPayMCOGross = lstCovAdd.SelectedItem.SubItems(27)
        txtSuppPayMCONet = lstCovAdd.SelectedItem.SubItems(28)
        txtRenewalDisc = lstCovAdd.SelectedItem.SubItems(29)
        txtPolicyDisc = lstCovAdd.SelectedItem.SubItems(30)

    End If
End Sub

'  ################# END Suppplementary ######################

' ################################ Enable Tab / button   ######
Private Sub EnableEntries(status As Boolean)
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.Enabled = status
        ElseIf TypeOf ctl Is CommandButton Then
            ctl.Enabled = status
        End If
    Next ctl
End Sub

Private Sub InsertPrem()
Dim MBMInsertPrem As Double
Dim MBMPrem As Double
Dim PAYExpDate As String

MBMPrem = txtMCOFees
PAYExpDate = txtMBMPayorExpDate

    MBMInsertPrem = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(txtMBMPremium, "#,##0.00"))
    txtMBMPremium = ""
    txtMBMPremium = Format(MBMInsertPrem, "#,##0.00")
End Sub

Private Sub GetAccountSummary()
    Dim Acc As New ADODB.Recordset
    Dim strsql As String
    Dim masterlist As ListItem
    lstMBMAcc.listitems.Clear
    strsql = " select totalApprove, totalPayable2Hos, " & _
            " totalPaid2Hos, totalPayable2Mbm, totalPaidByMbm, MBMPatientCOVID, MBMPaidAmt," & _
            " totalPaid2MBM , totalExcess , totalReimburse from view_MBMTotal_summary where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    Acc.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If Not Acc.EOF Then
        Do Until Acc.EOF
        With Acc
        Set masterlist = lstMBMAcc.listitems.Add(, , txtMBMNumber)
            masterlist.SubItems(1) = Format(!mbmpatientcovid)
            masterlist.SubItems(2) = Format(!totalApprove, "#,##0.00")
            masterlist.SubItems(3) = Format(!totalPayable2Hos, "#,##0.00")
            If Trim(!totalPaid2Hos) <> "" Then
                masterlist.SubItems(4) = Format(!totalPaid2Hos, "#,##0.00")
            End If
            If IsNumeric(masterlist.SubItems(4)) = False Then
                masterlist.SubItems(4) = 0
            End If
            
            ' Out standing hos
            'txtHosOS = Format(CDbl(txtPayable2Hos) - CDbl(txtPaid2Hos), "#,#.00;(#,#.00)")
            
            masterlist.SubItems(5) = Format(!totalPayable2Mbm, "#,##0.00;(#,##0.00)")
            'If IsNumeric(masterlist.SubItems(3)) = False Then
            '    masterlist.SubItems(5) = 0
            'End If
            If Trim(!MBMPaidAmt) <> "" Then
                masterlist.SubItems(6) = Format(!MBMPaidAmt, "#,##0.00")
            End If
            If IsNumeric(masterlist.SubItems(6)) = False Then
                masterlist.SubItems(6) = 0
            End If
            If Trim(!totalPaidByMbm) <> "" Then
                masterlist.SubItems(7) = Format(!totalPaidByMbm, "#,##0.00")
            End If
            If IsNumeric(masterlist.SubItems(7)) = False Then
                masterlist.SubItems(7) = 0
            End If
            'txtMemOS = Format(CDbl(txtPayableMem) - CDbl(txtPaid2MBM) + CDbl(txtPaidByMem), "#,#.00;(#,#.00)")
            masterlist.SubItems(8) = Format(!TotalReimburse, "#,##0.00")
            masterlist.SubItems(9) = Format(!TotalExcess, "#,##0.00;(#,##0.00)")
            masterlist.SubItems(10) = Format(!totalApprove, "#,##0.00;(#,##0.00)")
            'masterlist.SubItems(11) = Format(!CLMExcess, "#,##0.00;(#,##0.00)")

        End With
        Acc.MoveNext
        Loop
    End If
    Acc.Close
    Set Acc = Nothing

End Sub



Private Sub lstLimitation_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstLimitation, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstLimitation, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstLimitation, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
    Case 4
        SortListView lstLimitation, ColumnHeader.Index, ldtDatetime, m_blnDirection(4)
    Case 5
        SortListView lstLimitation, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
    Case 6
        SortListView lstLimitation, ColumnHeader.Index, ldtDatetime, m_blnDirection(6)
    Case 7
        SortListView lstLimitation, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
    Case 8
        SortListView lstLimitation, ColumnHeader.Index, ldtDatetime, m_blnDirection(8)
    Case 9
        SortListView lstLimitation, ColumnHeader.Index, ldtString, m_blnDirection(9)
    End Select
End Sub

Private Sub SGCProcListing()
    Dim rst2 As New ADODB.Recordset
    Dim SGC As ListItem
    Dim sql As String
    
    If lstSurgicalProc.listitems.count = 0 And Len(Trim(cboHLTCode)) <> 0 _
     And Len(Trim(cboINSCode)) <> 0 And Len(Trim(cboPLNCode)) <> 0 And Len(Trim(cboPAYCode)) <> 0 Then
        sql = " select SGCProcCode, SGCProcCodeDesc , " & _
          " SGCProcCatCode, SGCProcCATDesc, SGCProcCodePerc " & _
          " from View_SGCProc where PLNCOde ='" & Trim(cboPLNCode) & "'" & _
          " AND HLTCode ='" & Trim(cboHLTCode) & "'" & _
          " AND INSType ='" & Trim(cboINSCode) & "'" & _
          " AND PAYCode = '" & Trim(cboPAYCode) & "'"
        rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
            Set SGC = lstSurgicalProc.listitems.Add(, , rst2!SGCProcCatCode)
                
                If Trim(rst2!SGCProcCATDesc) <> "" Then
                    SGC.SubItems(1) = rst2!SGCProcCATDesc
                End If
                
                If rst2!SGCProcCode <> "" Then
                    SGC.SubItems(2) = rst2!SGCProcCode
                End If
                
                If Trim(rst2!SGCProcCodeDesc) <> "" Then
                    SGC.SubItems(3) = rst2!SGCProcCodeDesc
                End If
                
                If Trim(rst2!SGCProcCodePerc) <> "" Then
                    SGC.SubItems(4) = rst2!SGCProcCodePerc
                End If
                
                rst2.MoveNext
        Loop
       
        rst2.Close
        Set rst2 = Nothing
    End If
End Sub

Private Sub txtMBMNumber_GotFocus()
cmdShow.Default = True
End Sub

Private Sub txtMBMNumber_LostFocus()
txtMBMNumber = ChkStr(txtMBMNumber)
cmdShow.Default = False
End Sub
