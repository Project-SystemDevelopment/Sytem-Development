VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMBMInquiry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member Inquiry"
   ClientHeight    =   7875
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11880
   LinkTopic       =   "mdiMain"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   5325
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9393
      _Version        =   393216
      Tabs            =   8
      Tab             =   7
      TabsPerRow      =   8
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
      TabPicture(0)   =   "frmMBMInquiry.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Other Details"
      TabPicture(1)   =   "frmMBMInquiry.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Supp&lementary"
      TabPicture(2)   =   "frmMBMInquiry.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstCovAdd"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&Adjustment History"
      TabPicture(3)   =   "frmMBMInquiry.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstAdjustmentHis"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "HIS &Benefits"
      TabPicture(4)   =   "frmMBMInquiry.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lstBenefitDetails"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&Case"
      TabPicture(5)   =   "frmMBMInquiry.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lstCaseList"
      Tab(5).Control(1)=   "Label7"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Mem &History"
      TabPicture(6)   =   "frmMBMInquiry.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lstMBMHis"
      Tab(6).Control(1)=   "Label26"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Account"
      TabPicture(7)   =   "frmMBMInquiry.frx":00C4
      Tab(7).ControlEnabled=   -1  'True
      Tab(7).Control(0)=   "lstMBMAcc"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Frame9"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Frame11"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "Frame12"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "Frame13"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "cmdTotal"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).ControlCount=   6
      Begin VB.CommandButton cmdTotal 
         Caption         =   "Total"
         Height          =   375
         Left            =   10560
         TabIndex        =   212
         Top             =   2280
         Width           =   855
      End
      Begin VB.Frame Frame13 
         Caption         =   "Worksheet Summary"
         Height          =   1215
         Left            =   120
         TabIndex        =   207
         Top             =   480
         Width           =   5055
         Begin VB.TextBox txtApprovedAmt 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   210
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtAvailableAnnualLmt 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   209
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtAnnualLmt 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   208
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label88 
            Caption         =   "Balance Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   183
            Top             =   495
            Width           =   945
         End
         Begin VB.Label Label87 
            Caption         =   "Approved Amt"
            Height          =   255
            Left            =   120
            TabIndex        =   184
            Top             =   735
            Width           =   1065
         End
         Begin VB.Label Label86 
            Caption         =   "Annual Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   206
            Top             =   255
            Width           =   1065
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Payable to Member"
         Height          =   1335
         Left            =   5640
         TabIndex        =   197
         Top             =   480
         Width           =   4815
         Begin VB.TextBox txtMemOS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   201
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtPaidByMem 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   200
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtPaid2MBM 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   199
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtPayableMem 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   198
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label txtMBMPayable 
            Caption         =   "Payable to/(by) Mem"
            Height          =   255
            Left            =   120
            TabIndex        =   205
            Top             =   240
            Width           =   1545
         End
         Begin VB.Label Label84 
            Caption         =   "Received From Mem"
            Height          =   255
            Left            =   120
            TabIndex        =   204
            Top             =   720
            Width           =   1785
         End
         Begin VB.Label Label82 
            Caption         =   "O/S"
            Height          =   255
            Left            =   120
            TabIndex        =   203
            Top             =   960
            Width           =   1425
         End
         Begin VB.Label Label81 
            Caption         =   "Paid To Mem"
            Height          =   255
            Left            =   120
            TabIndex        =   202
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Others"
         Height          =   1215
         Left            =   5640
         TabIndex        =   192
         Top             =   1800
         Width           =   4815
         Begin VB.TextBox txtExcess 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   194
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtReimbursement 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   193
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label80 
            Caption         =   "Total Reimburse"
            Height          =   255
            Left            =   120
            TabIndex        =   196
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label78 
            Caption         =   "Total Excess"
            Height          =   255
            Left            =   120
            TabIndex        =   195
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Payable To Hospital"
         Height          =   1215
         Left            =   120
         TabIndex        =   185
         Top             =   1800
         Width           =   5055
         Begin VB.TextBox txtHosOS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   188
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtPaid2Hos 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   187
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtPayable2Hos 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   186
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label68 
            Caption         =   "Paid Amount"
            Height          =   255
            Left            =   240
            TabIndex        =   191
            Top             =   495
            Width           =   945
         End
         Begin VB.Label Label67 
            Caption         =   "O/S "
            Height          =   255
            Left            =   240
            TabIndex        =   190
            Top             =   735
            Width           =   825
         End
         Begin VB.Label Label59 
            Caption         =   "Bill Amt"
            Height          =   255
            Left            =   240
            TabIndex        =   189
            Top             =   255
            Width           =   1065
         End
      End
      Begin MSComctlLib.ListView lstMBMHis 
         Height          =   4335
         Left            =   -74520
         TabIndex        =   160
         Top             =   720
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   7646
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Mem No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pol No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Eff Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Exp Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Insured Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Plan Code"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame1 
         Height          =   4530
         Left            =   -74880
         TabIndex        =   46
         Top             =   360
         Width           =   11445
         Begin VB.ComboBox CboInsMode 
            Height          =   315
            Left            =   9600
            TabIndex        =   181
            Top             =   2700
            Width           =   1545
         End
         Begin VB.TextBox cboPLNCode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   9960
            TabIndex        =   180
            Top             =   210
            Width           =   1200
         End
         Begin VB.TextBox cboINSCode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6840
            TabIndex        =   179
            Top             =   210
            Width           =   1875
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
            TabIndex        =   60
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
            TabIndex        =   125
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
         Begin VB.TextBox txtMBMPolicyYear 
            Enabled         =   0   'False
            Height          =   315
            Left            =   9960
            MaxLength       =   25
            TabIndex        =   49
            Top             =   1010
            Width           =   1200
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
            TabIndex        =   59
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
         Begin VB.Frame Frame2 
            Height          =   1275
            Left            =   120
            TabIndex        =   62
            Top             =   3120
            Width           =   11145
            Begin VB.Label Label11 
               Caption         =   "Access Fee"
               Height          =   255
               Left            =   5880
               TabIndex        =   84
               Top             =   180
               Width           =   1140
            End
            Begin VB.Label lblMCOFee 
               Caption         =   "MCO Fees"
               Height          =   255
               Left            =   2955
               TabIndex        =   83
               Top             =   195
               Width           =   870
            End
            Begin VB.Label Label85 
               Caption         =   "Effective MCO Fee "
               Height          =   255
               Left            =   2955
               TabIndex        =   82
               Top             =   915
               Width           =   1455
            End
            Begin VB.Label Label83 
               Caption         =   "Effective Prem "
               Height          =   255
               Left            =   120
               TabIndex        =   81
               Top             =   915
               Width           =   1140
            End
            Begin VB.Label txtMCOFees 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   ""
                  Size            =   379848.2696
                  Charset         =   0
                  Weight          =   5971
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4440
               TabIndex        =   80
               Top             =   180
               Width           =   945
            End
            Begin VB.Label txtMBMPremiumEff 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               BeginProperty Font 
                  Name            =   ""
                  Size            =   379848.2696
                  Charset         =   0
                  Weight          =   5971
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   79
               Top             =   915
               Width           =   1035
            End
            Begin VB.Label txtMCOFeesEff 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   ""
                  Size            =   379848.2696
                  Charset         =   0
                  Weight          =   5971
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4440
               TabIndex        =   78
               Top             =   915
               Width           =   945
            End
            Begin VB.Label txtMBMProviderCharges 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   ""
                  Size            =   379848.2696
                  Charset         =   0
                  Weight          =   5971
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7560
               TabIndex        =   77
               Top             =   180
               Width           =   945
            End
            Begin VB.Label lblPremium 
               Caption         =   "Premium"
               Height          =   255
               Left            =   120
               TabIndex        =   76
               Top             =   195
               Width           =   960
            End
            Begin VB.Label txtMBMPremium 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   ""
                  Size            =   379848.2696
                  Charset         =   0
                  Weight          =   5971
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   75
               Top             =   180
               Width           =   1020
            End
            Begin VB.Label Label58 
               Caption         =   "Prem Refund"
               Height          =   255
               Left            =   120
               TabIndex        =   74
               Top             =   560
               Width           =   1020
            End
            Begin VB.Label txtMBMPremiumRed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               BeginProperty Font 
                  Name            =   ""
                  Size            =   379848.2696
                  Charset         =   0
                  Weight          =   5971
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   73
               Top             =   560
               Width           =   1035
            End
            Begin VB.Label txtMBMMCORed 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   ""
                  Size            =   379848.2696
                  Charset         =   0
                  Weight          =   5971
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4440
               TabIndex        =   72
               Top             =   555
               Width           =   945
            End
            Begin VB.Label Label61 
               Caption         =   "MCO Refund"
               Height          =   255
               Left            =   2955
               TabIndex        =   71
               Top             =   555
               Width           =   975
            End
            Begin VB.Label Label79 
               Caption         =   "Acces Fee Refund"
               Height          =   255
               Left            =   5880
               TabIndex        =   70
               Top             =   555
               Width           =   1380
            End
            Begin VB.Label Label77 
               Caption         =   "Effective Access Fee"
               Height          =   255
               Left            =   5880
               TabIndex        =   69
               Top             =   915
               Width           =   1620
            End
            Begin VB.Label txtMBMProviderEffCharges 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   ""
                  Size            =   379848.2696
                  Charset         =   0
                  Weight          =   5971
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7560
               TabIndex        =   68
               Top             =   915
               Width           =   945
            End
            Begin VB.Label txtMBMProviderRedCharges 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   ""
                  Size            =   379848.2696
                  Charset         =   0
                  Weight          =   5971
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7560
               TabIndex        =   67
               Top             =   555
               Width           =   945
            End
            Begin VB.Label Label5 
               Caption         =   "Annual Limit"
               Height          =   255
               Left            =   8880
               TabIndex        =   66
               Top             =   180
               Width           =   870
            End
            Begin VB.Label Label4 
               Caption         =   "Bal Limit"
               Height          =   255
               Left            =   8880
               TabIndex        =   65
               Top             =   555
               Width           =   885
            End
            Begin VB.Label txtAnnualLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   2
               EndProperty
               BeginProperty Font 
                  Name            =   ""
                  Size            =   379848.2696
                  Charset         =   0
                  Weight          =   5971
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9885
               TabIndex        =   64
               Top             =   180
               Width           =   960
            End
            Begin VB.Label txtAvailableAnnualLimit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   2
               EndProperty
               BeginProperty Font 
                  Name            =   ""
                  Size            =   379848.2696
                  Charset         =   0
                  Weight          =   5971
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9885
               TabIndex        =   63
               Top             =   555
               Width           =   960
            End
            Begin VB.Line Line6 
               X1              =   120
               X2              =   11160
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Line Line1 
               X1              =   120
               X2              =   11160
               Y1              =   840
               Y2              =   840
            End
         End
         Begin VB.TextBox txtMBMAdd1 
            Height          =   285
            Left            =   1080
            MaxLength       =   40
            TabIndex        =   53
            Top             =   195
            Width           =   4065
         End
         Begin VB.TextBox txtMBMAdd2 
            Height          =   285
            Left            =   1080
            MaxLength       =   40
            TabIndex        =   54
            Top             =   450
            Width           =   4065
         End
         Begin VB.TextBox txtMBMAdd3 
            Height          =   285
            Left            =   1080
            MaxLength       =   40
            TabIndex        =   55
            Top             =   710
            Width           =   4065
         End
         Begin VB.TextBox txtMBMPostCode 
            Height          =   285
            Left            =   1080
            MaxLength       =   5
            TabIndex        =   56
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox cboCTYCode 
            Height          =   315
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   51
            Top             =   960
            Width           =   1545
         End
         Begin VB.ComboBox cboSTACode 
            Height          =   315
            Left            =   1080
            TabIndex        =   52
            Top             =   1220
            Width           =   4065
         End
         Begin VB.ComboBox cboMBMSex 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   50
            Top             =   1505
            Width           =   1335
         End
         Begin VB.ComboBox cboMBMTakeOver 
            Height          =   315
            Left            =   3600
            TabIndex        =   105
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
            TabIndex        =   61
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
            TabIndex        =   48
            Top             =   1770
            Width           =   1545
         End
         Begin VB.ComboBox cboRACCode 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   57
            Top             =   2030
            Width           =   1335
         End
         Begin VB.ComboBox cboNATCode 
            Height          =   315
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   58
            Top             =   2030
            Width           =   1545
         End
         Begin VB.TextBox cboBRNCode 
            Height          =   315
            Left            =   1080
            TabIndex        =   124
            Top             =   2300
            Width           =   4065
         End
         Begin VB.TextBox txtMBMAgentCode 
            Height          =   315
            Left            =   3600
            MaxLength       =   15
            TabIndex        =   122
            Top             =   2580
            Width           =   1545
         End
         Begin VB.ComboBox cboMBMRenewal 
            Height          =   315
            ItemData        =   "frmMBMInquiry.frx":00E0
            Left            =   1080
            List            =   "frmMBMInquiry.frx":00E2
            Sorted          =   -1  'True
            TabIndex        =   176
            Top             =   2580
            Width           =   1290
         End
         Begin VB.TextBox txtMBMDateJoin 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6840
            MaxLength       =   25
            TabIndex        =   129
            Top             =   1020
            Width           =   1875
         End
         Begin VB.TextBox txtMBMEmployeeNo 
            Height          =   315
            Left            =   6840
            MaxLength       =   15
            TabIndex        =   131
            Top             =   1290
            Width           =   4335
         End
         Begin VB.TextBox txtDeptName 
            Height          =   315
            Left            =   6840
            MaxLength       =   150
            TabIndex        =   170
            Top             =   1580
            Width           =   4335
         End
         Begin VB.TextBox txtMBMGroupCompany 
            Height          =   315
            Left            =   6840
            MaxLength       =   150
            TabIndex        =   132
            Top             =   1860
            Width           =   4335
         End
         Begin VB.TextBox txtAGECode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   6840
            TabIndex        =   178
            Top             =   2140
            Width           =   4335
         End
         Begin VB.ComboBox cboSRVCode 
            Height          =   315
            Left            =   6840
            Sorted          =   -1  'True
            TabIndex        =   135
            Top             =   2410
            Width           =   4335
         End
         Begin VB.ComboBox cboMBMRelCode 
            Height          =   315
            Left            =   6840
            TabIndex        =   47
            Top             =   2700
            Width           =   1815
         End
         Begin VB.Label Label55 
            Caption         =   "Ins mode"
            Height          =   255
            Left            =   8880
            TabIndex        =   182
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label txtMBMBatchNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   9960
            TabIndex        =   127
            Top             =   780
            Width           =   1200
         End
         Begin VB.Label Label45 
            Caption         =   "Renewal "
            Height          =   240
            Left            =   120
            TabIndex        =   177
            Top             =   2640
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Dept. Name"
            Height          =   255
            Left            =   5640
            TabIndex        =   171
            Top             =   1680
            Width           =   1545
         End
         Begin VB.Label Label22 
            Caption         =   "Service Provider"
            Height          =   255
            Left            =   5640
            TabIndex        =   136
            Top             =   2480
            Width           =   1485
         End
         Begin VB.Label Label18 
            Caption         =   "Group Company"
            Height          =   255
            Left            =   5640
            TabIndex        =   134
            Top             =   1935
            Width           =   1545
         End
         Begin VB.Label Label19 
            Caption         =   "Employee Num"
            Height          =   255
            Left            =   5640
            TabIndex        =   133
            Top             =   1400
            Width           =   1545
         End
         Begin VB.Label Label10 
            Caption         =   "Date Join"
            Height          =   255
            Left            =   5640
            TabIndex        =   130
            Top             =   1080
            Width           =   780
         End
         Begin VB.Label Label8 
            Caption         =   "Batch Num"
            Height          =   255
            Left            =   8880
            TabIndex        =   128
            Top             =   760
            Width           =   885
         End
         Begin VB.Label Label17 
            Caption         =   "Bordx. Date"
            Height          =   255
            Left            =   5640
            TabIndex        =   126
            Top             =   780
            Width           =   1275
         End
         Begin VB.Label Label14 
            Caption         =   "Agent Code"
            Height          =   255
            Left            =   2520
            TabIndex        =   123
            Top             =   2640
            Width           =   1140
         End
         Begin VB.Label Label20 
            Caption         =   "Ins Type"
            Height          =   255
            Left            =   5640
            TabIndex        =   108
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label21 
            Caption         =   "Plan Code"
            Height          =   255
            Left            =   8880
            TabIndex        =   107
            Top             =   210
            Width           =   825
         End
         Begin VB.Label Label29 
            Caption         =   "Age Band"
            Height          =   255
            Left            =   5640
            TabIndex        =   106
            Top             =   2170
            Width           =   735
         End
         Begin VB.Label Label72 
            Caption         =   "PostCode"
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   1015
            Width           =   1140
         End
         Begin VB.Label Label69 
            Caption         =   "Birthdate"
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label Label60 
            Caption         =   "State"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   1320
            Width           =   465
         End
         Begin VB.Label Label57 
            Caption         =   "City/Town"
            Height          =   255
            Left            =   2520
            TabIndex        =   96
            Top             =   1020
            Width           =   1020
         End
         Begin VB.Label Label54 
            Caption         =   "Address"
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label56 
            Caption         =   "Sex "
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   1560
            Width           =   825
         End
         Begin VB.Label Label9 
            Caption         =   "Pol Year"
            Height          =   255
            Left            =   8880
            TabIndex        =   93
            Top             =   1080
            Width           =   1050
         End
         Begin VB.Label Label16 
            Caption         =   "Exp Date"
            Height          =   255
            Left            =   8880
            TabIndex        =   92
            Top             =   450
            Width           =   795
         End
         Begin VB.Label Label15 
            Caption         =   "Eff Date"
            Height          =   255
            Left            =   5640
            TabIndex        =   91
            Top             =   510
            Width           =   1275
         End
         Begin VB.Label Label63 
            Caption         =   "TakeOver"
            Height          =   255
            Left            =   2520
            TabIndex        =   90
            Top             =   1620
            Width           =   825
         End
         Begin VB.Label Label12 
            Caption         =   "Branch"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   2400
            Width           =   960
         End
         Begin VB.Label Label70 
            Caption         =   "Race"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   2140
            Width           =   420
         End
         Begin VB.Label Label71 
            Caption         =   "Nationality"
            Height          =   275
            Left            =   2520
            TabIndex        =   87
            Top             =   2080
            Width           =   960
         End
         Begin VB.Label Label27 
            Caption         =   "Marital Status"
            Height          =   255
            Left            =   2520
            TabIndex        =   86
            Top             =   1860
            Width           =   975
         End
         Begin VB.Label Label30 
            Caption         =   "Supp Type"
            Height          =   255
            Left            =   5640
            TabIndex        =   85
            Top             =   2760
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4005
         Left            =   -74760
         TabIndex        =   30
         Top             =   360
         Width           =   11220
         Begin VB.ComboBox cboMBMOccupation 
            Height          =   315
            Left            =   1560
            TabIndex        =   147
            Top             =   195
            Width           =   3585
         End
         Begin VB.TextBox txtMBMExclusion 
            Height          =   1905
            Left            =   5640
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   32
            Top             =   480
            Width           =   5340
         End
         Begin VB.TextBox txtMBMAllergic 
            Height          =   1020
            Left            =   5640
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   31
            Top             =   2760
            Width           =   5340
         End
         Begin VB.ComboBox txtMBMWorkNature 
            Height          =   315
            Left            =   1560
            TabIndex        =   149
            Top             =   480
            Width           =   3585
         End
         Begin VB.TextBox txtMBMTelNoH 
            Height          =   315
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   146
            Top             =   760
            Width           =   3585
         End
         Begin VB.TextBox txtMBMTelNoM 
            Height          =   315
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   139
            Top             =   1040
            Width           =   3585
         End
         Begin VB.TextBox txtMBMTelNoO 
            Height          =   315
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   140
            Top             =   1320
            Width           =   3585
         End
         Begin VB.TextBox txtMBMEmail 
            Height          =   315
            Left            =   1560
            TabIndex        =   138
            Top             =   1610
            Width           =   3585
         End
         Begin VB.TextBox txtMBMWeight 
            Height          =   315
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   148
            Top             =   1880
            Width           =   3585
         End
         Begin VB.TextBox txtMBMHeight 
            Height          =   315
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   141
            Top             =   2160
            Width           =   3585
         End
         Begin VB.ComboBox cboLGNCode 
            Height          =   315
            Left            =   1560
            TabIndex        =   145
            Top             =   2450
            Width           =   3585
         End
         Begin VB.ComboBox cboMBMBlood 
            Height          =   315
            Left            =   1560
            TabIndex        =   144
            Top             =   2730
            Width           =   3585
         End
         Begin VB.ComboBox cboMBMSmoking 
            Height          =   315
            Left            =   1560
            TabIndex        =   143
            Top             =   3000
            Width           =   3585
         End
         Begin VB.ComboBox cboMBMAlcohol 
            Height          =   315
            Left            =   1560
            TabIndex        =   142
            Top             =   3290
            Width           =   3585
         End
         Begin VB.Label Email 
            Caption         =   "Email"
            Height          =   255
            Left            =   240
            TabIndex        =   121
            Top             =   1640
            Width           =   1095
         End
         Begin VB.Label Label28 
            Caption         =   "Tel Num. (H)"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   790
            Width           =   1005
         End
         Begin VB.Label Label32 
            Caption         =   "H/P Num"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   1070
            Width           =   1005
         End
         Begin VB.Label Label33 
            Caption         =   "Tel Num (O)"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   1350
            Width           =   1005
         End
         Begin VB.Label Label34 
            Caption         =   "Weight"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   1910
            Width           =   1005
         End
         Begin VB.Label Label35 
            Caption         =   "Height"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   2190
            Width           =   1005
         End
         Begin VB.Label Label36 
            Caption         =   "Pref. Language"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   2480
            Width           =   1140
         End
         Begin VB.Label Label40 
            Caption         =   "Warranty Exclusion"
            Height          =   255
            Left            =   5640
            TabIndex        =   39
            Top             =   240
            Width           =   1590
         End
         Begin VB.Label Label41 
            Caption         =   "Allergic"
            Height          =   255
            Left            =   5640
            TabIndex        =   38
            Top             =   2520
            Width           =   1590
         End
         Begin VB.Label Label23 
            Caption         =   "Work Nature"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   510
            Width           =   1140
         End
         Begin VB.Label Label39 
            Caption         =   "Alcohol (Y/N)"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   3320
            Width           =   1275
         End
         Begin VB.Label Label38 
            Caption         =   "Smoking (Y/N)"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   3060
            Width           =   1275
         End
         Begin VB.Label Label37 
            Caption         =   "Blood Type"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   2760
            Width           =   1005
         End
         Begin VB.Label Label13 
            Caption         =   "Occupation"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   225
            Width           =   1005
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   15
         Top             =   360
         Width           =   11100
         Begin VB.TextBox txtCovExclusion 
            Height          =   930
            Left            =   6000
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   22
            Top             =   2160
            Width           =   4935
         End
         Begin VB.TextBox txtCovAllergic 
            Height          =   930
            Left            =   240
            MaxLength       =   60
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   2160
            Width           =   5175
         End
         Begin VB.ComboBox cboCovSalutation 
            Height          =   315
            Left            =   1320
            TabIndex        =   20
            Top             =   220
            Width           =   1245
         End
         Begin VB.TextBox txtCovName 
            Height          =   330
            Left            =   3915
            MaxLength       =   50
            TabIndex        =   19
            Top             =   195
            Width           =   7020
         End
         Begin VB.ComboBox cboCovBlood 
            Height          =   315
            Left            =   7320
            TabIndex        =   150
            Top             =   500
            Width           =   3615
         End
         Begin VB.ComboBox cboOccupation 
            Height          =   315
            Left            =   7320
            TabIndex        =   165
            Top             =   780
            Width           =   3615
         End
         Begin VB.ComboBox cboWorkNature 
            Height          =   315
            Left            =   7320
            TabIndex        =   164
            Top             =   1070
            Width           =   3615
         End
         Begin VB.ComboBox cboSmoking 
            Height          =   315
            Left            =   7320
            TabIndex        =   163
            Top             =   1350
            Width           =   3615
         End
         Begin VB.ComboBox cboAlcohol 
            Height          =   315
            Left            =   7320
            TabIndex        =   162
            Top             =   1640
            Width           =   3615
         End
         Begin VB.TextBox txtCovIcBcPp 
            Height          =   285
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   18
            Top             =   490
            Width           =   4125
         End
         Begin MSMask.MaskEdBox txtCovBirthdate 
            Height          =   285
            Left            =   1320
            TabIndex        =   23
            Top             =   750
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboCovRelationship 
            Height          =   315
            Left            =   1320
            TabIndex        =   137
            Top             =   1000
            Width           =   4125
         End
         Begin VB.ComboBox cboCovAdultChild 
            Height          =   315
            Left            =   4110
            TabIndex        =   16
            Top             =   1280
            Width           =   1335
         End
         Begin VB.ComboBox cboCovSex 
            Height          =   315
            Left            =   1320
            TabIndex        =   17
            Top             =   1280
            Width           =   1245
         End
         Begin VB.TextBox txtCovAnnualLimit 
            Height          =   315
            Left            =   1320
            TabIndex        =   151
            Top             =   1560
            Width           =   1245
         End
         Begin VB.TextBox txtAvailLimit 
            Height          =   285
            Left            =   4110
            TabIndex        =   157
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label44 
            Caption         =   "Occupation"
            Height          =   255
            Left            =   6000
            TabIndex        =   169
            Top             =   810
            Width           =   1095
         End
         Begin VB.Label Label46 
            Caption         =   "Work Nature"
            Height          =   255
            Left            =   6000
            TabIndex        =   168
            Top             =   1100
            Width           =   1095
         End
         Begin VB.Label Label43 
            Caption         =   "Smoking"
            Height          =   255
            Left            =   6000
            TabIndex        =   167
            Top             =   1380
            Width           =   1095
         End
         Begin VB.Label Label62 
            Caption         =   "Alcohol"
            Height          =   240
            Left            =   6000
            TabIndex        =   166
            Top             =   1677
            Width           =   855
         End
         Begin VB.Label Label42 
            Caption         =   "Mem Name"
            Height          =   255
            Left            =   2640
            TabIndex        =   161
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Balance Limit"
            Height          =   255
            Left            =   2880
            TabIndex        =   156
            Top             =   1595
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Supp Limit"
            Height          =   255
            Left            =   240
            TabIndex        =   155
            Top             =   1610
            Width           =   855
         End
         Begin VB.Label Label53 
            Caption         =   "Blood"
            Height          =   255
            Left            =   6000
            TabIndex        =   154
            Top             =   530
            Width           =   1005
         End
         Begin VB.Label Label52 
            Caption         =   "R'ship"
            Height          =   255
            Left            =   240
            TabIndex        =   153
            Top             =   1050
            Width           =   1005
         End
         Begin VB.Label Label47 
            Caption         =   "Warranty Exclusion"
            Height          =   255
            Left            =   6000
            TabIndex        =   152
            Top             =   1920
            Width           =   1500
         End
         Begin VB.Label Label48 
            Caption         =   "Allergic"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1920
            Width           =   600
         End
         Begin VB.Label Label66 
            Caption         =   "Salutation"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label65 
            Caption         =   "IC Num"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   520
            Width           =   1125
         End
         Begin VB.Label Label49 
            Caption         =   "Birthdate"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   785
            Width           =   1005
         End
         Begin VB.Label Label50 
            Caption         =   "Sex (M/F)"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1340
            Width           =   825
         End
         Begin VB.Label Label51 
            Caption         =   "Adult/Child"
            Height          =   255
            Left            =   2880
            TabIndex        =   24
            Top             =   1340
            Width           =   825
         End
      End
      Begin MSComctlLib.ListView lstCovAdd 
         Height          =   1545
         Left            =   -74760
         TabIndex        =   100
         Top             =   3720
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   2725
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
         NumItems        =   18
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
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "AnnualLimit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Available Limit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Occupation"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Work Nature"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Smaking"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "Alcohol"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstCaseList 
         Height          =   4155
         Left            =   -74760
         TabIndex        =   8
         Top             =   600
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   7329
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Case Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Mem No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Patient"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Relationship"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DOA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "DOD"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Pre Diagnosis"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Case Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstAdjustmentHis 
         Height          =   4335
         Left            =   -74640
         TabIndex        =   6
         Top             =   480
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   7646
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
         Height          =   4395
         Left            =   -74640
         TabIndex        =   7
         Top             =   480
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   7752
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
            Text            =   "Benefits Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Benefit Description"
            Object.Width           =   9701
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "No Of Days"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "RM per Day"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Per Disability"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstMBMAcc 
         Height          =   2055
         Left            =   120
         TabIndex        =   211
         Top             =   3120
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   3625
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
         NumItems        =   10
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
      End
      Begin VB.Label Label26 
         Caption         =   "Double Click to Open Old Membership data (Members With Same IC are Listed Below)"
         BeginProperty Font 
            Name            =   ""
            Size            =   379848.2696
            Charset         =   0
            Weight          =   5971
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   175
         Top             =   480
         Width           =   8055
      End
      Begin VB.Label Label7 
         Caption         =   "Double click to open the case"
         BeginProperty Font 
            Name            =   ""
            Size            =   379848.2696
            Charset         =   0
            Weight          =   5971
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   158
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   12015
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   15
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Member Inquiry"
         BeginProperty Font 
            Name            =   ""
            Size            =   362234.8296
            Charset         =   0
            Weight          =   -9868
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   120
         Width           =   6615
      End
   End
   Begin VB.Frame Frame7 
      Height          =   525
      Left            =   120
      TabIndex        =   101
      Top             =   7320
      Width           =   11655
      Begin VB.TextBox txtFirstMeMNo 
         Height          =   285
         Left            =   4080
         TabIndex        =   172
         Text            =   "FirstMeMNo"
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   10200
         TabIndex        =   9
         Top             =   120
         Width           =   960
      End
      Begin VB.TextBox txtMBMAdultChild 
         Height          =   285
         Left            =   840
         TabIndex        =   104
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtLastEndStatus 
         Height          =   285
         Left            =   1680
         TabIndex        =   103
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtLastEdtDate 
         Height          =   285
         Left            =   2760
         TabIndex        =   102
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   11655
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
         Height          =   285
         Left            =   5280
         TabIndex        =   2
         Top             =   225
         Width           =   810
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&Pre"
         Height          =   285
         Left            =   4440
         TabIndex        =   1
         Top             =   225
         Width           =   810
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
         Left            =   6120
         TabIndex        =   4
         Top             =   225
         Width           =   810
      End
      Begin VB.TextBox txtMBMNumber 
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   0
         Top             =   225
         Width           =   2175
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Go"
         Default         =   -1  'True
         Height          =   285
         Left            =   3600
         TabIndex        =   3
         Top             =   225
         Width           =   810
      End
      Begin VB.Label txtMBMStatus 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   ""
            Size            =   379848.2696
            Charset         =   0
            Weight          =   5971
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   9120
         TabIndex        =   174
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label25 
         Caption         =   "Status"
         Height          =   330
         Left            =   7800
         TabIndex        =   173
         Top             =   255
         Width           =   915
      End
      Begin VB.Label Label76 
         Caption         =   "Mem Num."
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   210
         Width           =   1275
      End
   End
   Begin VB.Frame Frame10 
      Height          =   975
      Left            =   120
      TabIndex        =   109
      Top             =   1080
      Width           =   11655
      Begin VB.TextBox txtMBMIcBcPP 
         Height          =   285
         Left            =   9120
         MaxLength       =   15
         TabIndex        =   114
         Top             =   187
         Width           =   2175
      End
      Begin VB.ComboBox cboMBMSalutation 
         Height          =   315
         Left            =   1200
         TabIndex        =   113
         Top             =   195
         Width           =   915
      End
      Begin VB.TextBox txtMBMName 
         Height          =   330
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   111
         Top             =   187
         Width           =   3675
      End
      Begin VB.ComboBox cboPAYCode 
         Height          =   315
         Left            =   1200
         TabIndex        =   110
         Top             =   500
         Width           =   915
      End
      Begin VB.TextBox txtMBMPolNo 
         Height          =   330
         Left            =   3240
         MaxLength       =   25
         TabIndex        =   112
         Top             =   500
         Width           =   3675
      End
      Begin VB.Label cboHLTCode 
         BackColor       =   &H80000009&
         Height          =   285
         Left            =   9120
         TabIndex        =   118
         Top             =   500
         Width           =   2175
      End
      Begin VB.Label Label24 
         Caption         =   "Salutation"
         Height          =   255
         Left            =   240
         TabIndex        =   159
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label75 
         Caption         =   "Health Type"
         Height          =   255
         Left            =   7800
         TabIndex        =   120
         Top             =   520
         Width           =   915
      End
      Begin VB.Label Label64 
         Caption         =   "IC Num"
         Height          =   330
         Left            =   7800
         TabIndex        =   119
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label31 
         Caption         =   "Mem Name"
         Height          =   255
         Left            =   2280
         TabIndex        =   117
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label73 
         Caption         =   "Policy Num"
         Height          =   255
         Left            =   2280
         TabIndex        =   116
         Top             =   525
         Width           =   900
      End
      Begin VB.Label Label74 
         Caption         =   "Payor"
         Height          =   255
         Left            =   240
         TabIndex        =   115
         Top             =   520
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmMBMInquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bExit As Boolean
Public bDataStatus As Boolean

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
Dim totalExcess As Variant
Dim TotalReimburse As Variant
Dim MBMTotalAnnualLmt As Variant
Dim MBMTotalAvailLmt As Variant

    If lstMBMAcc.listitems.count <> 0 Then
        strsqlMBM = "Select * from View_MBM_MemberType where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        
        
    If MBM.recordCount Then
        If MBM!MBMMemberType = "P" Then
            txtAnnualLmt = txtAnnualLimit
            txtAvailableAnnualLmt = txtAvailableAnnualLimit
        Else
            Do Until MBM.EOF
                MBMTotalAnnualLmt = MBMTotalAnnualLmt + MBM!MBMCAnnualLimit
                MBMTotalAvailLmt = MBMTotalAvailLmt + MBM!MBMCAvailableLimit
            MBM.MoveNext
            Loop
                txtAnnualLmt = MBMTotalAnnualLmt + txtAnnualLimit
                txtAvailableAnnualLmt = MBMTotalAvailLmt + txtAvailableAnnualLimit
        End If
        MBM.Close
        Set MBM = Nothing
    Else
    
        strsqlMBMArc = "Select * from View_MBM_MemTypeArc where MBMNumber = '" & Trim(txtMBMNumber) & "'"
        MBMArc.Open strsqlMBMArc, medixmasterconn, adOpenKeyset, adLockOptimistic

        If MBMArc.recordCount Then
            If MBMArc!MBMMemberType = "P" Then
                txtAnnualLmt = txtAnnualLimit
                txtAvailableAnnualLmt = txtAvailableAnnualLimit
            Else
                Do Until MBM.EOF
                    MBMTotalAnnualLmt = MBMTotalAnnualLmt + MBMArc!MBMCAnnualLimit
                    MBMTotalAvailLmt = MBMTotalAvailLmt + MBMArc!MBMCAvailableLimit
                MBMArc.MoveNext
                Loop
                    txtAnnualLmt = MBMTotalAnnualLmt + txtAnnualLimit
                    txtAvailableAnnualLmt = MBMTotalAvailLmt + txtAvailableAnnualLimit
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
                totalApprove = totalApprove + !totalApprove
                totalPayable2Hos = totalPayable2Hos + !totalPayable2Hos
                If Trim(!totalPaid2Hos) <> "" Then
                    totalPaid2Hos = totalPaid2Hos + !totalPaid2Hos
                End If
                totalPayable2Mbm = !totalPayable2Mbm
                If Trim(!totalPaid2MBM) <> "" Then
                    totalPaid2MBM = totalPaid2MBM + totalPaid2MBM
                End If
                If Trim(!totalPaidByMbm) <> "" Then
                    totalPaidByMbm = totalPaidByMbm + totalPaidByMbm
                End If
                'txtMemOS = Format(CDbl(txtPayableMem) - CDbl(txtPaid2MBM) + CDbl(txtPaidByMem), "#,#.00;(#,#.00)")
                TotalReimburse = TotalReimburse + TotalReimburse
                totalExcess = totalExcess + !totalExcess
            End With
            Acc.MoveNext
            Loop
            
            txtApprovedAmt = totalApprove
            
            txtPayable2Hos = totalPayable2Hos
            If IsNumeric(txtPayable2Hos) = False Then
                txtPayable2Hos = 0
            End If
            
            txtPaid2Hos = totalPaid2Hos
            If IsNumeric(txtPaid2Hos) = False Then
                txtPaid2Hos = 0
            End If
            ' Out standing hos
            txtHosOS = Format(CDbl(txtPayable2Hos) - CDbl(txtPaid2Hos), "#,#.00;(#,#.00)")
            
            txtPayableMem = totalPayable2Mbm
            
            txtPaid2MBM = totalPaid2MBM
            If IsNumeric(txtPaid2MBM) = False Then
                txtPaid2MBM = 0
            End If
                
            txtPaidByMem = totalPaidByMbm
            If IsNumeric(txtPaidByMem) = False Then
                txtPaidByMem = 0
            End If
            ' Out standing MBM
            txtMemOS = Format(CDbl(txtPayableMem) - CDbl(txtPaid2MBM) + CDbl(txtPaidByMem), "#,#.00;(#,#.00)")
            txtReimbursement = Format(TotalReimburse, "#,#.00;(#,#.00)")
            txtExcess = Format(totalExcess, "#,#.00;(#,#.00)")
        End If
        Acc.Close
        Set Acc = Nothing
    End If
    'End If
    
End Sub

Private Sub Form_Load()
    Call ComboListing
    Call EnableEntries(False)
    bDataStatus = True
    counter = 0
    cmdShow.Enabled = True
    txtMBMNumber.Enabled = True
    cmdSearchMBM.Enabled = True
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim RecordExit
    If bExit = False Then
        RecordExit = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordExit = vbYes Then
            Unload Me
        Else
            Cancel = True
        End If
    End If
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
               ctl.Enabled = False
            End If
            'End If
        ElseIf TypeOf ctl Is ComboBox Then
            'If ctl.Enabled = True Then
                ctl.Enabled = True
                ctl.Text = ""
                ctl.Enabled = False
            'End If
        ElseIf TypeOf ctl Is CommandButton Then
            ctl.Enabled = True
            
        ElseIf TypeOf ctl Is MaskEdBox Then
            'If ctl.Enabled = True Then
                ctl.Enabled = True
                ctl.Mask = ""
                ctl.Text = ""
                ctl.Mask = "##/##/####"
                ctl.Enabled = False
            'End If
        End If
    Next
End Sub


Private Sub ComboListing()
    cboMBMTakeover.AddItem "N - None"
    cboMBMTakeover.AddItem "Y - Take Over"
    cboMBMTakeover.AddItem "C - Conversion"
    
    CboInsMode.AddItem "Y - Yearly"
    CboInsMode.AddItem "Q - Quaterly"
    CboInsMode.AddItem "B - By Yearly"
End Sub

Private Sub BNFListing()
    Dim rst2 As New ADODB.Recordset
    Dim BNF As ListItem
    Dim Sql As String
    If lstBenefitDetails.listitems.count = 0 And Len(Trim(cboHLTCode)) <> 0 _
     And Len(Trim(cboINSCode)) <> 0 And Len(Trim(cboPLNCode)) <> 0 Then
        Sql = " select BNFCode, BNFNoOFDays , " & _
          " BNFclaimableAmount, BNFdescription, BNFRMPerDIs " & _
          " from View_BNF where PLNCOde ='" & Trim(cboPLNCode) & "'" & _
          " AND HLTCode ='" & Trim(cboHLTCode) & "'" & _
          " AND INSCODe ='" & Trim(cboINSCode) & "'"
        rst2.Open Sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        Do Until rst2.EOF
            Set BNF = lstBenefitDetails.listitems.Add(, , rst2!BNFCode)
                If Trim(rst2!BNFdescription) <> "" Then
                    BNF.SubItems(1) = rst2!BNFdescription
                End If
                
                BNF.SubItems(2) = IIf(IsNull(rst2!BNFNoOFDays), 0, rst2!BNFNoOFDays)
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
    Dim HISList As ListItem
    Dim Sql As String
    If Len(Trim(txtMBMIcBcPP)) <> 0 Then
        If lstMBMHis.listitems.count = 0 Then
            Sql = " select MBMNumber, MBMPolicyNo, " & _
                " MBMPayorEffDate, MBMPayorExpDate, " & _
                " INSCode, PLNCode from View_MBM_MBMOthers " & _
                " where MBMIcBcPp ='" & txtMBMIcBcPP & "'" '& _
                '" and MBMFirstJoinedDate ='" & Format(txtMBMDateJoin, "mm/dd/yyyy") & "'"
'                " where MBMFirstMemNo = '" & RemStr(txtFirstMeMNo) & "'"

                 
            
            MBMHIS.Open Sql, medixmasterconn, adOpenKeyset, adLockOptimistic
            Do Until MBMHIS.EOF
                Set HISList = lstMBMHis.listitems.Add(, , MBMHIS!MBMNumber)
                    With MBMHIS
                        HISList.SubItems(1) = !MBMPolicyNo
                        HISList.SubItems(2) = !MBMPayorEffDate
                        HISList.SubItems(3) = !MBMPayorEXPDate
                        HISList.SubItems(4) = !INSCode
                        HISList.SubItems(5) = !PLNCode
                        .MoveNext
                    End With
            Loop
           
            MBMHIS.Close
            Set MBMHIS = Nothing
        End If
    End If
End Sub


Private Sub ADJListing()
    Dim ADJ As New ADODB.Recordset
    Dim ADJList As ListItem
    Dim Sql As String
    
    If Len(Trim(txtMBMNumber)) Then
         Sql = " select MBMAmendID , MBMNumber," & _
           " MBMAmendEdtType , MBMAmendDate , MBMAENDtEffDate ,   " & _
           " MBMAEndtBordxDate,  MBMAEndtBatchNo  " & _
           " from MBMAmendHistory where MBMnumber ='" & Trim(txtMBMNumber) & "'"
         ADJ.Open Sql, medixmasterconn, adOpenKeyset, adLockOptimistic
         
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


' ######## List out all the cases which have been discharged
Private Sub ListCases()
    Dim CASList As ListItem
    Dim strsql As String
    Dim CAS As New ADODB.Recordset

    strsql = "select CASNumber, MBMCName, " & _
     " DIAPrelimDiag1 , " & _
    " CASDischargeDate, CASDateAdmitted, " & _
    " CASStatus,   MBMName, MBMNumber  "
    strsql = strsql & " from VIEW_MBMInquiry_CASe where 0=0"
    strsql = strsql & " AND  MBMNumber = '" & RemStr(txtMBMNumber) & "' "
            
    CAS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If CAS.recordCount <> 0 Then
        Screen.MousePointer = vbHourglass
        lstCaseList.listitems.clear
        Do Until CAS.EOF
            Set CASList = lstCaseList.listitems.Add(, , CAS!CASNumber)
                CASList.SubItems(1) = CAS!MBMNumber
                If MBMCName = "" Then
                    CASList.SubItems(2) = "Self"
                    CASList.SubItems(3) = "Self"
                Else
                    CASList.SubItems(2) = CAS!MBMCName
                    CASList.SubItems(3) = CAS!MBMCRELCode
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
                If Trim(CAS!DIAPrelimDiag1) <> "" Then
                    CASList.SubItems(6) = CAS!DIAPrelimDiag1
                End If
                
                CASList.SubItems(7) = CAS!CASStatus
                'C'ASList.SubItems(8) = CAS!PLNCode
                'If Trim(CAS!CASRemark) <> "" Then
                 '   CASList.SubItems(9) = CAS!CASRemark
                'End If
                
                
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
        And Len(Trim(txtMBMBirthdate)) <> 0 _
        And Trim(txtMBMBirthdate) <> "__/__/____" _
        And Len(Trim(cboHLTCode)) <> 0 Then
        If IsDate(txtMBMBirthdate) Then
            SelectAge = GetAge(txtMBMPayorEffDate, txtMBMBirthdate)
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
        txtMCOFees = Format(GetMCO(Trim(Mid(cboINSCode, 1, 1)), Trim(Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1)), txtMBMAdultChild, CDate(txtMBMPayorEffDate)), "#,##0.00")
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
    Dim Sql As String
    If Trim(txtMBMNumber) <> "" Then
        Sql = " select MBMNumber from MBM where MBMnumber < '" & RemStr(txtMBMNumber) & "' " & _
            " Order by MBMnumber DESC "
        MBM.MaxRecords = 1
        MBM.Open Sql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        If MBM.recordCount Then
            txtMBMNumber.Text = MBM!MBMNumber
            cmdshow_Click
        Else
            MsgBox "First Record Reached! " & vbNewLine & "Not a valid Membership Number Format! ", vbCritical + vbOKOnly
        End If
        MBM.Close
        Set MBM = Nothing
    End If
End Sub

Private Sub CmdNext_Click()
    Dim MBM As New ADODB.Recordset
    Dim Sql As String
    Sql = " select MBMNumber from MBM where MBMnumber > '" & RemStr(txtMBMNumber) & "' " & _
        " Order by MBMnumber "
    MBM.MaxRecords = 1
    MBM.Open Sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If MBM.recordCount Then
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

Private Sub lstMBMAcc_Click()
Dim strsql As String
Dim MBM As New ADODB.Recordset
    
    If lstMBMAcc.listitems.count <> 0 Then
    
        strsql = "Select * from View_MBM_MemberType where MBMNumber = '" & Trim(txtMBMNumber) & "'"
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
        
        txtApprovedAmt = Format(lstMBMAcc.SelectedItem.SubItems(2), "#,#.00")
        txtPayable2Hos = Format(lstMBMAcc.SelectedItem.SubItems(3), "#,#.00")
        If IsNumeric(txtPayable2Hos) = False Then
            txtPayable2Hos = 0
        End If
        If Trim(lstMBMAcc.SelectedItem.SubItems(4)) <> "" Then
            txtPaid2Hos = Format(lstMBMAcc.SelectedItem.SubItems(4), "#,#.00")
        End If
        If IsNumeric(txtPaid2Hos) = False Then
            txtPaid2Hos = 0
        End If
        
        ' Out standing hos
        txtHosOS = Format(CDbl(txtPayable2Hos) - CDbl(txtPaid2Hos), "#,#.00;(#,#.00)")
        
        txtPayableMem = Format(lstMBMAcc.SelectedItem.SubItems(5), "#,#.00;(#,#.00)")
        If IsNumeric(txtPayableMem) = False Then
            txtPayableMem = 0
        End If
        If Trim(lstMBMAcc.SelectedItem.SubItems(6)) <> "" Then
            txtPaid2MBM = Format(lstMBMAcc.SelectedItem.SubItems(6), "#,#.00")
        End If
        If IsNumeric(txtPaid2MBM) = False Then
            txtPaid2MBM = 0
        End If
        If Trim(lstMBMAcc.SelectedItem.SubItems(7)) <> "" Then
            txtPaidByMem = Format(lstMBMAcc.SelectedItem.SubItems(7), "#,#.00")
        End If
        If IsNumeric(txtPaidByMem) = False Then
            txtPaidByMem = 0
        End If
        txtMemOS = Format(CDbl(txtPayableMem) - CDbl(txtPaid2MBM) + CDbl(txtPaidByMem), "#,#.00;(#,#.00)")
        txtReimbursement = Format(lstMBMAcc.SelectedItem.SubItems(8), "#,#.00")
        txtExcess = Format(lstMBMAcc.SelectedItem.SubItems(9), "#,#.00;(#,#.00)")
    End If
    End If
End Sub

Private Sub lstMBMHis_DblClick()
    'Dim f As New frmMBMInquiry
    'f.show
    'f.txtMBMNumber = lstMBMHis.SelectedItem.Text
    'f.cmdshow_Click
    txtMBMNumber = lstMBMHis.SelectedItem.Text
    cmdshow_Click
    SSTab1.Tab = 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 2 Then
        If Mid(cboINSCode, 1, 1) = "I" Or Mid(cboINSCode, 1, 1) = "G" Then
            MsgBox "This Plan does not have supplementary", vbOKOnly + vbCritical, DialogBoxTitle
        SSTab1.Tab = 0
        Else
        SSTab1.Tab = 2
        End If
        
    End If
    
    If SSTab1.Tab = 4 Then
        If lstBenefitDetails.listitems.count = 0 Then
            Call BNFListing
        End If
    End If
    
    If SSTab1.Tab = 3 Then
        lstAdjustmentHis.listitems.clear
        If lstAdjustmentHis.listitems.count = 0 Then
            Call ADJListing
        End If
    End If
    
    If SSTab1.Tab = 5 Then
        If Trim(txtMBMNumber) <> "" And lstCaseList.listitems.count = 0 Then
            Call ListCases
        End If
    End If
    
    If SSTab1.Tab = 6 Then
        If lstMBMHis.listitems.count = 0 Then
            Call MBMHISListing
        End If
    End If
    
    If SSTab1.Tab = 7 Then
        'If lstMBMAcc.listitems.count = 0 Then
        lstMBMAcc.listitems.clear
            Call GetAccountSummary
        'End If
    End If
End Sub

Private Sub lstCaseList_DblClick()
    frmCaseInquiry.txtCaseNo = ChkStr(lstCaseList.SelectedItem.Text)
    Call frmCaseInquiry.cmdshow_Click
    frmCaseInquiry.show
End Sub


Public Sub cmdshow_Click()
    Dim getRecord As Integer
    
    getRecord = showMBM
    Screen.MousePointer = vbHourglass
    If getRecord = 1 Then
        If txtMBMNumber = "" Then
            MsgBox "Please Enter Membership No", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            lstBenefitDetails.listitems.clear
            lstCaseList.listitems.clear
            lstMBMHis.listitems.clear
            Call showMBMCov
'            Call ProviderCallFees
            Call CallLimit
            'Call MemberCallFees
            Call BNFListing
            Call ListCases
            Call MBMHISListing
            Call GetAccountSummary
            'txtMBMPremiumEff = Format(CDbl(txtMBMPremium) + CDbl(MCOPremiumReduction(Format(txtMBMPayorExpDate, "mm/dd/yyyy"), Date, txtMBMPremium)), "######.99")
            'txtMCOFeesEff = Format(CDbl(txtMCOFees) + CDbl(MCOPremiumReduction(Format(txtMBMPayorExpDate, "mm/dd/yyyy"), Date, txtMCOFees)), "######.99")
        End If
    End If
    Screen.MousePointer = Default
End Sub

Private Function showMBM() As Integer
Dim sqlstr, sqlPolStatus   As String
Dim sqlstqMBMRef As String
Dim MBM As New ADODB.Recordset
Dim MBMRef As New ADODB.Recordset
Dim PolStatus As New ADODB.Recordset
Dim MBMNumber As String

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
    
    sqlstr = "Select MBMNumber, HLTCode, MBMPolicyNo,  MBMName , MBMIcBcPp,  " & _
            " MBMMemberType ,  MBMStatus, " & _
            " INSCode, PLNCode, MBMRenewFlag, " & _
            " MBMAddress1, MBMAddress2, " & _
            " MBMAddress3, MBMPostCode, CTYCode, " & _
            " MBMTelnoH, MBMEmail, MBMSalutation, " & _
            " MBMSex, SRVCodeList, MBMTelNoM, " & _
            " MBMTelNoO, MBMOccupation, " & _
            " MBMHeight, MBMWeight, " & _
            " MBMWorkNature," & _
            " MBMGroupCompany, " & _
            " MBMEmployeeNo, MBMBlood, " & _
            " MBMSmoking, MBMAlcohol, " & _
            " MBMExclusion, MBMAllergic, " & _
            " MBMMaritalStatus, MBMAccessFee, MBMPremium , MBMMCOfee , " & _
            " MBMFirstJoinedDate , BRNCode, " & _
            " MBMDOB, NATCode, RACCode," & _
            " LGNCode, STACode, MBMFirstPolNo , MBMFirstMEMNo,  " & _
            " BRNCode, MBMTakeOver, MBMAgentCode , MBMPolicyYear"
        sqlstr = sqlstr & ", PAYCode, MBMBatchNo, " & _
                " MBMBordxDate, AGECode ,  MBMPrevMemNo, MBMPrevPolNo ," & _
                " MBMPayorEffDate, MBMPayorExpDate , MBMAdultChild, MBMAvailableLimit, " & _
                " MBMLastEndorseStatus, MBMENDtBordxDate, " & _
                " MBMReinstate , MBMLastAdjustmentDate, MBMInstallment "
    sqlstr = sqlstr & " From View_MBM_MBMOthers Where MBMNumber ='" & Trim(txtMBMNumber) & "'"

    If bDataStatus = True Then
        MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    Else
        MBM.Open sqlstr, medixmasterconnA, adOpenKeyset, adLockOptimistic
    End If
    Call ClearEntries
    
    If MBM.recordCount = 0 Then
       MsgBox "No record found!", vbCritical
    Else
        showMBM = 1
        With MBM
            txtMBMName = IIf(!MBMStatus = "S", Trim(!MBMName) & " (Suspended)", Trim(!MBMName))
            txtMBMIcBcPP = Trim(!MBMIcBcPp)
            
            txtMBMStatus = Trim(!MBMStatus)
            
            sqlPolStatus = " select pstatuscode ,pstatusDesc from PolicyStatus where PStatusCode ='" & Trim(txtMBMStatus) & "'"
            PolStatus.Open sqlPolStatus, medixmasterconn, adOpenKeyset, adLockOptimistic
            If PolStatus.recordCount Then
                txtMBMStatus = PolStatus!pstatusDesc
            End If
            
            PolStatus.Close
            Set PolStatus = Nothing
            
            
            txtMBMPolNo = Trim(!MBMPolicyNo)
            cboINSCode = Trim(!INSCode)
            cboPLNCode = Trim(!PLNCode)
            txtMBMAdd1 = Trim(!MBMAddress1)
            txtMBMAdd2 = Trim(!MBMAddress2)
            txtMBMAdd3 = Trim(!MBMAddress3)
            txtMBMPostCode = Trim(!MBMPostCode)
            cboCTYCode = Trim(!CTYCode)
            If Len(!MBMRenewFlag) Then
                cboMBMRenewal = !MBMRenewFlag
            End If
            
            If Trim(!MBMTelnoH) <> "" Then
                txtMBMTelNoH = Trim(!MBMTelnoH)
            End If
            If Trim(!MBMEmail) <> "" Then
                txtMBMEmail = Trim(!MBMEmail)
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
            If Trim(!MBMTelNoM) <> "" Then
                txtMBMTelNoM = Trim(!MBMTelNoM)
            End If
            If Trim(!MBMTelNoO) <> "" Then
                txtMBMTelNoO = Trim(!MBMTelNoO)
            End If
            If Trim(!MBMOccupation) <> "" Then
                cboMBMOccupation = Trim(!MBMOccupation)
            End If
            If Trim(!MBMHeight) <> "" Then
                txtMBMHeight = Trim(!MBMHeight)
            End If
            If Trim(!MBMWeight) <> "" Then
                txtMBMWeight = Trim(!MBMWeight)
            End If
            If Trim(!MBMWorkNature) <> "" Then
                txtMBMWorkNature = Trim(!MBMWorkNature)
            End If
            If Trim(!MBMGroupCompany) <> "" Then
                txtMBMGroupCompany = Trim(!MBMGroupCompany)
            End If
            If Trim(!MBMEmployeeNo) <> "" Then
                txtMBMEmployeeNo = Trim(!MBMEmployeeNo)
            End If
            If Trim(!MBMBlood) <> "" Then
                cboMBMBlood = Trim(!MBMBlood)
            End If
            If Trim(!MBMSmoking) <> "" Then
                cboMBMSmoking = Trim(!MBMSmoking)
            End If
            If Trim(!MBMAlcohol) <> "" Then
                cboMBMAlcohol = Trim(!MBMAlcohol)
            End If
            If Trim(!MBMExclusion) <> "" Then
                txtMBMExclusion = Trim(!MBMExclusion)
            End If
            If Trim(!MBMAllergic) <> "" Then
                txtMBMAllergic = Trim(!MBMAllergic)
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
            If Trim(!LGNCode) <> "" Then
                cboLGNCode = Trim(!LGNCode)
            End If
            If Trim(!STACode) <> "" Then
                cboSTACode = Trim(!STACode)
            End If
            'If Trim(!MBMDepartment) <> "" Then
            '    txtDeptName = Trim(!MBMDepartment)
            'End If
            'If Trim(!MBMFirstPolNo) <> "" Then
            '    txtFirstPolNo = Trim(!MBMFirstPolNo)
            'End If
            'If Trim(!MBMPrevMemNo) <> "" Then
            '    txtMBMPrevMEMNo = Trim(!MBMPrevMemNo)
            'End If
            'If Trim(!MBMPrevPolNo) <> "" Then
            '    txtMBMPrevPolNo = Trim(!MBMPrevPolNo)
            'End If
            If Trim(!MBMDOB) <> "" Then
                txtMBMBirthdate.Mask = ""
                txtMBMBirthdate.Text = Trim(!MBMDOB)
            End If
            If Trim(!MBMFirstMEMNo) <> "" Then
                txtFirstMemNo = Trim(!MBMFirstMEMNo)
            End If
            If Trim(!BRNCode) <> "" Then
                cboBRNCode = Trim(!BRNCode)
            End If
            If Trim(!MBMTakeOver) = "C" Then
                cboMBMTakeover.ListIndex = 2
            ElseIf Trim(!MBMTakeOver) = "Y" Then
                cboMBMTakeover.ListIndex = 1
            Else
                cboMBMTakeover.ListIndex = 0
            End If
            If Trim(!MBMAgentCode) = "C" Then
                txtMBMAgentCode = Trim(!MBMAgentCode)
            End If
            If Trim(!MBMMemberType) = "P" Then
                cboMBMRelCode = "C - Combined Limit"
            ElseIf Trim(!MBMMemberType) = "Q" Then
                cboMBMRelCode = "S - Separated Limit"
            Else
                cboMBMRelCode = Trim(!MBMMemberType)
            End If
            
            cboPAYCode = Trim(!PAYCode)
            txtAGECode = Trim(!AgeCode)
            If Len(!MBMBordxDate) Then
                txtMBMBordxDate.Mask = ""
                txtMBMBordxDate = Trim(!MBMBordxDate)
            End If
            If Len(!MBMBatchNo) Then
                txtMBMBatchNo = Trim(!MBMBatchNo)
            End If
            cboHLTCode = Trim(!HLTCode)
            txtMBMPayorEffDate.Mask = ""
            txtMBMPayorEffDate = Trim(!MBMPayorEffDate)
            txtMBMPayorExpDate.Mask = ""
            txtMBMPayorExpDate = Trim(!MBMPayorEXPDate)
            txtMBMPolicyYear = Trim(!MBMPolicyYear)
            txtAvailableAnnualLimit = Format(Trim(!MBMAvailableLimit), "#,##0.00")
            txtMBMAdultChild = Trim(!MBMAdultChild)
        
            
            txtMBMPremium = Format(IIf(IsNull(!MBMPremium) = True, Format(GetPremium(Trim(cboINSCode), Trim(cboPAYCode), Trim(txtAGECode), Trim(cboPLNCode), Trim(cboHLTCode), txtMBMPayorEffDate), "#,##0.00"), !MBMPremium), "####0.00")
            txtMCOFees = Format(IIf(IsNull(!MBMMCOfee) = True, Format(GetMCO(Trim(cboINSCode), Trim(cboHLTCode), txtMBMAdultChild, txtMBMPayorEffDate), "#,##0.00"), !MBMMCOfee), "####0.00")
            If IsNull(!MBMAccessFee) = True Then
                ProviderCallFees
            Else
                txtMBMProviderCharges = Format(!MBMAccessFee, "####0.00")
               ' GetProString = GetProviderCharges(cboSRVCode, cboINSCode, cboHLTCode, txtMBMAdultChild)
             '   txtMBMCaseCharges = Format(Mid(GetProString, InStr(1, GetProString, "-") + 1), "#,##0.00")
            End If
  
        End With
        
            If MBM!MBMStatus = "C" Then
        
            MBMNumber = Trim(MBM!MBMNumber)
            sqlstqMBMRef = "Select * from MBMRefund where MBMNumber =  '" & Trim(MBMNumber) & "'"
            MBMRef.Open sqlstqMBMRef, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            With MBMRef
                    txtMBMPremiumRed = Format(Round(!MBMPremReduction, 0), "#,##0.00")
                    txtMBMMCORed = Format(Round(!MBMMCORefund, 0), "#,##0.00")
                    txtMBMProviderRedCharges = Format(Round(!MBMProviderRef, 2), "#,##0.00")
                    txtMBMProviderEffCharges = Format(Round(!MBMProviderEff, 2), "#,##0.00")
                    txtMCOFeesEff = Format(Round(!MBMMCOEffective, 0), "#,##0.00")
                    txtMBMPremiumEff = Format(Round(!MBMPREMEffective, 0), "#,##0.00")
            End With
            MBMRef.Close
            Set MBMRef = Nothing
        End If

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
        "  MBMCOccupation, " & _
        " MBMCWorkNature, MBMCSmoking, MBMCAlcohol "
    
    If bDataStatus = True Then
        sqlstr = sqlstr & " from  MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
        MBMCov.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    Else
        sqlstr = sqlstr & " from  MBMCoveredPersonsArchive where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
        MBMCov.Open sqlstr, medixmasterconnA, adOpenKeyset, adLockOptimistic
    End If

    lstCovAdd.listitems.clear
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
                    MemberConveredPerson.SubItems(8) = !MBMCAllergic
                End If
                
                If Len(Trim(!MBMCExclusion)) Then
                    MemberConveredPerson.SubItems(9) = !MBMCExclusion
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
            End With
            MBMCov.MoveNext
        Loop
    End If
    MBMCov.Close
    Set MBMCov = Nothing
End Sub

Private Sub cmdEditMBM_Click()
    EnableEntries (True)
    Frame10.Enabled = True
    Frame1.Enabled = False
    SSTab1.Enabled = True
    CmdAdd.Enabled = False
    CmdUpdate.Enabled = False
    CmdRemove.Enabled = False
    'cmdEditMBM.Enabled = False
    cboEndorseStatus.Enabled = True
End Sub

Private Sub CmdExit_Click()
Dim RecordExit
    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
    If RecordExit = vbYes Then
        bExit = True
        Unload Me
        Set frmMBMInquiry = Nothing
    End If
End Sub

Private Sub cmdMBMPolicy_Click()
    frmSearchMBM.txtMBMPolicyNo1.value = True
    frmSearchMBM.txtMBMNo1.value = False
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
        txtCovBirthdate.Mask = ""
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        cboCovBlood = lstCovAdd.SelectedItem.SubItems(7)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(10)
        txtCovAnnualLimit = Format(lstCovAdd.SelectedItem.SubItems(12), "###,##0.00")
        txtAvailLimit = lstCovAdd.SelectedItem.SubItems(13)
        cboOccupation = lstCovAdd.SelectedItem.SubItems(14)
        cboWorkNature = lstCovAdd.SelectedItem.SubItems(15)
        cboSmoking = lstCovAdd.SelectedItem.SubItems(16)
        cboAlcohol = lstCovAdd.SelectedItem.SubItems(17)

    End If
    Call GetAccountSummary
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

' ################## Start processing


'  ############### SAVE DATA INTO DATABSE #####################
Public Sub MBMRefund()
    Dim ServiceChargeRefundValue As Double
    Dim MBMRefund As New ADODB.Recordset
    Dim strsql As String
    Dim MCORefundValue As String
    Dim MBMMCOFees As Double
    Dim PremiumRefundValue As String
    Dim MCOEffective As String
    Dim PremiumEffective As String
    Dim FindAge As String
    Dim MBMPremium As Double
    Dim PlanCode As String
    Dim InsuredCode As String
    Dim HLTCode As String
    Dim payor As String
    Dim AgeValue As String
    Dim PayEffDate As String
    Dim PayExpDate As String
    Dim MBMDOB As String
    Dim ProviderRefValue As Double
    Dim ProviderEff As String
    
    HLTCode = Mid(cboHLTCode, 1, 1)
    If InStr(1, cboPLNCode, "-") > 1 Then
        PlanCode = Trim(Mid(cboPLNCode, 1, InStr(1, cboPLNCode, "-") - 1))
    Else
        PlanCode = Trim(Mid(PlanCode, 1, 4))
    End If
    InsuredCode = Mid(cboINSCode, 1, 1)
    payor = Mid(cboPAYCode, 1, 2)
    MBMDOB = txtMBMBirthdate
    PayEffDate = txtMBMPayorEffDate
    PayExpDate = txtMBMPayorExpDate
        
    strsql = "Select * from MBMRefund Where MBMNumber = '" & txtMBMNumber & "'"
    MBMRefund.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic

    FindAge = GetAge(CDate(PayEffDate), CDate(MBMDOB))
    AgeValue = Mid(GetAgeBand(HLTCode, FindAge), 1, 2)
    
    MBMPremium = GetPremium(InsuredCode, payor, AgeValue, PlanCode, HLTCode, CDate(PayEffDate))

    If AgeValue = "01" Then
        MBMMCOFees = GetMCO(InsuredCode, HLTCode, "C", CDate(PayEffDate))
    Else
        MBMMCOFees = GetMCO(InsuredCode, HLTCode, "A", CDate(PayEffDate))
    End If
    
    MCORefundValue = MCORefund(CDate(PayExpDate), CDate(txtDate), CDate(PayEffDate), MBMMCOFees)
    PremiumRefundValue = MCOPremiumReduction(CDate(PayExpDate), CDate(txtDate), CDate(PayEffDate), MBMPremium)
    ProviderRefValue = EAFeesReduction(CDate(PayExpDate), CDate(txtDate), CDate(PayEffDate), txtMBMProviderCharges)
    
    PremiumEffective = CDbl(MBMPremium) - CDbl(PremiumRefundValue)
    MCOEffective = CDbl(MBMMCOFees) - CDbl(MCORefundValue)
    ProviderEff = CDbl(txtMBMProviderCharges) - CDbl(ProviderRefValue)
        MBMRefund.AddNew
        With MBMRefund
            !MBMNumber = txtMBMNumber
            !MBMPREMEffective = PremiumEffective
            !MBMPremReduction = PremiumRefundValue
            !MBMMCOEffective = MCOEffective
            !MBMMCORefund = MCORefundValue
            !MBMProviderRef = ProviderRefValue
            !MBMProviderEff = ProviderEff
        End With
        MBMRefund.Update
    txtMBMPremiumEff = Format(PremiumEffective, "#,##0.00")
    txtMCOFeesEff = Format(MCOEffective, "#,##0.00")
    txtMBMPremiumRed = Format(PremiumRefundValue, "#,##0.00")
    txtMBMMCORed = Format(MCORefundValue, "#,##0.00")
    txtMBMProviderRedCharges = Format(ProviderRefValue, "#,##0.00")
    txtMBMProviderEffCharges = Format(ProviderEff, "#,##0.00")
End Sub

Private Sub InsertPrem()
Dim MBMInsertPrem As Double
Dim MBMPrem As Double
Dim PayExpDate As String

MBMPrem = txtMCOFees
PayExpDate = txtMBMPayorExpDate

    MBMInsertPrem = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(txtMBMPremium, "#,##0.00"))
    txtMBMPremium = ""
    txtMBMPremium = Format(MBMInsertPrem, "#,##0.00")
End Sub

Private Sub GetAccountSummary()
    Dim Acc As New ADODB.Recordset
    Dim strsql As String
    Dim masterlist As ListItem
    
    strsql = " select totalApprove, totalPayable2Hos, " & _
            " totalPaid2Hos, totalPayable2Mbm, totalPaidByMbm, MBMPatientCOVID, " & _
            " totalPaid2MBM , totalExcess , totalReimburse from view_MBMTotal_summary where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    Acc.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If Not Acc.EOF Then
        Do Until Acc.EOF
        With Acc
        Set masterlist = lstMBMAcc.listitems.Add(, , txtMBMNumber)
            masterlist.SubItems(1) = Format(!mbmpatientcovid)
            masterlist.SubItems(2) = Format(!totalApprove, "#,#.00")
            masterlist.SubItems(3) = Format(!totalPayable2Hos, "#,#.00")
            'If IsNumeric(txtPayable2Hos) = False Then
            '    txtPayable2Hos = 0
            'End If
            If Trim(!totalPaid2Hos) <> "" Then
                masterlist.SubItems(4) = Format(!totalPaid2Hos, "#,#.00")
            End If
            If IsNumeric(masterlist.SubItems(4)) = False Then
                masterlist.SubItems(4) = 0
            End If
            
            ' Out standing hos
            'txtHosOS = Format(CDbl(txtPayable2Hos) - CDbl(txtPaid2Hos), "#,#.00;(#,#.00)")
            
            masterlist.SubItems(5) = Format(!totalPayable2Mbm, "#,#.00;(#,#.00)")
            If IsNumeric(masterlist.SubItems(3)) = False Then
                masterlist.SubItems(5) = 0
            End If
            If Trim(!totalPaid2MBM) <> "" Then
                masterlist.SubItems(6) = Format(!totalPaid2MBM, "#,#.00")
            End If
            If IsNumeric(masterlist.SubItems(6)) = False Then
                masterlist.SubItems(6) = 0
            End If
            If Trim(!totalPaidByMbm) <> "" Then
                masterlist.SubItems(7) = Format(!totalPaidByMbm, "#,#.00")
            End If
            If IsNumeric(masterlist.SubItems(7)) = False Then
                masterlist.SubItems(7) = 0
            End If
            'txtMemOS = Format(CDbl(txtPayableMem) - CDbl(txtPaid2MBM) + CDbl(txtPaidByMem), "#,#.00;(#,#.00)")
            masterlist.SubItems(8) = Format(!TotalReimburse, "#,#.00")
            masterlist.SubItems(9) = Format(!totalExcess, "#,#.00;(#,#.00)")
        End With
        Acc.MoveNext
        Loop
    End If
    Acc.Close
    Set Acc = Nothing

End Sub
