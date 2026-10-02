VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMBMAdjustment 
   Caption         =   "Member Adjustment"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   -240
      TabIndex        =   83
      Top             =   0
      Width           =   12120
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   15
         Left            =   240
         TabIndex        =   84
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Member Adjustment"
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
         Height          =   255
         Left            =   360
         TabIndex        =   85
         Top             =   0
         Width           =   6615
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4965
      Left            =   120
      TabIndex        =   92
      Top             =   2400
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8758
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
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
      TabPicture(0)   =   "frmMBMAdjustment.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Supplementary"
      TabPicture(1)   =   "frmMBMAdjustment.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstCovAdd"
      Tab(1).Control(1)=   "Frame6"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Adjustment History"
      TabPicture(2)   =   "frmMBMAdjustment.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstAdjustmentHis"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Fees"
      TabPicture(3)   =   "frmMBMAdjustment.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame6 
         Height          =   3075
         Left            =   -74880
         TabIndex        =   203
         Top             =   360
         Width           =   11460
         Begin VB.TextBox txtCovAllergic 
            Height          =   915
            Left            =   4680
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   72
            Top             =   2040
            Width           =   2415
         End
         Begin VB.TextBox txtCovExclusion 
            Height          =   915
            Left            =   7200
            MaxLength       =   2500
            MultiLine       =   -1  'True
            TabIndex        =   73
            Top             =   2040
            Width           =   2775
         End
         Begin VB.ComboBox cboCovSalutation 
            Height          =   315
            Left            =   1320
            TabIndex        =   44
            Text            =   "MS"
            Top             =   195
            Width           =   975
         End
         Begin VB.TextBox txtCovName 
            Height          =   330
            Left            =   3480
            MaxLength       =   60
            TabIndex        =   45
            Top             =   180
            Width           =   3660
         End
         Begin VB.Frame Frame9 
            Height          =   1695
            Left            =   10320
            TabIndex        =   171
            Top             =   120
            Width           =   1065
            Begin VB.CommandButton cmdRemove 
               Caption         =   "&Remove"
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   75
               Top             =   600
               Width           =   840
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "&Add"
               Height          =   315
               Left            =   120
               TabIndex        =   74
               Top             =   240
               Width           =   840
            End
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "Update"
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   76
               Top             =   960
               Width           =   840
            End
            Begin VB.CommandButton cmdClear 
               Caption         =   "Clear"
               Height          =   315
               Left            =   120
               TabIndex        =   77
               Top             =   1320
               Width           =   840
            End
         End
         Begin VB.TextBox txtCovIcBcPp 
            Height          =   285
            Left            =   7800
            MaxLength       =   30
            TabIndex        =   46
            Top             =   180
            Width           =   2055
         End
         Begin VB.ComboBox cboCovRelationship 
            Height          =   315
            Left            =   1320
            TabIndex        =   47
            Text            =   "W -Wife"
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox cboCovAdultChild 
            Height          =   315
            Left            =   1320
            TabIndex        =   51
            Text            =   "A"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtPrevPLNCode 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5760
            TabIndex        =   49
            Top             =   480
            Width           =   1380
         End
         Begin VB.TextBox txtAvailableLimit 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   54
            Top             =   1010
            Width           =   975
         End
         Begin VB.TextBox txtPrevINSCom 
            Height          =   330
            Left            =   7800
            MaxLength       =   60
            TabIndex        =   50
            Top             =   440
            Width           =   2055
         End
         Begin VB.TextBox txtRBAmt 
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   58
            Top             =   1290
            Width           =   975
         End
         Begin VB.TextBox txtSuppPayPremGross 
            Height          =   285
            Left            =   1320
            TabIndex        =   63
            Top             =   1545
            Width           =   975
         End
         Begin VB.TextBox txtCOVPayRemarks 
            Height          =   285
            Left            =   5760
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   53
            Top             =   720
            Width           =   4095
         End
         Begin VB.TextBox txtSuppGenWP 
            Height          =   285
            Left            =   5760
            MaxLength       =   9
            TabIndex        =   56
            Top             =   960
            Width           =   945
         End
         Begin VB.TextBox txtSuppPREWP 
            Height          =   285
            Left            =   5760
            MaxLength       =   9
            TabIndex        =   60
            Top             =   1200
            Width           =   945
         End
         Begin VB.TextBox TxtSuppSpeWP 
            Height          =   300
            Left            =   5760
            MaxLength       =   9
            TabIndex        =   65
            Top             =   1455
            Width           =   945
         End
         Begin MSMask.MaskEdBox txtCovBirthdate 
            Height          =   285
            Left            =   3480
            TabIndex        =   48
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
            TabIndex        =   62
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
            TabIndex        =   67
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
            TabIndex        =   57
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
            TabIndex        =   61
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
            TabIndex        =   66
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
            TabIndex        =   52
            Text            =   "M"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtCovAnnualLimit 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   55
            Top             =   1005
            Width           =   1095
         End
         Begin VB.ComboBox cboSuppStaffDisc 
            Height          =   315
            Left            =   3480
            Sorted          =   -1  'True
            TabIndex        =   59
            Top             =   1260
            Width           =   1095
         End
         Begin VB.TextBox txtSuppPayMCOGross 
            Height          =   285
            Left            =   1320
            TabIndex        =   68
            Top             =   1805
            Width           =   975
         End
         Begin VB.TextBox txtRenewalDisc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   70
            Top             =   2055
            Width           =   975
         End
         Begin VB.TextBox txtSuppPayPremNet 
            Height          =   285
            Left            =   3480
            TabIndex        =   64
            Top             =   1550
            Width           =   1095
         End
         Begin VB.TextBox txtSuppPayMCONet 
            Height          =   285
            Left            =   3480
            TabIndex        =   69
            Top             =   1805
            Width           =   1095
         End
         Begin VB.TextBox txtPolicyDisc 
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   71
            Top             =   2055
            Width           =   1095
         End
         Begin VB.Label Label32 
            Caption         =   "Policy Disc."
            Height          =   255
            Left            =   2475
            TabIndex        =   202
            Top             =   2040
            Width           =   1365
         End
         Begin VB.Label Label28 
            Caption         =   "Renewal Disc."
            Height          =   255
            Left            =   120
            TabIndex        =   201
            Top             =   2055
            Width           =   1485
         End
         Begin VB.Label Label23 
            Caption         =   "Pay MCO Net"
            Height          =   255
            Left            =   2475
            TabIndex        =   200
            Top             =   1815
            Width           =   1365
         End
         Begin VB.Label Label13 
            Caption         =   "Pay MCO Gross"
            Height          =   255
            Left            =   120
            TabIndex        =   199
            Top             =   1815
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Pay Prem Net"
            Height          =   255
            Left            =   2480
            TabIndex        =   198
            Top             =   1560
            Width           =   1365
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            Caption         =   "To"
            Height          =   255
            Left            =   8400
            TabIndex        =   197
            Top             =   1440
            Width           =   285
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            Caption         =   "From"
            Height          =   255
            Left            =   8400
            TabIndex        =   196
            Top             =   1245
            Width           =   405
         End
         Begin VB.Label Label111 
            Caption         =   "Pay Prem Gross"
            Height          =   255
            Left            =   120
            TabIndex        =   195
            Top             =   1560
            Width           =   1365
         End
         Begin VB.Label Label37 
            Caption         =   "Clm Non-Pay Date"
            Height          =   375
            Left            =   8450
            TabIndex        =   194
            Top             =   1000
            Width           =   1365
         End
         Begin VB.Label Label53 
            Caption         =   "EOP"
            Height          =   255
            Left            =   6840
            TabIndex        =   193
            Top             =   1245
            Width           =   1005
         End
         Begin VB.Label Label104 
            Caption         =   "Spe. Ill. W. P."
            Height          =   255
            Left            =   4680
            TabIndex        =   192
            Top             =   1530
            Width           =   1065
         End
         Begin VB.Label Label103 
            Caption         =   "Pre W. P."
            Height          =   255
            Left            =   4680
            TabIndex        =   191
            Top             =   1245
            Width           =   780
         End
         Begin VB.Label Label78 
            Caption         =   "Gen W. P."
            Height          =   255
            Left            =   4680
            TabIndex        =   190
            Top             =   990
            Width           =   900
         End
         Begin VB.Label Label33 
            Caption         =   "EOP"
            Height          =   255
            Left            =   6840
            TabIndex        =   189
            Top             =   1000
            Width           =   1005
         End
         Begin VB.Label Label45 
            Caption         =   "EOP"
            Height          =   255
            Left            =   6840
            TabIndex        =   188
            Top             =   1500
            Width           =   1005
         End
         Begin VB.Label Label36 
            Caption         =   "Disc Req"
            Height          =   255
            Left            =   2480
            TabIndex        =   187
            Top             =   1320
            Width           =   1065
         End
         Begin VB.Label Label92 
            Caption         =   "Pay Remarks"
            Height          =   255
            Left            =   4680
            TabIndex        =   186
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label91 
            Caption         =   "T/O Ins "
            Height          =   255
            Left            =   7200
            TabIndex        =   185
            Top             =   480
            Width           =   1140
         End
         Begin VB.Label Label90 
            Caption         =   "R&&B Amt"
            Height          =   255
            Left            =   120
            TabIndex        =   184
            Top             =   1320
            Width           =   1140
         End
         Begin VB.Label Label82 
            Caption         =   "P. Pln Code"
            Height          =   255
            Left            =   4680
            TabIndex        =   183
            Top             =   525
            Width           =   1125
         End
         Begin VB.Label Label38 
            Caption         =   "Name"
            Height          =   255
            Left            =   2480
            TabIndex        =   182
            Top             =   195
            Width           =   1140
         End
         Begin VB.Label Label51 
            Caption         =   "Adult/Child"
            Height          =   255
            Left            =   120
            TabIndex        =   181
            Top             =   795
            Width           =   825
         End
         Begin VB.Label Label50 
            Caption         =   "Sex "
            Height          =   255
            Left            =   2480
            TabIndex        =   180
            Top             =   735
            Width           =   825
         End
         Begin VB.Label Label48 
            Caption         =   "Allergic"
            Height          =   255
            Left            =   4680
            TabIndex        =   179
            Top             =   1800
            Width           =   600
         End
         Begin VB.Label Label39 
            Caption         =   "Salutation"
            Height          =   255
            Left            =   120
            TabIndex        =   178
            Top             =   225
            Width           =   1140
         End
         Begin VB.Label Label41 
            Caption         =   "IC Num"
            Height          =   255
            Left            =   7200
            TabIndex        =   177
            Top             =   165
            Width           =   1125
         End
         Begin VB.Label Label49 
            Caption         =   "DOB"
            Height          =   255
            Left            =   2480
            TabIndex        =   176
            Top             =   495
            Width           =   1020
         End
         Begin VB.Label Label46 
            Caption         =   "Supp Lmt"
            Height          =   255
            Left            =   2480
            TabIndex        =   175
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label52 
            Caption         =   "R'Ship"
            Height          =   255
            Left            =   120
            TabIndex        =   174
            Top             =   525
            Width           =   1005
         End
         Begin VB.Label Label47 
            Caption         =   "Supplementary Warranty "
            Height          =   255
            Left            =   7200
            TabIndex        =   173
            Top             =   1800
            Width           =   2100
         End
         Begin VB.Label Label6 
            Caption         =   "Bal. Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   172
            Top             =   1020
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1245
         Left            =   -74880
         TabIndex        =   148
         Top             =   360
         Width           =   11145
         Begin VB.Line Line1 
            X1              =   120
            X2              =   10920
            Y1              =   840
            Y2              =   840
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9765
            TabIndex        =   170
            Top             =   525
            Width           =   1200
         End
         Begin VB.Label txtMBMProviderEffCharges 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            TabIndex        =   169
            Top             =   885
            Width           =   1185
         End
         Begin VB.Label txtMBMProviderRedCharges 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            TabIndex        =   168
            Top             =   555
            Width           =   1185
         End
         Begin VB.Label txtMCOFees 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
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
            Left            =   3720
            TabIndex        =   167
            Top             =   180
            Width           =   1185
         End
         Begin VB.Label txtMBMPremium 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
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
            Left            =   1320
            TabIndex        =   166
            Top             =   180
            Width           =   1020
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   165
            Top             =   915
            Width           =   1035
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   164
            Top             =   525
            Width           =   1035
         End
         Begin VB.Label lblMCOFee 
            Caption         =   "MCO Fees"
            Height          =   255
            Left            =   2595
            TabIndex        =   163
            Top             =   180
            Width           =   870
         End
         Begin VB.Label Label85 
            Caption         =   "Effective MCO"
            Height          =   255
            Left            =   2595
            TabIndex        =   162
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label Label83 
            Caption         =   "Effective Prem "
            Height          =   255
            Left            =   120
            TabIndex        =   161
            Top             =   930
            Width           =   1140
         End
         Begin VB.Label txtMCOFeesEff 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Left            =   3720
            TabIndex        =   160
            Top             =   915
            Width           =   1185
         End
         Begin VB.Label txtMBMProviderCharges 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
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
            TabIndex        =   159
            Top             =   180
            Width           =   1185
         End
         Begin VB.Label lblPremium 
            Caption         =   "Premium"
            Height          =   255
            Left            =   120
            TabIndex        =   158
            Top             =   180
            Width           =   960
         End
         Begin VB.Label Label58 
            Caption         =   "Prem Refund"
            Height          =   255
            Left            =   120
            TabIndex        =   157
            Top             =   525
            Width           =   1020
         End
         Begin VB.Label txtMBMMCORed 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Left            =   3720
            TabIndex        =   156
            Top             =   525
            Width           =   1185
         End
         Begin VB.Label Label61 
            Caption         =   "MCO Refund"
            Height          =   255
            Left            =   2595
            TabIndex        =   155
            Top             =   525
            Width           =   1095
         End
         Begin VB.Label Label79 
            Caption         =   "Access Red "
            Height          =   255
            Left            =   5475
            TabIndex        =   154
            Top             =   570
            Width           =   1500
         End
         Begin VB.Label Label77 
            Caption         =   "Acess Eff "
            Height          =   255
            Left            =   5475
            TabIndex        =   153
            Top             =   885
            Width           =   780
         End
         Begin VB.Label Label4 
            Caption         =   "Bal. Limit"
            Height          =   255
            Left            =   8760
            TabIndex        =   152
            Top             =   570
            Width           =   885
         End
         Begin VB.Label txtAnnualLimit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "0"
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9765
            TabIndex        =   151
            Top             =   195
            Width           =   1200
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   10920
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label11 
            Caption         =   "Access Fee"
            Height          =   255
            Left            =   5475
            TabIndex        =   150
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label5 
            Caption         =   "Annual Limit"
            Height          =   255
            Left            =   8760
            TabIndex        =   149
            Top             =   240
            Width           =   870
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4530
         Left            =   120
         TabIndex        =   93
         Top             =   360
         Width           =   11325
         Begin VB.TextBox txtMBMExclusion 
            Height          =   1185
            Left            =   240
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Top             =   3240
            Width           =   4860
         End
         Begin VB.TextBox txtRemarks 
            Height          =   1185
            Left            =   5880
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Top             =   3240
            Width           =   5295
         End
         Begin VB.TextBox txtMBMPolicyYear 
            Enabled         =   0   'False
            Height          =   285
            Left            =   9720
            MaxLength       =   25
            TabIndex        =   35
            Top             =   920
            Width           =   1515
         End
         Begin VB.ComboBox cboINSCode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7200
            Sorted          =   -1  'True
            TabIndex        =   29
            Text            =   "I"
            Top             =   120
            Width           =   1440
         End
         Begin VB.TextBox txtMBMAdd1 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   14
            Top             =   160
            Width           =   3705
         End
         Begin VB.TextBox txtMBMAdd2 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   15
            Top             =   420
            Width           =   3705
         End
         Begin VB.TextBox txtMBMAdd3 
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   16
            Top             =   680
            Width           =   3705
         End
         Begin VB.ComboBox cboCTYCode 
            Height          =   315
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   930
            Width           =   1450
         End
         Begin VB.TextBox txtMBMPostCode 
            Height          =   285
            Left            =   1320
            MaxLength       =   6
            TabIndex        =   17
            Top             =   930
            Width           =   1095
         End
         Begin VB.ComboBox cboSTACode 
            Height          =   315
            Left            =   1320
            TabIndex        =   19
            Top             =   1190
            Width           =   3735
         End
         Begin VB.ComboBox cboMBMSex 
            Height          =   315
            Left            =   1320
            Sorted          =   -1  'True
            TabIndex        =   20
            Text            =   "M"
            Top             =   1475
            Width           =   1125
         End
         Begin VB.ComboBox cboMBMTakeOver 
            Height          =   315
            Left            =   3600
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1475
            Width           =   1455
         End
         Begin VB.ComboBox cboMBMMaritalStatus 
            Height          =   315
            Left            =   3600
            TabIndex        =   23
            Text            =   "M"
            Top             =   1760
            Width           =   1455
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
            TabIndex        =   22
            Top             =   1760
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboNATCode 
            Height          =   315
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   25
            Text            =   "MY - Malaysia"
            Top             =   2040
            Width           =   1455
         End
         Begin VB.ComboBox cboRACCode 
            Height          =   315
            Left            =   1320
            Sorted          =   -1  'True
            TabIndex        =   24
            Text            =   "C - Chinese"
            Top             =   2040
            Width           =   1095
         End
         Begin VB.ComboBox cboBRNCode 
            Height          =   315
            Left            =   1320
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   2330
            Width           =   3735
         End
         Begin VB.TextBox txtMBMAgentCode 
            Height          =   315
            Left            =   3480
            MaxLength       =   15
            TabIndex        =   28
            Top             =   2615
            Width           =   1560
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
            Left            =   7200
            TabIndex        =   30
            Top             =   410
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboPLNCode 
            Enabled         =   0   'False
            Height          =   315
            Left            =   9720
            Sorted          =   -1  'True
            TabIndex        =   31
            Top             =   120
            Width           =   1515
         End
         Begin MSMask.MaskEdBox txtMBMBordxDate 
            Height          =   285
            Left            =   7200
            TabIndex        =   33
            Top             =   650
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   503
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
            Left            =   9720
            TabIndex        =   32
            Top             =   410
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboSRVCode 
            Height          =   315
            Left            =   7200
            Sorted          =   -1  'True
            TabIndex        =   39
            Text            =   "EA - Europe Assistance"
            Top             =   2160
            Width           =   4065
         End
         Begin VB.ComboBox cboMBMRelCode 
            Height          =   315
            Left            =   7200
            TabIndex        =   40
            Text            =   "C - Combined Limit"
            Top             =   2400
            Width           =   1545
         End
         Begin MSMask.MaskEdBox txtMBMDateJoin 
            Height          =   285
            Left            =   7200
            TabIndex        =   34
            Top             =   900
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtMBMEmployeeNo 
            Height          =   300
            Left            =   7200
            MaxLength       =   15
            TabIndex        =   36
            Top             =   1150
            Width           =   4035
         End
         Begin VB.TextBox txtDeptName 
            Height          =   300
            Left            =   7200
            MaxLength       =   50
            TabIndex        =   37
            Top             =   1400
            Width           =   4035
         End
         Begin VB.TextBox txtMBMGroupCompany 
            Height          =   285
            Left            =   7200
            MaxLength       =   150
            TabIndex        =   38
            Top             =   1660
            Width           =   4035
         End
         Begin VB.ComboBox cboMBMRenewal 
            Height          =   315
            ItemData        =   "frmMBMAdjustment.frx":0070
            Left            =   1320
            List            =   "frmMBMAdjustment.frx":0072
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   2615
            Width           =   1050
         End
         Begin VB.ComboBox CboInsMode 
            Height          =   315
            Left            =   9840
            TabIndex        =   41
            Top             =   2400
            Width           =   1425
         End
         Begin VB.Label Label40 
            Caption         =   "Principal Warranty "
            Height          =   255
            Left            =   240
            TabIndex        =   145
            Top             =   3000
            Width           =   1590
         End
         Begin VB.Label Label84 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   5880
            TabIndex        =   144
            Top             =   3000
            Width           =   1590
         End
         Begin VB.Label Label42 
            Caption         =   "Instl mode"
            Height          =   255
            Left            =   9000
            TabIndex        =   136
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label26 
            Caption         =   "Renewal "
            Height          =   240
            Left            =   240
            TabIndex        =   135
            Top             =   2640
            Width           =   1065
         End
         Begin VB.Label Label25 
            Caption         =   "Dept. Name"
            Height          =   255
            Left            =   5880
            TabIndex        =   134
            Top             =   1400
            Width           =   945
         End
         Begin VB.Label Label72 
            Caption         =   "PostCode"
            Height          =   255
            Left            =   240
            TabIndex        =   120
            Top             =   960
            Width           =   1140
         End
         Begin VB.Label Label69 
            Caption         =   "Birthdate"
            Height          =   255
            Left            =   240
            TabIndex        =   119
            Top             =   1800
            Width           =   1005
         End
         Begin VB.Label Label60 
            Caption         =   "State"
            Height          =   255
            Left            =   240
            TabIndex        =   118
            Top             =   1240
            Width           =   465
         End
         Begin VB.Label Label57 
            Caption         =   "City/Town"
            Height          =   255
            Left            =   2520
            TabIndex        =   117
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label54 
            Caption         =   "Address"
            Height          =   255
            Left            =   240
            TabIndex        =   116
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label56 
            Caption         =   "Sex "
            Height          =   255
            Left            =   240
            TabIndex        =   115
            Top             =   1505
            Width           =   825
         End
         Begin VB.Label Label9 
            Caption         =   "Pol. Year"
            Height          =   255
            Left            =   8760
            TabIndex        =   114
            Top             =   980
            Width           =   690
         End
         Begin VB.Label Label63 
            Caption         =   "TakeOver"
            Height          =   255
            Left            =   2520
            TabIndex        =   113
            Top             =   1560
            Width           =   825
         End
         Begin VB.Label Label12 
            Caption         =   "Branch"
            Height          =   255
            Left            =   240
            TabIndex        =   112
            Top             =   2400
            Width           =   960
         End
         Begin VB.Label Label70 
            Caption         =   "Race"
            Height          =   255
            Left            =   240
            TabIndex        =   111
            Top             =   2080
            Width           =   420
         End
         Begin VB.Label Label71 
            Caption         =   "Nationality"
            Height          =   255
            Left            =   2520
            TabIndex        =   110
            Top             =   2080
            Width           =   960
         End
         Begin VB.Label Label27 
            Caption         =   "Marital Status"
            Height          =   255
            Left            =   2520
            TabIndex        =   109
            Top             =   1820
            Width           =   1095
         End
         Begin VB.Label Label30 
            Caption         =   "Supp Type"
            Height          =   255
            Left            =   5880
            TabIndex        =   108
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Label Label29 
            Caption         =   "Age Band"
            Height          =   255
            Left            =   5880
            TabIndex        =   106
            Top             =   1940
            Width           =   735
         End
         Begin VB.Label Label20 
            Caption         =   "Ins Type"
            Height          =   255
            Left            =   5880
            TabIndex        =   105
            Top             =   120
            Width           =   960
         End
         Begin VB.Label Label21 
            Caption         =   "Plan Code"
            Height          =   255
            Left            =   8760
            TabIndex        =   104
            Top             =   180
            Width           =   825
         End
         Begin VB.Label Label10 
            Caption         =   "Date Join"
            Height          =   255
            Left            =   5880
            TabIndex        =   103
            Top             =   900
            Width           =   780
         End
         Begin VB.Label Label19 
            Caption         =   "E'yee Num"
            Height          =   255
            Left            =   5880
            TabIndex        =   102
            Top             =   1150
            Width           =   1545
         End
         Begin VB.Label Label18 
            Caption         =   "Grp. Company"
            Height          =   255
            Left            =   5880
            TabIndex        =   101
            Top             =   1660
            Width           =   1185
         End
         Begin VB.Label Label14 
            Caption         =   "Agent Code"
            Height          =   255
            Left            =   2520
            TabIndex        =   100
            Top             =   2670
            Width           =   1140
         End
         Begin VB.Label Label22 
            Caption         =   "Serv. Provider"
            Height          =   255
            Left            =   5880
            TabIndex        =   99
            Top             =   2160
            Width           =   1485
         End
         Begin VB.Label Label17 
            Caption         =   "Bordx. Date"
            Height          =   255
            Left            =   5880
            TabIndex        =   98
            Top             =   650
            Width           =   1275
         End
         Begin VB.Label Label8 
            Caption         =   "Batch Num"
            Height          =   255
            Left            =   8760
            TabIndex        =   96
            Top             =   705
            Width           =   885
         End
         Begin VB.Label Label15 
            Caption         =   "Eff Date"
            Height          =   255
            Left            =   5880
            TabIndex        =   95
            Top             =   410
            Width           =   1275
         End
         Begin VB.Label Label16 
            Caption         =   "Exp Date"
            Height          =   255
            Left            =   8760
            TabIndex        =   94
            Top             =   425
            Width           =   915
         End
         Begin VB.Label txtMBMBatchNo 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   9720
            TabIndex        =   97
            Top             =   690
            Width           =   1515
         End
         Begin VB.Label txtAGECode 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   7200
            TabIndex        =   107
            Top             =   1940
            Width           =   4035
         End
      End
      Begin MSComctlLib.ListView lstCovAdd 
         Height          =   1305
         Left            =   -74880
         TabIndex        =   146
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
         NumItems        =   40
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
            Text            =   "Blood Type"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Allergic"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Exclusion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Status"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Annual Limit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Occupation"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Work Nature"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Smoking"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Alcohol"
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
            Text            =   "PrevPLNStatus"
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
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Text            =   "PayorPremNet"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Text            =   "PayorMCOGross"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Text            =   "PayorMCONet"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   38
            Text            =   "RenewalDisc"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   39
            Text            =   "PolDisc"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstAdjustmentHis 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   147
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
   End
   Begin VB.Frame Frame7 
      Height          =   525
      Left            =   120
      TabIndex        =   127
      Top             =   7440
      Width           =   11655
      Begin VB.CommandButton CmdMBMOthers 
         Caption         =   "Others Details"
         Enabled         =   0   'False
         Height          =   315
         Left            =   7440
         TabIndex        =   78
         Top             =   120
         Width           =   1200
      End
      Begin VB.TextBox txtcount 
         Height          =   285
         Left            =   6240
         TabIndex        =   142
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   10500
         TabIndex        =   81
         Top             =   120
         Width           =   960
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9555
         TabIndex        =   80
         Top             =   120
         Width           =   960
      End
      Begin VB.CommandButton cmdEditMBM 
         Caption         =   "E&dit Status"
         Enabled         =   0   'False
         Height          =   315
         Left            =   8640
         TabIndex        =   79
         Top             =   120
         Width           =   960
      End
      Begin VB.TextBox txtMBMAdultChild 
         Height          =   285
         Left            =   120
         TabIndex        =   130
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtLastEndStatus 
         Height          =   285
         Left            =   840
         TabIndex        =   129
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtLastEdtDate 
         Height          =   285
         Left            =   1440
         TabIndex        =   128
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame16 
      Height          =   540
      Left            =   120
      TabIndex        =   86
      Top             =   280
      Width           =   11775
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&Pre"
         Height          =   285
         Left            =   4680
         TabIndex        =   133
         Top             =   160
         Width           =   810
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "&Next"
         Height          =   285
         Left            =   5520
         TabIndex        =   132
         Top             =   160
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
         Left            =   6360
         TabIndex        =   87
         Top             =   160
         Width           =   810
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Go"
         Height          =   285
         Left            =   3840
         TabIndex        =   82
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
      Begin VB.Label Label76 
         Caption         =   "Mem Num."
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   160
         Width           =   1275
      End
   End
   Begin VB.Frame Frame10 
      Height          =   850
      Left            =   120
      TabIndex        =   89
      Top             =   760
      Width           =   11730
      Begin VB.ComboBox cboExpStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMBMAdjustment.frx":0074
         Left            =   8160
         List            =   "frmMBMAdjustment.frx":0076
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   180
         Width           =   1080
      End
      Begin VB.ComboBox cboEndorseStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMBMAdjustment.frx":0078
         Left            =   1440
         List            =   "frmMBMAdjustment.frx":007A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   1935
      End
      Begin VB.ComboBox cboStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMBMAdjustment.frx":007C
         Left            =   5280
         List            =   "frmMBMAdjustment.frx":007E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Width           =   1680
      End
      Begin VB.ComboBox cboDataStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMBMAdjustment.frx":0080
         Left            =   8160
         List            =   "frmMBMAdjustment.frx":0082
         TabIndex        =   4
         Text            =   "A - Active"
         Top             =   435
         Width           =   1065
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   460
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
         Left            =   5280
         TabIndex        =   6
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
      Begin VB.TextBox txtBatchNo 
         Height          =   285
         Left            =   10440
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "1"
         Top             =   225
         Width           =   1200
      End
      Begin VB.Label Label89 
         Caption         =   "Exp. Status"
         Height          =   255
         Left            =   7200
         TabIndex        =   143
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label59 
         Caption         =   "Bat. Num"
         Height          =   255
         Left            =   9360
         TabIndex        =   141
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Endt Bordx Date"
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
         Left            =   3840
         TabIndex        =   140
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblAdj 
         Caption         =   "Endt Eff Date"
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
         TabIndex        =   139
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label txtLastBatch 
         BackColor       =   &H80000009&
         Height          =   240
         Left            =   10080
         TabIndex        =   138
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label62 
         Caption         =   "Data Status"
         Height          =   255
         Left            =   7200
         TabIndex        =   137
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label24 
         Caption         =   "Endt. Status"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label55 
         Caption         =   "Pol. Status"
         Height          =   255
         Left            =   3840
         TabIndex        =   90
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   121
      Top             =   1560
      Width           =   11655
      Begin VB.ComboBox cboMBMSalutation 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Text            =   "MR"
         Top             =   165
         Width           =   690
      End
      Begin VB.TextBox txtMBMIcBcPP 
         Height          =   285
         Left            =   8760
         MaxLength       =   30
         TabIndex        =   10
         Top             =   165
         Width           =   2535
      End
      Begin VB.TextBox txtMBMName 
         Height          =   330
         Left            =   2760
         MaxLength       =   60
         TabIndex        =   9
         Top             =   165
         Width           =   4575
      End
      Begin VB.ComboBox cboHLTCode 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8760
         Sorted          =   -1  'True
         TabIndex        =   13
         Text            =   "S"
         Top             =   420
         Width           =   2530
      End
      Begin VB.TextBox txtMBMPolNo 
         Height          =   330
         Left            =   2760
         MaxLength       =   25
         TabIndex        =   12
         Top             =   460
         Width           =   4575
      End
      Begin VB.ComboBox cboPAYCode 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   1  'Simple Combo
         TabIndex        =   11
         Top             =   440
         Width           =   675
      End
      Begin VB.Label Label80 
         Caption         =   "Mem Name"
         Height          =   255
         Left            =   1800
         TabIndex        =   131
         Top             =   210
         Width           =   1140
      End
      Begin VB.Label Label73 
         Caption         =   "Policy Num"
         Height          =   255
         Left            =   1800
         TabIndex        =   126
         Top             =   495
         Width           =   900
      End
      Begin VB.Label Label31 
         Caption         =   "Salutation"
         Height          =   255
         Left            =   120
         TabIndex        =   125
         Top             =   195
         Width           =   1140
      End
      Begin VB.Label Label64 
         Caption         =   "IC. Num"
         Height          =   255
         Left            =   7560
         TabIndex        =   124
         Top             =   165
         Width           =   855
      End
      Begin VB.Label Label75 
         Caption         =   "Health Type"
         Height          =   255
         Left            =   7560
         TabIndex        =   123
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label74 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   122
         Top             =   510
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmMBMAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim startTrans As Boolean
Dim counter As Integer
Dim bExit As Boolean
Dim bTabDone As Boolean
Public MBMPayorEXPDate As Integer
Dim tempSPremium As String
Dim tempSAnnLmt As String
Dim ScanCode As String
Dim txtTempProvider As Double
Dim txtTempPrem As Double
Dim txtTempMCOFees As Double
Dim txtTempExpDate  As String
Dim txtTempEffDate As String
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset
Dim Disc As Boolean
Public PolDics As String
Public RenDisc As String


Private Sub cboCovAdultChild_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboCovAdultChild.Text, cboCovAdultChild.SelStart) + Chr(KeyAscii) + Mid(cboCovAdultChild.Text, cboCovAdultChild.SelStart + cboCovAdultChild.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboCovAdultChild.ListCount - 1
        
        If NewText = LCase(Left(cboCovAdultChild.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCovAdultChild.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboCovAdultChild.Text = ValidValue
               cboCovAdultChild.SelStart = 0
               cboCovAdultChild.SelLength = Len(ValidValue)
             End If
        End If
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

Private Sub cboHLTCode_LostFocus()
    Clipboard.Clear
End Sub

Private Sub cboMBMRelCode_GotFocus()
    If Mid(cboMBMRelCode, 1, 1) = "S" Then
        Call SearchSuppLmt
    End If
End Sub

Private Sub cboMBMRelCode_LostFocus()
    Clipboard.Clear
End Sub

Private Sub CboInsMode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    Clipboard.Clear
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(CboInsMode.Text, CboInsMode.SelStart) + Chr(KeyAscii) + Mid(CboInsMode.Text, CboInsMode.SelStart + CboInsMode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To CboInsMode.ListCount - 1
            
            If NewText = LCase(Left(CboInsMode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = CboInsMode.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   CboInsMode.Text = ValidValue
                   CboInsMode.SelStart = 0
                   CboInsMode.SelLength = Len(ValidValue)
                 End If
            End If
End Sub

Private Sub cboMBMRenewal_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboMBMRenewal.Text, cboMBMRenewal.SelStart) + Chr(KeyAscii) + Mid(cboMBMRenewal.Text, cboMBMRenewal.SelStart + cboMBMRenewal.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboMBMRenewal.ListCount - 1
            
            If NewText = LCase(Left(cboMBMRenewal.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboMBMRenewal.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   cboMBMRenewal.Text = ValidValue
                   cboMBMRenewal.SelStart = 0
                   cboMBMRenewal.SelLength = Len(ValidValue)
                 End If
            End If
End Sub


Private Sub cboMBMRenewal_LostFocus()
    Clipboard.Clear
End Sub

Private Sub cboMBMSex_LostFocus()
    Clipboard.Clear
End Sub

'Private Sub cboMBMSmoking_LostFocus()
    'Clipboard.Clear
'End Sub

Private Sub cboNATCode_LostFocus()
    Clipboard.Clear
End Sub
Private Sub cboPLNCode_LostFocus()
    Clipboard.Clear
End Sub

Private Sub cboRACCode_LostFocus()
    Clipboard.Clear
End Sub

Private Sub CmdMBMOthers_Click()
    FrmMBMOthers.Source = 1
    If FrmMBMOthers.Visible = False Then
        Load FrmMBMOthers
    End If
    FrmMBMOthers.show
    'FrmMBMOthers.cmdClose_Click
    
    FrmMBMOthers.Frame5.Enabled = True
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
    FrmMBMOthers.txtGENWPEnd.Enabled = False
    FrmMBMOthers.txtPREWPEnd.Enabled = False
    FrmMBMOthers.txtSPEWPEnd.Enabled = False
    cmdSave.Enabled = False
    'Call TabEnabled(False)
    counter = 0
    cmdRemove.Enabled = False
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    bExit = False
    bTabDone = False
    SSTab1.Tab = 0
    cboINSCode.Enabled = False
    cboPLNCode.Enabled = False
    txtTempProvider = 0
    txtTempMCOFees = 0
    txtTempPrem = 0
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
    'Dim Work As New ADODB.Recordset
    Dim RAC As New ADODB.Recordset
    Dim NAT As New ADODB.Recordset
    Dim REL As New ADODB.Recordset
    'Dim LGN As New ADODB.Recordset
    'Dim CTY As New ADODB.Recordset
    Dim STA As New ADODB.Recordset
    'Dim BRN As New ADODB.Recordset
    'Dim PAY As New ADODB.Recordset
    'Dim OCP As New ADODB.Recordset
    Dim HLT As New ADODB.Recordset
    Dim PLN As New ADODB.Recordset
    Dim INS As New ADODB.Recordset
    Dim sqlPAY As String
    
    SRV.Open "select SRVCode, SRVName from  SRV", medixmasterconn, adOpenKeyset, adLockOptimistic
    RAC.Open "select RACCode, RACName from   RAC", medixmasterconn, adOpenKeyset, adLockOptimistic
    NAT.Open "select NATCode, NATName from  NAT", medixmasterconn, adOpenKeyset, adLockOptimistic
    REL.Open "select RELCode, RELDescription from  REL", medixmasterconn, adOpenKeyset, adLockOptimistic
    'OCP.Open "select OCPCode, OCPDescription from OCP", medixmasterconn, adOpenKeyset, adLockOptimistic
    'LGN.Open "select LGNCode, LGNDescription from LGN", medixmasterconn, adOpenKeyset, adLockOptimistic
    'CTY.Open "select ctyCode , CTYDescription from CTY", medixmasterconn, adOpenKeyset, adLockOptimistic
    STA.Open "select STACOde, STADescription from  STA", medixmasterconn, adOpenKeyset, adLockOptimistic
    'BRN.Open "Select BRNCOde, BRNNAme from BRN", medixmasterconn, adOpenKeyset, adLockOptimistic
    'Work.Open "Select  WorkID, WorkDescription from WorkNature ", medixmasterconn, adOpenKeyset, adLockOptimistic
    sqlPAY = "select PAYCode, PAYCompanyName from PAY "
    'PAY.Open sqlPAY, medixmasterconn, adOpenKeyset, adLockOptimistic
    HLT.Open "select HLTCode, HLTDescription from HLT", medixmasterconn, adOpenKeyset, adLockOptimistic
    'PLN.Open "select PLNCode, PLNDescription from PLN", medixmasterconn, adOpenKeyset, adLockOptimistic
    'INS.Open "select INSCode, INSDescription from INS", medixmasterconn, adOpenKeyset, adLockOptimistic

    'Do Until PAY.EOF
    '    cboPAYCode.AddItem ChkStr(PAY!PAYCode) '& " - " & ChkStr(PAY!PAYCompanyName)
    '    PAY.MoveNext
    'Loop
    
    'Do Until LGN.EOF
        'cboLGNCode.AddItem ChkStr(LGN!LGNCode) & " - " & LGN!LGNDescription
        'LGN.MoveNext
    'Loop

    Do Until SRV.EOF
        cboSRVCode.AddItem ChkStr(SRV!SRVCode) '& " - " & ChkStr(SRV!SRVname)
        SRV.MoveNext
    Loop

    Do Until RAC.EOF
        cboRACCode.AddItem ChkStr(RAC!RACCode) '& " - " & ChkStr(RAC!RACName)
        RAC.MoveNext
    Loop

    Do Until NAT.EOF
        cboNATCode.AddItem ChkStr(NAT!NATCode) '& " - " & ChkStr(NAT!NATName)
        NAT.MoveNext
    Loop

    Do Until REL.EOF
        cboCovRelationship.AddItem ChkStr(REL!RELCODE) '& " - " & REL!RELdescription
        REL.MoveNext
    Loop

    'Do Until OCP.EOF
        'cboMBMOccupation.AddItem ChkStr(OCP!OCPCode) & " - " & OCP!OCPDescription
        'OCP.MoveNext
    'Loop

    'Do Until Work.EOF
        'txtMBMWorkNature.AddItem Work!WorkID & " - " & Work!WorkDescription
        'Work.MoveNext
    'Loop
    
    'Do Until CTY.EOF
    '    cboBRNCode.AddItem CTY!CTYDescription
    '    cboCTYCode.AddItem CTY!CTYDescription
    '    CTY.MoveNext
    'Loop

    Do Until STA.EOF
        cboSTACode.AddItem STA!STADescription
        STA.MoveNext
    Loop

  '  Do Until BRN.EOF
   '     cboBRNCode.AddItem ChkStr(BRN!BRNName)
    '    BRN.MoveNext
    'Loop

    Do Until HLT.EOF
        cboHLTCode.AddItem HLT!HLTCode '& " - " & HLT!HLTDescription
        HLT.MoveNext
    Loop

    'Do Until PLN.EOF
    '    cboPLNCode.AddItem PLN!PLNCode '& " - " & PLN!PLNDescription
    '    PLN.MoveNext
    'Loop
    
    'Do Until INS.EOF
    '    cboINSCode.AddItem INS!INSCode '& " - " & INS!INSDescription
    '    INS.MoveNext
    'Loop
    cboMBMRelCode.Clear
    cboMBMRelCode.AddItem "C - Combined Limit"
    cboMBMRelCode.AddItem "S - Separate Limit"


    cboMBMSex.AddItem "M"
    cboMBMSex.AddItem "F"
    
    cboCovSex.AddItem "M"
    cboCovSex.AddItem "F"
    
    cboMBMMaritalStatus.AddItem "S"
    cboMBMMaritalStatus.AddItem "M"
    
    'cboMBMBlood.AddItem "A"
    'cboMBMBlood.AddItem "B"
    'cboMBMBlood.AddItem "AB"
    'cboMBMBlood.AddItem "O"
    
    'cboAlcohol.AddItem "Y"
    'cboAlcohol.AddItem "N"
    'cboSmoking.AddItem "Y"
    'cboSmoking.AddItem "N"

    
    'cboCovBlood.AddItem "A"
    'cboCovBlood.AddItem "B"
    'cboCovBlood.AddItem "AB"
    'cboCovBlood.AddItem "O"
    
    cboCovAdultChild.AddItem "A"
    cboCovAdultChild.AddItem "C"
    
    cboMBMSalutation.AddItem "MR"
    cboMBMSalutation.AddItem "MS"
    cboMBMSalutation.AddItem "MRS"
    
    cboCovSalutation.AddItem "MR"
    cboCovSalutation.AddItem "MS"
    cboCovSalutation.AddItem "MRS"
    
    'cboMBMSmoking.AddItem "Y"
    'cboMBMSmoking.AddItem "N"
    
    FrmMBMOthers.cboStaffDisc.AddItem "N - No"
    FrmMBMOthers.cboStaffDisc.AddItem "Y - Yes"
    
    cboSuppStaffDisc.AddItem "N - No"
    cboSuppStaffDisc.AddItem "Y - Yes"
    
    
    cboEndorseStatus.AddItem "A" & " - " & "Amendment"
    cboEndorseStatus.AddItem "C" & " - " & "Cancellation"
    cboEndorseStatus.AddItem "S" & " - " & "Add Supplementary"
    cboEndorseStatus.AddItem "T" & " - " & "Delete Supplementary"
    cboEndorseStatus.AddItem "X" & " - " & "Suspend Membership"
    cboEndorseStatus.AddItem "Y" & " - " & "Re-Activate Membership"
    cboEndorseStatus.AddItem "I" & " - " & "Insert New Member"
    cboEndorseStatus.AddItem "D" & " - " & "Delete Policy"
    
    
    cboStatus.AddItem "A" & " - " & "Active"
    cboStatus.AddItem "C" & " - " & "Cancel"
    cboStatus.AddItem "D" & " - " & "Deleted"
    cboStatus.AddItem "S" & " - " & "Suspended"
    'cboStatus.AddItem "E" & " - " & "Expired"
                
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
  
    'cboMBMAlcohol.AddItem "Y"
    'cboMBMAlcohol.AddItem "N"
    
    CboInsMode.AddItem "A - Annually"
    CboInsMode.AddItem "Q - Quarterly"
    CboInsMode.AddItem "S - Semi Annual"
    CboInsMode.AddItem "M - Monthly"

    FrmMBMOthers.cboBTBG.AddItem "Y - Yes"
    FrmMBMOthers.cboBTBG.AddItem "N - No"

    'PAY.Close
    'Set PAY = Nothing
    SRV.Close
    Set SRV = Nothing
    'OCP.Close
    'Set OCP = Nothing
    RAC.Close
    Set RAC = Nothing
    NAT.Close
    Set NAT = Nothing
    REL.Close
    Set REL = Nothing
    'LGN.Close
    'Set LGN = Nothing
    'CTY.Close
    'Set CTY = Nothing
    STA.Close
    Set STA = Nothing
    'BRN.Close
    'Set BRN = Nothing
    HLT.Close
    Set HLT = Nothing
    'INS.Close
    'Set INS = Nothing
    'PLN.Close
    'Set PLN = Nothing
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
             Set ADJList = lstAdjustmentHis.listitems.Add(, , ADJ!MBMAENDtBordxDate)
        Else
        Set ADJList = lstAdjustmentHis.listitems.Add(, , " ")
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
    If Trim(txtMBMPayorEffDate) <> "__/__/____" _
        And Trim(txtMBMBirthDate) <> "__/__/____" _
        And Len(Trim(cboHLTCode)) <> 0 Then
        If IsDate(txtMBMBirthDate) Then
            SelectAge = GetAge(CDate(txtMBMPayorEffDate), CDate(txtMBMBirthDate))
            txtAGECode = GetAgeBand(Trim(Mid(cboHLTCode, 1, 1)), SelectAge)
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
    Dim tempPremium, tempMCOFee As Single
    Dim Confirm1, Confirm2
    If Len(Trim(cboPAYCode)) <> 0 Then
        tempPremium = Format(GetPremium(Trim(Mid(cboINSCode, 1, 1)), Trim(Mid(cboPAYCode, 1, 2)), Trim(Mid(txtAGECode, 1, 2)), Trim(Mid(cboPLNCode, 1, 4)), Trim(Mid(cboHLTCode, 1, 1)), txtMBMPayorEffDate), "#,##0.00")
    End If
    
    strsqlPLN = "select * from DiscPlan where PLNCode = '" & Trim(Mid(cboPLNCode, 1, 4)) & "'"
    PLN.Open strsqlPLN, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    If PLN.recordCount Then
        If PolDics > 0 Then
            tempPremium = tempPremium * (100 - PolDics) / 100
        End If
        If RenDisc > 0 Then
            tempPremium = tempPremium * (100 - RenDisc) / 100
        End If
    End If
    PLN.Close
    Set PLN = Nothing

    txtTempPrem = tempPremium
    
    tempMCOFee = Format(GetMCO(Trim(Mid(cboINSCode, 1, 1)), Trim(Mid(cboHLTCode, 1, 1)), txtMBMAdultChild, txtMBMPayorEffDate, Trim(Mid(cboPLNCode, 1, 4))), "#,##0.00")
    txtTempMCOFees = tempMCOFee
    If Format(txtMBMPremium, "#,##0.00") <> tempPremium Then
        Confirm1 = MsgBox("Do you want to change member's premium amount?", vbCritical + vbYesNo)
        If Confirm1 = vbYes Then
           txtMBMPremium = tempPremium
           
        End If
    End If
    If Format(txtMCOFees, "#,##0.00") <> tempMCOFee Then
        Confirm2 = MsgBox(" Do you want to change member's MCO Fee?", vbCritical + vbYesNo)
        If Confirm2 = vbYes Then
            txtMCOFees = tempMCOFee
        End If
    End If
End Sub

Public Sub ProviderCallFees()
    Dim GetProString As String
    
    If Trim(cboINSCode) <> "" And Trim(cboSRVCode) <> "" And Trim(cboHLTCode) <> "" Then
        txtMBMProviderCharges = ""
        'GetProString = GetProviderCharges(Trim(Mid(cboSRVCode, 1, InStr(1, cboSRVCode, "-") - 1)), Trim(Mid(cboINSCode, 1, InStr(1, cboINSCode, "-") - 1)), Trim(Mid(txtHLTCode, 1, InStr(1, txtHLTCode, "-") - 1)), txtMBMAdultChild)
        If InStr(1, cboSRVCode, "-") <> 0 Then
            GetProString = GetProviderCharges(Trim(Mid(cboSRVCode, 1, InStr(1, cboSRVCode, "-") - 1)), Trim(cboINSCode), Trim(cboHLTCode), txtMBMAdultChild)
            txtMBMProviderCharges = GetProString
            txtTempProvider = GetProString
        Else
            GetProString = GetProviderCharges(Trim(cboSRVCode), Trim(cboINSCode), Trim(cboHLTCode), txtMBMAdultChild)
            txtMBMProviderCharges = GetProString
            txtTempProvider = GetProString
        End If
        
        'If InStr(1, GetProString, "-") - 1 > 0 Then
        '    txtMBMProviderCharges = Format(Mid(GetProString, 1, InStr(1, GetProString, "-") - 1), "#,##0.00")
        'End If
        'txtMBMCaseCharges = Format(Mid(GetProString, InStr(1, GetProString, "-") + 1), "#,##0.00")
    Else
        txtMBMProviderCharges = 0
     '   txtMBMCaseCharges = 0
    End If
            
End Sub

Public Sub CallLimit()
    txtAnnualLimit = ""
    If Len(Trim(cboPAYCode)) <> 0 And cboINSCode <> "" And txtAGECode <> "" And cboHLTCode <> "" And txtMBMPayorEffDate <> "__/__/____" Then
        txtAnnualLimit = Format(GetAnnualLimit(cboINSCode, cboPLNCode, cboHLTCode, txtMBMPayorEffDate), "#,##0.00")
    End If

    
    'txtAnnualLimit = ""
    'If Len(Trim(cboPLNCode)) And txtMBMPayorEffDate <> "__/__/____" Then
    '    txtAnnualLimit = Format(GetAnnualLimit(Mid(cboINSCode, 1, 1), Mid(cboPLNCode, 1, 2), Mid(cboHLTCode, 1, 1), txtMBMPayorEffDate), "#,##0.00")
    'End If
End Sub


' ######################### END CALCULATIOn###########
' ######################### START CHECKING ###########


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

Private Sub cboEndorseStatus_Change()
    If cboEndorseStatus.ListIndex = 0 Then
        If Mid(cboStatus, 1, 1) = "C" Or Mid(cboStatus, 1, 1) = "D" Then
            cboStatus.Enabled = False
        Else
            Call EnableSupplementary(True)
            Frame1.Enabled = True
            txtDate.Enabled = True
            lstCovAdd.Enabled = True
            txtDate.mask = ""
            txtDate = ""
            txtDate.mask = "##/##/####"
            txtEndtBordxDate.Enabled = True
            Call TabEnabled(True)
            cmdAdd.Enabled = False
            cmdRemove.Enabled = False
            cmdUpdate.Enabled = True
            cmdClear.Enabled = True
            cboStatus.Enabled = True
            'Frame1.Enabled = False
            'Frame4.Enabled = False
            'Frame5.Enabled = False
            Frame6.Enabled = True
        End If
        
    ElseIf cboEndorseStatus.ListIndex = 1 Then

            cmdExit.Enabled = True
            cmdAdd.Enabled = False
            cmdRemove.Enabled = False
            lstCovAdd.Enabled = False
            txtDate.Enabled = True
            txtEndtBordxDate.Enabled = True
            cmdSearchMBM.Enabled = True
            cmdEditMBM.Enabled = True
            Frame1.Enabled = False
            Frame4.Enabled = False
            'Frame5.Enabled = False
            Frame6.Enabled = False
    ElseIf cboEndorseStatus.ListIndex = 2 Then
        If Mid(cboStatus, 1, 1) = "C" Or Mid(cboStatus, 1, 1) = "D" Then
            cboStatus.Enabled = False
        ElseIf Mid(cboINSCode, 1, 1) = "I" Or Mid(cboINSCode, 1, 1) = "G" Then
            SSTab1.Tab = 0
        'End If
        Else
            'lblAdj.Caption = ""
            'lblAdj.Caption = "Deletion Date"
            'Call EnableEntries(False)
            Call EnableSupplementary(True)
            cmdExit.Enabled = True
            'cmdEditMBM.Enabled = False
            cmdClear.Enabled = True
            cmdAdd.Enabled = True
            cmdRemove.Enabled = False
            lstCovAdd.Enabled = True
            cboEndorseStatus.Enabled = False
            'Call TabEnabled(True)
            SSTab1.Tab = 2
            'txtDate =
            txtDate.Enabled = True
            'Frame5.Enabled = False
            txtDate.SetFocus
            txtEndtBordxDate.Enabled = True
            cmdEditMBM.Enabled = True
            Call ClearCoveredEntry
        End If
    ElseIf cboEndorseStatus.ListIndex = 4 Then
        If Mid(cboStatus, 1, 1) = "C" Or Mid(cboStatus, 1, 1) = "D" Then
            cboStatus.Enabled = False
        Else
            'lblAdj.Caption = ""
            'lblAdj.Caption = "Suspension Date"
            Call EnableEntries(False)
            Call EnableSupplementary(True)
            cmdExit.Enabled = True
            cmdClear.Enabled = True
            'cmdEditMBM.Enabled = True
            cmdAdd.Enabled = False
            cmdRemove.Enabled = True
            lstCovAdd.Enabled = True
            'Call TabEnabled(False)
            'SSTab1.Tab = 2
            'txtDate = Date
            txtDate.Enabled = True
            txtEndtBordxDate.Enabled = True
            'txtDate.SetFocus
            cmdEditMBM.Enabled = True
        End If
        
    ElseIf cboEndorseStatus.ListIndex = 3 Then
        'Call EnableSupplementary(False)
        'lblAdj.Caption = ""
        'lblAdj.Caption = "Adjustment Date"
        lstCovAdd.Enabled = True
        cmdRemove.Enabled = True
        cmdClear.Enabled = True
        SSTab1.Tab = 2
        
    ElseIf cboEndorseStatus.ListIndex = 6 Then
        'lblAdj.Caption = ""
        'lblAdj.Caption = "Insertion Date"
         Call ClearEntries
         cboINSCode.Enabled = True
         cboPLNCode.Enabled = True
         txtMBMBordxDate.Enabled = True
        lstCovAdd.Enabled = True
        cmdRemove.Enabled = True
        cmdClear.Enabled = True
        Frame1.Enabled = True
        Frame4.Enabled = True
        cboStatus.ListIndex = 0
        txtDate.Enabled = True
        txtEndtBordxDate.Enabled = True
        txtDate.SetFocus
    ElseIf Mid(cboEndorseStatus, 1, 1) = "D" Then
            Frame1.Enabled = False
            Frame4.Enabled = False
            'Frame5.Enabled = False
            Frame6.Enabled = False
            txtDate.mask = ""
            txtDate = ""
            txtDate.mask = "##/##/####"

        'Call DeletePrincipals
    Else ' Reactivate
        'lblAdj.Caption = ""
        'lblAdj.Caption = "Reactivation Date"
        Call TabEnabled(True)
        'txtDate.Mask = "__/__/____"
        txtDate.Enabled = True
        txtEndtBordxDate.Enabled = True
        'txtDate.SetFocus
        cmdEditMBM.Enabled = True
        txtDate.mask = ""
        txtDate = ""
        txtDate.mask = "##/##/####"


    End If
    cboEndorseStatus.Enabled = False
'    CmdSave.Enabled = True
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
    
End Sub

Private Sub cboEndorseStatus_Click()
    If cboEndorseStatus.ListIndex = 0 Then
        If Mid(cboStatus, 1, 1) = "C" Or Mid(cboStatus, 1, 1) = "D" Then
            cboStatus.Enabled = False
        'txtMBMEndorseStatus.Enabled = True
        Else
            'lblAdj.Caption = ""
            'lblAdj.Caption = "Adjustment Date"
            'Call EnableEntries(True)
            Call EnableSupplementary(True)
            Frame1.Enabled = True
            txtDate.Enabled = True
            lstCovAdd.Enabled = True
            txtDate.mask = ""
            txtDate = ""
            txtDate.mask = "##/##/####"
            txtEndtBordxDate.Enabled = True
            'txtDate = Date
            Call TabEnabled(True)
            cmdAdd.Enabled = False
            cmdRemove.Enabled = False
            cmdUpdate.Enabled = True
            cmdClear.Enabled = True
            cboStatus.Enabled = True
            'Frame1.Enabled = False
            'Frame4.Enabled = False
            'Frame5.Enabled = True
            'Frame6.Enabled = True

        End If
        
    ElseIf cboEndorseStatus.ListIndex = 1 Then
            'lblAdj.Caption = ""
            'lblAdj.Caption = "Cancelation Date"
            'Call EnableEntries(False)
            'Call EnableSupplementary(False)
            
            cmdExit.Enabled = True
            cmdAdd.Enabled = False
            cmdRemove.Enabled = False
            lstCovAdd.Enabled = False
            'txtDate.Mask = "__/__/____"
            'txtDate = Date
            txtDate.Enabled = True
            txtEndtBordxDate.Enabled = True
            'txtDate.SetFocus
            'If txtMBMNumber <> "" Then
            '    Frame1.Enabled = False
                'Frame4.Enabled = False
            'End If
            cmdSearchMBM.Enabled = True
            'cmdMBMPolicy.Enabled = True
            'cmdSearchIC.Enabled = True
            'cmdSearchName.Enabled = True
            cmdEditMBM.Enabled = True
            Frame1.Enabled = False
            Frame4.Enabled = False
            'Frame5.Enabled = False
            'Frame6.Enabled = False
    ElseIf cboEndorseStatus.ListIndex = 2 Then
        If Mid(cboStatus, 1, 1) = "C" Or Mid(cboStatus, 1, 1) = "D" Then
            cboStatus.Enabled = False
        ElseIf Mid(cboINSCode, 1, 1) = "I" Or Mid(cboINSCode, 1, 1) = "G" Then
            SSTab1.Tab = 0
        'End If
        Else
            'lblAdj.Caption = ""
            'lblAdj.Caption = "Deletion Date"
            'Call EnableEntries(False)
            Call EnableSupplementary(True)
            cmdExit.Enabled = True
            'cmdEditMBM.Enabled = False
            'cmdClear.Enabled = True
            cmdAdd.Enabled = True
            cmdRemove.Enabled = False
            lstCovAdd.Enabled = True
            cboEndorseStatus.Enabled = False
            'Call TabEnabled(True)
            SSTab1.Tab = 2
            'txtDate =
            txtDate.Enabled = True
            'Frame5.Enabled = False
            'Frame6.Enabled = True
            txtDate.SetFocus
            txtEndtBordxDate.Enabled = True
            cmdEditMBM.Enabled = True
            Call ClearCoveredEntry
        End If
    ElseIf cboEndorseStatus.ListIndex = 4 Then
        If Mid(cboStatus, 1, 1) = "C" Or Mid(cboStatus, 1, 1) = "D" Then
            cboStatus.Enabled = False
        Else
            'lblAdj.Caption = ""
            'lblAdj.Caption = "Suspension Date"
            Call EnableEntries(False)
            Call EnableSupplementary(True)
            cmdExit.Enabled = True
            'cmdClear.Enabled = True
            'cmdEditMBM.Enabled = True
            cmdAdd.Enabled = False
            cmdRemove.Enabled = True
            lstCovAdd.Enabled = True
            'Call TabEnabled(False)
            'SSTab1.Tab = 2
            'txtDate = Date
            txtDate.Enabled = True
            txtEndtBordxDate.Enabled = True
            'txtDate.SetFocus
            cmdEditMBM.Enabled = True
        End If
        
    ElseIf cboEndorseStatus.ListIndex = 3 Then
        'Call EnableSupplementary(False)
        'lblAdj.Caption = ""
        'lblAdj.Caption = "Adjustment Date"
        lstCovAdd.Enabled = True
        cmdRemove.Enabled = True
        'cmdClear.Enabled = True
        SSTab1.Tab = 2
        
    ElseIf cboEndorseStatus.ListIndex = 6 Then
        'lblAdj.Caption = ""
        'lblAdj.Caption = "Insertion Date"
         Call ClearEntries
         cboINSCode.Enabled = True
         cboPLNCode.Enabled = True
         txtMBMBordxDate.Enabled = True
        lstCovAdd.Enabled = True
        cmdRemove.Enabled = True
        'cmdClear.Enabled = True
        Frame1.Enabled = True
        Frame4.Enabled = True
        cboStatus.ListIndex = 0
        txtDate.Enabled = True
        txtEndtBordxDate.Enabled = True
        txtDate.SetFocus
    ElseIf Mid(cboEndorseStatus, 1, 1) = "D" Then
            Frame1.Enabled = False
            Frame4.Enabled = False
            'Frame5.Enabled = False
            'Frame6.Enabled = False
        Call DeletePrincipals
    Else ' Reactivate
        'lblAdj.Caption = ""
        'lblAdj.Caption = "Reactivation Date"
        Call TabEnabled(True)
        'txtDate.Mask = "__/__/____"
        txtDate.Enabled = True
        txtEndtBordxDate.Enabled = True
        'txtDate.SetFocus
        cmdEditMBM.Enabled = True
    End If
    cboEndorseStatus.Enabled = False
'    CmdSave.Enabled = True
    cmdPrev.Enabled = True
    CmdNext.Enabled = True
End Sub

Private Sub cboINSCode_LostFocus()
    Clipboard.Clear
    'If Len(Trim(cboinsCode.Text)) Then
    '    Call LoadPlanType(cboinsCode.Text)
    'End If
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

Private Sub cboMBMSalutation_Click()
    If ChkStr(cboMBMSalutation) = "MR" Then
        cboMBMSex = "M"
    Else
        cboMBMSex = "F"
    End If
End Sub

Private Sub cboMBMSalutation_CHANGE()
    If ChkStr(cboMBMSalutation) = "MR" Then
        cboMBMSex = "M"
    Else
        cboMBMSex = "F"
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

Private Sub lstCaseList_DblClick()
    'frmCaseInquiry.txtCaseNo = ChkStr(lstCaseList.SelectedItem.Text)
    'Call frmCaseInquiry.CmdShow_Click
    'frmCaseInquiry.show
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
        'cboOccupation = lstCovAdd.SelectedItem.SubItems(13)
        'cboWorkNature = lstCovAdd.SelectedItem.SubItems(14)
        'cboSmoking = lstCovAdd.SelectedItem.SubItems(15)
        'cboAlcohol = lstCovAdd.SelectedItem.SubItems(16)
        txtAvailableLimit = Format(lstCovAdd.SelectedItem.SubItems(17), "#,##0.00")
        'txtSuppPrem = Format(lstCovAdd.SelectedItem.SubItems(18), "#,##0.00")
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
        cboSuppStaffDisc = lstCovAdd.SelectedItem.SubItems(30)
        'If Len(lstCovAdd.SelectedItem.SubItems(31)) Then
        TxtSuppClmNonPayFr.mask = ""
        TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(31)
        'End If
        'If Len(lstCovAdd.SelectedItem.SubItems(32)) Then
        TxtSuppClmNonPayTo.mask = ""
        TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(32)
        txtSuppPayPremGross = lstCovAdd.SelectedItem.SubItems(34)
        txtSuppPayPremNet = lstCovAdd.SelectedItem.SubItems(35)
        txtSuppPayMCOGross = lstCovAdd.SelectedItem.SubItems(36)
        txtSuppPayMCONet = lstCovAdd.SelectedItem.SubItems(37)
        txtRenewalDisc = lstCovAdd.SelectedItem.SubItems(38)
        txtPolicyDisc = lstCovAdd.SelectedItem.SubItems(39)
        'End If
    End If
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
        'cboOccupation = lstCovAdd.SelectedItem.SubItems(13)
        'cboWorkNature = lstCovAdd.SelectedItem.SubItems(14)
        'cboSmoking = lstCovAdd.SelectedItem.SubItems(15)
        'cboAlcohol = lstCovAdd.SelectedItem.SubItems(16)
        txtAvailableLimit = Format(lstCovAdd.SelectedItem.SubItems(17), "#,##0.00")
        'txtSuppPrem = Format(lstCovAdd.SelectedItem.SubItems(18), "#,##0.00")
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
        cboSuppStaffDisc = lstCovAdd.SelectedItem.SubItems(30)
        'If Len(lstCovAdd.SelectedItem.SubItems(31)) Then
        TxtSuppClmNonPayFr.mask = ""
        TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(31)
        'End If
        'If Len(lstCovAdd.SelectedItem.SubItems(32)) Then
        TxtSuppClmNonPayTo.mask = ""
        TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(32)
        txtSuppPayPremGross = lstCovAdd.SelectedItem.SubItems(34)
        txtSuppPayPremNet = lstCovAdd.SelectedItem.SubItems(35)
        txtSuppPayMCOGross = lstCovAdd.SelectedItem.SubItems(36)
        txtSuppPayMCONet = lstCovAdd.SelectedItem.SubItems(37)
        txtRenewalDisc = lstCovAdd.SelectedItem.SubItems(38)
        txtPolicyDisc = lstCovAdd.SelectedItem.SubItems(39)
        
        'End If
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
End Sub

Private Sub txtCovBirthdate_LostFocus()
    Dim MBMValidDate As New ADODB.Recordset
    Dim strsql As String
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
End Sub

Private Sub txtCovExclusion_Click()
    KeyPreview = False
End Sub

Private Sub txtCovExclusion_LostFocus()
    KeyPreview = True
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

Private Sub txtEndtBordxDate_LostFocus()
 Dim BAT As New ADODB.Recordset
    Dim strsql As String
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
End Sub

Private Sub txtMBMBirthdate_LostFocus()
    Dim MBMValidDate As New ADODB.Recordset
    Dim strsql As String
    Dim DateExit
    
    If txtMBMBirthDate <> "__/__/____" Then
        If IsDate(txtMBMBirthDate) Then
           strsql = "Select ValidID ,MBMDOBStart , MBMDOBEnd  From MBMValidDateRange "
           MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If MBMValidDate.recordCount Then
                If DateDiff("d", MBMValidDate!MBMDOBStart, txtMBMBirthDate) >= 0 And DateDiff("d", MBMValidDate!MBMDOBEnd, txtMBMBirthDate) <= 0 Then
                    If txtMBMPayorEffDate <> "__/__/____" Then
                        Call MemberCallAge
                        Call CallLimit
                        If Len(Trim(txtAGECode)) <> 0 Then
                            If txtMBMPremium.Caption = "" Or txtMBMPremium.Caption = "0" Then
                                Call MemberCallFees
                                'Call ProviderCallFees
                            End If
                        End If
                    End If
                Else
                    DateExit = MsgBox("Uncertain Birthdate, please select OK to continue " & vbNewLine & _
                                " Otherwise press Cancel to make changes! ", vbExclamation + vbOKCancel + vbDefaultButton2)
                            '" It must be in between " & MBMValidDate!MBMDOBStart & " and " & MBMValidDate!MBMDOBEnd, vbExclamation + vbOKCancel + vbDefaultButton2)
                    If DateExit = vbCancel Then
                        txtMBMBirthDate.SetFocus
                    Else
                        Call MemberCallAge
                        MBMValidDate.Close
                        Set MBMValidDate = Nothing
                        Exit Sub
                    End If
                End If
            End If
            MBMValidDate.Close
            Set MBMValidDate = Nothing
        Else
            MsgBox "Invalid date format! ", vbCritical + vbOKCancel
            txtMBMBirthDate.SetFocus
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


Private Sub txtMBMNumber_GotFocus()
cmdShow.Default = True
End Sub

Private Sub txtMBMNumber_LostFocus()
txtMBMNumber = ChkStr(txtMBMNumber)
cmdShow.Default = False
End Sub

Private Sub txtMBMPayorEffDate_gotfocus()
    txtMBMPayorEffDate.SelStart = 0
'    Call MemberCallAge
End Sub

Private Sub txtMBMPayorEffDate_LostFocus()
    Dim MBMValidDate As New ADODB.Recordset
    Dim strsql As String
    Dim CheckDays As Integer
    Dim tempdate As Date
    Dim answer
    'Dim Confirm1
    'Dim Confirm2
    
    'CheckDays = 0
    
    If (txtMBMPayorEffDate) <> "__/__/____" Then
        If IsDate(txtMBMPayorEffDate) Then
           
           strsql = "Select ValidID , MBMEffStart, MBMEffEnd From MBMValidDateRange "
           MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                           
           If MBMValidDate.recordCount Then
                If DateDiff("d", MBMValidDate!MBMEffStart, txtMBMPayorEffDate) >= 0 And DateDiff("d", MBMValidDate!MBMEffEnd, txtMBMPayorEffDate) <= 0 Then
                    tempdate = Format(DateAdd("yyyy", 1, txtMBMPayorEffDate), "dd/mm/yyyy")
                    txtMBMPayorExpDate = Format(DateAdd("d", -1, tempdate), "dd/mm/yyyy")
                    'txtMBMPayorEffDate.SetFocus
                    txtMBMEmployeeNo.SetFocus

                    If Mid(cboEndorseStatus, 1, 1) <> "I" Then
                        If txtMBMBirthDate <> "__/__/____" Then
                            Call MemberCallAge
                                If Len(Trim(txtAGECode)) <> 0 Then
                                    Call MemberCallFees
                                    'Call ProviderCallFees
                                End If
                            Call CallLimit
                        End If
                    End If
                Else
                    MsgBox "Policy Effective Date is out of acceptable range! ", vbCritical + vbOKOnly
                    txtMBMPayorEffDate.SetFocus
                End If
           Else
                tempdate = Format(DateAdd("yyyy", 1, txtMBMPayorEffDate), "dd/mm/yyyy")
                txtMBMPayorExpDate = Format(DateAdd("d", -1, tempdate), "dd/mm/yyyy")
                'txtMBMPayorEffDate.SetFocus
                txtMBMEmployeeNo.SetFocus
                If Mid(cboEndorseStatus, 1, 1) <> "I" Then
                    If txtMBMBirthDate <> "__/__/____" Then
                        Call MemberCallAge
                        Call MemberCallFees
                        'Call ProviderCallFees
                        Call CallLimit
                    End If
                End If
           End If
           MBMValidDate.Close
           Set MBMValidDate = Nothing
        Else
            answer = MsgBox("Invalid Date Format!", vbCritical + vbOKCancel)
            If answer = vbOK Then
                txtMBMPayorEffDate.SetFocus
            Else
                txtMBMPayorExpDate.SetFocus
            End If
        End If
    End If
    'End If
End Sub

' ######################### END CHECKING ###########
' ######################### Start Tab Index ###########
Private Sub cboMBMRelCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
        SSTab1.Tab = 1
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    cboCovSalutation = "MR"
    txtCovName = ""
    txtCovIcBcPp = ""
    txtCovBirthdate.mask = "__/__/____"
    txtCovBirthdate.mask = "##/##/####"
    
    'cboMBMRelCode.text = "S - Shared Limit"
    cboCovRelationship = "W - Wife"
    cboCovSex = "M"
    cboCovAdultChild = "A"
    txtCovAnnualLimit = ""
    txtAvailableLimit = ""
    'cboCovBlood = ""
    'cboOccupation = ""
    'cboWorkNature = ""
    'cboSmoking = ""
    'cboAlcohol = ""
    txtCovAllergic = ""
    txtCovExclusion = ""
    txtCovAnnualLimit = tempSAnnLmt
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
        If Mid(cboINSCode, 1, 1) = "I" Or Mid(cboINSCode, 1, 1) = "G" Then
            MsgBox "This Plan does not have supplementary", vbOKOnly + vbCritical, DialogBoxTitle
            SSTab1.Tab = 0
        Else
            SSTab1.Tab = 1
            cboCovSalutation.Enabled = True
            'If lstCovAdd.listitems.count = 0 Then
            '    Call showMBMCov
            'End If
        End If
    ElseIf SSTab1.Tab = 2 Then
        lstAdjustmentHis.listitems.Clear
        If lstAdjustmentHis.listitems.count = 0 Then
            Call ADJListing
        End If
        lstAdjustmentHis.SetFocus
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

Private Sub Combo1_GotFocus()
    If bTabDone = False Then
        cmdEditMBM.SetFocus
        bTabDone = True
    Else
        'txtMBMAllergic.SetFocus
        'bTabDone = False
    End If
End Sub

Private Sub Combo2_GotFocus()
    If bTabDone = False Then
        cmdEditMBM.SetFocus
        bTabDone = True
    Else
       ' lstCovAdd.SetFocus
        bTabDone = False
    End If

End Sub

Private Sub lstAdjustmentHis_GotFocus()
    bTabDone = False
End Sub

Private Sub SSTab1_GotFocus()
    If SSTab1.Tab = 0 Then
        'txtMBMAdd1.SetFocus
    End If
End Sub

Private Sub Combo3_GotFocus()
    If bTabDone = False Then
        cmdEditMBM.SetFocus
        bTabDone = True
    Else
        lstAdjustmentHis.SetFocus
    End If
End Sub


Private Sub cmdExit_LostFocus()
    If Me.ActiveControl.TabIndex = 44 Then
        txtMBMNumber.SetFocus
    End If
End Sub
' ######################### START BUTTON CLICKING ###########


Public Sub cmdshow_Click()
Dim MBMNumber As String
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
    'If getRecord = 1 Then
        'If txtMBMNumber = "" Then
            'MsgBox "Please Enter Membership No", vbOKOnly + vbCritical, DialogBoxTitle
        'Else
            'cmdEditMBM.Enabled = True
            'Call showMBMCov
            'Call CallLimit
            
            '================== Original Dim ============================
            'Call ProviderCallFees
            'Call MemberCallFees
            'Call BNFListing
            'Call ListCases
            'txtMBMPremiumEff = Format(CDbl(txtMBMPremium) + CDbl(MCOPremiumReduction(Format(txtMBMPayorExpDate, "mm/dd/yyyy"), Date, txtMBMPremium)), "######.99")
            'txtMCOFeesEff = Format(CDbl(txtMCOFees) + CDbl(MCOPremiumReduction(Format(txtMBMPayorExpDate, "mm/dd/yyyy"), Date, txtMCOFees)), "######.99")
            '================== Original Dim ============================
            
        'End If
    'End If
    
    ' ================ Add by CK =============================
    
    If txtMBMNumber = "" Then
        MsgBox "Please Enter Membership No", vbOKOnly + vbCritical, DialogBoxTitle
        txtMBMNumber.SetFocus
    Else
        
        'txtMBMNumber = Trim(txtMBMNumber)
        ScanCode = Trim(txtMBMNumber)
        If Mid(ScanCode, 1, 2) = "LO" Then
            txtMBMNumber = ""
            Call SearchForm(ScanCode)
        Else
            Call FindMBM
        End If
        'getRecord = showMBM
        'cmdEditMBM.Enabled = True
        
        'Call CallLimit
    End If
    ' ================ Add by CK =============================
    
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
        cmdEditMBM.Enabled = True
        'Call showMBMCov
        Call showMBMOthers
        Call showMBMCov
        Call CallLimit
    End If
    rst2.Close
    Set rst2 = Nothing
    
End Function

Private Function showMBM() As Integer
Dim sqlstr, GetProString As String
Dim MBM As New ADODB.Recordset
Dim sqlstqMBMRef As String
Dim MBMRef As New ADODB.Recordset
Dim MBMNumber As String
Dim strCount As String
Dim rstCount As New ADODB.Recordset
    
    txtMBMPremiumRed = ""
    txtMBMMCORed = ""
    txtMBMProviderRedCharges = ""
    txtMBMProviderEffCharges = ""
    txtMCOFeesEff = ""
    txtMBMPremiumEff = ""
    cboEndorseStatus.ListIndex = 0
    sqlstr = "Select MBMNumber, HLTCode, MBMPolicyNo, MBMName, MBMIcBcPp, " & _
            " MBMMemberType, MBMStatus, INSCode, PLNCode, MBMRenewFlag, " & _
            " MBMAddress1, MBMAddress2, MBMAddress3, MBMPostCode, CTYCode, " & _
            " MBMTelnoH, MBMEmail, MBMSalutation, MBMSex, SRVCodeList, MBMTelNoM, " & _
            " MBMTelNoO, MBMOccupation, MBMHeight, MBMWeight, MBMWorkNature," & _
            " MBMGroupCompany, MBMMemberType, MBMEmployeeNo, MBMBlood, " & _
            " MBMSmoking, MBMAlcohol, MBMExclusion, MBMAllergic, " & _
            " MBMMaritalStatus, MBMAccessFee, MBMPremium, MBMMCOfee, " & _
            " MBMFirstJoinedDate, BRNCode, MBMDOB, NATCode, RACCode, " & _
            " LGNCode, STACode, MBMFirstPolNo , MBMFirstMEMNo, MBMENDtEffDate, MBMPOLSubNo1, " & _
            " BRNCode, MBMTakeOver, MBMAgentCode , MBMPolicyYear, mbmDataStatus, LabelStatus"
    sqlstr = sqlstr & ", PAYCode, MBMBatchNo, " & _
            " MBMBordxDate, AGECode, MBMPrevMemNo, MBMPrevPolNo ," & _
            " MBMPayorEffDate, MBMPayorExpDate , MBMAdultChild, MBMAvailableLimit, " & _
            " MBMLastEndorseStatus , MBMENDtBordxDate , " & _
            " MBMReinstate, MBMLastAdjustmentDate, MBMInstallment,  MBMExpiredStatus, " & _
            " MBMPrevPLNCode, MBMRBAmt, MBMPAYRemarks, MBMPrevInsCom, MBMDepartment, MBMRemarks, MBMCoPayRB, " & _
            " MBMGenWP, MBMGenEndWP, MBMPreExWP, MBMPreExEndWP, MBMSpeIllWP, MBMPolicyDisc, MBMRenewalDisc, " & _
            "MBMSpeIllEndWP, MBMStaffDec, MBMCLMNonPayFr, MBMCLMNonPayTo,MBMNewPlanCode,MBMNewPolEffDate , MBMBTBG, MBMPayorPrem "
    sqlstr = sqlstr & " From View_MBM_MBMOthers Where MBMNumber ='" & Trim(txtMBMNumber) & "'"
    MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'strCount = "select count(*) as totalRecord From View_MBM_MBMOthers  where 0=0 "
    'strCount = strCount + sqlstr

    'Set rstCount = medixmasterconn.Execute(strCount)
    
    'If rstCount!totalrecord = 0 Then
    '    MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    'Else
    

        showMBM = 1
        With MBM
            txtMBMName = Trim(!MBMName)
            If Len(!MBMBatchNo) Then
                txtBatchNo = Trim(!MBMBatchNo)
            End If
            If Len(!MBMIcBcPp) Then
                txtMBMIcBcPP = Trim(!MBMIcBcPp)
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
            If Len(!MBMBatchNo) Then
                txtMBMBatchNo = Trim(!MBMBatchNo)
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
            If Trim(!SRVCodeList) <> "" Then
                cboSRVCode = Trim(!SRVCodeList)
            End If
            If Trim(!MBMGroupCompany) <> "" Then
                txtMBMGroupCompany = Trim(!MBMGroupCompany)
            End If
            If Trim(!MBMEmployeeNo) <> "" Then
                txtMBMEmployeeNo = Trim(!MBMEmployeeNo)
            End If
            If Trim(!MBMInstallment) <> "" Then
                CboInsMode = Trim(!MBMInstallment)
            End If
            If Trim(!MBMExclusion) <> "" Then
                txtMBMExclusion = Trim(!MBMExclusion)
            End If
            If Trim(!MBMMaritalStatus) <> "" Then
                cboMBMMaritalStatus = Trim(!MBMMaritalStatus)
            End If
            If Trim(!MBMFirstJoinedDate) <> "" Then
                txtMBMDateJoin = Format(!MBMFirstJoinedDate, "dd/mm/yyyy")
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
                txtMBMBirthDate = Format(!MBMDOB, "dd/mm/yyyy")
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
            If Len(!PAYCode) Then
                cboPAYCode = Trim(!PAYCode)
            End If
            If Len(!AgeCode) Then
                txtAGECode = Trim(!AgeCode)
            End If
            If Trim(!MBMBordxDate) <> "" Then
                txtMBMBordxDate = Format(!MBMBordxDate, "dd/mm/yyyy")
            End If
            If Len(!HLTCode) Then
                cboHLTCode = Trim(!HLTCode)
            End If
            If Trim(!MBMPayorEffDate) <> "" Then
                txtMBMPayorEffDate = Format(!MBMPayorEffDate, "dd/mm/yyyy")
            End If
            If Trim(!MBMPayorEXPDate) <> "" Then
                txtMBMPayorExpDate = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
            End If
            If Len(!MBMPolicyYear) Then
                txtMBMPolicyYear = Trim(!MBMPolicyYear)
            End If
            If Len(!MBMDataStatus) Then
                cboDataStatus = Trim(!MBMDataStatus)
            End If
            If Len(!MBMRemarks) Then
                txtRemarks = Trim(!MBMRemarks)
            End If
            If Len(!MBMAvailableLimit) Then
                txtAvailableAnnualLimit = Format(Trim(!MBMAvailableLimit), "#,##0.00")
            End If
            If Len(!MBMAdultChild) Then
                txtMBMAdultChild = Trim(!MBMAdultChild)
            End If
            If ChkStr(MBM!MBMMemberType) = "P" Then
                cboMBMRelCode = "C - Combined Limit"
            Else
                cboMBMRelCode = "S - Separated Limit"
            End If
            
            If Len(Trim(cboPAYCode)) <> 0 And Len(cboINSCode) <> 0 And txtAGECode <> "" And Len(cboHLTCode) <> 0 And txtMBMPayorEffDate <> "__/__/____" Then
                txtMBMPremium = Format(IIf(IsNull(!MBMPremium) = True, Format(GetPremium(Trim(cboINSCode), Trim(cboPAYCode), Trim(txtAGECode), Trim(cboPLNCode), Trim(cboHLTCode), txtMBMPayorEffDate), "#,##0.00"), !MBMPremium), "#,##0.00")
            End If
            If Len(Trim(cboPAYCode)) <> 0 And Len(cboINSCode) <> 0 And txtAGECode <> "" And Len(cboHLTCode) <> 0 And txtMBMPayorEffDate <> "__/__/____" Then

                txtMCOFees = Format(IIf(IsNull(!MBMMCOFee) = True, Format(GetMCO(Trim(cboINSCode), Trim(cboHLTCode), txtMBMAdultChild, txtMBMPayorEffDate, Mid(cboPLNCode, 1, 4)), "#,##0.00"), !MBMMCOFee), "#,##0.00")
            End If
            
            If Len(cboSRVCode) <> 0 And Len(cboINSCode) <> 0 And Len(cboHLTCode) <> 0 And txtMBMAdultChild <> "" Then
                txtMBMProviderCharges = Format(IIf(IsNull(!MBMAccessFee) = True, Format(GetProviderCharges(Trim(Mid(cboSRVCode, 1, 2)), Trim(cboINSCode), Trim(cboHLTCode), Trim(txtMBMAdultChild)), "####0.00"), !MBMAccessFee), "#,##0.00")
            End If
            txtTempMCOFees = txtMCOFees
            txtTempPrem = txtMBMPremium
            txtTempProvider = txtMBMProviderCharges
            If Trim(!MBMPayorEffDate) <> "" Then
                txtTempEffDate = Format(!MBMPayorEffDate, "dd/mm/yyyy")
            End If
            If Trim(!MBMPayorEXPDate) <> "" Then
                txtTempExpDate = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
            End If
            If Trim(!MBMStatus) = "A" Then
                cboStatus.ListIndex = 0
            ElseIf Trim(!MBMStatus) = "C" Then
                cboStatus.ListIndex = 1
            ElseIf Trim(!MBMStatus) = "D" Then
                cboStatus.ListIndex = 2
            ElseIf Trim(!MBMStatus) = "S" Then
                cboStatus.ListIndex = 3
            End If
            If Trim(!MBMExpiredStatus) = "N" Then
                cboExpStatus.ListIndex = 0
            Else
                cboExpStatus.ListIndex = 1
            End If

            If Len(!MBMPolicyDisc) Then
                PolDics = !MBMPolicyDisc
            End If
            
            If Len(!MBMRenewalDisc) Then
                RenDisc = !MBMRenewalDisc
            End If
    
                
            End With
    
            If MBM!MBMStatus = "C" Then
            
                MBMNumber = Trim(MBM!MBMNumber)
                sqlstqMBMRef = "Select * from MBMRefund where MBMNumber =  '" & MBMNumber & "'"
                MBMRef.Open sqlstqMBMRef, medixmasterconn, adOpenKeyset, adLockOptimistic
                If MBMRef.recordCount Then
                With MBMRef
                    If MBMRef!MBMPremReduction <> "" Then
                        txtMBMPremiumRed = Format(Round(!MBMPremReduction, 0), "#,##0.00")
                    End If
                    If Len(!MBMMCORefund) Then
                        txtMBMMCORed = Format(Round(!MBMMCORefund, 0), "#,##0.00")
                    End If
                    If Len(!MBMProviderRef) Then
                        txtMBMProviderRedCharges = Format(Round(!MBMProviderRef, 1), "#,##0.00")
                    End If
                    If Len(!MBMProviderEff) Then
                        txtMBMProviderEffCharges = Format(Round(!MBMProviderEff, 1), "#,##0.00")
                    End If
                    If Len(!MBMMCOEffective) Then
                        txtMCOFeesEff = Format(Round(!MBMMCOEffective, 0), "#,##0.00")
                    End If
                    If Len(!MBMPREMEffective) Then
                        txtMBMPremiumEff = Format(Round(!MBMPREMEffective, 0), "#,##0.00")
                    End If
                End With
                MBMRef.Close
                Set MBMRef = Nothing
            End If
            End If
            'End If
    MBM.Close
    Set MBM = Nothing
End Function

Private Sub showMBMCov()
    Dim MBMCov As New ADODB.Recordset
    Dim sqlstr As String
                
    sqlstr = "select MBMCNumber, MBMCCoverID, MBMCName, MBMCSalutation," & _
        " MBMCICBCPPNo , MBMCDOB, MBMCSex, MBMCAdultChild, MBMCAgeband, " & _
        " MBMCRELCode, MBMCBloodType, MBMCAllergic,  " & _
        " MBMCExclusion, MBMCStatus , MBMCAnnualLimit, MBMCOccupation, " & _
        " MBMCWorkNature, MBMCSmoking, MBMCAlcohol, MBMCPremium, MBMCAvailableLimit, " & _
        " MBMCPAYRemarks, MBMCRBAmt, MBMCPrevInsCom, MBMCPrevPLNCode, " & _
        " MBMCGenWP, MBMCGenEndWP, MBMCPreExWP, MBMCPreExEndWP, " & _
        " MBMCSpeIllWP, MBMSpeIllEndWP, MBMCStaffDec, MBMCCLMNonPayFr, MBMCCLMNonPayTo, MBMPayorPrem, " & _
        " MBMCPayorPremGross, MBMCPayorPremNet, MBMCPayorMCOGross, MBMCPayorMCONet, MBMCPolicyDisc, MBMCRenewalDisc " & _
        " from  MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
    MBMCov.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic

    lstCovAdd.listitems.Clear
    If MBMCov.recordCount Then
        Dim MemberConveredPerson As ListItem
        Do While Not MBMCov.EOF
            Set MemberConveredPerson = lstCovAdd.listitems.Add(, , MBMCov!MBMCName)
            With MBMCov
                If Len(Trim(!MBMCSalutation)) <> 0 Then
                    MemberConveredPerson.SubItems(1) = ChkStr(!MBMCSalutation)
                End If
                If Len(Trim(!MBMCIcBcPpNo)) <> 0 Then
                    MemberConveredPerson.SubItems(2) = !MBMCIcBcPpNo
                End If
                If Len(Trim(!MBMCDOB)) <> "" Then
                    MemberConveredPerson.SubItems(3) = Format(!MBMCDOB, "dd/mm/yyyy")
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
                                
                If Len(Trim(!MBMCOccupation)) Then
                    MemberConveredPerson.SubItems(13) = !MBMCOccupation
                End If
 
                If Len(Trim(!MBMCWorkNature)) Then
                    MemberConveredPerson.SubItems(14) = !MBMCWorkNature
                End If
                
                If Len(Trim(!MBMCSmoking)) Then
                    MemberConveredPerson.SubItems(15) = !MBMCSmoking
                End If
                
                If Len(Trim(!MBMCAlcohol)) Then
                    MemberConveredPerson.SubItems(16) = !MBMCAlcohol
                End If
                
                If Len(Trim(!MBMCAvailableLimit)) Then
                    MemberConveredPerson.SubItems(17) = !MBMCAvailableLimit
                End If
                
                If Len(Trim(!MBMCPremium)) Then
                    MemberConveredPerson.SubItems(18) = !MBMCPremium
                End If
                If Len(Trim(!MBMCPrevPLNCode)) Then
                    MemberConveredPerson.SubItems(20) = !MBMCPrevPLNCode
                End If
                If Len(Trim(!MBMCRBAmt)) Then
                    MemberConveredPerson.SubItems(21) = !MBMCRBAmt
                End If
                If Len(Trim(!MBMCPAYRemarks)) Then
                    MemberConveredPerson.SubItems(22) = !MBMCPAYRemarks
                End If
                If Len(Trim(!MBMCPrevInsCom)) Then
                    MemberConveredPerson.SubItems(23) = !MBMCPrevInsCom
                End If
                
                If Len(!MBMCGenWP) Then
                    MemberConveredPerson.SubItems(24) = !MBMCGenWP
                End If
                If Len(!MBMCGenEndWP) Then
                    MemberConveredPerson.SubItems(25) = !MBMCGenEndWP
                End If
                If Len(!MBMCPreExWP) Then
                    MemberConveredPerson.SubItems(26) = !MBMCPreExWP
                End If
                If Len(!MBMCPreExEndWP) Then
                    MemberConveredPerson.SubItems(27) = !MBMCPreExEndWP
                End If
                If Len(!MBMCSpeIllWP) Then
                    MemberConveredPerson.SubItems(28) = !MBMCSpeIllWP
                End If
                If Len(!MBMSpeIllEndWP) Then
                    MemberConveredPerson.SubItems(29) = !MBMSpeIllEndWP
                End If
                If Len(!MBMCStaffDec) Then
                   MemberConveredPerson.SubItems(30) = !MBMCStaffDec
                End If
                If Len(!MBMCCLMNonPayFr) Then
                    MemberConveredPerson.SubItems(31) = !MBMCCLMNonPayFr
                End If
                If Len(!MBMCCLMNonPayTo) Then
                    MemberConveredPerson.SubItems(32) = !MBMCCLMNonPayTo
                End If
                If Len(!MBMCPayorPremGross) Then
                    MemberConveredPerson.SubItems(34) = !MBMCPayorPremGross
                End If
                If Len(!MBMCPayorPremNet) Then
                    MemberConveredPerson.SubItems(35) = !MBMCPayorPremNet
                End If
                
                If Len(!MBMCPayorMCOGross) Then
                    MemberConveredPerson.SubItems(36) = !MBMCPayorMCOGross
                End If
                
                If Len(!MBMCPayorMCONet) Then
                    MemberConveredPerson.SubItems(37) = !MBMCPayorMCONet
                End If
                
                If Len(!MBMCRenewalDisc) Then
                    MemberConveredPerson.SubItems(38) = !MBMCRenewalDisc
                End If
                
                If Len(!MBMCPolicyDisc) Then
                    MemberConveredPerson.SubItems(39) = !MBMCPolicyDisc
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
    cboINSCode.Enabled = False
    cboPLNCode.Enabled = False
    Frame10.Enabled = True
    Frame1.Enabled = False
    SSTab1.Enabled = True
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    cmdRemove.Enabled = False
    cboEndorseStatus.Enabled = True
End Sub

Private Sub cmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub cmdMBMPolicy_Click()
    frmSearchMBM.txtMBMPolicyNo1.value = True
    frmSearchMBM.TxtMBMNo1.value = False
    frmSearchMBM.Source = 2
    frmSearchMBM.show
End Sub

Private Sub cmdSearchMBM_Click()
    Call ClearEntries
    frmSearchMBM.Source = 2
    frmSearchMBM.show
End Sub

Private Sub cmdSearchIC_Click()
    frmSearchMBM.txtIcBcPc1.value = True
    frmSearchMBM.TxtMBMNo1.value = False
    frmSearchMBM.Source = 2
    frmSearchMBM.show
End Sub

Private Sub cmdSearchName_Click()
    frmSearchMBM.txtMBMName6.value = True
    frmSearchMBM.TxtMBMNo1.value = False
    frmSearchMBM.Source = 2
    frmSearchMBM.show
End Sub

'Private Sub cmdListBranch_Click()
    'frmMBMListBranch.Source = 2
    'frmMBMListBranch.show
'End Sub

'Private Sub cmdListCity_Click()
    'frmMBMListCity.Source = 2
    'frmMBMListCity.show
'End Sub

'Private Sub cmdListState_Click()
    'frmMBMListState.Source = 2
    'frmMBMListState.show
'End Sub

' ######## START SUPPLEMENTARY #####################

Private Sub cmdAdd_Click()
    Dim MemberConveredPerson As ListItem
    Dim RecordSaved

    If txtCovName = "" Then
        MsgBox "Please Enter Covered Person Name!", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf txtCovIcBcPp = "" Then
        MsgBox "Please Enter Covered Person Identification!", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboCovRelationship = "" Then
        MsgBox "Please Enter Covered Person Relationship!", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf txtCovBirthdate = "__/__/____/" Then
        MsgBox "Please Enter Covered Person Birth Date!", vbOKOnly + vbCritical, DialogBoxTitle
    'ElseIf txtCovBirthdate > Date Then
        'MsgBox "Birth Date can not greater than System Date!", vbOKOnly + vbCritical, DialogBoxTitle
    ElseIf cboCovSex = "" Then
        MsgBox "Please Enter Covered Person Sex!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            Set MemberConveredPerson = lstCovAdd.listitems.Add(, , txtCovName)
            MemberConveredPerson.SubItems(1) = cboCovSalutation
            MemberConveredPerson.SubItems(2) = txtCovIcBcPp
            MemberConveredPerson.SubItems(3) = txtCovBirthdate
            MemberConveredPerson.SubItems(4) = cboCovSex
            MemberConveredPerson.SubItems(5) = cboCovAdultChild
            MemberConveredPerson.SubItems(6) = cboCovRelationship
            'MemberConveredPerson.SubItems(7) = cboCovBlood
            MemberConveredPerson.SubItems(8) = txtCovAllergic
            MemberConveredPerson.SubItems(9) = txtCovExclusion
            MemberConveredPerson.SubItems(11) = "N"
            MemberConveredPerson.SubItems(12) = txtCovAnnualLimit
            'If cboOccupation <> "" Then
            '    MemberConveredPerson.SubItems(13) = ChkStr(Mid(cboOccupation, 1, InStr(1, cboOccupation, "-") - 1))
            'End If
            'If cboWorkNature <> "" Then
            '    MemberConveredPerson.SubItems(14) = ChkStr(Mid(cboWorkNature, 1, InStr(1, cboWorkNature, "-") - 1))
            'End If
            'If cboSmoking <> "" Then
            '    MemberConveredPerson.SubItems(15) = ChkStr(Trim(Mid(cboSmoking, 1, 1)))
            'End If
            'If cboAlcohol <> "" Then
            '    MemberConveredPerson.SubItems(16) = ChkStr(Trim(Mid(cboAlcohol, 1, 1)))
            'End If
            'If txtSuppPrem <> "" Then
            '    MemberConveredPerson.SubItems(18) = txtSuppPrem
            'End If
            'If cboPrevPLNStatus <> "" Then
            '    MemberConveredPerson.SubItems(19) = Mid(cboPrevPLNStatus, 1, 1)
            'End If
            If txtPrevPLNCode <> "" Then
                MemberConveredPerson.SubItems(20) = txtPrevPLNCode
            End If
            If txtRBAmt <> "" Then
                MemberConveredPerson.SubItems(21) = txtRBAmt
            End If
            If txtCOVPayRemarks <> "" Then
                MemberConveredPerson.SubItems(22) = txtCOVPayRemarks
            End If
            If txtPrevINSCom <> "" Then
                MemberConveredPerson.SubItems(23) = txtPrevINSCom
            End If
            If TxtSuppClmNonPayFr <> "__/__/____" Then
                MemberConveredPerson.SubItems(31) = TxtSuppClmNonPayFr
            End If
            If TxtSuppClmNonPayTo <> "__/__/____" Then
                MemberConveredPerson.SubItems(32) = TxtSuppClmNonPayTo
            End If
            If txtEndtBordxDate <> "__/__/____" Then
                MemberConveredPerson.SubItems(33) = txtEndtBordxDate
            End If

            If txtSuppPayPremGross <> "" Then
                MemberConveredPerson.SubItems(34) = txtSuppPayPremGross
            End If

            If txtSuppPayPremNet <> "" Then
                MemberConveredPerson.SubItems(35) = txtSuppPayPremNet
            End If
            
            If txtSuppPayMCOGross <> "" Then
                MemberConveredPerson.SubItems(36) = txtSuppPayMCOGross
            End If
            
            If txtSuppPayMCONet <> "" Then
                MemberConveredPerson.SubItems(37) = txtSuppPayMCONet
            End If
            
            If txtRenewalDisc <> "" Then
                MemberConveredPerson.SubItems(38) = txtRenewalDisc
            End If
            
            If txtPolicyDisc <> "" Then
                MemberConveredPerson.SubItems(39) = txtPolicyDisc
            End If

            cboCovSalutation.SetFocus
        End If
    End If
End Sub

Private Sub cmdRemove_Click()
    Dim ForLoop As Integer
    Dim ListCount As Integer
    Dim strsql As String
    Dim DEL As New ADODB.Recordset
    Dim RecordSaved
    ListCount = 0
    If txtCovName = "" Then
        MsgBox "Please Choose a Record", vbOKOnly + vbCritical, DialogBoxTitle
    Else
        RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            ListCount = lstCovAdd.listitems.count
            For ForLoop = 1 To ListCount
                If lstCovAdd.listitems(ForLoop).Selected = True Then
                    If Trim(lstCovAdd.SelectedItem.SubItems(10)) <> "" Then
                        lstCovAdd.SelectedItem.SubItems(11) = "R"
                    Else
                        lstCovAdd.listitems.Remove (ForLoop)
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
            For ForLoop = 1 To lstCovAdd.listitems.count
                If lstCovAdd.listitems(ForLoop).Selected = True Then
                        lstCovAdd.SelectedItem.Text = txtCovName
                        lstCovAdd.SelectedItem.SubItems(1) = Trim(cboCovSalutation)
                        lstCovAdd.SelectedItem.SubItems(2) = txtCovIcBcPp
                        lstCovAdd.SelectedItem.SubItems(3) = txtCovBirthdate
                        lstCovAdd.SelectedItem.SubItems(4) = cboCovSex
                        lstCovAdd.SelectedItem.SubItems(5) = cboCovAdultChild
                        lstCovAdd.SelectedItem.SubItems(6) = cboCovRelationship
                        'lstCovAdd.SelectedItem.SubItems(7) = cboCovBlood
                        lstCovAdd.SelectedItem.SubItems(8) = txtCovAllergic
                        lstCovAdd.SelectedItem.SubItems(9) = txtCovExclusion
                        lstCovAdd.SelectedItem.SubItems(11) = "E"
                        lstCovAdd.SelectedItem.SubItems(12) = txtCovAnnualLimit
                        'lstCovAdd.SelectedItem.SubItems(13) = ChkStr(Trim(Mid(cboOccupation, 1, 2)))
                        'lstCovAdd.SelectedItem.SubItems(14) = ChkStr(Trim(Mid(cboWorkNature, 1, 3)))
                        'lstCovAdd.SelectedItem.SubItems(15) = ChkStr(Mid(cboSmoking, 1, 1))
                        'If cboAlcohol <> "" Then
                        'lstCovAdd.SelectedItem.SubItems(16) = ChkStr(Trim(Mid(cboAlcohol, 1, 1)))
                        'End If
                        'If cboPrevPLNStatus <> "" Then
                        'lstCovAdd.SelectedItem.SubItems(18) = txtSuppPrem
                        'lstCovAdd.SelectedItem.SubItems(19) = Mid(cboPrevPLNStatus, 1, 1)
                        'End If
                        'If txtPrevPLNCode <> "" Then
                        lstCovAdd.SelectedItem.SubItems(20) = txtPrevPLNCode
                        'End If
                        lstCovAdd.SelectedItem.SubItems(21) = txtRBAmt
                        lstCovAdd.SelectedItem.SubItems(22) = txtCOVPayRemarks
                        lstCovAdd.SelectedItem.SubItems(23) = txtPrevINSCom
                        lstCovAdd.SelectedItem.SubItems(24) = txtSuppGenWP
                        lstCovAdd.SelectedItem.SubItems(25) = txtSuppGENWPEnd
                        lstCovAdd.SelectedItem.SubItems(26) = txtSuppPREWP
                        lstCovAdd.SelectedItem.SubItems(27) = txtSuppPREWPEnd
                        lstCovAdd.SelectedItem.SubItems(28) = TxtSuppSpeWP
                        lstCovAdd.SelectedItem.SubItems(29) = txtSuppSPEWPEnd
                        lstCovAdd.SelectedItem.SubItems(30) = cboSuppStaffDisc
                        lstCovAdd.SelectedItem.SubItems(34) = txtSuppPayPremGross
                        lstCovAdd.SelectedItem.SubItems(35) = txtSuppPayPremNet
                        lstCovAdd.SelectedItem.SubItems(36) = txtSuppPayMCOGross
                        lstCovAdd.SelectedItem.SubItems(37) = txtSuppPayMCONet
                        lstCovAdd.SelectedItem.SubItems(38) = txtRenewalDisc
                        lstCovAdd.SelectedItem.SubItems(39) = txtPolicyDisc
                    Exit For
                End If
            Next ForLoop
            cboCovSalutation.SetFocus
        End If
       Call ClearCoveredEntry
    End If
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
        'cboCovBlood = lstCovAdd.SelectedItem.SubItems(7)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(10)
        txtCovAnnualLimit = Format(lstCovAdd.SelectedItem.SubItems(12), "###,##0.00")
        'cboOccupation = lstCovAdd.SelectedItem.SubItems(13)
        'cboWorkNature = lstCovAdd.SelectedItem.SubItems(14)
        'cboSmoking = lstCovAdd.SelectedItem.SubItems(15)
        'cboAlcohol = lstCovAdd.SelectedItem.SubItems(16)
        txtAvailableLimit = Format(lstCovAdd.SelectedItem.SubItems(17), "#,##0.00")
        'txtSuppPrem = Format(lstCovAdd.SelectedItem.SubItems(18), "#,##0.00")
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
        
        'If lstCovAdd.SelectedItem.SubItems(31) <> "" Then
        TxtSuppClmNonPayFr.mask = ""
        TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(31)
        'End If
        'If lstCovAdd.SelectedItem.SubItems(32) <> "" Then
        TxtSuppClmNonPayTo.mask = ""
        TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(32)
        txtSuppPayPremGross = lstCovAdd.SelectedItem.SubItems(34)
        txtSuppPayPremNet = lstCovAdd.SelectedItem.SubItems(35)
        txtSuppPayMCOGross = lstCovAdd.SelectedItem.SubItems(36)
        txtSuppPayMCONet = lstCovAdd.SelectedItem.SubItems(37)
        txtRenewalDisc = lstCovAdd.SelectedItem.SubItems(38)
        txtPolicyDisc = lstCovAdd.SelectedItem.SubItems(39)
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

Private Sub CmdClear_Click()
    'Call EnableCovered(True)
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
    'cboCovBlood = ""
    txtCovAllergic = ""
    txtCovExclusion = ""
    txtCovAnnualLimit = ""
    'cboOccupation = ""
    'cboWorkNature = ""
    'cboSmoking = ""
    'cboAlcohol = ""
    'cboWorkNature = ""
    txtPrevPLNCode = ""
    txtRBAmt = ""
    txtPrevINSCom = ""
    txtCOVPayRemarks = ""
    'txtSuppPrem = ""
    txtCovAnnualLimit = ""
    txtAvailableLimit = ""
    txtCovAnnualLimit = tempSAnnLmt
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
    txtSuppPayPremGross = ""
    txtSuppPayPremNet = ""
    txtSuppPayMCOGross = ""
    txtSuppPayMCONet = ""
    txtRenewalDisc = ""
    txtPolicyDisc = ""
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
    'cboCovBlood.Enabled = status
    txtCovAllergic.Enabled = status
    txtCovExclusion.Enabled = status
    'txtCovAnnualLimit.Enabled = status
    cmdRemove.Enabled = status
    cmdAdd.Enabled = status
    'cboOccupation.Enabled = status
    'cboCovBlood.Enabled = status
    'cboWorkNature.Enabled = status
    'cboSmoking.Enabled = status
    
    
End Sub
'  ################# END Suppplementary ######################

' ################################ Enable Tab / button   ######
Public Sub EnableCovered(status As Boolean)
    cboCovSalutation.Enabled = status
    txtCovName.Enabled = status
    txtCovIcBcPp.Enabled = status
    txtCovBirthdate.Enabled = status
    cboCovSex.Enabled = status
    cboCovAdultChild.Enabled = status
    cboCovRelationship.Enabled = status
    'cboCovBlood.Enabled = status
    txtCovAllergic.Enabled = status
    txtCovExclusion.Enabled = status
    'txtCovAnnualLimit.Enabled = status
End Sub

Private Sub TabEnabled(status As Boolean)
    SSTab1.TabEnabled(0) = status
    SSTab1.TabEnabled(1) = status
    SSTab1.TabEnabled(2) = status
    SSTab1.TabEnabled(3) = status
'    SSTab1.TabEnabled(4) = status

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
    'cmdEditMBM.Enabled = False
    cboEndorseStatus.Enabled = False
    cboStatus.Enabled = False
    txtCovAnnualLimit.Enabled = False
    txtAvailableLimit.Enabled = False
    'txtSuppPrem.Enabled = False
    cboExpStatus.Enabled = False
    cboPAYCode.Enabled = False
    FrmMBMOthers.txtGENWPEnd.Enabled = False
    FrmMBMOthers.txtPREWPEnd.Enabled = False
    FrmMBMOthers.txtSPEWPEnd.Enabled = False
    txtSuppGENWPEnd.Enabled = False
    txtSuppPREWPEnd.Enabled = False
    txtSuppSPEWPEnd.Enabled = False
End Sub

' ################## Start processing

'Private Sub DeleteCoveredID(MembershipNo As String, coverID As Integer)
 '   Dim DELCOV As New ADODB.Recordset
  '  Dim delstr As String
   '
    'delstr = "Delete from MBMCoveredPersons Where MBMCCoverID =" & coverID & " And MBMNumber = '" & MembershipNo & "'"
    'DELCOV.Open delstr, medixmasterconn, adOpenKeyset, adLockOptimistic
'
 '   DELCOV.Close
  '  Set DELCOV = Nothing
'End Sub

Private Sub CancelPrincipalAdjustment()
Dim MBM As New ADODB.Recordset
Dim MemberAmendHis As New ADODB.Recordset
Dim strsql As String

    strsql = "Select * from MBM Where MBMNumber = '" & txtMBMNumber & "'"
    MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
    MemberAmendHis.Open "MBMAmendHistory", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    MemberAmendHis.AddNew
    With MemberAmendHis
        
        !MBMNumber = MBM!MBMNumber
        !MBMPolicyNo = MBM!MBMPolicyNo
        !MBMIcBcPp = MBM!MBMIcBcPp
        !MBMName = MBM!MBMName
        !MBMExpiryDate = MBM!MBMPayorEXPDate
        !MBMPlaNCode = MBM!PLNCode
        !MBMAmendEdtType = Mid(cboEndorseStatus, 1, 1)
        If IsDate(txtDate) Then
            !MBMAEndtEffDate = txtDate
        End If
        If IsDate(txtEndtBordxDate) Then
            !MBMAENDtBordxDate = txtEndtBordxDate
        End If
        !MBMAmendUser = Logon
        !MBMAEndtBatchNo = txtBatchNo
        MemberAmendHis.Update
    End With

        

                With MBM
                    !MBMLastEndorseStatus = "C"
                    !MBMStatus = "C"
                    !MBMENDtEffDate = txtDate
                    !MBMENDtBordxDate = txtEndtBordxDate
                End With
                MBM.Update
    MemberAmendHis.Close
    Set MemberAmendHis = Nothing
    MBM.Close
    Set MBM = Nothing
    'cboStatus.Enabled = True
    cboStatus.ListIndex = 1
    'cboStatus.Enabled = False
End Sub

Private Sub CancelSupplementaryAdjustment()
    Dim MBMCov As New ADODB.Recordset
    Dim strsql As String
        
    strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & txtMBMNumber & "'"
    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    'MemberHis.Open "MBA", medixmasterconn, adOpenKeyset, adLockOptimistic
    'MBMHIS.Open "MBMHIS", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until MBMCov.EOF
                With MBMCov
                    !MBMCStatus = "C"
                End With
                MBMCov.Update
        MBMCov.MoveNext
    Loop
    MBMCov.Close
End Sub

'  ############### SAVE DATA INTO DATABSE #####################
Private Sub cmdSave_Click()
On Error GoTo ErrorSave
Dim RecordSaved, confirmCancel
Dim strsql As String
Dim CAS As New ADODB.Recordset
Dim bSucess  As Boolean
    bSucess = False
    If Trim(txtMBMNumber) = "" And Mid(cboEndorseStatus, 1, 1) <> "I" Then
        MsgBox "Please select a record to edit!", vbCritical
        txtMBMNumber.SetFocus
    ElseIf Trim(txtMBMPolNo) = "" Then
        MsgBox "Please Enter Policy No!", vbCritical
        txtMBMPolNo.SetFocus
    ElseIf txtEndtBordxDate = "__/__/____" Then
        MsgBox "Please Enter Endt. Bordx Date!", vbCritical
    '    Frame10.Enabled = True
        txtEndtBordxDate.SetFocus
    'ElseIf txtEndtBordxDate > Date Then
        'MsgBox "Endt. Bordx Date can not greater than System Date!", vbCritical
        'txtEndtBordxDate.SetFocus
    ElseIf txtDate = "__/__/____" Then
        MsgBox "Please Enter Endt. Effective Date!", vbCritical
        txtDate.SetFocus
    'ElseIf txtDate > Date Then
        'MsgBox "Endt. Effective Date can not greater than System Date!", vbCritical
        'txtDate.SetFocus
    ElseIf txtMBMBirthDate = "__/__/____" Then
        MsgBox "Please Enter Birth Date!", vbCritical
        txtMBMBirthDate.SetFocus
    'ElseIf txtMBMBirthDate > Date Then
        'MsgBox "Birth Date can not greater than System Date!", vbCritical
        'txtMBMBirthDate.SetFocus
    ElseIf txtMBMPayorEffDate = "__/__/____" Then
        MsgBox "Please Enter Payor Eff. Date!", vbCritical
        txtMBMPayorEffDate.SetFocus
    'ElseIf txtMBMPayorEffDate > Date Then
        'MsgBox "Payor Eff. Date can not greater than System Date!", vbCritical
        'txtMBMPayorEffDate.SetFocus
    ElseIf txtMBMPayorExpDate = "__/__/____" Then
        MsgBox "Please Enter Payor Exp. Date!", vbCritical
        txtMBMPayorExpDate.SetFocus
    'ElseIf txtMBMPayorExpDate > Date Then
        'MsgBox "Payor Exp. Date can not greater than System Date!", vbCritical
        'txtMBMPayorExpDate.SetFocus
    ElseIf txtMBMBordxDate = "__/__/____" Then
        MsgBox "Please Enter Bordx Date!", vbCritical
        txtMBMBordxDate.SetFocus
    'ElseIf txtMBMBordxDate > Date Then
        'MsgBox "Bordx Date can not greater than System Date!", vbCritical
        'txtMBMBordxDate.SetFocus
    'ElseIf txtMBMDateJoin = "__/__/____" Then
        'MsgBox "Please Enter Date Joined!", vbCritical
        'txtMBMDateJoin.SetFocus
    'ElseIf txtMBMDateJoin > Date Then
        'MsgBox "Date Joined can not greater than System Date!", vbCritical
        'txtMBMDateJoin.SetFocus
    ElseIf txtMBMAdd1 = "" Then
        MsgBox "Please Enter Address!", vbCritical
        txtMBMAdd1.SetFocus
    ElseIf txtMBMPostCode = "" Then
        MsgBox "Please Enter Post Code!", vbCritical
        txtMBMPostCode.SetFocus
    ElseIf cboSTACode = "" Then
        MsgBox "Please Enter State!", vbCritical
        cboSTACode.SetFocus
    ElseIf cboINSCode = "" Then
        MsgBox "Please Enter Insured Type!", vbCritical
        cboINSCode.SetFocus
    ElseIf cboPLNCode = "" Then
        MsgBox "Please Enter Plan Code!", vbCritical
        cboPLNCode.SetFocus
    ElseIf txtMBMIcBcPP = "" Then
        MsgBox "Please Enter IC Number!", vbCritical
        txtMBMIcBcPP.SetFocus
    ElseIf cboMBMSex = "" Then
        MsgBox "Please Enter Sex!", vbCritical
        cboMBMSex.SetFocus
    ElseIf FrmMBMOthers.txtPLNCode <> "" And FrmMBMOthers.txtNewPolEffDate = "__/__/____" Then
        MsgBox "Please Enter New Policy Eff Date", vbOKOnly + vbCritical, DialogBoxTitle
        FrmMBMOthers.txtNewPolEffDate = "__/__/____"
        FrmMBMOthers.txtNewPolEffDate.SetFocus
    ElseIf FrmMBMOthers.txtPLNCode = "" And FrmMBMOthers.txtNewPolEffDate <> "__/__/____" Then
        MsgBox "Please Enter New Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
        FrmMBMOthers.txtPLNCode.SetFocus

    'ElseIf txtEndtBordxDate = "__/__/____" Then
    '    MsgBox "Please enter ENDtBordx Date!", vbCritical
    '    Frame10.Enabled = True
    '    txtEndtBordxDate.SetFocus
    'ElseIf Trim(txtLastBatch) = "" Then
    '    MsgBox "Please enter Batch No!", vbCritical
    '    txtBatchNo.SetFocus
    Else
            RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                'startTrans = True
                'medixmasterconn.BeginTrans
                Screen.MousePointer = vbHourglass
            strsqlPLN = "select * from DiscPlan where PLNCode = '" & Trim(Mid(cboPLNCode, 1, 4)) & "'"
            PLN.Open strsqlPLN, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
        
            If PLN.recordCount Then
                Disc = True
            Else
                Disc = False
            End If
    
            PLN.Close
            Set PLN = Nothing
                
                If cboEndorseStatus = "" Then
                    MsgBox "Please Choose An Endorsement Type", vbOKOnly + vbCritical, DialogBoxTitle
                    cboEndorseStatus.SetFocus
                ElseIf cboEndorseStatus.ListIndex = 1 Then
                    If txtDate = "__/__/____" Then
                        MsgBox "Please Enter EndtEff Date", vbOKOnly + vbCritical, DialogBoxTitle
                        txtDate.SetFocus
                    ElseIf txtEndtBordxDate = "__/__/____" Then
                        MsgBox "Please Enter EndtBordx Date", vbOKOnly + vbCritical, DialogBoxTitle
                        txtEndtBordxDate.SetFocus

                    ElseIf Mid(cboStatus, 1, 1) = "D" Then
                        MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                    ElseIf Mid(cboStatus, 1, 1) = "C" Then
                        MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                    ElseIf Mid(cboStatus, 1, 1) = "S" Then
                        MsgBox "Membership has been Suspended. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle

                    Else
                        strsql = "Select CASNumber from CAS Where MBMNumber ='" & txtMBMNumber & "' And CASStatus = 'O'"
                        CAS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                        If CAS.recordCount > 0 Then
                            confirmCancel = MsgBox("There is at least an outstanding case under this member!", vbYesNo + vbCritical)
                        
                            'If confirmCancel Then
                            '    Call CancelPrincipalAdjustment
                            '    Call MBMRefund
                            '    Call CancelSupplementaryAdjustment
                            'End If
                        Else
                            Call CancelPrincipalAdjustment
                            Call MBMRefund
                            Call CancelSupplementaryAdjustment
                            'Label7.Caption = "Cancelled"
                            MsgBox "This Membership has been cancelled successfully", vbOKOnly + vbInformation, DialogBoxTitle
                            bSucess = True
                        End If
                    End If
                   
                    ElseIf cboEndorseStatus.ListIndex = 6 Then
                        If txtDate = "__/__/____" Then
                            MsgBox "Please Enter EndtEff Date", vbOKOnly + vbCritical, DialogBoxTitle
                            txtDate.SetFocus
                        ElseIf txtEndtBordxDate = "__/__/____" Then
                            MsgBox "Please Enter EndtBordx Date", vbOKOnly + vbCritical, DialogBoxTitle
                            txtEndtBordxDate.SetFocus
                        ElseIf txtMBMName = "" Then
                            MsgBox "Please Enter Member Name", vbOKOnly + vbCritical, DialogBoxTitle
                            txtMBMName.SetFocus
                        ElseIf txtMBMIcBcPP = "" Then
                            MsgBox "Please Enter IC No", vbOKOnly + vbCritical, DialogBoxTitle
                            txtMBMIcBcPP.SetFocus
                        ElseIf cboPAYCode = "" Then
                            MsgBox "Please Enter Payor", vbOKOnly + vbCritical, DialogBoxTitle
                            cboPAYCode.SetFocus
                        ElseIf txtMBMPolNo = "" Then
                            MsgBox "Please Enter Policy No", vbOKOnly + vbCritical, DialogBoxTitle
                            txtMBMPolNo.SetFocus
                        ElseIf cboHLTCode = "" Then
                            MsgBox "Please Enter Health Code", vbOKOnly + vbCritical, DialogBoxTitle
                            cboHLTCode.SetFocus
                        ElseIf txtMBMAdd1 = "" Then
                            MsgBox "Please Enter Address", vbOKOnly + vbCritical, DialogBoxTitle
                            txtMBMAdd1.SetFocus
                        ElseIf txtMBMBirthDate = "" Then
                            MsgBox "Please Enter BirthDate", vbOKOnly + vbCritical, DialogBoxTitle
                            txtMBMBirthDate.SetFocus
                        ElseIf cboINSCode = "" Then
                            MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
                            cboINSCode.SetFocus
                        ElseIf cboPLNCode = "" Then
                            MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
                            cboPLNCode.SetFocus
                        ElseIf txtMBMPayorEffDate = "" Then
                            MsgBox "Please Enter Effective Date", vbOKOnly + vbCritical, DialogBoxTitle
                            txtMBMPayorEffDate.SetFocus
                        ElseIf txtMBMPayorExpDate = "" Then
                            MsgBox "Please Enter Expiry Date", vbOKOnly + vbCritical, DialogBoxTitle
                            txtMBMPayorExpDate.SetFocus
                        ElseIf txtMBMBordxDate = "" Then
                            MsgBox "Please Enter Bordx Date", vbOKOnly + vbCritical, DialogBoxTitle
                            txtMBMBordxDate.SetFocus
                                                    
                            
                       
                       
                       
                        ElseIf Mid(cboStatus, 1, 1) = "D" Then
                            MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                        ElseIf Mid(cboStatus, 1, 1) = "C" Then
                            MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                        Else
                            Call InsertNewRegistrationSub
                            bSucess = True
                        End If
                        
                        
                'ElseIf Mid(cboStatus, 1, 1) = "A" Then
                        ' delete Supplementary
                    ElseIf Mid(cboEndorseStatus, 1, 1) = "T" Then
                        If Mid(cboStatus, 1, 1) = "D" Then
                            MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                        ElseIf Mid(cboStatus, 1, 1) = "C" Then
                            MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                        ElseIf Mid(cboStatus, 1, 1) = "S" Then
                            MsgBox "Membership has been Suspended. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                    ElseIf txtDate = "__/__/____" Then
                        MsgBox "Please Enter EndtEff Date", vbOKOnly + vbCritical, DialogBoxTitle
                        txtDate.SetFocus
                    ElseIf txtEndtBordxDate = "__/__/____" Then
                        MsgBox "Please Enter EndtBordx Date", vbOKOnly + vbCritical, DialogBoxTitle
                        txtEndtBordxDate.SetFocus

                        Else
                            Call DeleteSupplementary
                            bSucess = True
                        End If
                        ' Add  Supplementary
                 ElseIf Mid(cboEndorseStatus, 1, 1) = "S" Then
                     If Mid(cboStatus, 1, 1) = "D" Then
                         MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     ElseIf Mid(cboStatus, 1, 1) = "C" Then
                         MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     ElseIf Mid(cboStatus, 1, 1) = "S" Then
                         MsgBox "Membership has been suspended. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     ElseIf txtDate = "__/__/____" Then
                         MsgBox "Please Enter EndtEffDate Date", vbOKOnly + vbCritical, DialogBoxTitle
                         txtDate.SetFocus
                    ElseIf txtEndtBordxDate = "__/__/____" Then
                        MsgBox "Please Enter EndtBordx Date", vbOKOnly + vbCritical, DialogBoxTitle
                        txtEndtBordxDate.SetFocus
                     Else
                '     RecordSaved = vbYes
                         Call AddSupplementaryAdjustment
                         bSucess = True
                     End If
                 ' Susspend MEMbership
                 ElseIf Mid(cboEndorseStatus, 1, 1) = "X" Then
                     If Mid(cboStatus, 1, 1) = "D" Then
                         MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     ElseIf Mid(cboStatus, 1, 1) = "C" Then
                         MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     ElseIf txtDate = "__/__/____" Then
                        MsgBox "Please Enter EndtEff Date", vbOKOnly + vbCritical, DialogBoxTitle
                        txtDate.SetFocus
                    'ElseIf txtEndtBordxDate = "__/__/____" Then
                    '    MsgBox "Please Enter EndtBordx Date", vbOKOnly + vbCritical, DialogBoxTitle
                    '    txtEndtBordxDate.SetFocus

                     Else
                        Call SuspendPrincipal
                        Call SuspendSupplementary
                        MsgBox "Record is Suspended", vbOKOnly + vbInformation, DialogBoxTitle
                        bSucess = True
                     End If
                 ElseIf Mid(cboEndorseStatus, 1, 1) = "Y" Then
                    If Mid(cboStatus, 1, 1) = "D" Then
                         MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                        ElseIf Mid(cboStatus, 1, 1) = "C" Then
                             MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                        ElseIf txtDate = "__/__/____" Then
                            MsgBox "Please Enter EndtEff Date", vbOKOnly + vbCritical, DialogBoxTitle
                        txtDate.SetFocus
                     Else
                        Call ReactivatePrincipal
                        Call ReactivateSupplementary
                        MsgBox "Record is Activated", vbOKOnly + vbInformation, DialogBoxTitle
                        bSucess = True
                    End If
                 ' Amendment
                 ElseIf Mid(cboEndorseStatus, 1, 1) = "A" Then
                     If txtEndtBordxDate = "__/__/____" Then
                        MsgBox "Please Enter EndtBordx Date", vbOKOnly + vbCritical, DialogBoxTitle
                        txtEndtBordxDate.SetFocus
                     ElseIf txtDate = "__/__/____" Then
                         MsgBox "Please Enter EndtEff Date", vbOKOnly + vbCritical, DialogBoxTitle
                         txtDate.SetFocus
                     ElseIf Mid(cboStatus, 1, 1) = "C" Then
                         MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     ElseIf Mid(cboStatus, 1, 1) = "D" Then
                         MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                    ElseIf Mid(cboStatus, 1, 1) = "S" Then
                        MsgBox "Membership has been Suspended. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     Else
                         Call SavePrincipalAdjustment
                         Call SaveSupplementaryAdjustment
                         MsgBox "Updated Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
                         bSucess = True
                     End If
              
                'ElseIf cboStatus = "D" Then
                '    strsql = "Select CASNumber from CAS Where MBMNumber ='" & txtMBMNumber & "' And CASStatus = 'O'"
                '    CAS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                
               '     If CAS.RecordCount > 0 Then
               '         confirmCancel = MsgBox("There is at least an outstanding case under this member!", vbYesNo + vbCritical)
               '         If confirmCancel Then
               '             Call DeleteSupplementary
               '         End If
               '     Else
                        'RecordSaved = vbYes
                        
               '     End If
                ElseIf Mid(cboEndorseStatus, 1, 1) = "D" Then
                     If txtEndtBordxDate = "__/__/____" Then
                            MsgBox "Please Enter EndtBordx Date", vbOKOnly + vbCritical, DialogBoxTitle
                         txtEndtBordxDate.SetFocus
                     ElseIf txtDate = "__/__/____" Then
                         MsgBox "Please Enter EndtEff Date", vbOKOnly + vbCritical, DialogBoxTitle
                         txtDate.SetFocus
                     ElseIf Mid(cboStatus, 1, 1) = "C" Then
                         MsgBox "Membership has been canceled. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                     ElseIf Mid(cboStatus, 1, 1) = "D" Then
                         MsgBox "Membership has been deleted. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle
                    ElseIf Mid(cboStatus, 1, 1) = "S" Then
                        MsgBox "Membership has been Suspended. No further changes are allowed!", vbOKOnly + vbCritical, DialogBoxTitle

                     Else
                        Call DeletePrincipals
                        bSucess = True
                        MsgBox "Deleted Sucessfully", vbOKOnly + vbInformation, DialogBoxTitle
                     End If
                Else
                    If cboStatus = "D" Then
                            MsgBox "This Membership is Deleted", vbOKOnly + vbInformation, DialogBoxTitle
                    ElseIf cboStatus = "C" Then
                            MsgBox "This Membership is Cancelled", vbOKOnly + vbInformation, DialogBoxTitle
                    ElseIf cboStatus = "S" And Mid(cboEndorseStatus, 1, 1) <> "C" Then
                            MsgBox "This Membership is Suspended", vbOKOnly + vbInformation, DialogBoxTitle
                    End If
                End If
            End If
            Screen.MousePointer = vbDefault
            If bSucess = True Then
                cmdEditMBM.Enabled = True
                cboEndorseStatus.Enabled = False
                cmdSave.Enabled = False
                Call ADJListing
            End If
            'If startTrans = True Then
            '    medixmasterconn.CommitTrans
            'End If
    'End If
    End If
    
   
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If startTrans = True Then
        medixmasterconn.RollbackTrans
    End If
    MsgBox ERR.Description
End Sub

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
    Dim Payor As String
    Dim AgeValue As String
    Dim PayEffDate As String
    Dim PAYExpDate As String
    Dim MBMDOB As String
    Dim ProviderRefVal As Double
    Dim ProviderEff As String
    
    HLTCode = Mid(cboHLTCode, 1, 1)
    If InStr(1, cboPLNCode, "-") > 1 Then
        PlanCode = Mid(cboPLNCode, 1, InStr(1, cboPLNCode, "-") - 1)
    Else
        PlanCode = Trim(Mid(cboPLNCode, 1, 4))
    End If
    InsuredCode = Mid(cboINSCode, 1, 1)
    Payor = Mid(cboPAYCode, 1, 2)
    MBMDOB = txtMBMBirthDate
    PayEffDate = txtMBMPayorEffDate
    PAYExpDate = txtMBMPayorExpDate
        
    strsql = "Select * from MBMRefund Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    MBMRefund.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic

    FindAge = GetAge(CDate(PayEffDate), CDate(MBMDOB))
    AgeValue = Mid(GetAgeBand(HLTCode, FindAge), 1, 2)
    MBMPremium = GetPremium(InsuredCode, Payor, AgeValue, PlanCode, HLTCode, CDate(PayEffDate))
   ' MBMPremium = txtMBMPremium
    If AgeValue = "01" Then
        MBMMCOFees = GetMCO(InsuredCode, HLTCode, "C", CDate(PayEffDate), Trim(Mid(cboPLNCode, 1, 4)))
    Else
        MBMMCOFees = GetMCO(InsuredCode, HLTCode, "A", CDate(PayEffDate), Trim(Mid(cboPLNCode, 1, 4)))
    End If
    
    ServiceChargeRefundValue = GetProviderCharges(Trim(Mid(cboSRVCode, 1, 2)), Trim(cboINSCode), Trim(cboHLTCode), txtMBMAdultChild)
    
    MCORefundValue = MCORefund(CDate(PAYExpDate), CDate(txtDate), CDate(PayEffDate), MBMMCOFees)
    PremiumRefundValue = MCOPremiumReduction(CDate(PAYExpDate), CDate(txtDate), CDate(PayEffDate), MBMPremium)
    ProviderRefVal = EAFeesReduction(CDate(PAYExpDate), CDate(txtDate), CDate(PayEffDate), ServiceChargeRefundValue)
    
    PremiumEffective = CDbl(MBMPremium) - CDbl(PremiumRefundValue)
    MCOEffective = CDbl(MBMMCOFees) - CDbl(MCORefundValue)
    ProviderEff = CDbl(txtMBMProviderCharges) - CDbl(ProviderRefVal)
        MBMRefund.AddNew
        With MBMRefund
            !MBMNumber = txtMBMNumber
            !MBMPREMEffective = PremiumEffective
            !MBMPremReduction = PremiumRefundValue
            !MBMMCOEffective = MCOEffective
            !MBMMCORefund = MCORefundValue
            !MBMProviderRef = ProviderRefVal
            !MBMProviderEff = ProviderEff
            !PAYCode = cboPAYCode
            !MBMBordxDate = txtMBMBordxDate
            !MBMBatchNo = txtBatchNo
            !HLTCode = Mid(cboHLTCode, 1, 1)
        End With
        MBMRefund.Update
    txtMBMPremiumEff = Format(Round(PremiumEffective, 0), "#,##0.00")
    txtMCOFeesEff = Format(Round(MCOEffective, 0), "#,##0.00")
    txtMBMPremiumRed = Format(Round(PremiumRefundValue, 0), "#,##0.00")
    txtMBMMCORed = Format(Round(MCORefundValue, 0), "#,##0.00")
    txtMBMProviderRedCharges = Format(Round(ProviderRefVal, 1), "#,##0.00")
    txtMBMProviderEffCharges = Format(Round(ProviderEff, 1), "#,##0.00")
End Sub

Private Sub SavePrincipalAdjustment()

    Dim MemberAmendHis As New ADODB.Recordset
    Dim MBM As New ADODB.Recordset
    Dim MBMOthers As New ADODB.Recordset
    Dim strsql As String
    Dim strsqlOthers As String
    Dim RecordSaved
    
    strsql = "Select * from MBM Where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    MemberAmendHis.Open "MBMAmendHistory", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    MemberAmendHis.AddNew
    With MemberAmendHis
        
        !MBMNumber = ChkStr(MBM!MBMNumber)
        If MBM!MBMPolicyNo <> "" Then
            !MBMPolicyNo = ChkStr(MBM!MBMPolicyNo)
        ElseIf txtMBMPolNo <> "" Then
            !MBMPolicyNo = ChkStr(txtMBMPolNo)
        End If
        !MBMIcBcPp = ChkStr(MBM!MBMIcBcPp)
        !MBMName = ChkStr(MBM!MBMName)
        !MBMExpiryDate = MBM!MBMPayorEXPDate
        !MBMPlaNCode = ChkStr(MBM!PLNCode)
        !MBMAmendEdtType = Mid(cboEndorseStatus, 1, 1)
        If IsDate(txtDate) Then
            !MBMAEndtEffDate = txtDate
        End If
        !MBMAmendUser = ChkStr(Logon)
        !MBMAEndtBatchNo = txtBatchNo
        If IsDate(txtEndtBordxDate) Then
            !MBMAENDtBordxDate = txtEndtBordxDate
        End If
        !MBMAmendDate = Date
        MemberAmendHis.Update
    End With
    
    With MBM
        !MBMSalutation = ChkStr(cboMBMSalutation)
        !MBMName = ChkStr(txtMBMName)
        !MBMPolicyNo = Trim(ChkStr(txtMBMPolNo))
        !MBMSex = ChkStr(cboMBMSex)
        !MBMName = ChkStr(txtMBMName)
        !MBMAddress1 = Trim(ChkStr(txtMBMAdd1))
        !MBMAddress2 = Trim(ChkStr(txtMBMAdd2))
        !MBMAddress3 = Trim(ChkStr(txtMBMAdd3))
        !CTYCode = Trim(ChkStr(cboCTYCode))
        !STACode = Trim(ChkStr(cboSTACode))
        !MBMPostCode = txtMBMPostCode
        '!MBMEmail = ChkStr(txtMBMEmail)
        !MBMIcBcPp = ChkStr(txtMBMIcBcPP)
        !MBMDOB = txtMBMBirthDate
        !RACCode = Trim(Mid(cboRACCode, 1, 2))
        !LGNCode = Trim(Mid(FrmMBMOthers.cboLGNCode, 1, 1))
        If InStr(1, cboNATCode, "-") > 0 Or Len(cboNATCode) = 2 Then
            If InStr(1, cboNATCode, "-") > 0 Then
                !NATCode = Trim(Mid(cboNATCode, 1, InStr(1, cboNATCode, "-") - 1))
            Else
                !NATCode = Trim(Mid(cboNATCode, 1, 2))
            End If
        End If
        
        If IsDate(txtMBMPayorEffDate) Then
            !MBMPayorEffDate = txtMBMPayorEffDate
        End If
        If IsDate(txtMBMPayorExpDate) Then
            !MBMPayorEXPDate = txtMBMPayorExpDate
        End If
        !BRNCode = ChkStr(cboBRNCode)
        !MBMTelnoH = FrmMBMOthers.txtMBMTelNoH
        !MBMTelNoM = FrmMBMOthers.txtMBMTelNoM
        !MBMTelNoO = FrmMBMOthers.txtMBMTelNoO
        !MBMLastAdjustmentDate = Date
        !MBMLastUpdateUser = Logon
        !MBMLastEndorseStatus = Trim(Mid(cboEndorseStatus, 1, InStr(1, cboEndorseStatus, "-") - 1))
        !MBMBordxDate = txtMBMBordxDate
        !MBMTakeOver = Mid(cboMBMTakeOver, 1, 1)
        !MBMRenewFlag = Mid(cboMBMRenewal, 1, 1)
        If Mid(ChkStr(cboMBMRelCode), 1, 1) = "C" Then
                !MBMMemberType = "P"
             Else
                !MBMMemberType = "Q"
        End If
        !MBMDataStatus = Mid(cboDataStatus, 1, 1)
        MBM.Update
    End With
    
        If MBMPayorEXPDate = 1 Then
            RecordSaved = MsgBox("Do you want to edit Expiry Date?", vbYesNo + vbExclamation)
            If RecordSaved = vbYes Then
                Call InsertMCO
                Call InsertPrem
                Call InsertProviderFees
            End If
         End If

    ' INsert into Membership Other details table
    strsqlOthers = "select * from MBMOthers where MBMNumber = '" & Trim(txtMBMNumber) & "'"
    MBMOthers.Open strsqlOthers, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MBMOthers.recordCount = 0 Then
            MBMOthers.AddNew
            MBMOthers!MBMNumber = txtMBMNumber
        End If
            With MBMOthers
                 !MBMWeight = ChkStr(FrmMBMOthers.txtMBMWeight)
                 !MBMHeight = ChkStr(FrmMBMOthers.txtMBMHeight)
                 If InStr(1, FrmMBMOthers.txtMBMWorkNature, "-") > 0 Then
                    !MBMWorkNature = Trim(Mid(FrmMBMOthers.txtMBMWorkNature, 1, InStr(1, FrmMBMOthers.txtMBMWorkNature, "-") - 1))
                 Else
                   !MBMWorkNature = Trim(ChkStr(FrmMBMOthers.txtMBMWorkNature))
                 End If
                 !MBMExclusion = ChkStr(txtMBMExclusion)
                 !MBMAllergic = ChkStr(FrmMBMOthers.txtMBMAllergic)
                 !MBMDepartment = ChkStr(txtDeptName)
                 !MBMGroupCompany = ChkStr(txtMBMGroupCompany)
                 !MBMPolSubNo1 = FrmMBMOthers.TxtPOLSubNo
                 If FrmMBMOthers.Check1.value = 1 Then
                    !LabelStatus = "Y"
                 Else
                    !LabelStatus = Null
                 End If
                 !MBMEmployeeNo = ChkStr(txtMBMEmployeeNo)
                 !MBMBlood = Trim(Mid(FrmMBMOthers.cboMBMBlood, 1, 1))
                 !MBMSmoking = Trim(Mid(FrmMBMOthers.cboMBMSmoking, 1, 1))
                 !MBMAlcohol = Mid(FrmMBMOthers.cboMBMAlcohol, 1, 1)
                 !MBMAgentCode = ChkStr(txtMBMAgentCode)
                 !RELCODE = "P"
                 !MBMInstallment = Trim(Mid(CboInsMode, 1, 1))
                 If Len(Trim(cboMBMMaritalStatus)) Then
                    !MBMMaritalStatus = Mid(cboMBMMaritalStatus, 1, 1)
                 End If
                 'If Len(FrmMBMOthers.txtMBMPrevPlan) Then
                 '   !MBMPrevPLNCode = ChkStr(FrmMBMOthers.txtMBMPrevPlan)
                 'End If
                 If Len(FrmMBMOthers.txtMBMRBAmt) Then
                    !MBMRBAmt = FrmMBMOthers.txtMBMRBAmt
                 End If
                 If Len(FrmMBMOthers.txtMBMPAYRemarks) Then
                    !MBMPAYRemarks = ChkStr(FrmMBMOthers.txtMBMPAYRemarks)
                 End If
                 If Len(FrmMBMOthers.txtMBMPrevINSCom) Then
                    !MBMPrevInsCom = ChkStr(FrmMBMOthers.txtMBMPrevINSCom)
                 End If

                '!MBMFirstMEMNo = Trim(txtFirstMemNo)
                '!MBMFirstPolNo = txtFirstPolNo
                '!MBMPrevPolNo = txtMBMPrevPolNo
                '!MBMPrevMemNo = txtMBMPrevMEMNo
                If Len(txtMCOFees) Then
                    !MBMMCOFee = txtMCOFees
                End If
                
                If Len(txtMBMPremium) Then
                    !MBMPremium = txtMBMPremium
                End If
                
                If Len(txtMBMProviderCharges) Then
                    !MBMAccessFee = txtMBMProviderCharges
                End If
                If txtMBMDateJoin <> "__/__/____" Then
                    !MBMFirstJoinedDate = txtMBMDateJoin
                End If
                If InStr(1, FrmMBMOthers.cboMBMOccupation, "-") > 0 Then
                    !MBMOccupation = Trim(Mid(FrmMBMOthers.cboMBMOccupation, 1, InStr(1, FrmMBMOthers.cboMBMOccupation, "-") - 1))
                Else
                    !MBMOccupation = Trim(ChkStr(FrmMBMOthers.cboMBMOccupation))
                End If
                If txtRemarks = "" Then
                    !MBMRemarks = ""
                Else
                    !MBMRemarks = ChkStr(txtRemarks)
                End If
                If Len(FrmMBMOthers.cboCOPayRB) Then
                    !MBMCoPayRB = FrmMBMOthers.cboCOPayRB
                Else
                    !MBMCoPayRB = ""
                End If
                
             '=================== Add By CK ================
                If Len(FrmMBMOthers.TxtGenWP) Then
                    !MBMGenWP = FrmMBMOthers.TxtGenWP
                Else
                    !MBMGenWP = Null
                End If
                If FrmMBMOthers.txtGENWPEnd <> "__/__/____" And FrmMBMOthers.txtGENWPEnd <> "" Then
                   !MBMGenEndWP = FrmMBMOthers.txtGENWPEnd
                Else
                    !MBMGenEndWP = Null
                End If
                If Len(FrmMBMOthers.txtPREWP) Then
                    !MBMPreExWP = FrmMBMOthers.txtPREWP
                Else
                    !MBMPreExWP = Null
                End If
                If FrmMBMOthers.txtPREWPEnd <> "__/__/____" And FrmMBMOthers.txtPREWPEnd <> "" Then
                   !MBMPreExEndWP = FrmMBMOthers.txtPREWPEnd
                Else
                    !MBMPreExEndWP = Null
                End If
                If Len(FrmMBMOthers.TxtSpeWP) Then
                    !MBMSpeIllWP = Trim(FrmMBMOthers.TxtSpeWP)
                Else
                    !MBMSpeIllWP = Null
                End If
                If FrmMBMOthers.txtSPEWPEnd <> "__/__/____" And FrmMBMOthers.txtSPEWPEnd <> "" Then
                   !MBMSpeIllEndWP = FrmMBMOthers.txtSPEWPEnd
                Else
                    !MBMSpeIllEndWP = Null
                End If
                If Len(FrmMBMOthers.cboStaffDisc) Then
                    !MBMStaffDec = Mid(FrmMBMOthers.cboStaffDisc, 1, 1)
                Else
                    !MBMStaffDec = ""
                End If
                If FrmMBMOthers.TxtClmNonPayFr <> "__/__/____" And FrmMBMOthers.TxtClmNonPayFr <> "" Then
                   !MBMCLMNonPayFr = FrmMBMOthers.TxtClmNonPayFr
                Else
                    !MBMCLMNonPayFr = Null
                End If
                If FrmMBMOthers.TxtClmNonPayTo <> "__/__/____" And FrmMBMOthers.TxtClmNonPayTo <> "" Then
                   !MBMCLMNonPayTo = FrmMBMOthers.TxtClmNonPayTo
                Else
                    !MBMCLMNonPayTo = Null
                End If
             '=================== Add By CK ================

                If Len(FrmMBMOthers.txtPLNCode) Then
                    !MBMNewPlanCode = FrmMBMOthers.txtPLNCode
                Else
                    !MBMNewPlanCode = Null
                End If
                If FrmMBMOthers.txtNewPolEffDate <> "__/__/____" And FrmMBMOthers.txtNewPolEffDate <> "" Then
                    !MBMNewPolEffDate = FrmMBMOthers.txtNewPolEffDate
                Else
                    !MBMNewPolEffDate = Null
                End If
                If Len(FrmMBMOthers.cboBTBG) Then
                    !MBMBTBG = Mid(FrmMBMOthers.cboBTBG, 1, 1)
                Else
                    !MBMBTBG = Null
                End If
                
                If Len(FrmMBMOthers.txtPayPrem) Then
                    !MBMPayorPremGross = FrmMBMOthers.txtPayPrem
                End If
                
                If Len(FrmMBMOthers.txtPayPremNett) Then
                    !MBMPayorPremNet = FrmMBMOthers.txtPayPremNett
                End If
                
                If Len(FrmMBMOthers.txtPayMCOGross) Then
                    !MBMPayorMCOGross = FrmMBMOthers.txtPayMCOGross
                End If
                
                If Len(FrmMBMOthers.txtPayMCONett) Then
                    !MBMPayorMCONet = FrmMBMOthers.txtPayMCONett
                End If

                If Len(FrmMBMOthers.txtPolicyDisc) Then
                    !MBMPolicyDisc = FrmMBMOthers.txtPolicyDisc
                End If

                If Len(FrmMBMOthers.txtRenewalDisc) Then
                    !MBMRenewalDisc = FrmMBMOthers.txtRenewalDisc
                End If

         End With
         MBMOthers.Update
            '!MBMUsualDoctor = txtMBMUsualDoctor
            '!MBMClinicName = txtMBMClinicName
            '!MBMClinicTelNo = txtMBMClinicTelNo
            '!MBMClininal = txtMBMClinical
            '!MBMClinicAdd = txtMBMClinicAdd
    'End If
        MemberAmendHis.Close
        Set MemberAmendHis = Nothing
        MBM.Close
        Set MBM = Nothing
        MBMOthers.Close
        Set MBMOthers = Nothing
End Sub

Private Sub SaveSupplementaryAdjustment()
'On Error Resume Next
Dim Arraycnt, arrayCurrRow As Integer
Dim strsql As String
Dim ListCount As Integer
Dim strsql2 As String
Dim MBMCov As New ADODB.Recordset
Dim LastCoveredID As Integer
Dim listloop As Integer
    
    strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    LastCoveredID = GetLastCoveredID(txtMBMNumber)
        ListCount = lstCovAdd.listitems.count
        For listloop = 1 To ListCount
        If lstCovAdd.listitems(listloop).SubItems(11) = "E" Then
            With MBMCov
                MBMCov.Filter = ("MBMCCoverID = '" & lstCovAdd.listitems(listloop).SubItems(10) & "'")
                  !MBMCName = ChkStr(lstCovAdd.listitems(listloop).Text)
                  !MBMCSalutation = ChkStr(lstCovAdd.listitems(listloop).SubItems(1))
                  !MBMCIcBcPpNo = ChkStr(lstCovAdd.listitems(listloop).SubItems(2))
                  !MBMCDOB = lstCovAdd.listitems(listloop).SubItems(3)
                  !MBMCSex = ChkStr(lstCovAdd.listitems(listloop).SubItems(4))
                  !MBMCAdultChild = ChkStr(lstCovAdd.listitems(listloop).SubItems(5))
                  !MBMCRELCode = Mid(lstCovAdd.listitems(listloop).SubItems(6), 1, 1)
                  !MBMCBloodType = ChkStr(lstCovAdd.listitems(listloop).SubItems(7))
                  !MBMCAllergic = ChkStr(lstCovAdd.listitems(listloop).SubItems(8))
                  !MBMCExclusion = ChkStr(lstCovAdd.listitems(listloop).SubItems(9))
                  If Len(lstCovAdd.listitems(listloop).SubItems(12)) Then
                    !MBMCAnnualLimit = lstCovAdd.listitems(listloop).SubItems(12)
                  End If
                  If Len(lstCovAdd.listitems(listloop).SubItems(13)) Then
                    !MBMCOccupation = ChkStr(lstCovAdd.listitems(listloop).SubItems(13))
                  End If
                  If Len(lstCovAdd.listitems(listloop).SubItems(14)) Then
                    !MBMCWorkNature = ChkStr(lstCovAdd.listitems(listloop).SubItems(14))
                  End If
                  If Len(lstCovAdd.listitems(listloop).SubItems(15)) Then
                    !MBMCSmoking = lstCovAdd.listitems(listloop).SubItems(15)
                  End If
                  If Len(lstCovAdd.listitems(listloop).SubItems(16)) Then
                    !MBMCAlcohol = lstCovAdd.listitems(listloop).SubItems(16)
                  End If
                  If Len(lstCovAdd.listitems(listloop).SubItems(17)) Then
                    !MBMCAvailableLimit = lstCovAdd.listitems(listloop).SubItems(17)
                  End If
                  If Len(lstCovAdd.listitems(listloop).SubItems(18)) Then
                       !MBMCPremium = lstCovAdd.listitems(listloop).SubItems(18)
                  End If
                  If Len(lstCovAdd.listitems(listloop).SubItems(20)) Then
                       !MBMCPrevPLNCode = ChkStr(lstCovAdd.listitems(listloop).SubItems(20))
                  End If
                  If Len(lstCovAdd.listitems(listloop).SubItems(21)) Then
                       !MBMCRBAmt = ChkStr(lstCovAdd.listitems(listloop).SubItems(21))
                  End If
                  If Len(lstCovAdd.listitems(listloop).SubItems(22)) Then
                       !MBMCPAYRemarks = ChkStr(lstCovAdd.listitems(listloop).SubItems(22))
                  End If
                  If Len(lstCovAdd.listitems(listloop).SubItems(23)) Then
                       !MBMCPrevInsCom = ChkStr(lstCovAdd.listitems(listloop).SubItems(23))
                  End If
                
                    If Len(lstCovAdd.listitems(listloop).SubItems(24)) Then
                        !MBMCGenWP = ChkStr(lstCovAdd.listitems(listloop).SubItems(24))
                    Else
                        !MBMCGenWP = Null
                    End If
                    If lstCovAdd.listitems(listloop).SubItems(25) <> "" And lstCovAdd.listitems(listloop).SubItems(25) <> "__/__/____" Then
                        !MBMCGenEndWP = ChkStr(lstCovAdd.listitems(listloop).SubItems(25))
                    Else
                        !MBMCGenEndWP = Null
                    End If
                    If Len(lstCovAdd.listitems(listloop).SubItems(26)) Then
                        !MBMCPreExWP = ChkStr(lstCovAdd.listitems(listloop).SubItems(26))
                    Else
                        !MBMCPreExWP = Null
                    End If
                    If lstCovAdd.listitems(listloop).SubItems(27) <> "" And lstCovAdd.listitems(listloop).SubItems(27) <> "__/__/____" Then
                        !MBMCPreExEndWP = ChkStr(lstCovAdd.listitems(listloop).SubItems(27))
                    Else
                        !MBMCPreExEndWP = Null
                    End If
                    If Len(lstCovAdd.listitems(listloop).SubItems(28)) Then
                        !MBMCSpeIllWP = ChkStr(lstCovAdd.listitems(listloop).SubItems(28))
                    Else
                        !MBMCSpeIllWP = Null
                    End If
                    If lstCovAdd.listitems(listloop).SubItems(29) <> "" And lstCovAdd.listitems(listloop).SubItems(29) <> "__/__/____" Then
                        !MBMSpeIllEndWP = ChkStr(lstCovAdd.listitems(listloop).SubItems(29))
                    Else
                        !MBMSpeIllEndWP = Null
                    End If
                    If Len(lstCovAdd.listitems(listloop).SubItems(30)) Then
                        !MBMCStaffDec = ChkStr(lstCovAdd.listitems(listloop).SubItems(30))
                    End If
                    If lstCovAdd.listitems(listloop).SubItems(31) <> "" And lstCovAdd.listitems(listloop).SubItems(31) <> "__/__/____" Then
                        !MBMCCLMNonPayFr = ChkStr(lstCovAdd.listitems(listloop).SubItems(31))
                    Else
                        !MBMCCLMNonPayFr = Null
                    End If
                    If lstCovAdd.listitems(listloop).SubItems(32) <> "" And lstCovAdd.listitems(listloop).SubItems(32) <> "__/__/____" Then
                        !MBMCCLMNonPayTo = ChkStr(lstCovAdd.listitems(listloop).SubItems(32))
                    Else
                        !MBMCCLMNonPayTo = Null
                    End If
                    
                    If lstCovAdd.listitems(listloop).SubItems(33) <> "" And lstCovAdd.listitems(listloop).SubItems(33) <> "__/__/____" Then
                        !MBMCBordxDate = ChkStr(lstCovAdd.listitems(listloop).SubItems(33))
                    Else
                        !MBMCBordxDate = Null
                    End If
                    
                    If lstCovAdd.listitems(listloop).SubItems(34) <> "" Then
                        !MBMCPayorPremGross = lstCovAdd.listitems(listloop).SubItems(34)
                    Else
                        !MBMCPayorPremGross = Null
                    End If
                    
                    If Len(lstCovAdd.listitems(listloop).SubItems(35)) Then
                        !MBMCPayorPremNet = lstCovAdd.listitems(listloop).SubItems(35)
                    End If
                    
                    If Len(lstCovAdd.listitems(listloop).SubItems(36)) Then
                        !MBMCPayorMCOGross = lstCovAdd.listitems(listloop).SubItems(36)
                    End If
                    
                    If Len(lstCovAdd.listitems(listloop).SubItems(37)) Then
                        !MBMCPayorMCONet = lstCovAdd.listitems(listloop).SubItems(37)
                    End If
                    
                    If Len(lstCovAdd.listitems(listloop).SubItems(38)) Then
                        !MBMCPolicyDisc = lstCovAdd.listitems(listloop).SubItems(38)
                    End If
                    
                    If Len(lstCovAdd.listitems(listloop).SubItems(39)) Then
                        !MBMCRenewalDisc = lstCovAdd.listitems(listloop).SubItems(39)
                    End If
                    
                  lstCovAdd.listitems(listloop).SubItems(11) = ""
                MBMCov.Update
            End With
            MBMCov.MoveNext
           'For arrayCurrRow = 1 To 10
           '     If lstString(arrayCurrRow) = 0 Then
           '         lstString(arrayCurrRow) = listloop
           '         Exit For
           '     End If
           ' Next

      End If
    Next listloop
    MBMCov.Close
    Set MBMCov = Nothing
    
End Sub

Private Sub DeleteSupplementary()
Dim MBM As New ADODB.Recordset
Dim strsql As String
Dim listrecord As Integer
Dim listloop As Integer
Dim ListCount As Integer
Dim DEL As New ADODB.Recordset
Dim LastCoveredID As Integer
Dim SelectAge As String
Dim strdelcov As String
Dim MBMCov As New ADODB.Recordset
Dim ForLoop As Integer
Dim planstring
Dim MBMCovPersons As New ADODB.Recordset


strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & txtMBMNumber & "'"
MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    LastCoveredID = GetLastCoveredID(txtMBMNumber)
        ListCount = lstCovAdd.listitems.count
        For listloop = 1 To ListCount
        If lstCovAdd.listitems(listloop).SubItems(11) = "R" Then
            strdelcov = "Delete from MBMCoveredPersons Where MBMCNumber ='" & txtMBMNumber & "' And MBMCCoverID ='" & lstCovAdd.listitems(listloop).SubItems(10) & "'"
            MBMCovPersons.Open strdelcov, medixmasterconn, adOpenKeyset, adLockOptimistic
            'lstCovAdd.ListItems(listloop).SubItems(11) = "W"
            'MBMCovPersons.MoveNext
      End If
    Next listloop
    
    Call DelSuppLoop
End Sub

Private Sub AddSupplementaryAdjustment()
    Dim listloop As Integer
    Dim ListCount As Integer
    Dim strsql As String
    Dim DEL As New ADODB.Recordset
    Dim LastCoveredID As Integer
    Dim SelectAge As String
    Dim strdelcov As String
    Dim MBMCov As New ADODB.Recordset
    Dim ForLoop As Integer
    Dim planstring
    Dim AgeValue As String
    Dim strsqlannlmt As String
    Dim AnnLmt As New ADODB.Recordset
    Dim strsqlPrem  As String
    Dim Prem As New ADODB.Recordset
    
    strsql = "Select * from MBMCoveredPersons "
    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
        LastCoveredID = GetLastCoveredID(txtMBMNumber)
        ListCount = lstCovAdd.listitems.count
        For listloop = 1 To ListCount
        If lstCovAdd.listitems(listloop).SubItems(11) = "N" Then
              MBMCov.AddNew
              With MBMCov
                  SelectAge = GetAge(txtMBMPayorEffDate, lstCovAdd.listitems(listloop).SubItems(3))
                  !MBMCAgeband = Mid(GetAgeBand("S", SelectAge), 1, 2)
                  AgeValue = Mid(GetAgeBand("S", SelectAge), 1, 2)
                  !MBMCStatus = Mid(cboStatus, 1, 1)
                  !MBMCBordxDate = txtMBMBordxDate
                  !MBMCNumber = ChkStr(txtMBMNumber)
                  !MBMCSalutation = ChkStr(lstCovAdd.listitems(listloop).SubItems(1))
                  !MBMCName = ChkStr(lstCovAdd.listitems(listloop).Text)
                  !MBMCIcBcPpNo = ChkStr(lstCovAdd.listitems(listloop).SubItems(2))
                  !MBMCDOB = lstCovAdd.listitems(listloop).SubItems(3)
                  !MBMCSex = ChkStr(lstCovAdd.listitems(listloop).SubItems(4))
                  !MBMCAdultChild = ChkStr(lstCovAdd.listitems(listloop).SubItems(5))
                  !MBMCRELCode = Mid(lstCovAdd.listitems(listloop).SubItems(6), 1, 1)
                  !MBMCBloodType = lstCovAdd.listitems(listloop).SubItems(7)
                  !MBMCAllergic = ChkStr(lstCovAdd.listitems(listloop).SubItems(8))
                  !MBMCExclusion = ChkStr(lstCovAdd.listitems(listloop).SubItems(9))
                  !MBMCCoverID = LastCoveredID
                  If Len(lstCovAdd.listitems(listloop).SubItems(12)) Then
                    !MBMCAnnualLimit = lstCovAdd.listitems(listloop).SubItems(12)
                  End If
                  If Len(lstCovAdd.listitems(listloop).SubItems(13)) Then
                    !MBMCOccupation = ChkStr(lstCovAdd.listitems(listloop).SubItems(13))
                  End If
                  If Len(lstCovAdd.listitems(listloop).SubItems(14)) Then
                    !MBMCWorkNature = ChkStr(lstCovAdd.listitems(listloop).SubItems(14))
                  End If
                  If Len(lstCovAdd.listitems(listloop).SubItems(15)) Then
                    !MBMCSmoking = lstCovAdd.listitems(listloop).SubItems(15)
                  End If
                  If Len(lstCovAdd.listitems(listloop).SubItems(16)) Then
                    !MBMCAlcohol = lstCovAdd.listitems(listloop).SubItems(16)
                  End If
                If lstCovAdd.listitems(listloop).SubItems(33) <> "" And lstCovAdd.listitems(listloop).SubItems(33) <> "__/__/____" Then
                    !MBMCBordxDate = ChkStr(lstCovAdd.listitems(listloop).SubItems(33))
                Else
                    !MBMCBordxDate = Null
                End If
                  
                !MBMCMemberType = "C"
                !MBMCRegDate = Date
                If LastCoveredID > 10 Then
                  '!MBMCCoverID = 0 & LastCoveredID
                  lstCovAdd.listitems(listloop).SubItems(11) = ""
                  lstCovAdd.listitems(listloop).SubItems(10) = LastCoveredID
                Else
                  !MBMCCoverID = 0 & LastCoveredID
                  lstCovAdd.listitems(listloop).SubItems(11) = ""
                  lstCovAdd.listitems(listloop).SubItems(10) = 0 & LastCoveredID
                End If
                If Len(lstCovAdd.listitems(listloop).SubItems(35)) Then
                    !MBMCPayorPremNet = lstCovAdd.listitems(listloop).SubItems(35)
                End If
                
                If Len(lstCovAdd.listitems(listloop).SubItems(36)) Then
                    !MBMCPayorMCOGross = lstCovAdd.listitems(listloop).SubItems(36)
                End If
                
                If Len(lstCovAdd.listitems(listloop).SubItems(37)) Then
                    !MBMCPayorMCONet = lstCovAdd.listitems(listloop).SubItems(37)
                End If
                
                If Len(lstCovAdd.listitems(listloop).SubItems(38)) Then
                    !MBMCPolicyDisc = lstCovAdd.listitems(listloop).SubItems(38)
                End If
                
                If Len(lstCovAdd.listitems(listloop).SubItems(39)) Then
                    !MBMCRenewalDisc = lstCovAdd.listitems(listloop).SubItems(39)
                End If
                  
                    strsqlannlmt = "Select SuppLimitStatus, PLNCode, SuppLimit from AnnualLimit where PLNCode = '" & Trim(cboPLNCode) & "'"
                    AnnLmt.Open strsqlannlmt, medixmasterconn, adOpenKeyset, adLockOptimistic

                    strsqlPrem = "Select SuppPrmStatus, PLNCode, SuppPrmAmt from PRM where PLNCode = '" & Trim(cboPLNCode) & "'"
                    Prem.Open strsqlPrem, medixmasterconn, adOpenKeyset, adLockOptimistic
                    
                    If AnnLmt.recordCount Then
                        If AnnLmt!SuppLimitStatus = "B" Then
                            If listloop = "1" Then
                               !MBMCAnnualLimit = AnnLmt!SuppLimit
                               !MBMCAvailableLimit = AnnLmt!SuppLimit
                            End If
                        ElseIf AnnLmt!SuppLimitStatus = "C" Then
                            If listloop = "1" Then
                                !MBMCAnnualLimit = AnnLmt!SuppLimit
                                !MBMCAvailableLimit = AnnLmt!SuppLimit
                            End If
                        End If
                    End If
                    
                    If Prem.recordCount Then
                        If Prem!SuppPrmStatus = "B" Then
                            If listloop = "1" Then
                                tempSPremium = Prem!SuppPrmAmt
                                 If Prem!SuppPrmAmt <> "" Then
                                     !MBMCBasicPremium = tempSPremium
                                 Else
                                     !MBMCBasicPremium = 0
                                 End If
                                 
                                 tempSPremium = Format(GetSPremium(Mid(cboINSCode, 1, 1), Mid(cboPAYCode, 1, 2), AgeValue, Trim(Mid(cboPLNCode, 1, 4)), Trim(cboHLTCode), CDate(txtMBMPayorEffDate)), "#,##0.00")
                                 tempSPremium = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(tempSPremium, "#,##0.00"))
                                 
                                 If Disc = True Then
                                     If lstCovAdd.listitems(listloop).SubItems(38) <> "" Then
                                         tempSPremium = tempSPremium * (100 - CDbl(lstCovAdd.listitems(listloop).SubItems(38))) / 100
                                     End If
                                     If lstCovAdd.listitems(listloop).SubItems(39) <> "" Then
                                         tempSPremium = tempSPremium * (100 - CDbl(lstCovAdd.listitems(listloop).SubItems(39))) / 100
                                     End If
                                 End If
                                 !MBMCPremium = tempSPremium
                            End If
                        ElseIf Prem!SuppPrmStatus = "C" Then
                            'If Len(Trim(lstCovAdd.listitems(listloop).SubItems(11))) Then
                                '!MBMCAnnualLimit = AnnLmt!SuppLimit
                                '!MBMCAvailableLimit = AnnLmt!SuppLimit
                                tempSPremium = Prem!SuppPrmAmt
                                If tempSPremium <> "" Then
                                    !MBMCBasicPremium = tempSPremium
                                Else
                                    !MBMCBasicPremium = 0
                                End If
                                
                                'tempSPremium = Format(GetSPremium(Mid(cboINSCode, 1, 1), Mid(txtPAYCode, 1, 2), AgeValue, Trim(Mid(cboPLNCode, 1, 4)), Trim(txtHLTCode), CDate(txtMBMPayorEffDate)), "#,##0.00")
                                tempSPremium = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(tempSPremium, "#,##0.00"))
                                
                                If Disc = True Then
                                    If lstCovAdd.listitems(listloop).SubItems(38) <> "" Then
                                        tempSPremium = tempSPremium * (100 - CDbl(lstCovAdd.listitems(listloop).SubItems(38))) / 100
                                    End If
                                    If lstCovAdd.listitems(listloop).SubItems(39) <> "" Then
                                        tempSPremium = tempSPremium * (100 - CDbl(lstCovAdd.listitems(listloop).SubItems(39))) / 100
                                    End If
                                End If
                                
                                !MBMCPremium = tempSPremium
                            End If
                        End If
                AnnLmt.Close
                Prem.Close
                Set AnnLmt = Nothing
                Set Prem = Nothing
                
                'If lstCovAdd.listitems(listloop).SubItems(34) <> "" Then
                '    !MBMPayorPrem = lstCovAdd.listitems(listloop).SubItems(34)
                'Else
                '    !MBMPayorPrem = Null
                'End If
                  
                  
                  LastCoveredID = LastCoveredID + 1
                 
             MBMCov.Update
             End With
      End If
    Next listloop
    
    MBMCov.Close
    Set MBMCov = Nothing
End Sub


' ################ START Key press effect ###################

'Private Sub cboMBMAlcohol_KeyPress(KeyAscii As Integer)
'On Error Resume Next
'Dim NewText As String
'Dim ValidCount As Integer
'Dim ValidValue As String
'Dim i As Integer
    
    'Clipboard.clear
    
    'If KeyAscii >= 32 And KeyAscii <> 127 Then
        'NewText = LCase(Left(cboMBMAlcohol.Text, cboMBMAlcohol.SelStart) + Chr(KeyAscii) + Mid(cboMBMAlcohol.Text, cboMBMAlcohol.SelStart + cboMBMAlcohol.SelLength + 1))
        
        'ValidCount = 0
        'ValidValue = ""
        
        'For i = 0 To cboMBMAlcohol.ListCount - 1
            
            'If NewText = LCase(Left(cboMBMAlcohol.List(i), Len(NewText))) Then
                    'ValidCount = ValidCount + 1
                    'ValidValue = cboMBMAlcohol.List(i)
            'End If
        'Next
    
    'If ValidCount <= 1 Then KeyAscii = 0
        
        'If ValidCount = 1 Then
            'cboMBMAlcohol.Text = ValidValue
            'cboMBMAlcohol.SelStart = 0
            'cboMBMAlcohol.SelLength = Len(ValidValue)
        'End If
    'End If
'End Sub


'Private Sub cboMBMSmoking_KeyPress(KeyAscii As Integer)
'On Error Resume Next
    'Dim NewText As String
    'Dim ValidCount As Integer
    'Dim ValidValue As String
    'Dim i As Integer
    
    'Clipboard.clear
    
    'If KeyAscii >= 32 And KeyAscii <> 127 Then
        'NewText = LCase(Left(cboMBMSmoking.Text, cboMBMSmoking.SelStart) + Chr(KeyAscii) + Mid(cboMBMSmoking.Text, cboMBMSmoking.SelStart + cboMBMSmoking.SelLength + 1))
        
        'ValidCount = 0
        'ValidValue = ""
        
        'For i = 0 To cboMBMSmoking.ListCount - 1
            
            'If NewText = LCase(Left(cboMBMSmoking.List(i), Len(NewText))) Then
                    'ValidCount = ValidCount + 1
                    'ValidValue = cboMBMSmoking.List(i)
            'End If
        'Next
    
    'If ValidCount <= 1 Then KeyAscii = 0
        
        'If ValidCount = 1 Then
            'cboMBMSmoking.Text = ValidValue
            'cboMBMSmoking.SelStart = 0
            'cboMBMSmoking.SelLength = Len(ValidValue)
        'End If
    'End If
'End Sub



'Private Sub cboMBMBlood_KeyPress(KeyAscii As Integer)
'On Error Resume Next
    'Dim NewText As String
    'Dim ValidCount As Integer
    'Dim ValidValue As String
    'Dim i As Integer
    
    'Clipboard.clear
    
    'If KeyAscii >= 32 And KeyAscii <> 127 Then
        'NewText = LCase(Left(cboMBMBlood.Text, cboMBMBlood.SelStart) + Chr(KeyAscii) + Mid(cboMBMBlood.Text, cboMBMBlood.SelStart + cboMBMBlood.SelLength + 1))
        
        'ValidCount = 0
        'ValidValue = ""
        
        'For i = 0 To cboMBMBlood.ListCount - 1
            
            'If NewText = LCase(Left(cboMBMBlood.List(i), Len(NewText))) Then
                    'ValidCount = ValidCount + 1
                    'ValidValue = cboMBMBlood.List(i)
            'End If
        'Next
    
    'If ValidCount <= 1 Then KeyAscii = 0
        
        'If ValidCount = 1 Then
            'cboMBMBlood.Text = ValidValue
            'cboMBMBlood.SelStart = 0
            'cboMBMBlood.SelLength = Len(ValidValue)
        'End If
    'End If
'End Sub



Private Sub cboHLTcode_KeyPress(KeyAscii As Integer)

On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    'Clipboard.clear
    
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

Private Sub cboPAYCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboPAYCode.Text, cboPAYCode.SelStart) + Chr(KeyAscii) + Mid(cboPAYCode.Text, cboPAYCode.SelStart + cboPAYCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboPAYCode.ListCount - 1
            
            If NewText = LCase(Left(cboPAYCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboPAYCode.List(i)
            End If
        Next
    
    If ValidCount <= 1 Then KeyAscii = 0
        
        If ValidCount = 1 Then
            cboPAYCode.Text = ValidValue
            cboPAYCode.SelStart = 0
            cboPAYCode.SelLength = Len(ValidValue)
        End If
    End If
End Sub


Private Sub cboMBMRelCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    'Clipboard.clear
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboMBMRelCode.Text, cboMBMRelCode.SelStart) + Chr(KeyAscii) + Mid(cboMBMRelCode.Text, cboMBMRelCode.SelStart + cboMBMRelCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboMBMRelCode.ListCount - 1
            
            If NewText = LCase(Left(cboMBMRelCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboMBMRelCode.List(i)
            End If
        Next
    
    If ValidCount <= 1 Then KeyAscii = 0
        
        If ValidCount = 1 Then
            cboMBMRelCode.Text = ValidValue
            cboMBMRelCode.SelStart = 0
            cboMBMRelCode.SelLength = Len(ValidValue)
        End If
    End If
 
End Sub

'************* CK Edit 27/06/2001 ********************8
'Private Sub cboBRNCode_KeyPress(KeyAscii As Integer)
'On Error Resume Next
    'Dim NewText As String
    'Dim ValidCount As Integer
    'Dim ValidValue As String
    'Dim i As Integer
    
    'If KeyAscii >= 32 And KeyAscii <> 127 Then
        'NewText = LCase(Left(cboBRNCode.text, cboBRNCode.SelStart) + Chr(KeyAscii) + Mid(cboBRNCode.text, cboBRNCode.SelStart + cboBRNCode.SelLength + 1))
        
        'ValidCount = 0
        'ValidValue = ""
        
        'For i = 0 To cboBRNCode.ListCount - 1
            
            'If NewText = LCase(Left(cboBRNCode.list(i), Len(NewText))) Then
                    'ValidCount = ValidCount + 1
                    'ValidValue = cboBRNCode.list(i)
            'End If
        'Next
    
    'If ValidCount <= 1 Then KeyAscii = 0
        
        'If ValidCount = 1 Then
            'cboBRNCode.text = ValidValue
            'cboBRNCode.SelStart = 0
            'cboBRNCode.SelLength = Len(ValidValue)
        'End If
    'End If
'End Sub
'************* CK Edit 27/06/2001 ********************8


Private Sub cboCovRelationship_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    'Clipboard.clear
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboCovRelationship.Text, cboCovRelationship.SelStart) + Chr(KeyAscii) + Mid(cboCovRelationship.Text, cboCovRelationship.SelStart + cboCovRelationship.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboCovRelationship.ListCount - 1
            
            If NewText = LCase(Left(cboCovRelationship.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboCovRelationship.List(i)
            End If
        Next
    
    If ValidCount <= 1 Then KeyAscii = 0
        
        If ValidCount = 1 Then
            cboCovRelationship.Text = ValidValue
            cboCovRelationship.SelStart = 0
            cboCovRelationship.SelLength = Len(ValidValue)
        End If
    End If
            
End Sub

Private Sub cboCovSalutation_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
    Clipboard.Clear
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboCovSalutation.Text, cboCovSalutation.SelStart) + Chr(KeyAscii) + Mid(cboCovSalutation.Text, cboCovSalutation.SelStart + cboCovSalutation.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboCovSalutation.ListCount - 1
            
            If NewText = LCase(Left(cboCovSalutation.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboCovSalutation.List(i)
            End If
        Next
    
    If ValidCount <= 1 Then KeyAscii = 0
        
        If ValidCount = 1 Then
            cboCovSalutation.Text = ValidValue
            cboCovSalutation.SelStart = 0
            cboCovSalutation.SelLength = Len(ValidValue)
        End If
    End If
            
End Sub

Private Sub cboCovSex_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    'Clipboard.clear
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboCovSex.Text, cboCovSex.SelStart) + Chr(KeyAscii) + Mid(cboCovSex.Text, cboCovSex.SelStart + cboCovSex.SelLength + 1))
   
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboCovSex.ListCount - 1
        
        If NewText = LCase(Left(cboCovSex.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboCovSex.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

             If ValidCount = 1 Then
               cboCovSex.Text = ValidValue
               cboCovSex.SelStart = 0
               cboCovSex.SelLength = Len(ValidValue)
             End If
        End If

End Sub


Private Sub cboINSCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
   
    'Clipboard.clear
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Right(cboINSCode.Text, cboINSCode.SelStart) + Chr(KeyAscii) + Mid(cboINSCode.Text, cboINSCode.SelStart + cboINSCode.SelLength + 1))
  
    ValidCount = 0
    ValidValue = ""
   
    For i = 0 To cboINSCode.ListCount - 1
        
        If NewText = LCase(Left(cboINSCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboINSCode.List(i)
        
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
             
             If ValidCount = 1 Then
               cboINSCode.Text = ValidValue
               cboINSCode.SelStart = 0
               cboINSCode.SelLength = Len(ValidValue)
             End If
        End If
End Sub

Private Sub cboPLNCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    'Clipboard.clear
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboPLNCode.Text, cboPLNCode.SelStart) + Chr(KeyAscii) + Mid(cboPLNCode.Text, cboPLNCode.SelStart + cboPLNCode.SelLength + 1))

    ValidCount = 0
    ValidValue = ""

    For i = 0 To cboPLNCode.ListCount - 1

        If NewText = LCase(Left(cboPLNCode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPLNCode.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0

           If ValidCount = 1 Then
             cboPLNCode.Text = ValidValue
             cboPLNCode.SelStart = 0
             cboPLNCode.SelLength = Len(ValidValue)
            End If
 
        End If
End Sub

Private Sub cboRACCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        
        Clipboard.Clear
        
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboRACCode.Text, cboRACCode.SelStart) + Chr(KeyAscii) + Mid(cboRACCode.Text, cboRACCode.SelStart + cboRACCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboRACCode.ListCount - 1
            
            If NewText = LCase(Left(cboRACCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboRACCode.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   cboRACCode.Text = ValidValue
                   cboRACCode.SelStart = 0
                   cboRACCode.SelLength = Len(ValidValue)
                 End If
            
            End If

End Sub

Private Sub cboSRVCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
        'Clipboard.clear
        
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboSRVCode.Text, cboSRVCode.SelStart) + Chr(KeyAscii) + Mid(cboSRVCode.Text, cboSRVCode.SelStart + cboSRVCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboSRVCode.ListCount - 1
            
            If NewText = LCase(Left(cboSRVCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboSRVCode.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                If ValidCount = 1 Then
                  cboSRVCode.Text = ValidValue
                  cboSRVCode.SelStart = 0
                  cboSRVCode.SelLength = Len(ValidValue)
                End If
            
            End If
End Sub


Private Sub cboMBMSalutation_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        Clipboard.Clear
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboMBMSalutation.Text, cboMBMSalutation.SelStart) + Chr(KeyAscii) + Mid(cboMBMSalutation.Text, cboMBMSalutation.SelStart + cboMBMSalutation.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboMBMSalutation.ListCount - 1
            
            If NewText = LCase(Left(cboMBMSalutation.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboMBMSalutation.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   cboMBMSalutation.Text = ValidValue
                   cboMBMSalutation.SelStart = 0
                   cboMBMSalutation.SelLength = Len(ValidValue)
                 End If
            
            End If
            
End Sub


Private Sub cboMBMSex_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
        
                
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboMBMSex.Text, cboMBMSex.SelStart) + Chr(KeyAscii) + Mid(cboMBMSex.Text, cboMBMSex.SelStart + cboMBMSex.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboMBMSex.ListCount - 1
            
            If NewText = LCase(Left(cboMBMSex.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboMBMSex.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   cboMBMSex.Text = ValidValue
                   cboMBMSex.SelStart = 0
                   cboMBMSex.SelLength = Len(ValidValue)
                 End If
            End If
End Sub

Private Sub cboNATCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
        
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboNATCode.Text, cboNATCode.SelStart) + Chr(KeyAscii) + Mid(cboNATCode.Text, cboNATCode.SelStart + cboNATCode.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboNATCode.ListCount - 1
            
            If NewText = LCase(Left(cboNATCode.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboNATCode.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   cboNATCode.Text = ValidValue
                   cboNATCode.SelStart = 0
                   cboNATCode.SelLength = Len(ValidValue)
                 End If
            End If
End Sub



Private Sub cboMBMMaritalStatus_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    
        Clipboard.Clear
    
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboMBMMaritalStatus.Text, cboMBMMaritalStatus.SelStart) + Chr(KeyAscii) + Mid(cboMBMMaritalStatus.Text, cboMBMMaritalStatus.SelStart + cboMBMMaritalStatus.SelLength + 1))
        
        ValidCount = 0
        ValidValue = ""
        
        For i = 0 To cboMBMMaritalStatus.ListCount - 1
            
            If NewText = LCase(Left(cboMBMMaritalStatus.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboMBMMaritalStatus.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0
            
                 If ValidCount = 1 Then
                   cboMBMMaritalStatus.Text = ValidValue
                   cboMBMMaritalStatus.SelStart = 0
                   cboMBMMaritalStatus.SelLength = Len(ValidValue)
                 End If
            End If
End Sub

Private Sub txtMBMPayorExpDate_LostFocus()
Dim MBMValidDate As New ADODB.Recordset
Dim strsql As String
Dim CheckDays As Integer
Dim Confirm1
Dim Confirm2

    CheckDays = 0
    
If txtMBMPayorExpDate <> "__/__/____" Then
    If IsDate(txtMBMPayorExpDate) Then
        If Mid(cboEndorseStatus, 1, 1) <> "I" Then
            strsql = "Select ValidID , MBMExpStart, MBMExpEnd From MBMValidDateRange "
            MBMValidDate.MaxRecords = 1
            MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
            If MBMValidDate.recordCount Then
                If DateDiff("d", MBMValidDate!MBMExpStart, txtMBMPayorExpDate) >= 0 And DateDiff("d", MBMValidDate!MBMExpEnd, txtMBMPayorExpDate) <= 0 Then
                    If txtMBMPayorEffDate <> "__/__/____" Then
                        
                        'CK added---------------------------------------------------
           
                        CheckDays = DateDiff("d", txtMBMPayorEffDate, txtMBMPayorExpDate)
           
                        If CheckDays > 365 Then
                            Confirm1 = MsgBox("Policy more than 365 days. Do you want to continue?", vbCritical + vbYesNo)
                            If Confirm1 = vbYes Then
                                Call InsertMCO
                                Call InsertPrem
                                Call InsertProviderFees
                            End If
                        ElseIf CheckDays <= 30 Then
                            Confirm2 = MsgBox("Policy less than 30 days. Do you want to continue?", vbCritical + vbYesNo)
                            If Confirm2 = vbYes Then
                                Call InsertMCO
                                Call InsertPrem
                                Call InsertProviderFees
                            End If
                        
                        Else
                            Call InsertMCO
                            Call InsertPrem
                            Call InsertProviderFees
                        End If
                        
                        'CK added---------------------------------------------------
                        
                        'Call InsertMCO
                        'Call InsertPrem
                        'Call InsertProviderFees
                    
                    End If
                
                Else
                    MsgBox "Policy Expiry Date is out of acceptable range! ", vbCritical + vbOKOnly
                    txtMBMPayorExpDate.SetFocus
                End If
            End If
            'End If
            MBMValidDate.Close
            Set MBMValidDate = Nothing
        End If
        
            If txtMBMPayorEffDate <> "__/__/____" And txtTempEffDate <> "__/__/____" And txtTempExpDate = "__/__/____" And txtMBMPayorExpDate <> "__/__/____" Then
                If CDate(txtMBMPayorEffDate) = CDate(txtTempEffDate) And CDate(txtTempExpDate) = CDate(txtMBMPayorExpDate) Then
                    txtMCOFees = txtTempMCOFees
                    txtMBMPremium = txtTempPrem
                    txtMBMProviderCharges = txtTempProvider
                End If
            End If
    Else
        If txtMBMPayorExpDate <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtMBMPayorExpDate.SetFocus
        End If
    End If
End If

End Sub

Private Sub txtMBMPolNo_LostFocus()
    'On Error Resume Next
    Dim CheckRecord
    Dim MBMPOLNO As New ADODB.Recordset
    Dim strsql As String
    Dim msgstr As String
    Dim RecordConfirm
    Dim CheckFlag As Boolean
    
    If Len(Trim(txtMBMPolNo)) Then
        Screen.MousePointer = vbHourglass
        If CheckFlag = False Then
            strsql = "Select MBMNumber From MBM Where MBMPolicyNo ='" & Trim(txtMBMPolNo) & "'"
            MBMPOLNO.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
           If MBMPOLNO.recordCount >= 1 Then
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
                   ' txtMBMAdd1.SetFocus
                    CheckFlag = True
                End If
            End If
            MBMPOLNO.Close
            Set MBMPOLNO = Nothing
        End If
        Screen.MousePointer = vbDefault
    End If

End Sub


Private Sub txtMBMRBAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub



'Private Sub txtMBMPayorExpDate_KeyPress(KeyAscii As Integer)
'Dim MBMMCOFees As String
'Dim MBMPremium As String
'Dim AccessFees As String
'
'    MBMMCOFees = InsertMCOFees(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), txtMCOFees, "#,##0.00")
'    MBMPremium = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), txtMBMPremium, "#,##0.00")
'    AccessFees = InsertProvider(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), txtMBMProviderCharges, "#,##0.00")'

    'txtMCOFees = MBMMCOFees
    'txtMBMPremium = MBMPremium
    'txtMBMProviderCharges = AccessFees
'End Sub

'Private Sub txtMBMWorkNature_KeyPress(KeyAscii As Integer)
'On Error Resume Next
    'Dim NewText As String
    'Dim ValidCount As Integer
    'Dim ValidValue As String
    'Dim i As Integer
        
        'Clipboard.clear
        
        'If KeyAscii >= 32 And KeyAscii <> 127 Then
        'NewText = LCase(Left(txtMBMWorkNature.Text, txtMBMWorkNature.SelStart) + Chr(KeyAscii) + Mid(txtMBMWorkNature.Text, txtMBMWorkNature.SelStart + txtMBMWorkNature.SelLength + 1))
        
        'ValidCount = 0
        'ValidValue = ""
        
        'For i = 0 To txtMBMWorkNature.ListCount - 1
            
            'If NewText = LCase(Left(txtMBMWorkNature.List(i), Len(NewText))) Then
                    'ValidCount = ValidCount + 1
                    'ValidValue = txtMBMWorkNature.List(i)
            'End If
            'Next
            'If ValidCount <= 1 Then KeyAscii = 0
            
                 'If ValidCount = 1 Then
                   'txtMBMWorkNature.Text = ValidValue
                   'txtMBMWorkNature.SelStart = 0
                   'txtMBMWorkNature.SelLength = Len(ValidValue)
                 'End If
            'End If
'End Sub

Private Sub SuspendPrincipal()
    Dim MemberAmendHis As New ADODB.Recordset
    Dim MBM As New ADODB.Recordset
    Dim MBMCov As New ADODB.Recordset
    Dim strsql As String
    Dim strsqlDEL As String
    Dim DEL As New ADODB.Recordset
    Dim SelectAge As String
    
    strsql = "Select * from MBM Where MBMNumber = '" & txtMBMNumber & "'"
    MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    MemberAmendHis.Open "MBMAmendHistory", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    MemberAmendHis.AddNew
    With MemberAmendHis
        !MBMNumber = MBM!MBMNumber
        !MBMPolicyNo = MBM!MBMPolicyNo
        !MBMIcBcPp = MBM!MBMIcBcPp
        !MBMName = MBM!MBMName
        !MBMExpiryDate = MBM!MBMPayorEXPDate
        !MBMPlaNCode = MBM!PLNCode
        !MBMAmendEdtType = Mid(cboEndorseStatus, 1, 1)
        If IsDate(txtDate) Then
            !MBMAEndtEffDate = txtDate
        End If
        'If IsDate(txtEndtBordxDate) Then
        '    !MBMAENDtBordxDate = txtEndtBordxDate
        'End If
        '!MBMValueDate = Date
        MemberAmendHis.Update
    End With

    
    'MemberHis.Open "MBA", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'Do Until MBM.EOF
                With MBM
                    !MBMLastEndorseStatus = Mid(cboEndorseStatus, 1, 1)
                    !MBMStatus = "S"
                    '!MBMENDtBordxDate = txtEndtBordxDate
                    !MBMENDtEffDate = txtDate
                    MBM.Update
                End With
        'MBM.MoveNext
    'Loop
    MBM.Close
    cboStatus.ListIndex = 3

    'Call EntriesEnable(False)
    'txtMBMEndorseStatus.Enabled = False
    'cmdSave.Enabled = False
    'txtMBACancelDate.Enabled = False
End Sub

Private Sub SuspendSupplementary()
    Dim Arraycnt, arrayCurrRow As Integer
    Dim MemberAmendHis As New ADODB.Recordset
    Dim MBMCov As New ADODB.Recordset
    Dim strsql As String
    Dim ListCount As Integer
    Dim listloop As Integer
    Dim LastCoveredID As Integer
    Dim lstString()
    Dim Item
    lstString = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

    strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & txtMBMNumber & "'"
    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    LastCoveredID = GetLastCoveredID(txtMBMNumber)

    ListCount = lstCovAdd.listitems.count
    For listloop = 1 To ListCount

                With MBMCov
                    MBMCov.Filter = ("MBMCCoverID = '" & lstCovAdd.listitems(listloop).SubItems(10) & "'")
                    !MBMCStatus = "S"
                    '!MBMCLastUpdateDate = txtDate
                    MBMCov.Update
                End With
            
            For arrayCurrRow = 1 To 10
                If lstString(arrayCurrRow) = 0 Then
                    lstString(arrayCurrRow) = listloop
                    Exit For
                End If
            Next
        Next listloop
    MBMCov.Close
    Set MBMCov = Nothing
    

    'Call EntriesEnable(False)
    'txtMBMEndorseStatus.Enabled = False
    'cmdSave.Enabled = False
    'txtMBACancelDate.Enabled = False
End Sub

Private Sub ReactivateSupplementary()
    Dim Arraycnt, arrayCurrRow As Integer
    Dim MemberHis As New ADODB.Recordset
    Dim MBMCov As New ADODB.Recordset
    Dim strsql As String
    Dim ListCount As Integer
    Dim listloop As Integer
    Dim LastCoveredID As Integer
    Dim lstString()
    Dim Item
    lstString = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

    strsql = "Select * from MBMCoveredPersons Where MBMCNumber = '" & txtMBMNumber & "'"
    MBMCov.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    'MemberHis.Open "MBA", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    LastCoveredID = GetLastCoveredID(txtMBMNumber)

    ListCount = lstCovAdd.listitems.count
    For listloop = 1 To ListCount

                With MBMCov
                    MBMCov.Filter = ("MBMCCoverID = '" & lstCovAdd.listitems(listloop).SubItems(10) & "'")
                    !MBMCStatus = "A"
                    '!MBMCLastUpdateDate = txtDate
                    MBMCov.Update
                End With
            
            For arrayCurrRow = 1 To 10
                If lstString(arrayCurrRow) = 0 Then
                    lstString(arrayCurrRow) = listloop
                    Exit For
                End If
            Next
        Next listloop
    MBMCov.Close
    Set MBMCov = Nothing
    

    'Call EntriesEnable(False)
    'txtMBMEndorseStatus.Enabled = False
    'cmdSave.Enabled = False
    'txtMBACancelDate.Enabled = False
End Sub

Private Sub ReactivatePrincipal()
    Dim MemberAmendHis As New ADODB.Recordset
    Dim MBM As New ADODB.Recordset
    Dim MBMCov As New ADODB.Recordset
    Dim strsql As String
    Dim strsqlDEL As String
    Dim DEL As New ADODB.Recordset
    Dim SelectAge As String
    
    strsql = "Select * from MBM Where MBMNumber = '" & txtMBMNumber & "'"
    MBM.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    MemberAmendHis.Open "MBMAmendHistory", medixmasterconn, adOpenKeyset, adLockOptimistic

    MemberAmendHis.AddNew
    With MemberAmendHis
        !MBMNumber = MBM!MBMNumber
        !MBMPolicyNo = MBM!MBMPolicyNo
        !MBMIcBcPp = MBM!MBMIcBcPp
        !MBMName = MBM!MBMName
        !MBMExpiryDate = MBM!MBMPayorEXPDate
        !MBMPlaNCode = MBM!PLNCode
        !MBMAmendEdtType = Mid(cboEndorseStatus, 1, 1)
        If IsDate(txtDate) Then
            !MBMAEndtEffDate = txtDate
        End If
        'If IsDate(txtEndtBordxDate) Then
        '    !MBMAENDtBordxDate = txtEndtBordxDate
        'End If

        '!MBMValueDate = Date
        MemberAmendHis.Update
    End With

    
    'MemberHis.Open "MBA", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'Do Until MBM.EOF
                With MBM
                    !MBMLastEndorseStatus = "Y"
                    !MBMStatus = "A"
                    '!MBMSuspendDate = txtDate
                    !MBMENDtEffDate = txtDate
                    '!MBMENDtBordxDate = txtEndtBordxDate
                    MBM.Update
                End With
        'MBM.MoveNext
    'Loop
    MBM.Close
    cboStatus.ListIndex = 0
    'Call EntriesEnable(False)
    'txtMBMEndorseStatus.Enabled = False
    'cmdSave.Enabled = False
    'txtMBACancelDate.Enabled = False
End Sub

Private Sub InsertNewRegistrationSub()
    Dim MBM As New ADODB.Recordset
    Dim MBMOthers As New ADODB.Recordset
    Dim MBMCoveredPersons As New ADODB.Recordset
    Dim planstring
    Dim listloop As Integer
    Dim listrecord As Integer
    Dim MembershipNo As String
    Dim SelectAge As String
    Dim RecordSaved
    Dim completeMBM, completeMBMOthers, completeMBMC As Boolean
    
    planstring = cboPLNCode
    ' INsert into Membership table
    MBM.Open "MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    ' Begin transaction
    medixmasterconn.beginTrans
    startTrans = True
    'MembershipNo = Trim(Mid(cboPAYCode, 1, 2)) & Trim(Mid(cboPLNCode, 1, 2)) & Trim(Mid(cboINSCode, 1, 1)) & Mid(BusinessYear, 3, 2) & GenerateMembershipNo(Trim(Mid(cboPAYCode, 1, 2)), Trim(Mid(cboPLNCode, 1, 2)), Trim(Mid(cboINSCode, 1, 1)))
    MembershipNo = Trim(Mid(cboPAYCode, 1, 2)) & Trim(Mid(cboPLNCode, 1, 2)) & Trim(Mid(cboINSCode, 1, 1)) & GenerateMembershipNo(Trim(Mid(cboPAYCode, 1, 2)), Trim(Mid(cboPLNCode, 1, 2)), Trim(Mid(cboINSCode, 1, 1)))
    MBM.AddNew
         With MBM
             !MBMNumber = MembershipNo
             !MBMPolicyNo = txtMBMPolNo
             !MBMIcBcPp = txtMBMIcBcPP
             !MBMName = txtMBMName
             !MBMSalutation = cboMBMSalutation
             !HLTCode = Trim(Mid(Trim(cboHLTCode), 1, 2))
             !PAYCode = Trim(Mid(cboPAYCode, 1, 2))
             !MBMAddress1 = txtMBMAdd1
             !MBMAddress2 = txtMBMAdd2
             !MBMAddress3 = txtMBMAdd3
             !CTYCode = cboCTYCode
             !MBMPostCode = txtMBMPostCode
             '!MBMTelnoH = txtMBMTelNoH
             '!MBMTelNoM = txtMBMTelNoM
             '!MBMPolSubNo1 = TxtPOLSubNo
             '!MBMTelNoO = txtMBMTelNoO
             '!MBMEmail = txtMBMEmail
             '!MBMReceiptDate = CDate(txtMBMBordxDate)
             If Len(Trim(cboBRNCode)) Then
                !BRNCode = Trim(Mid(cboBRNCode, 1, 4))
             End If
             If IsDate(txtMBMBirthDate) Then
                txtMBMBirthDate.mask = ""
                !MBMDOB = txtMBMBirthDate
                txtMBMBirthDate.mask = "##/##/####"
             End If
             !NATCode = Trim(Mid(cboNATCode, 1, 2))
             !RACCode = Trim(Mid(cboRACCode, 1, 2))
             !MBMSex = cboMBMSex
             !MBMPayorEffDate = CDate(txtMBMPayorEffDate)
             '!MBMDateJoin = CDate(txtMBMDateJoin)
             !MBMPayorEXPDate = CDate(txtMBMPayorExpDate)
             !AgeCode = Trim(Mid(txtAGECode, 1, 2))
             !INSCode = Trim(Mid(cboINSCode, 1, 2))
             !PLNCode = Trim(Mid(cboPLNCode, 1, 2))
             !MBMAdultChild = txtMBMAdultChild
             txtDate.mask = ""
             !MBMENDtEffDate = txtDate
             txtDate.mask = "##/##/####"
             txtEndtBordxDate.mask = ""
             !MBMENDtBordxDate = txtEndtBordxDate
             txtEndtBordxDate.mask = "##/##/####"
             'Call CallLimit
             '!MBMAvailableLimit = txtAvailableAnnualLimit
             !SRVCodeList = Trim(Mid(cboSRVCode, 1, InStr(1, cboSRVCode, "-") - 1))
             !STACode = Trim(cboSTACode)
             '!LGNCode = Trim(Mid(cboLGNCode, 1, 2))
             'txtMBMBordxDate.Mask = ""
             '!MBMBordxDate = CDate(txtMBMBordxDate)
             !MBMTakeOver = Mid(ChkStr(cboMBMTakeOver), 1, 1)
             !MBMRenewFlag = Mid(cboMBMRenewal, 1, 1)
             !BRNCode = cboBRNCode
             !MBMStatus = "A"
             !MBMValueDate = Date
             !MBMPolicyYear = 1
             !MBMTakeOver = Trim(Mid(cboMBMTakeOver, 1, 1))
             !MBMMemberType = "P"
             
            '!MBMBatchNo = txtMBMBatchNo
             Call MemberCallAge
             Call MemberCallFees
             'Call ProviderCallFees
             Call InsertMCO
             Call CallLimit
             Call InsertPrem
             Call InsertProviderFees
             .Update
     End With
    completeMBM = True
    ' INsert into Membership Other details table
     MBMOthers.Open "MBMOthers", medixmasterconn, adOpenKeyset, adLockOptimistic
     MBMOthers.AddNew
     With MBMOthers
             !MBMNumber = MembershipNo
             '!MBMWeight = txtMBMWeight
             '!MBMHeight = txtMBMHeight
             '!MBMOccupation = Trim(Mid(cboMBMOccupation, 1, 2))
             '!MBMWorkNature = txtMBMWorkNature
             !MBMExclusion = txtMBMExclusion
             '!MBMAllergic = txtMBMAllergic
             !MBMGroupCompany = txtMBMGroupCompany
             !MBMEmployeeNo = txtMBMEmployeeNo
             '!MBMBlood = cboMBMBlood
             '!MBMSmoking = cboMBMSmoking
             '!MBMAlcohol = cboMBMAlcohol
             !MBMAgentCode = txtMBMAgentCode
             
             If Len(Trim(cboMBMMaritalStatus)) Then
                !MBMMaritalStatus = Mid(cboMBMMaritalStatus, 1, 1)
             End If
             !MBMAccessFee = IIf(IsNumeric(txtMBMProviderCharges) = True, txtMBMProviderCharges, 0)
             !MBMPremium = txtMBMPremium
             !MBMMCOFee = txtMCOFees
             'If ChkStr(cboMBMRenewal) = "Y" Then
             '   !MBMFirstMEMNo = txtFirstMeMNo
             '   !MBMFirstPolNo = txtFirstPolNo
             '   !MBMPrevPolNo = txtMBMPrevPolNo
             '   !MBMPrevMemNo = txtMBMPrevMEMNo
             '   !MBMFirstJoinedDate = txtMBMDateJoin
             'Else
             '   !MBMFirstMEMNo = MembershipNo
             '   !MBMFirstPolNo = txtMBMPolicyNo
             '   !MBMFirstJoinedDate = txtMBMDateJoin
             'End If
             
     End With
     MBMOthers.Update
     completeMBMOthers = True
     
    If Trim(MBM!INSCode) = "I" Or Trim(MBM!INSCode) = "G" Then
        completeMBMC = False
     ElseIf Trim(MBM!INSCode) = "F" Or Trim(MBM!INSCode) = "H" Then
     listrecord = lstCovAdd.listitems.count
    ' Insert into Membership Covered Person's table
     If listrecord <> 0 Then
         MBMCoveredPersons.Open "MBMCOveredPersons", medixmasterconn, adOpenKeyset, adLockOptimistic
         For listloop = 1 To listrecord
             MBMCoveredPersons.AddNew
             With MBMCoveredPersons
                 !MBMCNumber = MembershipNo
                 !MBMCSalutation = lstCovAdd.listitems(listloop).SubItems(1)
                 !MBMCName = lstCovAdd.listitems(listloop).Text
                 !MBMCIcBcPpNo = lstCovAdd.listitems(listloop).SubItems(2)
                 !MBMCDOB = lstCovAdd.listitems(listloop).SubItems(3)
                 SelectAge = GetAge(txtMBMPayorEffDate, lstCovAdd.listitems(listloop).SubItems(3))
                 !MBMCAgeband = Mid(GetAgeBand(Trim(Mid(cboHLTCode, 1, InStr(1, cboHLTCode, "-") - 1)), SelectAge), 1, 2)
                 !MBMCSex = lstCovAdd.listitems(listloop).SubItems(4)
                 !MBMCAdultChild = lstCovAdd.listitems(listloop).SubItems(5)
                 !MBMCRELCode = Mid(lstCovAdd.listitems(listloop).SubItems(6), 1, 1)
                 !MBMCBloodType = lstCovAdd.listitems(listloop).SubItems(7)
                 !MBMCAllergic = lstCovAdd.listitems(listloop).SubItems(8)
                 !MBMCExclusion = lstCovAdd.listitems(listloop).SubItems(9)
                 If Len(Trim(lstCovAdd.listitems(listloop).SubItems(12))) Then
                    !MBMCAnnualLimit = lstCovAdd.listitems(listloop).SubItems(12)
                 End If
                 !MBMCCoverID = listloop
                 .Update
            End With
         Next listloop
         MBMCoveredPersons.Close
         Set MBMCoveredPersons = Nothing
    End If
   
    completeMBMC = True
     End If
          
    'If Trim(MBM!INSCode) = "F" Or Trim(MBM!INSCode) = "H" Then
       If completeMBM = True And completeMBMC = True And completeMBMOthers = True Then
            ' commit changes
            medixmasterconn.CommitTrans
        'End If
    'ElseIf Trim(MBM!INSCode) = "I" Or Trim(MBM!INSCode) = "G" Then
       ElseIf completeMBM = True And completeMBMOthers = True And completeMBMC = False Then
            medixmasterconn.CommitTrans
        'End If
        Else
            'Roll back
            medixmasterconn.RollbackTrans
        End If
    startTrans = False
    MBM.Close
    Set MBM = Nothing
   
    MBMOthers.Close
    Set MBMOthers = Nothing
     
     
     'cboMBMRenewal.Enabled = True
     MsgBox "Record Saved, Your Membership No is : " & MembershipNo, vbOKOnly + vbInformation, DialogBoxTitle
     ' Clering previous record
     'txtMBMPayorExpDate.Enabled = True
     'Call ClearEntries
     'txtMBMPayorExpDate.Enabled = False
    
     SSTab1.Tab = 0
     'cboMBMRenewal.SetFocus
End Sub

Private Sub InsertMCO()
Dim MBMInsertMCOFees As Double
'Dim MBMMCOFess As Double
Dim PAYExpDate As String

'MBMMCOFess = txtMCOFees
'PayExpDate = txtMBMPayorExpDate
    
    If txtTempMCOFees <> 0 Then
    MBMInsertMCOFees = InsertMCOFees(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(txtTempMCOFees, "#,##0.00"))
    txtMCOFees = ""
    'txtMCOFees = Round(Format(MBMInsertMCOFees, 0, "#,##0.00"))
    txtMCOFees = Format(Round(MBMInsertMCOFees, 0), "#,##0.00")
    End If
End Sub

Private Sub InsertPrem()
Dim MBMInsertPrem As Double
Dim MBMPrem As Double
Dim PAYExpDate As String

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
            cboSRVCode.Text = "EA - EUROPE ASSISTANCE"
            cboHLTCode.Text = ""
            'cboPatientList.Clear
        ElseIf TypeOf ctl Is ListView Then
            ctl.listitems.Clear
        ElseIf TypeOf ctl Is MaskEdBox Then
            ctl.mask = ""
            ctl.Text = ""
            ctl.mask = "##/##/####"
        ElseIf TypeOf ctl Is Label Then
            txtMBMPremium = ""
            txtMBMPremiumRed = ""
            txtMBMPremiumEff = ""
            txtMCOFees = ""
            txtMBMMCORed = ""
            txtMCOFeesEff = ""
            txtMBMProviderCharges = ""
            txtMBMProviderRedCharges = ""
            txtMBMProviderEffCharges = ""
            txtAnnualLimit = ""
            txtAvailableAnnualLimit = ""
        End If
    Next ctl
    FrmMBMOthers.Check1.value = 0
    '  Call ComboListing
      'cmdSaveKIV.Enabled = True
      'cmdGenerate.Enabled = True
    FrmMBMOthers.cboCOPayRB = ""
    FrmMBMOthers.cboBTBG = ""
End Sub

Private Sub DeletePrincipals()
Dim strsqlMBM As String
Dim MBM As New ADODB.Recordset
Dim MBMArchieve As New ADODB.Recordset
Dim strdelPrincipal As String
Dim DelPrincipal As New ADODB.Recordset
Dim DelOthers As String
Dim MBMOthers As New ADODB.Recordset
Dim MemberAmendHis As New ADODB.Recordset
            
    If txtMBMNumber <> "" Then
            
    strsqlMBM = "SELECT * From MBM where MBMNumber = '" & txtMBMNumber & "'"
    MBM.Open strsqlMBM, medixmasterconn, adOpenKeyset, adLockOptimistic

    MemberAmendHis.Open "MBMAmendHistory", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    MemberAmendHis.AddNew
    With MemberAmendHis
        
        !MBMNumber = MBM!MBMNumber
        !MBMPolicyNo = MBM!MBMPolicyNo
        !MBMIcBcPp = MBM!MBMIcBcPp
        !MBMName = MBM!MBMName
        !MBMExpiryDate = MBM!MBMPayorEXPDate
        !MBMPlaNCode = MBM!PLNCode
        !MBMAmendEdtType = Mid(cboEndorseStatus, 1, 1)
        If IsDate(txtDate) Then
            !MBMAEndtEffDate = txtDate
        End If
        !MBMAmendUser = Logon
        If Len(txtBatchNo) Then
            !MBMAEndtBatchNo = txtBatchNo
        End If
        If IsDate(txtEndtBordxDate) Then
            !MBMAENDtBordxDate = txtEndtBordxDate
        End If
        MemberAmendHis.Update
    End With

             MBM!MBMStatus = "D"
             MBM!MBMValueDate = Date
             txtDate.mask = ""
             If IsDate(txtDate) Then
                MBM!MBMENDtEffDate = txtDate
             End If
             If IsDate(txtEndtBordxDate) Then
             'txtEndtBordxDate.Mask = ""
                MBM!MBMENDtBordxDate = txtEndtBordxDate
             End If
        MBM.Update
            cboStatus.ListIndex = 2
        'txtMBMNumber = MBM!MBMNumber
        'txtMBMType = MBM!MBMMemberType
        
        If Trim(MBM!INSCode) = "F" Or Trim(MBM!INSCode) = "H" Then
            Call DeleteSupplementaryPol
        End If

    MBM.Close
    Set MBM = Nothing
    End If
End Sub

Private Sub DeleteSupplementaryPol()
Dim strsqlMBMCov As String
Dim MBMCov As New ADODB.Recordset
        

      strsqlMBMCov = "SELECT * From MBMCoveredPersons  Where MBMCNumber = '" & Trim(txtMBMNumber) & "'"
      MBMCov.Open strsqlMBMCov, medixmasterconn, adOpenKeyset, adLockOptimistic
      
        Do Until MBMCov.EOF
                With MBMCov
                    !MBMCStatus = "D"
                End With
                MBMCov.Update
        MBMCov.MoveNext
        Loop

    
    MBMCov.Close
    Set MBMCov = Nothing
End Sub

Private Sub DelSuppLoop()
Dim listrecord As Integer
Dim listloop As Integer
Dim ListCount As Integer
Dim ForLoop As Integer
Dim LastCoveredID As Integer

        ListCount = lstCovAdd.listitems.count
        For listloop = 1 To ListCount
        If listloop > ListCount Then
            Exit Sub
       ElseIf listloop <= ListCount Then
       If lstCovAdd.listitems(listloop).SubItems(11) = "R" Then
            lstCovAdd.listitems.Remove (listloop)
            'listloop = listloop - 1
            ListCount = ListCount - 1
          End If
          End If
        'If listloop <> ListCount Then
            'Exit Sub
        'Else
    Next listloop
      
    'Else
    'Exit Sub
        'End If
        
End Sub

Public Sub InsertProviderFees()
Dim MBMInsertProviderFees As Double
'Dim MBMProviderFees As Double
'Dim PayExpDate As String

'MBMPrem = txtMCOFees
'PayExpDate = txtMBMPayorExpDate
    If txtTempProvider <> 0 Then
        MBMInsertProviderFees = InsertProvider(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(txtTempProvider, "#,##0.00"))
        txtMBMProviderCharges = ""
        'txtMBMProviderCharges = Round(Format(MBMInsertProviderFees, 1, "#,##0.00"))
        txtMBMProviderCharges = Format(Round(MBMInsertProviderFees, 1), "#,##0.00")
    End If
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


' ################ END Key press effect ###################
'Private Sub txtMBMWorkNature_LostFocus()
    'Clipboard.Clear
'End Sub

Private Sub cboMBMRelCode_CHANGE()
    If Mid(cboMBMRelCode, 1, 1) = "S" Then
        Call SearchSuppLmt
    End If

'    If Trim(Mid(cboMBMRelCode, 1, 1)) = "C" Then
'        txtCovAnnualLimit.Enabled = False
'        txtAvailableLimit.Enabled = False
'        txtSuppPrem.Enabled = False
'    Else
'        txtCovAnnualLimit.Enabled = True
'        txtAvailableLimit.Enabled = False
'        txtSuppPrem.Enabled = False
 '   End If
End Sub


Private Sub cboMBMRelCode_Click()
    If Mid(cboMBMRelCode, 1, 1) = "S" Then
        Call SearchSuppLmt
    End If
'    If Trim(Mid(cboMBMRelCode, 1, 1)) = "C" Then
'        txtCovAnnualLimit.Enabled = False
'        txtAvailableLimit.Enabled = False
'        txtSuppPrem.Enabled = False
'    Else
'        txtCovAnnualLimit.Enabled = True
'        txtAvailableLimit.Enabled = False
'        txtSuppPrem.Enabled = False
'    End If
End Sub

Private Sub SearchSuppLmt()
Dim strsqlannlmt As String
Dim AnnLmt As New ADODB.Recordset
Dim strsqlPrem As String
Dim Prem As New ADODB.Recordset
    If txtMBMPayorEffDate <> "__/__/____" Then
    
    strsqlannlmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(cboPLNCode) & "' AND INSCode = '" & cboINSCode & "'"
    AnnLmt.Open strsqlannlmt, medixmasterconn, adOpenKeyset, adLockOptimistic

    strsqlPrem = "Select SuppPrmStatus, PLNCode, SuppPrmAmt from PRM where PLNCode = '" & Trim(cboPLNCode) & "' AND INSCode = '" & cboINSCode & "'"
    Prem.Open strsqlPrem, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'If Mid(cboSuppAnnualLimit, 1, 1) = "S" Then
        If AnnLmt!SuppLimitStatus = "B" Then
            tempSPremium = Prem!SuppPrmAmt
            tempSAnnLmt = Format(GetSuppLimit(cboINSCode, cboPLNCode, Mid(cboHLTCode, 1, 1), CDate(txtMBMPayorEffDate)), "#,##0.00")
    'Cater for SuppLimitStatus is "C"
        ElseIf AnnLmt!SuppLimitStatus = "C" Then
            'tempSPremium = Prem!SuppPrmAmt
            tempSAnnLmt = Format(GetSuppLimit(cboINSCode, cboPLNCode, Mid(cboHLTCode, 1, 1), CDate(txtMBMPayorEffDate)), "#,##0.00")
            txtCovAnnualLimit = tempSAnnLmt
        Else
            tempSAnnLmt = txtAnnualLimit
            txtCovAnnualLimit = txtAnnualLimit
        End If
    AnnLmt.Close
    Prem.Close
    Set AnnLmt = Nothing
    Set Prem = Nothing

    End If
    
    'AnnLmt.Close
    'Prem.Close
    'Set AnnLmt = Nothing
    'Set Prem = Nothing
End Sub



Private Sub txtRemarks_Click()
    KeyPreview = False
End Sub

Private Sub TxtRemarks_LostFocus()
    KeyPreview = True
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
    
    'sqlstr = sqlstr & " From View_MBM_MBMOthers Where MBMNumber ='" & Trim(txtMBMNumber) & "'"
    'MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    

    If FrmMBMOthers.Visible = False Then
        Load FrmMBMOthers
    End If
    'FrmMBMOthers.show

    'FrmMBMOthers.Visible
    Call ClearForm(FrmMBMOthers)
    'FrmMBMOthers.Visible
    Label4 = ""
    sqlstr = " Select MBMNumber, MBMTelnoH, MBMTelNoM, MBMPayorPrem, " & _
            " MBMTelNoO, MBMPrevPLNCode, MBMExclusion, MBMAllergic, " & _
            " LabelStatus, MBMPrevMemNo, MBMPrevPolNo, MBMPrevPLNCode, MBMRBAmt, " & _
            " MBMPAYRemarks, MBMPrevInsCom, MBMRemarks, MBMCoPayRB, MBMGenWP, MBMGenEndWP, " & _
            " MBMPreExWP, MBMPreExEndWP, MBMSpeIllWP, MBMSpeIllEndWP, MBMStaffDec, MBMValueDate, " & _
            " MBMCLMNonPayFr, MBMCLMNonPayTo, MBMNewPlanCode, MBMNewPolEffDate, MBMPOLSubNo1, MBMBTBG, " & _
            " MBMOccupation, MBMWorkNature, MBMWeight, MBMHeight, MBMEmail, MBMBTBG, LGNCode, MBMBlood, " & _
            " MBMAlcohol, MBMPrevTakeOverPLNCode,MBMPayorPremGross,MBMPayorPremNet,MBMPayorMCOGross,MBMPayorMCONet, " & _
            " MBMSmoking, MBMPolicyDisc, MBMRenewalDisc "
    
    'If Source = 1 Then
    
        sqlstr = sqlstr & " From View_MBM_MBMOthers Where MBMNumber ='" & Trim(txtMBMNumber) & "'"
        MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        
        'strCount = "select count(*) as totalRecord From View_MBM_MBMOthers  Where MBMNumber = '" & Trim(frmMBMAdjustment.txtMBMNumber) & "'"
        'strCount = strCount + sqlstr
        
        MBMNumber = txtMBMNumber
        With MBM
            
            FrmMBMOthers.Label4 = Trim(MBMNumber)
            
            If Trim(!MBMTelnoH) <> "" Then
                FrmMBMOthers.txtMBMTelNoH = Trim(!MBMTelnoH)
            End If
            
            If Trim(!MBMTelNoM) <> "" Then
                FrmMBMOthers.txtMBMTelNoM = Trim(!MBMTelNoM)
            End If
            
            If Trim(!MBMTelNoO) <> "" Then
                FrmMBMOthers.txtMBMTelNoO = Trim(!MBMTelNoO)
            End If
            
            If Trim(!MBMCLMNonPayFr) <> "" Then
                FrmMBMOthers.TxtClmNonPayFr = Format(!MBMCLMNonPayFr, "dd/mm/yyyy")
            End If
            If Trim(!MBMCLMNonPayTo) <> "" Then
                FrmMBMOthers.TxtClmNonPayTo = Format(!MBMCLMNonPayTo, "dd/mm/yyyy")
            End If
            
            If Trim(!LabelStatus) = "Y" Then
                FrmMBMOthers.Check1.value = 1
            End If
            
            If Trim(!MBMPolSubNo1) <> "" Then
                FrmMBMOthers.TxtPOLSubNo = Trim(!MBMPolSubNo1)
            End If
            
            If Trim(!MBMAllergic) <> "" Then
                FrmMBMOthers.txtMBMAllergic = Trim(!MBMAllergic)
            End If
            
            If Len(!MBMPrevPLNCode) Then
                FrmMBMOthers.txtMBMPrevPlan = !MBMPrevPLNCode
            End If
            
            If Len(!MBMRBAmt) Then
                FrmMBMOthers.txtMBMRBAmt = !MBMRBAmt
            End If
            
            If Len(!MBMPAYRemarks) Then
                FrmMBMOthers.txtMBMPAYRemarks = !MBMPAYRemarks
            End If
            
            If Len(!MBMPrevInsCom) Then
                FrmMBMOthers.txtMBMPrevINSCom = !MBMPrevInsCom
            End If
        
            If Len(!MBMCoPayRB) Then
                FrmMBMOthers.cboCOPayRB = !MBMCoPayRB
            End If
            
            If Len(!MBMGenWP) Then
                FrmMBMOthers.TxtGenWP = !MBMGenWP
            End If
            
            If Len(!MBMPreExWP) Then
                FrmMBMOthers.txtPREWP = !MBMPreExWP
            End If
            
            If Len(!MBMSpeIllWP) Then
                FrmMBMOthers.TxtSpeWP = !MBMSpeIllWP
            End If
            
            If Len(!MBMStaffDec) Then
                FrmMBMOthers.cboStaffDisc = !MBMStaffDec
            End If
            
            If Len(!MBMGenEndWP) Then
                FrmMBMOthers.txtGENWPEnd = !MBMGenEndWP
            End If
            
            If Len(!MBMPreExEndWP) Then
                FrmMBMOthers.txtPREWPEnd = !MBMPreExEndWP
            End If
            
            If Len(!MBMSpeIllEndWP) Then
                FrmMBMOthers.txtSPEWPEnd = !MBMSpeIllEndWP
            End If
                
            If Len(!MBMNewPlanCode) Then
                FrmMBMOthers.txtPLNCode = !MBMNewPlanCode
            End If
            
            If Len(!MBMNewPolEffDate) Then
                FrmMBMOthers.txtNewPolEffDate = Format(!MBMNewPolEffDate, "dd/mm/yyyy")
            End If
            
            If Len(!MBMBTBG) Then
                FrmMBMOthers.cboBTBG = !MBMBTBG
            End If
            
            If Len(!MBMOccupation) Then
                FrmMBMOthers.cboMBMOccupation = !MBMOccupation
            End If
            
            If Len(!MBMWorkNature) Then
                FrmMBMOthers.txtMBMWorkNature = !MBMWorkNature
            End If
            
            If Len(!MBMWeight) Then
                FrmMBMOthers.txtMBMWeight = !MBMWeight
            End If
            
            If Len(!MBMHeight) Then
                FrmMBMOthers.txtMBMHeight = !MBMHeight
            End If
            
            If Len(!MBMEmail) Then
                FrmMBMOthers.txtMBMEmail = !MBMEmail
            End If
            
            If Len(!LGNCode) Then
                FrmMBMOthers.cboLGNCode = !LGNCode
            End If
            
            If Len(!MBMBlood) Then
                FrmMBMOthers.cboMBMBlood = !MBMBlood
            End If
                         
            If Len(!MBMSmoking) Then
                FrmMBMOthers.cboMBMSmoking = !MBMSmoking
            End If
            
            If Len(!MBMAlcohol) Then
                FrmMBMOthers.cboMBMAlcohol = !MBMAlcohol
            End If
            
            If Len(!MBMPayorPremGross) Then
                FrmMBMOthers.txtPayPrem = !MBMPayorPremGross
            End If
            
            If Len(!MBMPayorPremNet) Then
                FrmMBMOthers.txtPayPremNett = !MBMPayorPremNet
            End If
            
            If Len(!MBMPayorMCOGross) Then
                FrmMBMOthers.txtPayMCOGross = !MBMPayorMCOGross
            End If
            
            If Len(!MBMPayorMCONet) Then
                FrmMBMOthers.txtPayMCONett = !MBMPayorMCONet
            End If
            
            If Len(!MBMPolicyDisc) Then
                FrmMBMOthers.txtPolicyDisc = !MBMPolicyDisc
            End If
            
            If Len(!MBMRenewalDisc) Then
                FrmMBMOthers.txtRenewalDisc = !MBMRenewalDisc
            End If
        End With
        MBM.Close
        Set MBM = Nothing
    'End If
End Function

