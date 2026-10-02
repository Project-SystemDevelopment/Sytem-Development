VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMemberRegistration 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member Registration"
   ClientHeight    =   7950
   ClientLeft      =   1050
   ClientTop       =   480
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   91
      Top             =   0
      Width           =   11985
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Member Registration"
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
         TabIndex        =   92
         Top             =   90
         Width           =   2535
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5565
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   9816
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Principal &Details"
      TabPicture(0)   =   "frmMemberRegistration.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Supplementary"
      TabPicture(1)   =   "frmMemberRegistration.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstCovAdd"
      Tab(1).Control(1)=   "Frame6"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame6 
         Height          =   3075
         Left            =   -74880
         TabIndex        =   144
         Top             =   480
         Width           =   11460
         Begin VB.TextBox txtCovAllergic 
            Height          =   795
            Left            =   4680
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   73
            Top             =   2160
            Width           =   1575
         End
         Begin VB.TextBox txtCLNCovExclusion 
            Enabled         =   0   'False
            Height          =   795
            Left            =   9600
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   75
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox txtCovExclusion 
            Height          =   795
            Left            =   6360
            MaxLength       =   2500
            MultiLine       =   -1  'True
            TabIndex        =   74
            Top             =   2160
            Width           =   3135
         End
         Begin VB.ComboBox cboCovSalutation 
            Height          =   315
            Left            =   1320
            TabIndex        =   45
            Top             =   195
            Width           =   975
         End
         Begin VB.TextBox txtCovName 
            Height          =   330
            Left            =   3480
            MaxLength       =   60
            TabIndex        =   46
            Top             =   180
            Width           =   3660
         End
         Begin VB.Frame Frame9 
            Height          =   1695
            Left            =   10320
            TabIndex        =   145
            Top             =   120
            Width           =   1065
            Begin VB.CommandButton cmdRemove 
               Caption         =   "&Remove"
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   77
               Top             =   600
               Width           =   840
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "&Add"
               Height          =   315
               Left            =   120
               TabIndex        =   76
               Top             =   240
               Width           =   840
            End
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "Update"
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   78
               Top             =   960
               Width           =   840
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Clear"
               Height          =   315
               Left            =   120
               TabIndex        =   79
               Top             =   1320
               Width           =   840
            End
         End
         Begin VB.TextBox txtCovIcBcPp 
            Height          =   285
            Left            =   7800
            MaxLength       =   30
            TabIndex        =   47
            Top             =   180
            Width           =   2055
         End
         Begin VB.ComboBox cboCovRelationship 
            Height          =   315
            Left            =   1320
            TabIndex        =   48
            Text            =   "W -Wife"
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox cboCovAdultChild 
            Height          =   315
            Left            =   1320
            TabIndex        =   52
            Text            =   "A"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtPrevPLNCode 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5760
            TabIndex        =   50
            Top             =   480
            Width           =   1380
         End
         Begin VB.TextBox txtAvailableLimit 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   55
            Top             =   1010
            Width           =   975
         End
         Begin VB.TextBox txtPrevINSCom 
            Height          =   330
            Left            =   7800
            MaxLength       =   60
            TabIndex        =   51
            Top             =   440
            Width           =   2055
         End
         Begin VB.TextBox txtRBAmt 
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   61
            Top             =   1290
            Width           =   975
         End
         Begin VB.TextBox txtSuppPayPremGross 
            Height          =   285
            Left            =   1320
            TabIndex        =   65
            Top             =   1545
            Width           =   975
         End
         Begin VB.TextBox txtCOVPayRemarks 
            Height          =   285
            Left            =   5760
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   54
            Top             =   720
            Width           =   4095
         End
         Begin VB.TextBox txtSuppGenWP 
            Height          =   285
            Left            =   5760
            MaxLength       =   9
            TabIndex        =   57
            Top             =   960
            Width           =   945
         End
         Begin VB.TextBox txtSuppPREWP 
            Height          =   285
            Left            =   5760
            MaxLength       =   9
            TabIndex        =   63
            Top             =   1200
            Width           =   945
         End
         Begin VB.TextBox TxtSuppSpeWP 
            Height          =   300
            Left            =   5760
            MaxLength       =   9
            TabIndex        =   67
            Top             =   1455
            Width           =   945
         End
         Begin VB.TextBox txtSuppPayMCOGross 
            Height          =   285
            Left            =   1320
            TabIndex        =   69
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox txtRenewalDisc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   71
            Top             =   2040
            Width           =   975
         End
         Begin MSMask.MaskEdBox txtCovBirthdate 
            Height          =   285
            Left            =   3480
            TabIndex        =   49
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
            TabIndex        =   59
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
            TabIndex        =   60
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
            Height          =   315
            Left            =   7320
            TabIndex        =   58
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
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
            Height          =   315
            Left            =   7320
            TabIndex        =   64
            Top             =   1245
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
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
            Height          =   315
            Left            =   7320
            TabIndex        =   68
            Top             =   1530
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
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
            TabIndex        =   53
            Text            =   "M"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtCovAnnualLimit 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   56
            Top             =   1005
            Width           =   1095
         End
         Begin VB.ComboBox cboSuppStaffDisc 
            Height          =   315
            Left            =   3480
            Sorted          =   -1  'True
            TabIndex        =   62
            Top             =   1260
            Width           =   1095
         End
         Begin VB.TextBox txtSuppPayPremNet 
            Height          =   285
            Left            =   3480
            TabIndex        =   66
            Top             =   1550
            Width           =   1095
         End
         Begin VB.TextBox txtSuppPayMCONet 
            Height          =   285
            Left            =   3480
            TabIndex        =   70
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox txtPolicyDisc 
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   72
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label48 
            Caption         =   "Allergic"
            Height          =   255
            Left            =   4680
            TabIndex        =   177
            Top             =   1920
            Width           =   600
         End
         Begin VB.Label Label108 
            Caption         =   "Supp. Clinic Warranty"
            Height          =   255
            Left            =   9600
            TabIndex        =   176
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label32 
            Caption         =   "Policy Disc."
            Height          =   255
            Left            =   2475
            TabIndex        =   175
            Top             =   2040
            Width           =   1365
         End
         Begin VB.Label Label28 
            Caption         =   "Renewal Disc."
            Height          =   255
            Left            =   120
            TabIndex        =   174
            Top             =   2055
            Width           =   1485
         End
         Begin VB.Label Label23 
            Caption         =   "Pay MCO Net"
            Height          =   255
            Left            =   2475
            TabIndex        =   173
            Top             =   1815
            Width           =   1365
         End
         Begin VB.Label Label13 
            Caption         =   "Pay MCO Gross"
            Height          =   255
            Left            =   120
            TabIndex        =   172
            Top             =   1815
            Width           =   1365
         End
         Begin VB.Label Label31 
            Caption         =   "Pay Prem Net"
            Height          =   255
            Left            =   2480
            TabIndex        =   171
            Top             =   1560
            Width           =   1365
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            Caption         =   "To"
            Height          =   255
            Left            =   8400
            TabIndex        =   170
            Top             =   1440
            Width           =   285
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            Caption         =   "From"
            Height          =   255
            Left            =   8400
            TabIndex        =   169
            Top             =   1245
            Width           =   405
         End
         Begin VB.Label Label111 
            Caption         =   "Pay Prem Gross"
            Height          =   255
            Left            =   120
            TabIndex        =   168
            Top             =   1560
            Width           =   1365
         End
         Begin VB.Label Label37 
            Caption         =   "Clm Non-Pay Date"
            Height          =   375
            Left            =   8450
            TabIndex        =   167
            Top             =   1000
            Width           =   1365
         End
         Begin VB.Label Label53 
            Caption         =   "EOP"
            Height          =   255
            Left            =   6840
            TabIndex        =   166
            Top             =   1275
            Width           =   1005
         End
         Begin VB.Label Label104 
            Caption         =   "Spe. Ill. W. P."
            Height          =   255
            Left            =   4680
            TabIndex        =   165
            Top             =   1530
            Width           =   1065
         End
         Begin VB.Label Label103 
            Caption         =   "Pre W. P."
            Height          =   255
            Left            =   4680
            TabIndex        =   164
            Top             =   1245
            Width           =   780
         End
         Begin VB.Label Label78 
            Caption         =   "Gen W. P."
            Height          =   255
            Left            =   4680
            TabIndex        =   163
            Top             =   990
            Width           =   900
         End
         Begin VB.Label Label33 
            Caption         =   "EOP"
            Height          =   255
            Left            =   6840
            TabIndex        =   162
            Top             =   1030
            Width           =   1005
         End
         Begin VB.Label Label45 
            Caption         =   "EOP"
            Height          =   255
            Left            =   6840
            TabIndex        =   161
            Top             =   1560
            Width           =   1005
         End
         Begin VB.Label Label36 
            Caption         =   "Disc Req"
            Height          =   255
            Left            =   2480
            TabIndex        =   160
            Top             =   1320
            Width           =   1065
         End
         Begin VB.Label Label92 
            Caption         =   "Pay Remarks"
            Height          =   255
            Left            =   4680
            TabIndex        =   159
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label91 
            Caption         =   "T/O Ins "
            Height          =   255
            Left            =   7200
            TabIndex        =   158
            Top             =   480
            Width           =   1140
         End
         Begin VB.Label Label90 
            Caption         =   "R&&B Amt"
            Height          =   255
            Left            =   120
            TabIndex        =   157
            Top             =   1320
            Width           =   1140
         End
         Begin VB.Label Label82 
            Caption         =   "P. Pln Code"
            Height          =   255
            Left            =   4680
            TabIndex        =   156
            Top             =   525
            Width           =   1125
         End
         Begin VB.Label Label38 
            Caption         =   "Name"
            Height          =   255
            Left            =   2480
            TabIndex        =   155
            Top             =   195
            Width           =   1140
         End
         Begin VB.Label Label51 
            Caption         =   "Adult/Child"
            Height          =   255
            Left            =   120
            TabIndex        =   154
            Top             =   795
            Width           =   825
         End
         Begin VB.Label Label50 
            Caption         =   "Sex "
            Height          =   255
            Left            =   2480
            TabIndex        =   153
            Top             =   735
            Width           =   825
         End
         Begin VB.Label Label39 
            Caption         =   "Salutation"
            Height          =   255
            Left            =   120
            TabIndex        =   152
            Top             =   225
            Width           =   1140
         End
         Begin VB.Label Label41 
            Caption         =   "IC Num"
            Height          =   255
            Left            =   7200
            TabIndex        =   151
            Top             =   165
            Width           =   1125
         End
         Begin VB.Label Label49 
            Caption         =   "DOB"
            Height          =   255
            Left            =   2480
            TabIndex        =   150
            Top             =   495
            Width           =   1020
         End
         Begin VB.Label Label46 
            Caption         =   "Supp Lmt"
            Height          =   255
            Left            =   2480
            TabIndex        =   149
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label52 
            Caption         =   "R'Ship"
            Height          =   255
            Left            =   120
            TabIndex        =   148
            Top             =   525
            Width           =   1005
         End
         Begin VB.Label Label47 
            Caption         =   "Supplementary Warranty "
            Height          =   255
            Left            =   6360
            TabIndex        =   147
            Top             =   1920
            Width           =   2100
         End
         Begin VB.Label Label56 
            Caption         =   "Bal. Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   146
            Top             =   1020
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5085
         Left            =   120
         TabIndex        =   83
         Top             =   360
         Width           =   11295
         Begin VB.Frame Frame4 
            Height          =   585
            Left            =   120
            TabIndex        =   135
            Top             =   4440
            Width           =   11055
            Begin VB.Label txtAnnualLimit 
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
               Left            =   7320
               TabIndex        =   138
               Top             =   240
               Width           =   1290
            End
            Begin VB.Label txtMBMPremium 
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
               Left            =   960
               TabIndex        =   140
               Top             =   240
               Width           =   810
            End
            Begin VB.Label Label25 
               Caption         =   "Premium"
               Height          =   255
               Left            =   240
               TabIndex        =   143
               Top             =   240
               Width           =   600
            End
            Begin VB.Label Label26 
               Caption         =   "MCO Fees"
               Height          =   240
               Left            =   2040
               TabIndex        =   142
               Top             =   240
               Width           =   870
            End
            Begin VB.Label Label27 
               Caption         =   "Annual Limit"
               Height          =   255
               Left            =   6360
               TabIndex        =   141
               Top             =   240
               Width           =   960
            End
            Begin VB.Label txtMCOFees 
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
               Left            =   2880
               TabIndex        =   139
               Top             =   240
               Width           =   1050
            End
            Begin VB.Label txtMBMProviderCharges 
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
               Left            =   5160
               TabIndex        =   137
               Top             =   240
               Width           =   960
            End
            Begin VB.Label Label30 
               Caption         =   "Access Fee"
               Height          =   255
               Left            =   4200
               TabIndex        =   136
               Top             =   240
               Width           =   1140
            End
         End
         Begin VB.TextBox txtMBMExclusion 
            Height          =   1170
            Left            =   240
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Top             =   3240
            Width           =   4890
         End
         Begin VB.TextBox txtRemarks 
            Height          =   1215
            Left            =   5880
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            Top             =   3240
            Width           =   5130
         End
         Begin VB.TextBox txtMBMAdd1 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   14
            Top             =   150
            Width           =   4065
         End
         Begin VB.ComboBox cboPLNCode 
            Height          =   315
            Left            =   9720
            Sorted          =   -1  'True
            TabIndex        =   29
            Top             =   120
            Width           =   1140
         End
         Begin VB.ComboBox cboINSCode 
            Height          =   315
            ItemData        =   "frmMemberRegistration.frx":0038
            Left            =   7080
            List            =   "frmMemberRegistration.frx":003A
            Sorted          =   -1  'True
            TabIndex        =   28
            Top             =   120
            Width           =   1215
         End
         Begin MSMask.MaskEdBox txtMBMPayorExpDate 
            Height          =   315
            Left            =   9720
            TabIndex        =   31
            Top             =   410
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtMBMAdd2 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   15
            Top             =   400
            Width           =   4065
         End
         Begin VB.TextBox txtMBMAdd3 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   16
            Top             =   660
            Width           =   4065
         End
         Begin VB.TextBox txtMBMPostCode 
            Height          =   315
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   17
            Top             =   910
            Width           =   1215
         End
         Begin VB.ComboBox cboCTYCode 
            Height          =   315
            ItemData        =   "frmMemberRegistration.frx":003C
            Left            =   3480
            List            =   "frmMemberRegistration.frx":003E
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   910
            Width           =   1680
         End
         Begin VB.ComboBox cboSTACode 
            Height          =   315
            Left            =   1080
            TabIndex        =   19
            Top             =   1200
            Width           =   4095
         End
         Begin VB.ComboBox cboMBMTakeover 
            Height          =   315
            Left            =   3525
            TabIndex        =   21
            Text            =   "N"
            Top             =   1480
            Width           =   1650
         End
         Begin VB.ComboBox cboMBMMaritalStatus 
            Height          =   315
            Left            =   3525
            TabIndex        =   23
            Text            =   "M"
            Top             =   1760
            Width           =   1650
         End
         Begin VB.ComboBox cboMBMSex 
            CausesValidation=   0   'False
            Enabled         =   0   'False
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   20
            Tag             =   "M"
            Text            =   "M"
            Top             =   1480
            Width           =   1215
         End
         Begin MSMask.MaskEdBox txtMBMBirthdate 
            Height          =   315
            Left            =   1080
            TabIndex        =   22
            Top             =   1760
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboRACCode 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   24
            Text            =   "C - Chinese"
            Top             =   2040
            Width           =   1215
         End
         Begin VB.ComboBox cboNATCode 
            Height          =   315
            Left            =   3525
            Sorted          =   -1  'True
            TabIndex        =   25
            Text            =   "MY-Malaysia"
            Top             =   2040
            Width           =   1650
         End
         Begin VB.ComboBox cboBRNCode 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   2320
            Width           =   4095
         End
         Begin VB.TextBox txtMBMAgentCode 
            Height          =   285
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   27
            Top             =   2600
            Width           =   4065
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
            Left            =   7080
            TabIndex        =   30
            Top             =   410
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtMBMBordxDate 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7080
            MaxLength       =   40
            TabIndex        =   32
            Top             =   690
            Width           =   1215
         End
         Begin MSMask.MaskEdBox txtMBMDateJoin 
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
            Left            =   7080
            TabIndex        =   34
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtMBMEmployeeNo 
            Height          =   315
            Left            =   7080
            MaxLength       =   15
            TabIndex        =   36
            Top             =   1230
            Width           =   3825
         End
         Begin VB.TextBox txtMBMDepartment 
            Height          =   315
            Left            =   7080
            MaxLength       =   50
            TabIndex        =   37
            Top             =   1520
            Width           =   3825
         End
         Begin VB.TextBox txtMBMGroupCompany 
            Height          =   315
            Left            =   7080
            MaxLength       =   150
            TabIndex        =   38
            Top             =   1800
            Width           =   3825
         End
         Begin VB.ComboBox cboSRVCode 
            Height          =   315
            Left            =   7080
            Sorted          =   -1  'True
            TabIndex        =   40
            Text            =   "EA-Europe Assistance"
            Top             =   2370
            Width           =   3825
         End
         Begin VB.ComboBox cboMBMRelCode 
            Height          =   315
            Left            =   7080
            TabIndex        =   41
            Text            =   "C - Combined Limit"
            Top             =   2650
            Width           =   1425
         End
         Begin VB.ComboBox CboInsMode 
            Height          =   315
            Left            =   9480
            TabIndex        =   42
            Text            =   "A - Annual"
            Top             =   2650
            Width           =   1425
         End
         Begin VB.Label Label40 
            Caption         =   "Principal Warranty "
            Height          =   255
            Left            =   240
            TabIndex        =   133
            Top             =   3000
            Width           =   1590
         End
         Begin VB.Label Label76 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   5880
            TabIndex        =   132
            Top             =   3000
            Width           =   1590
         End
         Begin VB.Label Label70 
            Caption         =   "Instl mode"
            Height          =   255
            Left            =   8640
            TabIndex        =   129
            Top             =   2685
            Width           =   855
         End
         Begin VB.Label txtMBMPolicyYear 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "1"
            Height          =   315
            Left            =   9735
            TabIndex        =   35
            Top             =   990
            Width           =   1140
         End
         Begin VB.Label txtAGECode 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7080
            TabIndex        =   39
            Top             =   2115
            Width           =   3825
         End
         Begin VB.Label Label59 
            Caption         =   "Branch"
            Height          =   255
            Left            =   180
            TabIndex        =   127
            Top             =   2320
            Width           =   600
         End
         Begin VB.Label Label10 
            Caption         =   "Race"
            Height          =   255
            Left            =   180
            TabIndex        =   126
            Top             =   2040
            Width           =   600
         End
         Begin VB.Label Label57 
            Caption         =   "Dept. Name"
            Height          =   255
            Left            =   5880
            TabIndex        =   125
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "E'yee Num"
            Height          =   255
            Left            =   5880
            TabIndex        =   123
            Top             =   1260
            Width           =   1545
         End
         Begin VB.Label Label22 
            Caption         =   "Serv. Provider"
            Height          =   255
            Left            =   5880
            TabIndex        =   121
            Top             =   2400
            Width           =   1005
         End
         Begin VB.Label txtMBMBatchNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   9735
            TabIndex        =   33
            Top             =   720
            Width           =   1140
         End
         Begin VB.Label Label17 
            Caption         =   "Bordx. Date"
            Height          =   255
            Left            =   5880
            TabIndex        =   120
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label44 
            Caption         =   "Batch Num"
            Height          =   255
            Left            =   8640
            TabIndex        =   119
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label21 
            Caption         =   "Plan Code"
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
            Left            =   8640
            TabIndex        =   118
            Top             =   210
            Width           =   945
         End
         Begin VB.Label Label16 
            Caption         =   "Exp Date"
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
            Left            =   8640
            TabIndex        =   117
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label Label63 
            Caption         =   "Take Over"
            Height          =   255
            Left            =   2520
            TabIndex        =   115
            Top             =   1480
            Width           =   825
         End
         Begin VB.Label Label24 
            Caption         =   "Marital Status"
            Height          =   255
            Left            =   2520
            TabIndex        =   112
            Top             =   1760
            Width           =   1095
         End
         Begin VB.Label Label42 
            Caption         =   "Date Join"
            Height          =   255
            Left            =   5880
            TabIndex        =   111
            Top             =   990
            Width           =   780
         End
         Begin VB.Label Label7 
            Caption         =   "State"
            Height          =   255
            Left            =   180
            TabIndex        =   110
            Top             =   1200
            Width           =   660
         End
         Begin VB.Label Label20 
            Caption         =   "Ins Type"
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
            Left            =   5880
            TabIndex        =   109
            Top             =   150
            Width           =   960
         End
         Begin VB.Label Label12 
            Caption         =   "PostCode"
            Height          =   255
            Left            =   180
            TabIndex        =   90
            Top             =   910
            Width           =   900
         End
         Begin VB.Label Label9 
            Caption         =   "Birthdate"
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
            Left            =   180
            TabIndex        =   89
            Top             =   1760
            Width           =   765
         End
         Begin VB.Label Label5 
            Caption         =   "Sex (M/F)"
            Height          =   255
            Left            =   180
            TabIndex        =   88
            Top             =   1480
            Width           =   825
         End
         Begin VB.Label Label2 
            Caption         =   "Address"
            Height          =   255
            Left            =   180
            TabIndex        =   87
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label15 
            Caption         =   "Eff Date"
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
            Left            =   5880
            TabIndex        =   86
            Top             =   440
            Width           =   1275
         End
         Begin VB.Label Label6 
            Caption         =   "City/Town"
            Height          =   255
            Left            =   2520
            TabIndex        =   114
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label11 
            Caption         =   "Nationality"
            Height          =   255
            Left            =   2520
            TabIndex        =   113
            Top             =   2040
            Width           =   960
         End
         Begin VB.Label Label14 
            Caption         =   "Agent "
            Height          =   255
            Left            =   180
            TabIndex        =   128
            Top             =   2600
            Width           =   660
         End
         Begin VB.Label Label43 
            Caption         =   "Pol Year"
            Height          =   255
            Left            =   8640
            TabIndex        =   116
            Top             =   1020
            Width           =   1035
         End
         Begin VB.Label Label4 
            Caption         =   "Supp Type"
            Height          =   255
            Left            =   5880
            TabIndex        =   107
            Top             =   2680
            Width           =   1095
         End
         Begin VB.Label Label18 
            Caption         =   "Grp Company"
            Height          =   255
            Left            =   5880
            TabIndex        =   84
            Top             =   1840
            Width           =   1545
         End
         Begin VB.Label Label29 
            Caption         =   "Age Band"
            Height          =   255
            Left            =   5880
            TabIndex        =   85
            Top             =   2145
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView lstCovAdd 
         Height          =   1755
         Left            =   -74880
         TabIndex        =   134
         Top             =   3720
         Width           =   11520
         _ExtentX        =   20320
         _ExtentY        =   3096
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
         NumItems        =   38
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Mem Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Salutation"
            Object.Width           =   882
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
            Object.Width           =   1411
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
            Text            =   "Covered ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Annual Limit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Occupation"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Work Nature"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Smoking"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Alcohol"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "Prem"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "PrevPlnStatus"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "PrevPlnCode"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "RBAmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "PAY Remarks"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "Prev Insurance Com"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Text            =   "Gen W.P"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "EOP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Text            =   "Pre W.P."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Text            =   "EOP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Text            =   "Spe Ill W. P."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Text            =   "EOP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Text            =   "Staff DR"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   29
            Text            =   "Clm Non-Payment Fr"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   30
            Text            =   "Clm Non-Payment To"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Text            =   "Clinic Warranty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Text            =   "PayorPrem"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   33
            Text            =   "PayorPremNet"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   34
            Text            =   "PayorMCOGross"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Text            =   "PayorMCONet"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Text            =   "RenewalDisc"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Text            =   "PolDisc"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   80
      TabIndex        =   96
      Top             =   360
      Width           =   11775
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Enabled         =   0   'False
         Height          =   300
         Left            =   6480
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Go"
         Enabled         =   0   'False
         Height          =   300
         Left            =   5400
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtMBMPrevMEMNo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cboMBMRenewal 
         Height          =   315
         ItemData        =   "frmMemberRegistration.frx":0040
         Left            =   1320
         List            =   "frmMemberRegistration.frx":0042
         Sorted          =   -1  'True
         TabIndex        =   0
         Text            =   "N"
         Top             =   180
         Width           =   690
      End
      Begin VB.Label txtMBMPrevPolNo 
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   8640
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label68 
         Caption         =   "Prev. Pol. Num"
         Height          =   255
         Left            =   7440
         TabIndex        =   99
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label67 
         Caption         =   "Prev Mem No"
         Height          =   240
         Left            =   2040
         TabIndex        =   98
         Top             =   225
         Width           =   1290
      End
      Begin VB.Label Label66 
         Caption         =   "Renewal (Y/N)"
         Height          =   240
         Left            =   120
         TabIndex        =   97
         Top             =   225
         Width           =   1065
      End
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   120
      TabIndex        =   93
      Top             =   7320
      Width           =   11715
      Begin VB.CommandButton CmdMBMOthers 
         Caption         =   "Others Details"
         Height          =   315
         Left            =   7200
         TabIndex        =   131
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox txtcount 
         Height          =   285
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   122
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtINSFlag 
         Height          =   285
         Left            =   3120
         TabIndex        =   106
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtMBMAdultChild 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   105
         Text            =   "A"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtTempBordx 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   104
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtTempBatch 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   103
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   9435
         TabIndex        =   81
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   10440
         TabIndex        =   82
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   8445
         TabIndex        =   80
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   80
      TabIndex        =   94
      Top             =   880
      Width           =   11775
      Begin VB.CommandButton cmdClinic 
         Caption         =   ">"
         Enabled         =   0   'False
         Height          =   255
         Left            =   11400
         TabIndex        =   12
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtMBMIcBcPP 
         Height          =   285
         Left            =   8400
         MaxLength       =   30
         TabIndex        =   7
         Top             =   160
         Width           =   3015
      End
      Begin VB.TextBox txtMBMName 
         Height          =   330
         Left            =   3120
         MaxLength       =   60
         TabIndex        =   6
         Top             =   160
         Width           =   4245
      End
      Begin VB.ComboBox cboMBMSalutation 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Text            =   "MR"
         Top             =   160
         Width           =   690
      End
      Begin VB.TextBox txtPAYCode 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   440
         Width           =   690
      End
      Begin VB.TextBox txtMBMPolicyNo 
         Height          =   330
         Left            =   3120
         MaxLength       =   25
         TabIndex        =   9
         Top             =   440
         Width           =   4245
      End
      Begin VB.TextBox txtHLTCode 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8400
         TabIndex        =   10
         Top             =   420
         Width           =   855
      End
      Begin VB.ComboBox cboHNC 
         Height          =   315
         ItemData        =   "frmMemberRegistration.frx":0044
         Left            =   9840
         List            =   "frmMemberRegistration.frx":0046
         Sorted          =   -1  'True
         TabIndex        =   11
         Text            =   "H - H&S"
         Top             =   420
         Width           =   1545
      End
      Begin VB.Label Label107 
         Caption         =   "H&&C"
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
         TabIndex        =   130
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label65 
         Caption         =   "Mem Name"
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
         Left            =   2040
         TabIndex        =   124
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label55 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   108
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label54 
         Caption         =   "Policy Num"
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
         Left            =   2040
         TabIndex        =   102
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label Label64 
         Caption         =   "Health Type"
         Height          =   255
         Left            =   7440
         TabIndex        =   101
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "IC Num"
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
         Left            =   7440
         TabIndex        =   100
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Salutation"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmMemberRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecordSaved As Boolean
Dim startTrans As Boolean
Public bDataStatus As Boolean
Dim bExit As Boolean
Dim tabIndexDone As Boolean
Dim MBMPayorEXPDate As Integer
Dim tempSPremium As String
Dim tempSAnnLmt As String
Dim txtTempProvider As Double
Dim txtTempPrem As Double
Dim txtTempMCOFees As Double
Dim txtTempExpDate  As String
Dim txtTempEffDate As String
Dim txtFirstMeMNo As String
Dim txtFirstPolNo As String
Dim strsqlPLN As String
Dim PLN As New ADODB.Recordset
Dim Disc As Boolean
Public PolDics As String
Public RenDisc As String

Private Sub cboHNC_Change()
    If Mid(cboHNC, 1, 1) = "C" Then
        cmdClinic.Enabled = True
        cboINSCode.Text = ""
        cboPLNCode.Text = ""
        txtMBMPayorEffDate = "__/__/____"
        txtMBMPayorExpDate = "__/__/____"
        cboINSCode.Enabled = False
        cboPLNCode.Enabled = False
        txtMBMPayorEffDate.Enabled = False
        txtMBMPayorExpDate.Enabled = False
        txtCLNCovExclusion.Enabled = True
    ElseIf Mid(cboHNC, 1, 1) = "H" Then
        cmdClinic.Enabled = False
        cboINSCode.Enabled = True
        cboPLNCode.Enabled = True
        txtMBMPayorEffDate.Enabled = True
        txtMBMPayorExpDate.Enabled = True
        txtCLNCovExclusion.Enabled = False
    Else
        cmdClinic.Enabled = True
        cboINSCode.Enabled = True
        cboPLNCode.Enabled = True
        txtMBMPayorEffDate.Enabled = True
        txtMBMPayorExpDate.Enabled = True
        txtCLNCovExclusion.Enabled = True
    End If
End Sub

Private Sub cboHNC_Click()
    If Mid(cboHNC, 1, 1) = "C" Then
        txtCLNCovExclusion.Enabled = True
        cmdClinic.Enabled = True
        cboINSCode.Text = ""
        cboPLNCode.Text = ""
        txtMBMPayorEffDate = "__/__/____"
        txtMBMPayorExpDate = "__/__/____"
        cboINSCode.Enabled = False
        cboPLNCode.Enabled = False
        txtMBMPayorEffDate.Enabled = False
        txtMBMPayorExpDate.Enabled = False
    ElseIf Mid(cboHNC, 1, 1) = "H" Then
        txtCLNCovExclusion.Enabled = False
        cmdClinic.Enabled = False
        cboINSCode.Enabled = True
        cboPLNCode.Enabled = True
        txtMBMPayorEffDate.Enabled = True
        txtMBMPayorExpDate.Enabled = True
    Else
        cmdClinic.Enabled = True
        txtCLNCovExclusion.Enabled = True
        cboINSCode.Enabled = True
        cboPLNCode.Enabled = True
        txtMBMPayorEffDate.Enabled = True
        txtMBMPayorExpDate.Enabled = True
    End If
End Sub

Private Sub cmdClinic_Click()
    frmMBMClinical.show
End Sub

Private Sub CmdMBMOthers_Click()
    FrmMBMOthers.Source = 3
    
    If FrmMBMOthers.Visible = False Then
        Load FrmMBMOthers
    End If
    FrmMBMOthers.show
    
    'Unload FrmMBMOthers
    'FrmMBMOthers.show
    FrmMBMOthers.Frame5.Enabled = True
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

Private Sub CboInsMode_GotFocus()
    KeyPreview = True
End Sub

Private Sub CboInsMode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SSTab1.Tab = 1
        tabIndexDone = True
    End If
End Sub

Private Sub cboMBMRelCode_CHANGE()
    If Mid(cboMBMRelCode, 1, 1) = "S" Then
        Call SearchSuppLmt
    End If
End Sub

Private Sub cboMBMRelCode_Click()
    If Mid(cboMBMRelCode, 1, 1) = "S" Then
        Call SearchSuppLmt
    End If
End Sub

Private Sub cboPLNCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 0
        Shift = 0
        Clipboard.Clear
    End If
End Sub

Private Sub Form_Load()
    Call ComboListing
    Call LoadInsuredType
    Call LoadPlanType("I")
    bDataStatus = True
    cboMBMRenewal.SelStart = 0
    bExit = False
    tabIndexDone = False
    txtTempProvider = 0
    txtTempPrem = 0
    txtTempMCOFees = 0
    PolDics = 0
    RenDisc = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub ComboListing()
    Dim SRV As New ADODB.Recordset
    Dim RAC As New ADODB.Recordset
    Dim NAT As New ADODB.Recordset
    Dim REL As New ADODB.Recordset
    Dim LGN As New ADODB.Recordset
    Dim CTY As New ADODB.Recordset
    Dim STA As New ADODB.Recordset
    Dim BRN As New ADODB.Recordset
    Dim OCP As New ADODB.Recordset
    Dim Work As New ADODB.Recordset
    
    
    SRV.Open "select SRVCode, SRVName from  SRV", medixmasterconn, adOpenKeyset, adLockOptimistic
    RAC.Open "select RACCode, RACName from   RAC", medixmasterconn, adOpenKeyset, adLockOptimistic
    NAT.Open "select NATCode, NATName from  NAT", medixmasterconn, adOpenKeyset, adLockOptimistic
    REL.Open "select RELCode, RELDescription from  REL", medixmasterconn, adOpenKeyset, adLockOptimistic
    OCP.Open "select OCPCode, OCPDescription from OCP", medixmasterconn, adOpenKeyset, adLockOptimistic
    LGN.Open "select LGNCode, LGNDescription from LGN", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    CTY.Open "select ctyCode , CTYDescription from CTY", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    STA.Open "select STACOde, STADescription from  STA", medixmasterconn, adOpenKeyset, adLockOptimistic
    BRN.Open "Select BRNCOde, BRNNAme from BRN", medixmasterconn, adOpenKeyset, adLockOptimistic
    Work.Open "Select  WorkID, WorkDescription from WorkNature ", medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    

    Do Until SRV.EOF
        cboSRVCode.AddItem ChkStr(SRV!SRVCode) & " - " & ChkStr(SRV!SRVname)
        SRV.MoveNext
    Loop

    Do Until RAC.EOF
        cboRACCode.AddItem ChkStr(RAC!RACCode) & " - " & ChkStr(RAC!RACName)
        RAC.MoveNext
    Loop

    Do Until NAT.EOF
        cboNATCode.AddItem ChkStr(NAT!NATCode) & " - " & ChkStr(NAT!NATName)
        NAT.MoveNext
    Loop


    Do Until REL.EOF
        cboCovRelationship.AddItem ChkStr(REL!RELCODE) & " - " & REL!RELDescription

        REL.MoveNext
    Loop

    Do Until CTY.EOF
        cboBRNCode.AddItem ChkStr(CTY!CTYDescription)
        cboCTYCode.AddItem ChkStr(CTY!CTYDescription)
        CTY.MoveNext
    Loop

    Do Until STA.EOF
        cboSTACode.AddItem ChkStr(STA!STADescription)
        STA.MoveNext
    Loop

    cboMBMRelCode.AddItem "C - Combined Limit"
    cboMBMRelCode.AddItem "S - Separated Limit"
    
    CboInsMode.AddItem "A - Annually"
    CboInsMode.AddItem "Q - Quarterly"
    CboInsMode.AddItem "S - Semi Annual"
    CboInsMode.AddItem "M - Monthly"
        
    cboMBMSex.AddItem "M"
    cboMBMSex.AddItem "F"
    cboMBMSex.AddItem ""
    
    cboCovSex.AddItem "M"
    cboCovSex.AddItem "F"
    cboCovSex.AddItem ""
    
    cboMBMMaritalStatus.AddItem "S"
    cboMBMMaritalStatus.AddItem "M"
    
    cboCovAdultChild.AddItem "A"
    cboCovAdultChild.AddItem "C"
    
    cboMBMSalutation.AddItem "  "
    cboMBMSalutation.AddItem "MR"
    cboMBMSalutation.AddItem "MS"
    cboMBMSalutation.AddItem "MDM"
    cboMBMSalutation.AddItem "MRS"
    
    cboCovSalutation.AddItem "  "
    cboCovSalutation.AddItem "MR"
    cboCovSalutation.AddItem "MS"
    cboCovSalutation.AddItem "MDM"
    cboCovSalutation.AddItem "MRS"

    cboMBMTakeover.AddItem "N - None"
    cboMBMTakeover.AddItem "Y - Take Over"
    cboMBMTakeover.AddItem "C - Conversion"
    
    cboSuppStaffDisc.AddItem "N - No"
    cboSuppStaffDisc.AddItem "Y - Yes"
        
    cboMBMRenewal.AddItem "N"
    cboMBMRenewal.AddItem "R"
    cboMBMRenewal.AddItem "T"
    cboMBMRenewal.AddItem "P"
    
    cboHNC.AddItem "H - H&S"
    cboHNC.AddItem "C - Clinical"
    cboHNC.AddItem "M - Mix"
    
    SRV.Close
    Set SRV = Nothing
    OCP.Close
    Set OCP = Nothing
    RAC.Close
    Set RAC = Nothing
    NAT.Close
    Set NAT = Nothing
    REL.Close
    Set REL = Nothing
    LGN.Close
    Set LGN = Nothing
    CTY.Close
    Set CTY = Nothing
    STA.Close
    Set STA = Nothing
    BRN.Close
    Set BRN = Nothing
End Sub

Private Sub LoadInsuredType()
    Dim strsql As String
    Dim INS As New ADODB.Recordset
   
    cboINSCode.Clear
    
    strsql = "Select INSCode, INSDescription from INS "
    strsql = strsql & " order by INSCode "
    INS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until INS.EOF
        cboINSCode.AddItem ChkStr(INS!INSCode)
        INS.MoveNext
    Loop
    INS.Close
    Set INS = Nothing
End Sub

Public Sub LoadPlanType(strINSType As String)
    Dim strsql As String
    Dim PLN As New ADODB.Recordset
    Dim planstring As String
    planstring = ""   ' Initialize string.
    cboPLNCode.Enabled = True
    cboPLNCode.Clear
    
    strsql = "Select PLNCode ,PLNDescription from PLN where  0=0 "
     
    If Mid(ChkStr(txtHLTCode), 1, 1) = "N" Then
        strsql = strsql & " AND PAYCode= '" & Trim(txtPAYCode) & "'"  'need to reverify
        strsql = strsql & " AND HLTCode = 'N'"
    Else
        strsql = strsql & " AND HLTCode = 'S'"
    End If
    PLN.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If PLN.recordCount <> 0 Then
        Do Until PLN.EOF
            planstring = ChkStr(PLN!PLNCode)
            cboPLNCode.AddItem planstring
            PLN.MoveNext
        Loop
    End If
    PLN.Close
    Set PLN = Nothing
End Sub

Private Sub showMBMCov()
    Dim MBMCov As New ADODB.Recordset
    Dim sqlstr As String
                
    sqlstr = "select MBMCNumber, MBMCCoverID, MBMCName, MBMCSalutation," & _
        " MBMCICBCPPNo , MBMCDOB, MBMCSex, MBMCAdultChild, MBMCAgeband, " & _
        " MBMCRELCode, MBMCBloodType, MBMCAllergic,  " & _
        " MBMCExclusion, MBMCStatus , MBMCAnnualLimit, MBMCOccupation, " & _
        " MBMCWorkNature, MBMCSmoking, MBMCAlcohol, " & _
        " MBMCPayorPremGross, MBMCPayorPremNet, MBMCPayorMCOGross, MBMCPayorMCONet, MBMCPolicyDisc, MBMCRenewalDisc "

    If bDataStatus = True Then
        sqlstr = sqlstr & " from  MBMCoveredPersons where MBMCNumber = '" & Trim(txtMBMPrevMEMNo) & "'"
        MBMCov.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    Else
        sqlstr = sqlstr & " from  MBMCoveredPersonsArchive where MBMCNumber = '" & Trim(txtMBMPrevMEMNo) & "'"
        MBMCov.Open sqlstr, medixmasterconnA, adOpenKeyset, adLockOptimistic
    End If
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
                    MemberConveredPerson.SubItems(11) = !MBMCAnnualLimit
                End If
                
                If Len(Trim(!MBMCOccupation)) Then
                    MemberConveredPerson.SubItems(12) = !MBMCOccupation
                End If
 
                If Len(Trim(!MBMCWorkNature)) Then
                    MemberConveredPerson.SubItems(13) = !MBMCWorkNature
                End If
                
                If Len(Trim(!MBMCSmoking)) Then
                    MemberConveredPerson.SubItems(14) = !MBMCSmoking
                End If
                
                If Len(Trim(!MBMCAlcohol)) Then
                    MemberConveredPerson.SubItems(15) = !MBMCAlcohol
                End If
            
                If Len(!MBMCPayorPremGross) Then
                    MemberConveredPerson.SubItems(32) = !MBMCPayorPremGross
                End If
                If Len(!MBMCPayorPremNet) Then
                    MemberConveredPerson.SubItems(33) = !MBMCPayorPremNet
                End If
                
                If Len(!MBMCPayorMCOGross) Then
                    MemberConveredPerson.SubItems(34) = !MBMCPayorMCOGross
                End If
                
                If Len(!MBMCPayorMCONet) Then
                    MemberConveredPerson.SubItems(35) = !MBMCPayorMCONet
                End If
                
                If Len(!MBMCRenewalDisc) Then
                    MemberConveredPerson.SubItems(36) = !MBMCRenewalDisc
                End If
                
                If Len(!MBMCPolicyDisc) Then
                    MemberConveredPerson.SubItems(37) = !MBMCPolicyDisc
                End If

            End With
            MBMCov.MoveNext
        Loop
    End If
    MBMCov.Close
    Set MBMCov = Nothing
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
        End If
    Next ctl
    If ChkStr(cboMBMRenewal) = "R" Then
        cmdShow.Enabled = True
        cmdSearch.Enabled = True
    Else
        cmdShow.Enabled = False
        cmdSearch.Enabled = False
    End If
    
    cboMBMSex.Enabled = False
    txtPAYCode.Enabled = False
    txtMBMBordxDate.Enabled = False
    txtHLTCode.Enabled = False
    txtMBMAdultChild.Enabled = False
    txtTempBordx.Enabled = False
    txtTempBatch.Enabled = False
    cmdUpdate.Enabled = False
    cmdRemove.Enabled = False
    txtCovAnnualLimit.Enabled = False
    'txtGENWPEnd.Enabled = False
    'txtPREWPEnd.Enabled = False
    'txtSPEWPEnd.Enabled = False
    txtSuppGENWPEnd.Enabled = False
    txtSuppPREWPEnd.Enabled = False
    txtSuppSPEWPEnd.Enabled = False
End Sub

Private Sub ClearEntries()
    On Error Resume Next
   Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.Enabled = True Then
               ctl.Text = ""
            End If
        ElseIf TypeOf ctl Is ComboBox Then
            If ctl.Enabled = True Then
                If ctl.Name <> "cboCovSex" Then
                    ctl.Text = ""
                End If
            End If
        ElseIf TypeOf ctl Is CommandButton Then
            ctl.Enabled = True
        ElseIf TypeOf ctl Is MaskEdBox Then
            If ctl.Enabled = True Then
                ctl.mask = ""
                ctl.Text = ""
                ctl.mask = "##/##/####"
            End If
        End If
    Next ctl
    cboCovRelationship = "W - Wife"
    cboCovSalutation = "MR"
    cboCovSex = "M"
    cboCovAdultChild = "A"
    txtCovAnnualLimit = 0
    cboMBMSalutation = "MR"
    cboMBMRenewal = "N"
    cboMBMSex = "M"
    cboMBMTakeover = "N - None"
    cboMBMMaritalStatus = "M"
    cboRACCode = "C - Chinese"
    cboNATCode = "MY - Malaysia"
    cboINSCode = "I"
    cboSRVCode = "EA - Europe Assistance"
    cboMBMRelCode = "C - Combined Limit"
    lstCovAdd.listitems.Clear
    txtAGECode = ""
    txtMBMPremium = ""
    txtMCOFees = ""
    txtMBMProviderCharges = ""
    txtAnnualLimit = ""
    FrmMBMOthers.txtGENWPEnd.mask = ""
    FrmMBMOthers.txtGENWPEnd.Text = ""
    FrmMBMOthers.txtGENWPEnd.mask = "##/##/####"
    FrmMBMOthers.txtPREWPEnd.mask = ""
    FrmMBMOthers.txtPREWPEnd.Text = ""
    FrmMBMOthers.txtPREWPEnd.mask = "##/##/####"
    FrmMBMOthers.txtSPEWPEnd.mask = ""
    FrmMBMOthers.txtSPEWPEnd.Text = ""
    FrmMBMOthers.txtSPEWPEnd.mask = "##/##/####"
    FrmMBMOthers.TxtClmNonPayFr.Text = ""
    FrmMBMOthers.TxtClmNonPayFr.mask = "##/##/####"
    FrmMBMOthers.TxtClmNonPayTo.Text = ""
    FrmMBMOthers.TxtClmNonPayTo.mask = "##/##/####"
    FrmMBMOthers.cboBTBG = "N - No"
    cboHNC = "H - H&S"
End Sub



Private Sub lstCovAdd_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If lstCovAdd.listitems.count Then
        cboCovSalutation = lstCovAdd.SelectedItem.SubItems(1)
        txtCovName = lstCovAdd.SelectedItem.Text
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(10)
        txtCovAnnualLimit = Format(lstCovAdd.SelectedItem.SubItems(11), "###,##0.00")
        txtPrevPLNCode = lstCovAdd.SelectedItem.SubItems(18)
        txtRBAmt = lstCovAdd.SelectedItem.SubItems(19)
        txtCOVPayRemarks = lstCovAdd.SelectedItem.SubItems(20)
        txtPrevINSCom = lstCovAdd.SelectedItem.SubItems(21)
        txtSuppGenWP = lstCovAdd.SelectedItem.SubItems(22)
        txtSuppGENWPEnd = lstCovAdd.SelectedItem.SubItems(23)
        txtSuppPREWP = lstCovAdd.SelectedItem.SubItems(24)
        txtSuppPREWPEnd = lstCovAdd.SelectedItem.SubItems(25)
        TxtSuppSpeWP = lstCovAdd.SelectedItem.SubItems(26)
        txtSuppSPEWPEnd = lstCovAdd.SelectedItem.SubItems(27)
        cboSuppStaffDisc = lstCovAdd.SelectedItem.SubItems(28)
        TxtSuppClmNonPayFr.mask = ""
        TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(29)
        TxtSuppClmNonPayTo.mask = ""
        TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(30)
        txtCLNCovExclusion = lstCovAdd.SelectedItem.SubItems(31)
        txtSuppPayPremGross = lstCovAdd.SelectedItem.SubItems(32)
        txtSuppPayPremNet = lstCovAdd.SelectedItem.SubItems(33)
        txtSuppPayMCOGross = lstCovAdd.SelectedItem.SubItems(34)
        txtSuppPayMCONet = lstCovAdd.SelectedItem.SubItems(35)
        txtRenewalDisc = lstCovAdd.SelectedItem.SubItems(36)
        txtPolicyDisc = lstCovAdd.SelectedItem.SubItems(37)

        Call EnableCovered(False)
        cmdUpdate.Enabled = True
        cmdRemove.Enabled = True
        cmdAdd.Enabled = False
    End If
    Call EnableCovered(True)
End Sub

Private Sub lstCovAdd_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If lstCovAdd.listitems.count Then
        cboCovSalutation = lstCovAdd.SelectedItem.SubItems(1)
        txtCovName = lstCovAdd.SelectedItem.Text
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(10)
        txtCovAnnualLimit = Format(lstCovAdd.SelectedItem.SubItems(11), "###,##0.00")
        txtPrevPLNCode = lstCovAdd.SelectedItem.SubItems(18)
        txtRBAmt = lstCovAdd.SelectedItem.SubItems(19)
        txtCOVPayRemarks = lstCovAdd.SelectedItem.SubItems(20)
        txtPrevINSCom = lstCovAdd.SelectedItem.SubItems(21)
        txtSuppGenWP = lstCovAdd.SelectedItem.SubItems(22)
        txtSuppGENWPEnd = lstCovAdd.SelectedItem.SubItems(23)
        txtSuppPREWP = lstCovAdd.SelectedItem.SubItems(24)
        txtSuppPREWPEnd = lstCovAdd.SelectedItem.SubItems(25)
        TxtSuppSpeWP = lstCovAdd.SelectedItem.SubItems(26)
        txtSuppSPEWPEnd = lstCovAdd.SelectedItem.SubItems(27)
        cboSuppStaffDisc = lstCovAdd.SelectedItem.SubItems(28)
        TxtSuppClmNonPayFr.mask = ""
        TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(29)
        TxtSuppClmNonPayTo.mask = ""
        TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(30)
        txtCLNCovExclusion = lstCovAdd.SelectedItem.SubItems(31)
        txtSuppPayPremGross = lstCovAdd.SelectedItem.SubItems(32)
        txtSuppPayPremNet = lstCovAdd.SelectedItem.SubItems(33)
        txtSuppPayMCOGross = lstCovAdd.SelectedItem.SubItems(34)
        txtSuppPayMCONet = lstCovAdd.SelectedItem.SubItems(35)
        txtRenewalDisc = lstCovAdd.SelectedItem.SubItems(36)
        txtPolicyDisc = lstCovAdd.SelectedItem.SubItems(37)
        
        Call EnableCovered(False)
        cmdUpdate.Enabled = True
        cmdRemove.Enabled = True
        cmdAdd.Enabled = False
    End If
    Call EnableCovered(True)
End Sub

'  Contorl of Tab index  *********************************88

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        If Trim(Mid(cboINSCode, 1, 1)) = "I" Or Trim(Mid(cboINSCode, 1, 1)) = "G" Then
            MsgBox "No Covered Person Needed For This Insured Type", vbOKOnly + vbCritical, DialogBoxTitle
            cboINSCode.SetFocus
            SSTab1.Tab = 0
        Else
            cboCovSalutation.Enabled = True
            cboCovSalutation.SetFocus
        End If
    ElseIf SSTab1.Tab = 0 Then
        If txtMBMName.Enabled = True Then
            cboMBMRenewal.SetFocus
        End If
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Call EnterToTab(KeyAscii)
End Sub

Private Sub cmdExit_LostFocus()
    If Me.ActiveControl.TabIndex = 39 Then
        cboMBMRenewal.SetFocus
    ElseIf Me.ActiveControl.TabIndex = 53 Then
        cboMBMRenewal.SetFocus
    End If
End Sub

Private Sub cboMBMRelCode_GotFocus()
    If Mid(cboMBMRelCode, 1, 1) = "S" Then
        Call SearchSuppLmt
    End If
End Sub
Private Sub SSTab1_GotFocus()
    If SSTab1.Tab = 0 Then
        txtMBMAdd1.SetFocus
    End If
End Sub

Private Sub SSTab1_LostFocus()
    If SSTab1.Tab = 0 Then
        txtMBMAdd1.SetFocus
    'ElseIf SSTab1.Tab = 1 Then
    '    cboMBMOccupation.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        cboCovSalutation.SetFocus
    End If
End Sub

Private Sub txtCovAllergic_Click()
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
                    MsgBox "Birthdate is out of acceptable range! " & vbNewLine & _
                            " It must be in between " & MBMValidDate!MBMDOBStart & " and " & MBMValidDate!MBMDOBEnd, vbCritical + vbOKOnly
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

Private Sub TxtGenWP_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
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

Private Sub txtMBMAllergic_Click()
    KeyPreview = False
End Sub

Private Sub txtMBMAllergic_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SSTab1.Tab = 2
        SSTab1.SetFocus
        tabIndexDone = True
    End If
End Sub

Private Sub EnterToTab(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Then
        If KeyAscii = 13 Then   'Turn off the beep
            KeyAscii = 0
        End If
        If tabIndexDone = False Then
            SendKeys "{tab}"
        End If
    End If
    tabIndexDone = False
End Sub

Private Sub Combo1_GotFocus()
    cmdClear.SetFocus
End Sub
Private Sub txtMBMAllergic_LostFocus()
    KeyPreview = True
End Sub

Private Sub txtMBMDateJoin_LostFocus()
    If txtMBMDateJoin <> "__/__/____" Then
        If IsDate(txtMBMDateJoin) = False Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtMBMDateJoin.SetFocus
        End If
    End If
End Sub

Private Sub txtMBMExclusion_Click()
    KeyPreview = False
End Sub

Private Sub txtMBMExclusion_LostFocus()
    KeyPreview = True
End Sub

' ******************************

' ######################### END INITIALIZATION ###########

' ######################### Start Checking  ###########
Private Sub txtMBMIcBcPP_LostFocus()
    Dim CheckRecord
    Dim MBMIC As New ADODB.Recordset
    Dim strsql As String
    Dim msgstr As String
    Dim RecordConfirm
    Dim CheckFlag As Boolean
    
    If Len(Trim(txtMBMIcBcPP)) Then
        Screen.MousePointer = vbHourglass
        If CheckFlag = False Then
            strsql = "Select MBMNumber From MBM Where MBMICBCPP ='" & Trim(txtMBMIcBcPP) & "'"
            MBMIC.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            
           If MBMIC.recordCount >= 1 Then
                msgstr = "Duplicate IC/BC/PP In The Membership" & Chr(13) & "Membership No: " & MBMIC!MBMNumber & vbNewLine & _
                        " Do you want to accept it? "
                RecordConfirm = MsgBox(msgstr, vbCritical + vbYesNoCancel, DialogBoxTitle)
                If RecordConfirm = vbNo Then
                    txtMBMIcBcPP.SetFocus
                    txtMBMIcBcPP.SelStart = 0
                    CheckFlag = False
                ElseIf RecordConfirm = vbCancel Then
                    txtMBMIcBcPP = ""
                    txtMBMIcBcPP.SetFocus
                    CheckFlag = False
                Else
                    txtMBMPolicyNo.SetFocus
                    CheckFlag = True
                End If
            End If
            MBMIC.Close
            Set MBMIC = Nothing
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cboMBMSalutation_Click()
  On Error Resume Next
    If ChkStr(cboMBMSalutation) = "MR" Then
        cboMBMSex.ListIndex = 2
    ElseIf ChkStr(cboMBMSalutation) = "" Then
        cboMBMSex.ListIndex = 0
    Else
        cboMBMSex.ListIndex = 1
    End If
End Sub

Private Sub cboMBMSalutation_CHANGE()
    If ChkStr(cboMBMSalutation) = "MR" Then
        cboMBMSex.ListIndex = 2
    ElseIf ChkStr(cboMBMSalutation) = "" Then
        cboMBMSex.ListIndex = 0
    Else
        cboMBMSex.ListIndex = 1
    End If
End Sub


Private Sub cboCovSalutation_Click()
    If ChkStr(cboCovSalutation) = "MR" Then
        cboCovSex = "M"
    ElseIf ChkStr(cboCovSalutation) = "" Then
        cboCovSex = ""
    Else
        cboCovSex = "F"
    End If
End Sub

Private Sub cboCovSalutation_Change()
    If ChkStr(cboCovSalutation) = "MR" Then
        cboCovSex = "M"
    ElseIf ChkStr(cboCovSalutation) = "" Then
        cboCovSex = ""
    Else
        cboCovSex = "F"
    End If
End Sub

Private Sub txtMBMBirthdate_LostFocus()
    Dim MBMValidDate As New ADODB.Recordset
    Dim strsql As String
    Dim DateExit
    
    If txtMBMBirthdate <> "__/__/____" Then
        If IsDate(txtMBMBirthdate) Then
           strsql = "Select ValidID ,MBMDOBStart , MBMDOBEnd  From MBMValidDateRange "
           MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
            If MBMValidDate.recordCount Then
                If DateDiff("d", MBMValidDate!MBMDOBStart, txtMBMBirthdate) >= 0 And DateDiff("d", MBMValidDate!MBMDOBEnd, txtMBMBirthdate) <= 0 Then
                    If txtMBMPayorEffDate <> "__/__/____" Then
                        Call MemberCallAge
                        Call CallLimit
                        'Call ProviderCallFees
                        Call MemberCallFees
                        If txtMBMPremium = "0" Or Trim(txtMBMPremium) = "" Then
                            Call MemberCallFees
                        End If
                    End If
                Else
                    DateExit = MsgBox("Uncertain Birthdate, please select OK to continue " & vbNewLine & _
                                " Otherwise press Cancel to make changes! ", vbExclamation + vbOKCancel + vbDefaultButton2)
                    If DateExit = vbCancel Then
                        txtMBMBirthdate.SetFocus
                    Else
                        Call MemberCallAge
                        Call CallLimit
                        'Call ProviderCallFees
                        Call MemberCallFees
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
            txtMBMBirthdate.SetFocus
        End If
    End If
    frmMBMClinical.txtMBMBirthdate = txtMBMBirthdate
End Sub

Private Sub cboPLNCode_Click()
    If Len(cboPLNCode) >= 2 Then
        Call showGroup
        If txtMBMPayorEffDate <> "__/__/____" Then
            Call CallLimit
            Call MemberCallFees
            Call showGroup
        End If
    End If
End Sub

Private Sub cboPLNCode_Change()
    If Len(cboPLNCode) >= 2 Then
        Call showGroup
        If txtMBMPayorEffDate <> "__/__/____" Then
            Call CallLimit
            Call MemberCallFees
        
        End If
    End If
End Sub

Private Sub cboINSCode_Click()
    If Len(Trim(cboINSCode)) Then
        If Trim(cboINSCode) = "F" Or Trim(cboINSCode) = "H" Then
            cboMBMMaritalStatus = "M"
            If Trim(cboINSCode) <> Trim(txtINSFlag) Then
                Call EnableSupplementary(True)
            Else
                Call EnableSupplementary(False)
            End If
        ElseIf Trim(cboINSCode) = "I" Or Trim(cboINSCode) = "G" Then
            cboMBMMaritalStatus = "S"
            Call EnableSupplementary(False)
        End If
    End If
    If Len(Trim(cboPLNCode)) = 0 Then
        Call LoadPlanType(cboINSCode.Text)
    End If
    If Len(cboPLNCode) >= 2 Then
        Call CallLimit
        Call MemberCallFees
    End If
End Sub

Private Sub cboINSCode_Change()
    If Len(Trim(cboINSCode)) Then
        If Trim(cboINSCode) = "F" Or Trim(cboINSCode) = "H" Then
            cboMBMMaritalStatus = "M"
            If Trim(cboINSCode) <> Trim(txtINSFlag) Then
                Call EnableSupplementary(True)
            Else
                Call EnableSupplementary(False)
            End If
        ElseIf Trim(cboINSCode) = "I" Or Trim(cboINSCode) = "G" Then
            Call EnableSupplementary(False)
        End If
    End If
    If Len(Trim(cboPLNCode)) = 0 Then
        Call LoadPlanType(cboINSCode.Text)
    End If
    If Len(cboPLNCode) >= 2 Then
        Call CallLimit
        Call MemberCallFees
    End If

End Sub

Private Sub cboMBMRenewal_GotFocus()
    Me.KeyPreview = True
End Sub

Private Sub cboMBMRenewal_Click()
    If ChkStr(cboMBMRenewal) = "R" Then
        cmdSearch.Enabled = True
        cmdShow.Enabled = True
        txtMBMPrevMEMNo.Enabled = True
    Else
        cmdSearch.Enabled = False
        cmdShow.Enabled = False
        txtMBMPrevMEMNo.Enabled = False
        txtMBMPrevMEMNo = ""
        txtMBMPrevPolNo = ""
    End If
End Sub

Private Sub cboMBMRenewal_Change()
    If ChkStr(cboMBMRenewal) = "R" Then
        cmdSearch.Enabled = True
        cmdShow.Enabled = True
        txtMBMPrevMEMNo.Enabled = True
    Else
        cmdSearch.Enabled = False
        cmdShow.Enabled = False
        txtMBMPrevMEMNo = ""
        txtMBMPrevPolNo = ""
        txtMBMPrevMEMNo.Enabled = False
    End If
End Sub

Private Sub cmdSearch_Click()
    Call ClearEntries
    cboMBMRenewal = "R"
    bDataStatus = True
    frmSearchMBM.Source = 1
    frmSearchMBM.show vbModal
End Sub


Private Sub txtMBMPayorEffDate_LostFocus()
Dim MBMValidDate As New ADODB.Recordset
Dim strsql As String
Dim tempdate As Date
Dim answer

    If (txtMBMPayorEffDate) <> "__/__/____" Then
        If IsDate(txtMBMPayorEffDate) Then
            If ChkStr(cboMBMRenewal) = "N" Then
                txtMBMDateJoin = txtMBMPayorEffDate
            Else
                If txtMBMDateJoin = "__/__/____" Then
                    txtMBMDateJoin = txtMBMPayorEffDate
                End If
            End If
           strsql = "Select ValidID , MBMEffStart, MBMEffEnd From MBMValidDateRange "
           MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                           
           If MBMValidDate.recordCount Then
                If DateDiff("d", MBMValidDate!MBMEffStart, txtMBMPayorEffDate) >= 0 And DateDiff("d", MBMValidDate!MBMEffEnd, txtMBMPayorEffDate) <= 0 Then
                    tempdate = Format(DateAdd("yyyy", 1, txtMBMPayorEffDate), "dd/mm/yyyy")
                    txtMBMPayorExpDate = Format(DateAdd("d", -1, tempdate), "dd/mm/yyyy")
                    'txtMBMPayorEffDate.SetFocus
                    txtMBMEmployeeNo.SetFocus

                If txtMBMBirthdate <> "__/__/____" Then
                    Call MemberCallAge
                    Call MemberCallFees
                    'Call ProviderCallFees
                    Call CallLimit
                End If
                    
                Else
                    MsgBox "Policy Effective Date is out of acceptable range! " & vbNewLine & _
                            "It must be in between " & MBMValidDate!MBMEffStart & " and " & MBMValidDate!MBMEffEnd & "! ", vbCritical + vbOKOnly
                    txtMBMPayorEffDate.SetFocus
                End If
           Else
                tempdate = Format(DateAdd("yyyy", 1, txtMBMPayorEffDate), "dd/mm/yyyy")
                txtMBMPayorExpDate = Format(DateAdd("d", -1, tempdate), "dd/mm/yyyy")
                
                If txtMBMBirthdate <> "__/__/____" Then
                    Call MemberCallAge
                    Call MemberCallFees
                    'Call ProviderCallFees
                    Call CallLimit
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
    txtTempEffDate = txtMBMPayorEffDate
    txtTempExpDate = txtMBMPayorExpDate
End Sub

Private Sub txtMBMPayorExpDate_KeyPress(KeyAscii As Integer)
    MBMPayorEXPDate = 0
    MBMPayorEXPDate = 1
End Sub

Private Sub txtMBMPayorExpDate_LostFocus()
    Dim MBMValidDate As New ADODB.Recordset
    Dim strsql As String
    Dim answer
    Dim CheckDays As Integer
    Dim Confirm1
    Dim Confirm2
    
    CheckDays = 0
    
    If txtMBMPayorExpDate <> "__/__/____" Then
        If IsDate(txtMBMPayorExpDate) = True Then
            strsql = "Select ValidID , MBMExpStart, MBMExpEnd From MBMValidDateRange "
            MBMValidDate.MaxRecords = 1
            MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                 
            If MBMValidDate.recordCount Then
                If DateDiff("d", MBMValidDate!MBMExpStart, txtMBMPayorExpDate) >= 0 And DateDiff("d", MBMValidDate!MBMExpEnd, txtMBMPayorExpDate) <= 0 Then
                Else
                    MsgBox "Policy Expiry Date is out of acceptable range! " & vbNewLine & _
                            " It must be in between " & MBMValidDate!MBMExpStart & " and " & MBMValidDate!MBMExpEnd, vbCritical + vbOKOnly
                    txtMBMPayorExpDate.SetFocus
                End If
            End If
            MBMValidDate.Close
            Set MBMValidDate = Nothing
        
            If txtMBMPayorEffDate <> "__/__/____" And txtTempEffDate <> "__/__/____" And txtTempExpDate <> "__/__/____" And txtMBMPayorExpDate <> "__/__/____" Then
                If IsDate(txtMBMPayorEffDate) = True And IsDate(txtMBMPayorExpDate) Then
                 
                 'CK added --------------------------------------------------
                 
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
                 
                 
                 'CK added --------------------------------------------------
                 
                 
                 'Call InsertMCO
                 'Call InsertPrem
                 'Call InsertProviderFees
                
                 If CDate(txtMBMPayorEffDate) = CDate(txtTempEffDate) And CDate(txtTempExpDate) = CDate(txtMBMPayorExpDate) Then
                     txtMCOFees = txtTempMCOFees
                     txtMBMPremium = txtTempPrem
                     txtMBMProviderCharges = txtTempProvider
                 End If
                End If
            End If
        Else
            MsgBox "Invalid Date Format!", vbCritical + vbOKOnly
'            If answer = vbOK Then
             txtMBMPayorExpDate.SetFocus
 '           End If
        End If
    End If
End Sub


Private Sub cboCovRelationship_cLICK()
    If Mid(ChkStr(cboCovRelationship), 1, 1) = "W" Or Mid(ChkStr(cboCovRelationship), 1, 1) = "H" Then
        cboCovAdultChild = "A"
    Else
        cboCovAdultChild = "C"
    End If
End Sub

Private Sub cboCovRelationship_CHANGE()
    If Mid(ChkStr(cboCovRelationship), 1, 1) = "W" Or Mid(ChkStr(cboCovRelationship), 1, 1) = "H" Then
        cboCovAdultChild = "A"
    Else
        cboCovAdultChild = "C"
    End If
End Sub

' ####################### Calculation's Function ###############
Private Sub showGroup()
    Dim strsql As String
    Dim PLN As New ADODB.Recordset
    strsql = "Select PLNCode ,  GRPCompany from PLN where  0=0 "
    strsql = strsql & " AND PLNCode= '" & Trim(cboPLNCode) & "' AND grpcOMPANY IS NOT nuLL"  'need to reverify
    PLN.MaxRecords = 1
    PLN.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If Not PLN.EOF Then
        txtMBMGroupCompany = PLN!GrpCompany
    End If
    PLN.Close
    Set PLN = Nothing
End Sub

Private Sub MemberCallAge()
    Dim SelectAge As String
    If Trim(txtMBMPayorEffDate) <> "__/__/____" _
        And Len(Trim(txtMBMBirthdate)) <> 0 _
        And Trim(txtMBMBirthdate) <> "__/__/____" _
        And Len(Trim(txtHLTCode)) <> 0 Then
        If IsDate(txtMBMBirthdate) And IsDate(txtMBMPayorEffDate) Then
            SelectAge = GetAge(txtMBMPayorEffDate, txtMBMBirthdate)
            txtAGECode = GetAgeBand(Trim(txtHLTCode), SelectAge)
            If Len(Trim(txtAGECode)) <> 0 Then
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

    If Trim(txtPAYCode) <> "" And Trim(txtAGECode) <> "" And Trim(cboPLNCode) <> "" _
        And Trim(txtHLTCode) <> "" Then
        txtMBMPremium = "0"
        txtMCOFees = "0"
        If Len(Trim(txtPAYCode)) <> 0 Then
            txtMBMPremium = Format(GetPremium(Trim(Mid(cboINSCode, 1, 1)), Trim(txtPAYCode), Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)), _
            Trim(cboPLNCode), Trim(txtHLTCode), txtMBMPayorEffDate), "#,##0.00")
            
            txtTempPrem = txtMBMPremium

            'strsqlPLN = "select * from DiscPlan where PLNCode = '" & Trim(Mid(cboPLNCode, 1, 4)) & "'"
            'PLN.Open strsqlPLN, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
            'If PLN.recordCount Then
                If PolDics > 0 Then
                    txtMBMPremium = txtMBMPremium * (100 - PolDics) / 100
                End If
                If RenDisc > 0 Then
                    txtMBMPremium = txtMBMPremium * (100 - RenDisc) / 100
                End If
            'End If
            'PLN.Close
            'Set PLN = Nothing
        End If
        
        txtMCOFees = Format(GetMCO(Trim(Mid(cboINSCode, 1, 1)), Trim(txtHLTCode), txtMBMAdultChild, txtMBMPayorEffDate, Trim(Mid(cboPLNCode, 1, 4))), "#,##0.00")
        txtTempMCOFees = txtMCOFees
    
    End If
End Sub

Public Sub ProviderCallFees()
    Dim GetProString As String
    
    If Trim(cboINSCode) <> "" And Trim(cboSRVCode) <> "" And Trim(txtHLTCode) <> "" Then
        txtMBMProviderCharges = ""
        If InStr(1, cboSRVCode, "-") <> 0 Then
            GetProString = GetProviderCharges(Trim(Mid(cboSRVCode, 1, InStr(1, cboSRVCode, "-") - 1)), Trim(cboINSCode), Trim(txtHLTCode), txtMBMAdultChild)
        Else
            GetProString = GetProviderCharges(Trim(cboSRVCode), Trim(cboINSCode), Trim(txtHLTCode), txtMBMAdultChild)
        End If

        txtMBMProviderCharges = Format(GetProString, "#,##0.00")
    Else
        txtMBMProviderCharges = 0
    End If
      txtTempProvider = txtMBMProviderCharges
End Sub

Public Sub CallLimit()
        txtAnnualLimit = ""
        If Len(Trim(cboPLNCode)) And IsDate(txtMBMPayorEffDate) Then
            txtAnnualLimit = Format(GetAnnualLimit(Trim(cboINSCode), Trim(cboPLNCode), Trim(txtHLTCode), txtMBMPayorEffDate), "#,###,##0.00")
        End If
End Sub

' ######################### END CALCULATIOn###########

' ######################### BUTTON  CLICKING ###########
Public Sub cmdshow_Click()
    On Error Resume Next
    Dim sqlstr As String
    Dim MBM As New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    sqlstr = "Select MBMNumber, HLTCode, MBMPolicyNo,  MBMName , MBMIcBcPp,  " & _
            " MBMMemberType ,  MBMStatus, " & _
            " INSCode, PLNCode, " & _
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
            " MBMMaritalStatus, " & _
            " MBMFirstJoinedDate , BRNCode, " & _
            " MBMDOB, NATCode, RACCode," & _
            " LGNCode, STACode, MBMFirstPolNo , MBMFirstMEMNo,  " & _
            " BRNCode, MBMTakeOver, MBMAgentCode , MBMPolicyYear , " & _
            " MBMDepartment,MBMPolicyDisc, MBMRenewalDisc,MBMRemarks"
    sqlstr = sqlstr & " From View_MBM_MBMOthers Where 0=0 "
    sqlstr = sqlstr & " AND  MBMNumber  = '" & Trim(txtMBMPrevMEMNo) & "'"
    If bDataStatus = True Then
        MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    Else
        MBM.Open sqlstr, medixmasterconnA, adOpenKeyset, adLockOptimistic
    End If
    
    If MBM.recordCount = 0 Then
       MsgBox "No record found!", vbCritical + vbOKOnly
    Else
        txtMBMName = Trim(MBM!MBMName)
        txtMBMIcBcPP = Trim(MBM!MBMIcBcPp)
        txtMBMPrevPolNo = Trim(MBM!MBMPolicyNo)
        If Len(Trim(MBM!INSCode)) Then
            cboINSCode = Trim(MBM!INSCode)
        End If
        If Len(Trim(MBM!PLNCode)) Then
            cboPLNCode = Trim(MBM!PLNCode)
        End If
        If Len(Trim(MBM!MBMAddress1)) Then
            txtMBMAdd1 = Trim(MBM!MBMAddress1)
        End If
        If Len(Trim(MBM!MBMAddress2)) Then
            txtMBMAdd2 = Trim(MBM!MBMAddress2)
        End If
        If Len(Trim(MBM!MBMAddress3)) Then
            txtMBMAdd3 = Trim(MBM!MBMAddress3)
        End If
        If Len(Trim(MBM!MBMPostCode)) Then
            txtMBMPostCode = Trim(MBM!MBMPostCode)
        End If
        If Len(Trim(MBM!CTYCode)) Then
            cboCTYCode = Trim(MBM!CTYCode)
        End If
        If Len(Trim(MBM!MBMSalutation)) Then
            cboMBMSalutation = Trim(MBM!MBMSalutation)
        End If
        If Len(Trim(MBM!MBMSex)) Then
            cboMBMSex = Trim(MBM!MBMSex)
        End If
        If Len(Trim(MBM!SRVCodeList)) Then
            cboSRVCode = Trim(MBM!SRVCodeList)
        End If
        If Len(Trim(MBM!MBMGroupCompany)) Then
            txtMBMGroupCompany = Trim(MBM!MBMGroupCompany)
        End If
        If Len(Trim(MBM!MBMEmployeeNo)) Then
            txtMBMEmployeeNo = Trim(MBM!MBMEmployeeNo)
        End If
        If Len(MBM!MBMRemarks) Then
            txtRemarks = MBM!MBMRemarks
        End If
        If Len(Trim(MBM!MBMExclusion)) Then
            txtMBMExclusion = Trim(MBM!MBMExclusion)
        End If
        If Len(Trim(MBM!MBMMaritalStatus)) Then
            cboMBMMaritalStatus = Trim(MBM!MBMMaritalStatus)
        End If
        If Len(Trim(MBM!BRNCode)) Then
            cboBRNCode = Trim(MBM!BRNCode)
        End If
        If Len(Trim(MBM!MBMDOB)) Then
            txtMBMBirthdate.mask = ""
            txtMBMBirthdate = Trim(MBM!MBMDOB)
            txtMBMBirthdate.mask = "##/##/####"
        End If
        If Len(Trim(MBM!NATCode)) Then
            cboNATCode = Trim(MBM!NATCode)
        End If
        If Len(Trim(MBM!RACCode)) Then
            cboRACCode = Trim(MBM!RACCode)
        End If
        If Len(Trim(MBM!STACode)) Then
            cboSTACode = Trim(MBM!STACode)
        End If
        If Len(Trim(MBM!MBMFirstPolNo)) Then
            txtFirstPolNo = Trim(MBM!MBMFirstPolNo)
        End If
        If Len(Trim(MBM!MBMFirstMEMNo)) Then
            txtFirstMeMNo = Trim(MBM!MBMFirstMEMNo)
        End If
        
        If Len(Trim(MBM!MBMTakeOver)) Then
            cboMBMTakeover.Enabled = True
            cboMBMTakeover = Trim(MBM!MBMTakeOver)
        End If
        If Len(Trim(MBM!MBMAgentCode)) Then
            txtMBMAgentCode = Trim(MBM!MBMAgentCode)
        End If
        
        ' P if it is Normal , Q if it is Johor Port
        If ChkStr(MBM!MBMMemberType) = "P" Then
            cboMBMRelCode = "C - Combined Limit"
        Else
            cboMBMRelCode = "S - Separated Limit"
        End If
        If Trim(MBM!MBMFirstJoinedDate) <> "" Then
            txtMBMDateJoin = MBM!MBMFirstJoinedDate
        End If
        
        txtMBMPolicyYear = IIf(Trim(MBM!MBMPolicyYear) = "", 1, MBM!MBMPolicyYear + 1)
        If Trim(MBM!MBMDepartment) <> "" Then
            txtMBMDepartment = MBM!MBMDepartment
        End If
                
        If Len(MBM!MBMPolicyDisc) Then
            PolDics = MBM!MBMPolicyDisc
        End If
        
        If Len(MBM!MBMRenewalDisc) Then
            RenDisc = MBM!MBMRenewalDisc
        End If
        Call showMBMOthers
        Call showMBMCov
        Call CallLimit
        Call MemberCallFees

    End If
    'Call ProviderCallFees
    Screen.MousePointer = Default
End Sub

Private Sub txtCovBirthdate_gotfocus()
    txtCovBirthdate.SelStart = 0
End Sub

Private Sub txtMBMBirthdate_gotfocus()
    txtMBMBirthdate.SelStart = 0
End Sub

Private Sub txtMBMPayorEffDate_gotfocus()
    txtMBMPayorEffDate.SelStart = 0
End Sub

Private Sub CmdClear_Click()
On Error Resume Next
Dim RecordSaved
      RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            txtMBMBordxDate = txtTempBordx
            txtMBMBatchNo = txtTempBatch
            Call ClearEntries
            Call EnableEntries(True)
            cboMBMRenewal.SetFocus
            cboMBMRenewal = "N"
            txtMBMPolicyYear = 1
            
        End If
End Sub

Private Sub cmdExit_Click()

    Unload frmMemberRegistration
    Set frmMemberRegistration = Nothing

End Sub


Private Sub cmdSave_Click()
On Error GoTo ErrorSave
    Dim RecordSaved
    If cboMBMRenewal = "R" And _
       Len(Trim(txtMBMPrevMEMNo)) = 0 And _
       Len(Trim(txtMBMPrevPolNo)) = 0 Then
        MsgBox "Please Enter Previous Membership No or Previous Policy No", vbOKOnly + vbCritical, DialogBoxTitle
        cmdShow.SetFocus
    Else
        If Len(Trim(txtMBMName)) = 0 Then
            MsgBox "Please Enter Membership Name", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMName.SetFocus
        ElseIf Len(Trim(txtMBMPolicyNo)) = 0 Then
            MsgBox "Please Enter Policy No", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMPolicyNo.SetFocus
        
        '=== Clinical ======================================================================================
        
        ElseIf (Mid(cboHNC, 1, 1) = "H" And txtMBMPayorEffDate = "__/__/____") Or _
            (Mid(cboHNC, 1, 1) = "H" And Len(Trim(txtMBMPayorEffDate)) = 0) Or _
            (Mid(cboHNC, 1, 1) = "M" And txtMBMPayorEffDate = "__/__/____") Or _
            (Mid(cboHNC, 1, 1) = "M" And Len(Trim(txtMBMPayorEffDate)) = 0) Then
            MsgBox "Please Enter Payor Effective Date", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMPayorEffDate.SetFocus
        ElseIf (Mid(cboHNC, 1, 1) = "H" And txtMBMPayorExpDate = "__/__/____") Or _
            (Mid(cboHNC, 1, 1) = "H" And Len(Trim(txtMBMPayorExpDate)) = 0) Or _
            (Mid(cboHNC, 1, 1) = "M" And txtMBMPayorExpDate = "__/__/____") Or _
            (Mid(cboHNC, 1, 1) = "M" And Len(Trim(txtMBMPayorExpDate)) = 0) Then
            MsgBox "Please Enter Payor Expiry Date", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMPayorExpDate.SetFocus
        ElseIf (Mid(cboHNC, 1, 1) = "H" And Len(Trim(cboINSCode)) = 0) Or _
            (Mid(cboHNC, 1, 1) = "M" And Len(Trim(cboINSCode)) = 0) Then
            MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
            cboINSCode.SetFocus
        ElseIf (Mid(cboHNC, 1, 1) = "H" And cboPLNCode = "") Or _
            (Mid(cboHNC, 1, 1) = "M" And cboPLNCode = "") Then
            MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
            cboPLNCode.SetFocus
        
        '=== Clinical ======================================================================================
        
        '=== H & S Only ======================================================================================
        'ElseIf txtMBMPayorEffDate = "__/__/____" Then 'And Len(Trim(txtMBMPayorEffDate)) = 0 Then
            'MsgBox "Please Enter Payor Effective Date", vbOKOnly + vbCritical, DialogBoxTitle
            'txtMBMPayorEffDate.SetFocus
        'ElseIf txtMBMPayorExpDate = "__/__/____" Then 'And Len(Trim(txtMBMPayorExpDate)) = 0 Then
            'MsgBox "Please Enter Payor Expiry Date", vbOKOnly + vbCritical, DialogBoxTitle
            'txtMBMPayorExpDate.SetFocus
        'ElseIf Len(Trim(cboINSCode)) = 0 Then
            'MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
            'cboINSCode.SetFocus
        'ElseIf cboPLNCode = "" Then
            'MsgBox "Please Enter Plan Code", vbOKOnly + vbCritical, DialogBoxTitle
            'cboPLNCode.SetFocus
        '=== H & S Only ======================================================================================
        
        ElseIf Len(Trim(txtMBMIcBcPP)) = 0 Then
            MsgBox "Please Enter IC/BC/PP", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMIcBcPP.SetFocus
        ElseIf Len(Trim(txtHLTCode)) = 0 Then
            MsgBox "Please Enter Health Code", vbOKOnly + vbCritical, DialogBoxTitle
            txtHLTCode.SetFocus
        ElseIf Len(Trim(txtPAYCode)) = 0 Then
            MsgBox "Please Enter Payor", vbOKOnly + vbCritical, DialogBoxTitle
            txtPAYCode.SetFocus
        ElseIf Len(Trim(txtMBMAdd1)) = 0 Then
            MsgBox "Please Enter Address", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMAdd1.SetFocus
        ElseIf Len(Trim(txtMBMPostCode)) = 0 Then
            MsgBox "Please Enter Post Code", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMPostCode.SetFocus
        ElseIf Len(Trim(cboSTACode)) = 0 Then
            MsgBox "Please Enter State", vbOKOnly + vbCritical, DialogBoxTitle
            cboSTACode.SetFocus
        ElseIf txtMBMBirthdate = "__/__/____" Then
            MsgBox "Please Enter Birthdate", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMBirthdate.SetFocus
        ElseIf cboSRVCode = "" Then
            MsgBox "Please Enter Service Provider Code", vbOKOnly + vbCritical, DialogBoxTitle
            cboSRVCode.SetFocus
        ElseIf IsDate(txtMBMBirthdate) = False Then
            MsgBox "Invalid BirthDate", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMBirthdate = "__/__/____"
            txtMBMBirthdate.SetFocus
        
        '=== Clinical ======================================================================================
        ElseIf cboHNC.Text = "" Then
            MsgBox "Please Enter H&C Status", vbOKOnly + vbCritical, DialogBoxTitle
            cboHNC.SetFocus
        ElseIf Mid(cboHNC, 1, 1) = "C" And frmMBMClinical.cboPLNCode = "" And frmMBMClinical.cboINSCode = "" And frmMBMClinical.txtMBMPayorEffDate = "__/__/____" And frmMBMClinical.txtMBMPayorExpDate = "__/__/____" And _
               frmMBMClinical.lblAnnualLimit = "" And frmMBMClinical.lblMBMPremium = "" And frmMBMClinical.lblMCOFees = "" Then
               MsgBox "Please complete Member Clinical Form!", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf Mid(cboHNC, 1, 1) = "M" And frmMBMClinical.cboPLNCode = "" And frmMBMClinical.cboINSCode = "" And frmMBMClinical.txtMBMPayorEffDate = "__/__/____" And frmMBMClinical.txtMBMPayorExpDate = "__/__/____" And _
               frmMBMClinical.lblAnnualLimit = "" And frmMBMClinical.lblMBMPremium = "" And frmMBMClinical.lblMCOFees = "" Then
               MsgBox "Please complete Member Clinical Form!", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf Len(Trim(txtAGECode)) = 0 Then
            MsgBox "Invalid Age Group", vbOKOnly + vbCritical, DialogBoxTitle
        '=== Clinical ======================================================================================
            
        
        Else
            RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
              If RecordSaved = vbYes Then
                  Screen.MousePointer = vbHourglass
                  Call NewRegistrationSub
                  txtMBMPolicyYear = 1
                  Screen.MousePointer = vbDefault
              End If
        End If
    End If
    
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If startTrans = True Then
        medixmasterconn.RollbackTrans
    End If
    MsgBox ERR.Description
End Sub

' ######### To save the membership data ---> cater for new registration ############3
Private Sub NewRegistrationSub()
'On Error Resume Next
Dim MBM As New ADODB.Recordset
Dim planstring
Dim MembershipNo As String
Dim RecordSaved
Dim completeMBM, completeMBMOthers, completeMBMC, completeMBMCLNOthers As Boolean
    
    strsqlPLN = "select * from DiscPlan where PLNCode = '" & Trim(Mid(cboPLNCode, 1, 4)) & "'"
    PLN.Open strsqlPLN, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    
    If PLN.recordCount Then
        Disc = True
    End If
    
    PLN.Close
    Set PLN = Nothing
    
    planstring = cboPLNCode
    ' INsert into Membership table
    'MBM.Open "Select MBMNumber, * from MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
    MBM.Open "MBM", medixmasterconn, adOpenKeyset, adLockOptimistic
    ' Begin transaction
    medixmasterconn.beginTrans
    startTrans = True
    
    
    If Mid(cboHNC, 1, 1) = "C" Then
        MembershipNo = Trim(txtPAYCode) & Trim(Mid(frmMBMClinical.cboPLNCode, 1, IIf(InStr(1, frmMBMClinical.cboPLNCode, "-") = 0, 4, InStr(1, frmMBMClinical.cboPLNCode, "-") - 1))) & Trim(Mid(frmMBMClinical.cboINSCode, 1, 1)) & GenerateMembershipNo(Trim(txtPAYCode), Trim(Mid(frmMBMClinical.cboPLNCode, 1, IIf(InStr(1, frmMBMClinical.cboPLNCode, "-") = 0, 4, InStr(1, frmMBMClinical.cboPLNCode, "-") - 1))), Trim(Mid(frmMBMClinical.cboINSCode, 1, 1)))
    Else
        MembershipNo = Trim(txtPAYCode) & Trim(cboPLNCode) & Trim(cboINSCode) & GenerateMembershipNo(Trim(txtPAYCode), Trim(cboPLNCode), Trim(cboINSCode))
    End If
    
    
    MBM.AddNew
         With MBM
             !MBMNumber = ChkStr(MembershipNo)
             !MBMCOVID = "00"
             !MBMPolicyNo = ChkStr(txtMBMPolicyNo)
             !MBMIcBcPp = ChkStr(txtMBMIcBcPP)
             !MBMName = ChkStr(txtMBMName)
             !MBMSalutation = ChkStr(cboMBMSalutation)
             !HLTCode = Trim(Mid(Trim(txtHLTCode), 1, 2))
             !PAYCode = Trim(Mid(txtPAYCode, 1, 2))
             !MBMAddress1 = ChkStr(txtMBMAdd1)
             !MBMAddress2 = ChkStr(txtMBMAdd2)
             !MBMAddress3 = ChkStr(txtMBMAdd3)
             !CTYCode = ChkStr(cboCTYCode)
             !MBMPostCode = txtMBMPostCode
             !MBMTelnoH = FrmMBMOthers.txtMBMTelNoH
             !MBMTelNoM = FrmMBMOthers.txtMBMTelNoM
             !MBMTelNoO = FrmMBMOthers.txtMBMTelNoO
             !MBMEmail = ChkStr(FrmMBMOthers.txtMBMEmail)
             !MBMDOB = txtMBMBirthdate
             !NATCode = Trim(Mid(cboNATCode, 1, 2))
             !RACCode = Trim(Mid(cboRACCode, 1, 2))
             !MBMSex = cboMBMSex
             
             '======== cLINICAL ==========================
             If Trim(Mid(cboHNC, 1, 1)) = "H" Or Trim(Mid(cboHNC, 1, 1)) = "M" Then
                !MBMPayorEffDate = CDate(txtMBMPayorEffDate)
                !MBMPayorEXPDate = CDate(txtMBMPayorExpDate)
                !INSCode = Trim(Mid(cboINSCode, 1, 2))
                !PLNCode = Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") > 0, InStr(1, cboPLNCode, "-") - 1, 4)))
                !MBMAvailableLimit = txtAnnualLimit
             End If
             
             If Trim(Mid(cboHNC, 1, 1)) = "C" Or Trim(Mid(cboHNC, 1, 1)) = "M" Then
                !MBMCLNBalLmt = CDbl(frmMBMClinical.lblAnnualLimit)
                !MBMCLNVisitLmt = CDbl(frmMBMClinical.LblVisitLimit)
             End If
             '======== cLINICAL ==========================
             
             !AgeCode = Trim(Mid(txtAGECode, 1, 2))
             !MBMAdultChild = ChkStr(txtMBMAdultChild)
             !SRVCodeList = Trim(Mid(cboSRVCode, 1, IIf(InStr(1, cboSRVCode, "-") > 0, InStr(1, cboSRVCode, "-") - 1, 4)))
             !STACode = Trim(ChkStr(cboSTACode))
             !LGNCode = Trim(Mid(FrmMBMOthers.cboLGNCode, 1, 2))
             !MBMBordxDate = txtMBMBordxDate
             !MBMTakeOver = Mid(ChkStr(cboMBMTakeover), 1, 1)
             !MBMRenewFlag = cboMBMRenewal
             !BRNCode = ChkStr(cboBRNCode)
             !MBMStatus = "A"
             !MBMValueDate = Date
             !MBMPolicyYear = txtMBMPolicyYear
             !MBMTakeOver = Trim(Mid(cboMBMTakeover, 1, 1))
             ' Need to check whether the member is from Johor port or not
             If Mid(ChkStr(cboMBMRelCode), 1, 1) = "C" Then
                !MBMMemberType = "P"
             Else
                !MBMMemberType = "Q"
             End If
             !MBMExpiredStatus = "N"
             '========Clinical====================
             !HSClnInd = Mid(cboHNC, 1, 1)
             '========Clinical====================
             !MBMBatchNo = txtMBMBatchNo
             !MBMDataStatus = "A"
             !MBMRegUser = Logon
             !MBMLastUpdateUser = Logon
             .Update
     End With
    completeMBM = True
        
        'If MBMPayorEXPDate = 1 Then
        '    RecordSaved = MsgBox("Do you want to edit Expiry Date?", vbYesNo + vbExclamation)
        '    If RecordSaved = vbYes Then
        '        Call InsertMCO
        '        Call InsertPrem
        '        Call InsertProviderFees
        '    End If
        ' End If
    Call savembmothers(MembershipNo)
    completeMBMOthers = True
    
    If ChkStr(cboINSCode) = "F" Or ChkStr(cboINSCode) = "H" Then
        Call saveCoverPersons(MembershipNo, Mid(ChkStr(cboMBMRelCode), 1, 1))
    End If
    
    '========Clinical=======================
    If Mid(cboHNC, 1, 1) = "C" Or Mid(cboHNC, 1, 1) = "M" Then
        Call SaveMBMCLNOthers(MembershipNo)
    End If
    completeMBMCLNOthers = True
    '========Clinical=======================
    
    Call SaveBatch
    completeMBMC = True

    If completeMBM = True And completeMBMC = True And completeMBMOthers = True And completeMBMCLNOthers = True Then
    'If completeMBM = True And completeMBMC = True And completeMBMOthers = True Then
        ' commit changes
        medixmasterconn.CommitTrans
    Else
        'Roll back
        medixmasterconn.RollbackTrans
    End If
    startTrans = False
    MBM.Close
    Set MBM = Nothing
     
     
     cboMBMRenewal.Enabled = True
     MsgBox "Record Saved, Your Membership Number is : " & MembershipNo, vbOKOnly + vbInformation, DialogBoxTitle
     ' Clering previous record
     txtMBMPayorExpDate.Enabled = True
     Call ClearEntries
   
     SSTab1.Tab = 0
     cboMBMRenewal.SetFocus
End Sub

Private Sub SearchSuppLmt()
Dim strsqlannlmt As String
Dim AnnLmt As New ADODB.Recordset
Dim strsqlPrem As String
Dim Prem As New ADODB.Recordset
    If txtMBMPayorEffDate <> "__/__/____" Then
    
    strsqlannlmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(cboPLNCode) & "'"
    AnnLmt.Open strsqlannlmt, medixmasterconn, adOpenKeyset, adLockOptimistic

    strsqlPrem = "Select SuppPrmStatus, PLNCode, SuppPrmAmt from PRM where PLNCode = '" & Trim(cboPLNCode) & "' And PAYCode = '" & Trim(txtPAYCode) & "' And INSCode = '" & Trim(cboINSCode) & "' And HLTCode = '" & Trim(txtHLTCode) & "'"
    Prem.Open strsqlPrem, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    'Cater for SuppLimitStatus is "C"
    If AnnLmt.recordCount Then
        If AnnLmt!SuppLimitStatus <> "B" Then
            If AnnLmt!SuppLimitStatus = "C" Then
                tempSAnnLmt = Format(GetSuppLimit(cboINSCode, cboPLNCode, Mid(txtHLTCode, 1, 1), CDate(txtMBMPayorEffDate)), "#,##0.00")
                txtCovAnnualLimit = tempSAnnLmt
                tempSPremium = Format(IIf(IsNull(Prem!SuppPrmAmt), 0, Prem!SuppPrmAmt), "#,##0.00")
            Else
                tempSAnnLmt = txtAnnualLimit
                txtCovAnnualLimit = txtAnnualLimit
            End If
        End If
    End If
    AnnLmt.Close
    Prem.Close
    Set AnnLmt = Nothing
    Set Prem = Nothing

    End If

End Sub

Private Sub saveCoverPersons(MembershipNo As String, MBMtype As String)
Dim MBMCoveredPersons As New ADODB.Recordset
Dim listrecord As Integer
Dim listloop As Integer
Dim SelectAge As String
Dim strsqlannlmt As String
Dim strsqlPrem As String
Dim Prem As New ADODB.Recordset
Dim AnnLmt As New ADODB.Recordset
Dim MemberDOB As String
Dim FindAge As String
Dim AgeValue As String
Dim tempSPremium As String
Dim tempSAnnLmt As String
    
    listrecord = lstCovAdd.listitems.count
    ' Insert into Membership Covered Person's table
    
     If listrecord <> 0 Then
         MBMCoveredPersons.Open "MBMCOveredPersons", medixmasterconn, adOpenKeyset, adLockOptimistic
         For listloop = 1 To listrecord
             MBMCoveredPersons.AddNew
             With MBMCoveredPersons
                 !MBMCNumber = ChkStr(MembershipNo)
                 !MBMCSalutation = ChkStr(lstCovAdd.listitems(listloop).SubItems(1))
                 !MBMCName = ChkStr(lstCovAdd.listitems(listloop).Text)
                 !MBMCIcBcPpNo = ChkStr(lstCovAdd.listitems(listloop).SubItems(2))
                 !MBMCDOB = ChkStr(lstCovAdd.listitems(listloop).SubItems(3))
                 SelectAge = GetAge(txtMBMPayorEffDate, lstCovAdd.listitems(listloop).SubItems(3))
                 AgeValue = ChkStr(Mid(GetAgeBand(Trim(txtHLTCode), SelectAge), 1, 2))
                 !MBMCAgeband = AgeValue
                 !MBMCSex = ChkStr(lstCovAdd.listitems(listloop).SubItems(4))
                 !MBMCAdultChild = ChkStr(lstCovAdd.listitems(listloop).SubItems(5))
                 !MBMCRELCode = ChkStr(Mid(lstCovAdd.listitems(listloop).SubItems(6), 1, 1))
                 !MBMCBloodType = ChkStr(lstCovAdd.listitems(listloop).SubItems(7))
                 !MBMCAllergic = ChkStr(lstCovAdd.listitems(listloop).SubItems(8))
                 !MBMCExclusion = ChkStr(lstCovAdd.listitems(listloop).SubItems(9))
                 If MBMtype = "C" Then
                    !MBMCMemberType = "C"
                 Else
                    !MBMCMemberType = "S"
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
                            !MBMCAnnualLimit = AnnLmt!SuppLimit
                            !MBMCAvailableLimit = AnnLmt!SuppLimit
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
                                 
                                 tempSPremium = Format(GetSPremium(Mid(cboINSCode, 1, 1), Mid(txtPAYCode, 1, 2), AgeValue, Trim(Mid(cboPLNCode, 1, 4)), Trim(txtHLTCode), CDate(txtMBMPayorEffDate)), "#,##0.00")
                                 tempSPremium = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(tempSPremium, "#,##0.00"))
                                 
                                 If Disc = True Then
                                     If lstCovAdd.listitems(listloop).SubItems(36) <> "" Then
                                         tempSPremium = tempSPremium * (100 - CDbl(lstCovAdd.listitems(listloop).SubItems(36))) / 100
                                     End If
                                     If lstCovAdd.listitems(listloop).SubItems(37) <> "" Then
                                         tempSPremium = tempSPremium * (100 - CDbl(lstCovAdd.listitems(listloop).SubItems(37))) / 100
                                     End If
                                 End If
                                 !MBMCPremium = tempSPremium
                            End If
                        ElseIf Prem!SuppPrmStatus = "C" Then
                            tempSPremium = Prem!SuppPrmAmt
                            If tempSPremium <> "" Then
                                !MBMCBasicPremium = tempSPremium
                            Else
                                !MBMCBasicPremium = 0
                            End If
                            tempSPremium = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(tempSPremium, "#,##0.00"))
                            
                            If Disc = True Then
                                If lstCovAdd.listitems(listloop).SubItems(36) <> "" Then
                                    tempSPremium = tempSPremium * (100 - CDbl(lstCovAdd.listitems(listloop).SubItems(36))) / 100
                                End If
                                If lstCovAdd.listitems(listloop).SubItems(37) <> "" Then
                                    tempSPremium = tempSPremium * (100 - CDbl(lstCovAdd.listitems(listloop).SubItems(37))) / 100
                                End If
                            End If
                                
                            !MBMCPremium = tempSPremium
                            End If
                        End If
                AnnLmt.Close
                Prem.Close
                Set AnnLmt = Nothing
                Set Prem = Nothing
                'End If
                
                
                If Len(lstCovAdd.listitems(listloop).SubItems(12)) Then
                    !MBMCOccupation = Mid(ChkStr(lstCovAdd.listitems(listloop).SubItems(12)), 1, 2)
                End If
                If Len(lstCovAdd.listitems(listloop).SubItems(13)) Then
                    !MBMCWorkNature = Mid(ChkStr(lstCovAdd.listitems(listloop).SubItems(13)), 1, 3)
                End If
                If Len(lstCovAdd.listitems(listloop).SubItems(14)) Then
                    !MBMCSmoking = Mid(ChkStr(lstCovAdd.listitems(listloop).SubItems(14)), 1, 1)
                End If
                If Len(lstCovAdd.listitems(listloop).SubItems(15)) Then
                    !MBMCAlcohol = Mid(ChkStr(lstCovAdd.listitems(listloop).SubItems(15)), 1, 1)
                End If
                'Cater for different benefit
                If Len(lstCovAdd.listitems(listloop).SubItems(18)) Then
                    !MBMCPrevPLNCode = ChkStr(lstCovAdd.listitems(listloop).SubItems(18))
                End If
                If Len(lstCovAdd.listitems(listloop).SubItems(19)) Then
                    !MBMCRBAmt = ChkStr(lstCovAdd.listitems(listloop).SubItems(19))
                End If
                If Len(lstCovAdd.listitems(listloop).SubItems(20)) Then
                    !MBMCPAYRemarks = ChkStr(lstCovAdd.listitems(listloop).SubItems(20))
                End If
                If Len(lstCovAdd.listitems(listloop).SubItems(21)) Then
                    !MBMCPrevInsCom = ChkStr(lstCovAdd.listitems(listloop).SubItems(21))
                End If
                
                If Len(lstCovAdd.listitems(listloop).SubItems(22)) Then
                    !MBMCGenWP = ChkStr(lstCovAdd.listitems(listloop).SubItems(22))
                End If
                If lstCovAdd.listitems(listloop).SubItems(23) <> "__/__/____" And lstCovAdd.listitems(listloop).SubItems(23) <> "" Then
                    !MBMCGenEndWP = ChkStr(lstCovAdd.listitems(listloop).SubItems(23))
                End If
                If Len(lstCovAdd.listitems(listloop).SubItems(24)) Then
                    !MBMCPreExWP = ChkStr(lstCovAdd.listitems(listloop).SubItems(24))
                End If
                If lstCovAdd.listitems(listloop).SubItems(25) <> "__/__/____" And lstCovAdd.listitems(listloop).SubItems(25) <> "" Then
                    !MBMCPreExEndWP = ChkStr(lstCovAdd.listitems(listloop).SubItems(25))
                End If
                If Len(lstCovAdd.listitems(listloop).SubItems(26)) Then
                    !MBMCSpeIllWP = ChkStr(lstCovAdd.listitems(listloop).SubItems(26))
                End If
                If lstCovAdd.listitems(listloop).SubItems(27) <> "__/__/____" And lstCovAdd.listitems(listloop).SubItems(27) <> "" Then
                    !MBMSpeIllEndWP = ChkStr(lstCovAdd.listitems(listloop).SubItems(27))
                End If
                If Len(lstCovAdd.listitems(listloop).SubItems(28)) Then
                    !MBMCStaffDec = Mid(ChkStr(lstCovAdd.listitems(listloop).SubItems(28)), 1, 1)
                End If
                If lstCovAdd.listitems(listloop).SubItems(29) <> "__/__/____" And lstCovAdd.listitems(listloop).SubItems(29) <> "" Then
                    !MBMCCLMNonPayFr = ChkStr(lstCovAdd.listitems(listloop).SubItems(29))
                End If
                If lstCovAdd.listitems(listloop).SubItems(30) <> "__/__/____" And lstCovAdd.listitems(listloop).SubItems(30) <> "" Then
                    !MBMCCLMNonPayTo = ChkStr(lstCovAdd.listitems(listloop).SubItems(30))
                End If
                
                If listloop < "10" Then
                    !MBMCCoverID = 0 & listloop
                Else
                    !MBMCCoverID = listloop
                End If
                
                
                !MBMCBatchNo = ChkStr(txtMBMBatchNo)
                !MBMCBordxDate = ChkStr(txtMBMBordxDate)
                
                If Len(lstCovAdd.listitems(listloop).SubItems(32)) Then
                    !MBMCPayorPremGross = ChkStr(lstCovAdd.listitems(listloop).SubItems(32))
                End If
                
                If Len(lstCovAdd.listitems(listloop).SubItems(33)) Then
                    !MBMCPayorPremNet = ChkStr(lstCovAdd.listitems(listloop).SubItems(33))
                End If
                
                If Len(lstCovAdd.listitems(listloop).SubItems(34)) Then
                    !MBMCPayorMCOGross = ChkStr(lstCovAdd.listitems(listloop).SubItems(34))
                End If
                
                If Len(lstCovAdd.listitems(listloop).SubItems(35)) Then
                    !MBMCPayorMCONet = ChkStr(lstCovAdd.listitems(listloop).SubItems(35))
                End If
                
                If Len(lstCovAdd.listitems(listloop).SubItems(36)) Then
                    !MBMCPolicyDisc = ChkStr(lstCovAdd.listitems(listloop).SubItems(36))
                End If
                
                If Len(lstCovAdd.listitems(listloop).SubItems(37)) Then
                    !MBMCRenewalDisc = ChkStr(lstCovAdd.listitems(listloop).SubItems(37))
                End If
                
                .Update
            End With
         Next listloop
         MBMCoveredPersons.Close
         Set MBMCoveredPersons = Nothing
    End If
End Sub

Private Sub saveCLNCoverPersons(MembershipNo As String, MBMtype As String)
Dim MBMCLNCoveredPersons As New ADODB.Recordset
Dim listrecord As Integer
Dim listloop As Integer

listrecord = lstCovAdd.listitems.count

    ' Insert into Membership Clinic Covered Person's table
    
     If listrecord <> 0 Then
         MBMCLNCoveredPersons.Open "MBMCLNCovPerson", medixmasterconn, adOpenKeyset, adLockOptimistic
         For listloop = 1 To listrecord
             MBMCLNCoveredPersons.AddNew
             With MBMCLNCoveredPersons
                !MBMCNumber = ChkStr(MembershipNo)
                !MBMCExclusion = ChkStr(lstCovAdd.listitems(listloop).SubItems(31))
                 
                If MBMtype = "C" Then
                    !MBMCMemberType = "C"
                Else
                    !MBMCMemberType = "S"
                End If
                
                If listloop < "10" Then
                    !MBMCCoverID = 0 & listloop
                Else
                    !MBMCCoverID = listloop
                End If
                .Update
            End With
         Next listloop
         MBMCLNCoveredPersons.Close
         Set MBMCLNCoveredPersons = Nothing
    End If
End Sub

Private Sub savembmothers(MembershipNo As String)
     Dim MBMOthers As New ADODB.Recordset
    ' INsert into Membership Other details table
     'MBMOthers.Open "select MBMNumber , * from MBMOthers", medixmasterconn, adOpenKeyset, adLockOptimistic
     MBMOthers.Open "MBMOthers", medixmasterconn, adOpenKeyset, adLockOptimistic
     MBMOthers.AddNew
     With MBMOthers
             !MBMNumber = ChkStr(MembershipNo)
             !MBMWeight = ChkStr(FrmMBMOthers.txtMBMWeight)
             !MBMHeight = ChkStr(FrmMBMOthers.txtMBMHeight)
             !MBMOccupation = Trim(Mid(FrmMBMOthers.cboMBMOccupation, 1, 2))
             !MBMWorkNature = Mid(Trim(FrmMBMOthers.txtMBMWorkNature), 1, 3)
             !MBMExclusion = ChkStr(txtMBMExclusion)
             !MBMAllergic = ChkStr(FrmMBMOthers.txtMBMAllergic)
             !MBMGroupCompany = ChkStr(txtMBMGroupCompany)
             !MBMEmployeeNo = ChkStr(txtMBMEmployeeNo)
             !MBMBlood = ChkStr(FrmMBMOthers.cboMBMBlood)
             !MBMSmoking = FrmMBMOthers.cboMBMSmoking
             !MBMAlcohol = FrmMBMOthers.cboMBMAlcohol
             !MBMAgentCode = ChkStr(txtMBMAgentCode)
             '!MBMPolSubNo1 = ChkStr(FrmMBMOthers.TxtPOLSubNo)
             
             '====== Clinical ================================
             'If Trim(Mid(cboHNC, 1, 1)) = "H" Or Trim(Mid(cboHNC, 1, 1)) = "M" Then
                !MBMAccessFee = IIf(IsNumeric(txtMBMProviderCharges) = True, txtMBMProviderCharges, 0)
                !MBMPremium = Trim(txtMBMPremium)
                !MBMMCOFee = Trim(txtMCOFees)
                !MBMBasicMCO = txtTempMCOFees
                !MBMBasicPrem = txtTempPrem
                !MBMBasicEA = txtTempProvider
             'End If
             '====== Clinical ================================
             
             !MBMDepartment = ChkStr(txtMBMDepartment)
             !MBMInstallment = Trim(Mid(CboInsMode, 1, 1))
             !MBMPrevPLNCode = FrmMBMOthers.txtMBMPrevPlan
             If Len(FrmMBMOthers.txtMBMRBAmt) Then
                !MBMRBAmt = FrmMBMOthers.txtMBMRBAmt
             End If
             If Len(FrmMBMOthers.txtMBMPAYRemarks) Then
                !MBMPAYRemarks = ChkStr(FrmMBMOthers.txtMBMPAYRemarks)
             End If
             If Len(FrmMBMOthers.txtMBMPrevINSCom) Then
                !MBMPrevInsCom = ChkStr(FrmMBMOthers.txtMBMPrevINSCom)
             End If

                !RELCODE = "P"
             
             If Len(Trim(cboMBMMaritalStatus)) Then
                !MBMMaritalStatus = Mid(cboMBMMaritalStatus, 1, 1)
             End If
             
             If ChkStr(cboMBMRenewal) = "R" Then
                !MBMFirstMEMNo = ChkStr(txtFirstMeMNo)
                !MBMFirstPolNo = ChkStr(txtFirstPolNo)
                !MBMPrevPolNo = ChkStr(txtMBMPrevPolNo)
                !MBMPrevMemNo = ChkStr(txtMBMPrevMEMNo)
                !MBMFirstJoinedDate = txtMBMDateJoin
             Else
                !MBMFirstMEMNo = ChkStr(MembershipNo)
                !MBMFirstPolNo = ChkStr(txtMBMPolicyNo)
                !MBMFirstJoinedDate = txtMBMDateJoin
             End If
             If Len(txtRemarks) Then
                !MBMRemarks = ChkStr(txtRemarks)
             End If
             If Len(FrmMBMOthers.cboCOPayRB) Then
                !MBMCoPayRB = FrmMBMOthers.cboCOPayRB
             End If
             '=================== Add By CK ================
                If Len(FrmMBMOthers.TxtGenWP) Then
                   !MBMGenWP = FrmMBMOthers.TxtGenWP
                End If
                If FrmMBMOthers.txtGENWPEnd <> "__/__/____" Then
                   !MBMGenEndWP = FrmMBMOthers.txtGENWPEnd
                End If
                If Len(FrmMBMOthers.txtPREWP) Then
                   !MBMPreExWP = FrmMBMOthers.txtPREWP
                End If
                If FrmMBMOthers.txtPREWPEnd <> "__/__/____" Then
                   !MBMPreExEndWP = FrmMBMOthers.txtPREWPEnd
                End If
                If Len(FrmMBMOthers.TxtSpeWP) Then
                   !MBMSpeIllWP = Trim(FrmMBMOthers.TxtSpeWP)
                End If
                If FrmMBMOthers.txtSPEWPEnd <> "__/__/____" Then
                   !MBMSpeIllEndWP = FrmMBMOthers.txtSPEWPEnd
                End If
                If FrmMBMOthers.TxtClmNonPayFr <> "__/__/____" Then
                   !MBMCLMNonPayFr = FrmMBMOthers.TxtClmNonPayFr
                End If
                If FrmMBMOthers.TxtClmNonPayTo <> "__/__/____" Then
                   !MBMCLMNonPayTo = FrmMBMOthers.TxtClmNonPayTo
                End If
                If Len(FrmMBMOthers.cboStaffDisc) Then
                   !MBMStaffDec = Mid(FrmMBMOthers.cboStaffDisc, 1, 1)
                End If
             
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
             '=================== Add By CK ================
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
    MBMOthers.Close
    Set MBMOthers = Nothing
End Sub

Public Sub SaveBatch()
    Dim BAT As New ADODB.Recordset
    Dim strsql As String
    
    strsql = "Select * from BAT Where PAYCode ='" & txtPAYCode & "' " & _
         " And BATBordxDate = '" & Format(txtMBMBordxDate, "MM/DD/YYYY") & "' " & _
         " And BordxType = 'N'"
    BAT.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                                                                                    
    If BAT.recordCount = 0 Then
        BAT.AddNew
    End If
    With BAT
        !BordxType = "N"
        !PAYCode = txtPAYCode
        !BATBordxDate = txtMBMBordxDate
        !BATSequence = txtMBMBatchNo
    End With
    BAT.Update
  

    BAT.Close
    Set BAT = Nothing
End Sub
    
Private Sub UpdateBordxPriting()
Dim BordxPrinting As New ADODB.Recordset

    BordxPrinting.Open "BordxPrinting", medixmasterconn, adOpenKeyset, adLockOptimistic
    BordxPrinting.AddNew
    With BordxPrinting
        !MBMBordxDate = txtMBMBordxDate
        !PAYCode = txtPAYCode
        !MBMBatchNo = txtMBMBatchNo
        !MBMBordxType = "N"
    End With
    BordxPrinting.Update
    BordxPrinting.Close
    Set BordxPrinting = Nothing
End Sub

'  ################# Start Suppplementary ######################


Private Sub cmdAdd_Click()
On Error Resume Next
Dim MemberConveredPerson As ListItem
Dim CountList As Integer
'Dim RecordSaved

    If txtCovName = "" Then
        MsgBox "Please Enter Covered Person Name", vbOKOnly + vbCritical, DialogBoxTitle
        txtCovName.SetFocus
    ElseIf txtCovIcBcPp = "" Then
        MsgBox "Please Enter Covered Person Identification", vbOKOnly + vbCritical, DialogBoxTitle
        txtCovIcBcPp.SetFocus
    ElseIf txtCovBirthdate = "__/__/____" Then
        MsgBox "Please Enter Birthdate", vbOKOnly + vbCritical, DialogBoxTitle
        txtCovBirthdate.SetFocus
    ElseIf cboCovRelationship = "" Then
        MsgBox "Please Enter Covered Person Relationship", vbOKOnly + vbCritical, DialogBoxTitle
        cboCovRelationship.SetFocus
    Else
        CountList = lstCovAdd.listitems.count + 1
        Set MemberConveredPerson = lstCovAdd.listitems.Add(, , txtCovName)
        With MemberConveredPerson
            .SubItems(1) = cboCovSalutation
            .SubItems(2) = txtCovIcBcPp
            .SubItems(3) = txtCovBirthdate
            .SubItems(4) = cboCovSex
            .SubItems(5) = cboCovAdultChild
            .SubItems(6) = cboCovRelationship
            .SubItems(8) = txtCovAllergic
            .SubItems(9) = txtCovExclusion
            .SubItems(10) = CountList
            .SubItems(11) = IIf(Trim(txtCovAnnualLimit) = "", 0, txtCovAnnualLimit)
            .SubItems(18) = txtPrevPLNCode
            .SubItems(19) = txtRBAmt
            .SubItems(20) = txtCOVPayRemarks
            .SubItems(21) = txtPrevINSCom
            .SubItems(22) = txtSuppGenWP
            .SubItems(23) = txtSuppGENWPEnd
            .SubItems(24) = txtSuppPREWP
            .SubItems(25) = txtSuppPREWPEnd
            .SubItems(26) = TxtSuppSpeWP
            .SubItems(27) = txtSuppSPEWPEnd
            .SubItems(28) = cboSuppStaffDisc
            .SubItems(29) = TxtSuppClmNonPayFr
            .SubItems(30) = TxtSuppClmNonPayTo
            .SubItems(31) = txtCLNCovExclusion
            .SubItems(32) = txtSuppPayPremGross
            .SubItems(33) = txtSuppPayPremNet
            .SubItems(34) = txtSuppPayMCOGross
            .SubItems(35) = txtSuppPayMCONet
            .SubItems(36) = txtRenewalDisc
            .SubItems(37) = txtPolicyDisc
            End With
            Call ClearCoveredEntry
            If cboCovSalutation.Enabled = True Then
                cboCovSalutation.SetFocus
            End If
            cboCovRelationship = "S - SON"
    End If

End Sub

Private Sub cmdCovClear_Click()
    Call EnableCovered(True)
    Call ClearCoveredEntry
    cboCovSalutation.SetFocus
    cmdRemove.Enabled = False
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = False
End Sub

Private Sub cmdRemove_Click()
On Error Resume Next
    Dim ForLoop As Integer
    Dim ListCount As Integer
    Dim RecordSaved
    
      RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
        If RecordSaved = vbYes Then
            ListCount = lstCovAdd.listitems.count
            For ForLoop = 1 To ListCount
                If lstCovAdd.listitems(ForLoop).Selected = True Then
                    lstCovAdd.listitems.Remove (ForLoop)
                    Exit For
                End If
            Next ForLoop
            Call EnableCovered(True)
            Call ClearCoveredEntry
            cboCovSalutation.SetFocus
        End If
    
End Sub


Private Sub cmdUpdate_Click()
On Error Resume Next
Dim ForLoop As Integer
            For ForLoop = 1 To lstCovAdd.listitems.count
                If lstCovAdd.listitems(ForLoop - 1).Selected = True Then
                    If lstCovAdd.SelectedItem.SubItems(10) = txtcount Then
                        lstCovAdd.SelectedItem.Text = txtCovName
                        lstCovAdd.SelectedItem.SubItems(1) = cboCovSalutation
                        lstCovAdd.SelectedItem.SubItems(2) = txtCovIcBcPp
                        lstCovAdd.SelectedItem.SubItems(3) = txtCovBirthdate
                        lstCovAdd.SelectedItem.SubItems(4) = cboCovSex
                        lstCovAdd.SelectedItem.SubItems(5) = cboCovAdultChild
                        lstCovAdd.SelectedItem.SubItems(6) = cboCovRelationship
                        lstCovAdd.SelectedItem.SubItems(8) = txtCovAllergic
                        lstCovAdd.SelectedItem.SubItems(9) = txtCovExclusion
                        lstCovAdd.SelectedItem.SubItems(11) = txtCovAnnualLimit
                        lstCovAdd.SelectedItem.SubItems(18) = txtPrevPLNCode
                        lstCovAdd.SelectedItem.SubItems(19) = txtRBAmt
                        lstCovAdd.SelectedItem.SubItems(20) = txtCOVPayRemarks
                        lstCovAdd.SelectedItem.SubItems(21) = txtPrevINSCom
                        lstCovAdd.SelectedItem.SubItems(22) = txtSuppGenWP
                        lstCovAdd.SelectedItem.SubItems(23) = txtSuppGENWPEnd
                        lstCovAdd.SelectedItem.SubItems(24) = txtSuppPREWP
                        lstCovAdd.SelectedItem.SubItems(25) = txtSuppPREWPEnd
                        lstCovAdd.SelectedItem.SubItems(26) = TxtSuppSpeWP
                        lstCovAdd.SelectedItem.SubItems(27) = txtSuppSPEWPEnd
                        lstCovAdd.SelectedItem.SubItems(28) = cboSuppStaffDisc
                        lstCovAdd.SelectedItem.SubItems(29) = TxtSuppClmNonPayFr
                        lstCovAdd.SelectedItem.SubItems(30) = TxtSuppClmNonPayTo
                        lstCovAdd.SelectedItem.SubItems(31) = txtCLNCovExclusion
                        lstCovAdd.SelectedItem.SubItems(32) = txtSuppPayPremGross
                        lstCovAdd.SelectedItem.SubItems(33) = txtSuppPayPremNet
                        lstCovAdd.SelectedItem.SubItems(34) = txtSuppPayMCOGross
                        lstCovAdd.SelectedItem.SubItems(35) = txtSuppPayMCONet
                        lstCovAdd.SelectedItem.SubItems(36) = txtRenewalDisc
                        lstCovAdd.SelectedItem.SubItems(37) = txtPolicyDisc
                        
                        Call ClearCoveredEntry
                        Exit For
                    End If
                End If
            Next ForLoop
            
            cboCovSalutation.SetFocus
            cmdRemove.Enabled = True
            cmdAdd.Enabled = True
            cmdUpdate.Enabled = False
        'End If
End Sub

Private Sub lstCovAdd_Click()
On Error Resume Next
    If lstCovAdd.listitems.count Then
        cboCovSalutation = lstCovAdd.SelectedItem.SubItems(1)
        txtCovName = lstCovAdd.SelectedItem.Text
        txtCovIcBcPp = lstCovAdd.SelectedItem.SubItems(2)
        txtCovBirthdate = lstCovAdd.SelectedItem.SubItems(3)
        cboCovSex = lstCovAdd.SelectedItem.SubItems(4)
        cboCovAdultChild = lstCovAdd.SelectedItem.SubItems(5)
        cboCovRelationship = lstCovAdd.SelectedItem.SubItems(6)
        txtCovAllergic = lstCovAdd.SelectedItem.SubItems(8)
        txtCovExclusion = lstCovAdd.SelectedItem.SubItems(9)
        txtcount = lstCovAdd.SelectedItem.SubItems(10)
        txtCovAnnualLimit = Format(lstCovAdd.SelectedItem.SubItems(11), "###,##0.00")
        txtPrevPLNCode = lstCovAdd.SelectedItem.SubItems(18)
        txtRBAmt = lstCovAdd.SelectedItem.SubItems(19)
        txtCOVPayRemarks = lstCovAdd.SelectedItem.SubItems(20)
        txtPrevINSCom = lstCovAdd.SelectedItem.SubItems(21)
        txtSuppGenWP = lstCovAdd.SelectedItem.SubItems(22)
        txtSuppGENWPEnd = lstCovAdd.SelectedItem.SubItems(23)
        txtSuppPREWP = lstCovAdd.SelectedItem.SubItems(24)
        txtSuppPREWPEnd = lstCovAdd.SelectedItem.SubItems(25)
        TxtSuppSpeWP = lstCovAdd.SelectedItem.SubItems(26)
        txtSuppSPEWPEnd = lstCovAdd.SelectedItem.SubItems(27)
        cboSuppStaffDisc = lstCovAdd.SelectedItem.SubItems(28)
        TxtSuppClmNonPayFr.mask = ""
        TxtSuppClmNonPayFr = lstCovAdd.SelectedItem.SubItems(29)
        TxtSuppClmNonPayTo.mask = ""
        TxtSuppClmNonPayTo = lstCovAdd.SelectedItem.SubItems(30)
        txtCLNCovExclusion = lstCovAdd.SelectedItem.SubItems(31)
        txtSuppPayPremGross = lstCovAdd.SelectedItem.SubItems(32)
        txtSuppPayPremNet = lstCovAdd.SelectedItem.SubItems(33)
        txtSuppPayMCOGross = lstCovAdd.SelectedItem.SubItems(34)
        txtSuppPayMCONet = lstCovAdd.SelectedItem.SubItems(35)
        txtRenewalDisc = lstCovAdd.SelectedItem.SubItems(36)
        txtPolicyDisc = lstCovAdd.SelectedItem.SubItems(37)
        
        Call EnableCovered(False)
        cmdUpdate.Enabled = True
        cmdRemove.Enabled = True
        cmdAdd.Enabled = False
    End If
    Call EnableCovered(True)
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


Public Sub EnableCovered(status As Boolean)
    txtCovAnnualLimit.Enabled = status
    cboCovSalutation.Enabled = status
    txtCovName.Enabled = status
    txtCovIcBcPp.Enabled = status
    txtCovBirthdate.Enabled = status
    cboCovAdultChild.Enabled = status
    cboCovRelationship.Enabled = status
    txtCovAllergic.Enabled = status
    txtCovExclusion.Enabled = status
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
    txtcount = ""
    txtPrevPLNCode = ""
    txtRBAmt = ""
    txtPrevINSCom = ""
    txtCOVPayRemarks = ""
    txtCovAnnualLimit = ""
    If Mid(cboMBMRelCode, 1, 1) = "S" Then
        txtCovAnnualLimit = tempSAnnLmt
    End If
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

Private Sub DeleteCoveredID(MembershipNo As String)
Dim DELCOV As New ADODB.Recordset
Dim delstr As String

    delstr = "Delete from MBM Where MBMMemberType = 'C' And MBMNumber = '" & MembershipNo & "'"
    DELCOV.Open delstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    DELCOV.Close
End Sub

Private Sub EnableSupplementary(status As Boolean)
    lstCovAdd.Enabled = status
    cboCovSalutation.Enabled = status
    txtCovName.Enabled = status
    txtCovIcBcPp.Enabled = status
    txtCovBirthdate.Enabled = status
    cboCovAdultChild.Enabled = status
    cboCovRelationship.Enabled = status
    txtCovAllergic.Enabled = status
    txtCovExclusion.Enabled = status
    txtCovAnnualLimit.Enabled = status
    cmdRemove.Enabled = status
    cmdAdd.Enabled = status
End Sub

'  ################# END Suppplementary ######################


' ################ START Key press effect ###################
Private Sub cboCovAdultChild_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    Clipboard.Clear
    
    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboCovAdultChild.Text, cboCovAdultChild.SelStart) + Chr(KeyAscii) + Mid(cboCovAdultChild.Text, cboCovAdultChild.SelStart + cboCovAdultChild.SelLength + 1))
        
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


Private Sub cboCovRelationship_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    Clipboard.Clear
    
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
    Clipboard.Clear
   
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
    Clipboard.Clear

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
Clipboard.Clear
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
    Clipboard.Clear
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

Private Sub cboMBMRenewal_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    Clipboard.Clear
    
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
            
    If KeyAscii = 13 Then
        If cboMBMRenewal = "N" Then
            txtMBMPolicyNo.SetFocus
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
    Clipboard.Clear
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


Private Sub cboNATCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    Clipboard.Clear
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

Private Sub cboMBmtakeover_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    Clipboard.Clear
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboMBMTakeover.Text, cboMBMTakeover.SelStart) + Chr(KeyAscii) + Mid(cboMBMTakeover.Text, cboMBMTakeover.SelStart + cboMBMTakeover.SelLength + 1))

        ValidCount = 0
        ValidValue = ""

        For i = 0 To cboMBMTakeover.ListCount - 1

            If NewText = LCase(Left(cboMBMTakeover.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboMBMTakeover.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0

                 If ValidCount = 1 Then
                   cboMBMTakeover.Text = ValidValue
                   cboMBMTakeover.SelStart = 0
                   cboMBMTakeover.SelLength = Len(ValidValue)
                 End If
            End If
End Sub


Private Sub cboMBMRelCode_KeyPress(KeyAscii As Integer)
On Error Resume Next
    Dim NewText As String
    Dim ValidCount As Integer
    Dim ValidValue As String
    Dim i As Integer
    Clipboard.Clear
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
Private Sub txtMBMRBAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

' Data Validation
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

Private Sub InsertMCO()
Dim MBMInsertMCOFees As Double
Dim MBMMCOFess As Double
Dim PAYExpDate As String

    If Len(txtMCOFees) Then
        
        MBMInsertMCOFees = InsertMCOFees(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(txtMCOFees, "#,##0.00"))
    txtMCOFees = ""
    txtMCOFees = Round(MBMInsertMCOFees, 0)
    End If
End Sub

Private Sub InsertPrem()
Dim MBMInsertPrem As Double
Dim MBMPrem As Double
Dim PAYExpDate As String

    If Trim(cboPLNCode) <> "" Then
        If Len(txtMBMPremium) And IsDate(txtMBMPayorEffDate) Then
            MBMInsertPrem = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(txtMBMPremium, "#,##0.00"))
            txtMBMPremium = ""
            txtMBMPremium = Round(MBMInsertPrem, 0)
            
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
    End If
End Sub

Public Sub InsertProviderFees()
Dim MBMInsertProviderFees As Double
    If Len(txtMBMProviderCharges) Then
        MBMInsertProviderFees = InsertProvider(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(txtMBMProviderCharges, "#,##0.00"))
        txtMBMProviderCharges = ""
        txtMBMProviderCharges = Round(MBMInsertProviderFees, 1)
    End If
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
    End If
End Sub

Private Sub SaveMBMCLNOthers(MembershipNo As String)
Dim MBMCLNOthers As New ADODB.Recordset

     MBMCLNOthers.Open "select MBMNumber , * from MBMCLNOthers", medixmasterconn, adOpenKeyset, adLockOptimistic
     MBMCLNOthers.AddNew
     With MBMCLNOthers
             !MBMNumber = ChkStr(MembershipNo)
             !MBMPolNumber = ChkStr(txtMBMPolicyNo)
             !PLNType = Mid(frmMBMClinical.cboPLNCode, 1, IIf(InStr(1, frmMBMClinical.cboPLNCode, "-") > 0, InStr(1, frmMBMClinical.cboPLNCode, "-") - 1, 4))
             !InsType = Mid(frmMBMClinical.cboINSCode, 1, 1)
             !MBMPolEffDate = Format(frmMBMClinical.txtMBMPayorEffDate, "dd/mm/yyyy")
             !MBMPolExpDate = Format(frmMBMClinical.txtMBMPayorExpDate, "dd/mm/yyyy")
             !MBMRemarks = Trim(frmMBMClinical.txtRemarks)
             !MBMCLNPremium = CDbl(frmMBMClinical.lblMBMPremium)
             !MBMCLNBasicPrem = CDbl(frmMBMClinical.txtTempPrem)
             !MBMCLNMCOFee = CDbl(frmMBMClinical.lblMCOFees)
             !MBMCLNBasicMCO = CDbl(frmMBMClinical.txtTempMCOFees)
             
             If Mid(ChkStr(frmMBMClinical.cboMBMRelCode), 1, 1) = "C" Then
                !MBMCLNMemberType = "P"
             Else
                !MBMCLNMemberType = "Q"
             End If
             
             !MBMCLNExpiryStatus = "N"
             !MBMCLNDataStatus = "A"
             !MBMCLNStatus = "A"
     End With
     MBMCLNOthers.Update
   
    MBMCLNOthers.Close
    Set MBMCLNOthers = Nothing

End Sub

Private Sub cboHNC_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    
    Clipboard.Clear
        If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboHNC.Text, cboHNC.SelStart) + Chr(KeyAscii) + Mid(cboHNC.Text, cboHNC.SelStart + cboHNC.SelLength + 1))

        ValidCount = 0
        ValidValue = ""

        For i = 0 To cboHNC.ListCount - 1

            If NewText = LCase(Left(cboHNC.List(i), Len(NewText))) Then
                    ValidCount = ValidCount + 1
                    ValidValue = cboHNC.List(i)
            End If
            Next
            If ValidCount <= 1 Then KeyAscii = 0

                 If ValidCount = 1 Then
                   cboHNC.Text = ValidValue
                   cboHNC.SelStart = 0
                   cboHNC.SelLength = Len(ValidValue)
                 End If
            End If
End Sub

Public Function showMBMOthers() As Integer
Dim sqlstr As String
Dim MBM As New ADODB.Recordset
Dim strCount As String
Dim rstCount As New ADODB.Recordset
Dim MBMNumber As String
    
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
    
        sqlstr = sqlstr & " From View_MBM_MBMOthers Where MBMNumber ='" & Trim(txtMBMPrevMEMNo) & "'"
        MBM.Open sqlstr, medixmasterconn, adOpenKeyset, adLockOptimistic
    
        'strCount = "select count(*) as totalRecord From View_MBM_MBMOthers  Where MBMNumber = '" & Trim(frmMBMAdjustment.txtMBMNumber) & "'"
        'strCount = strCount + sqlstr
        
        MBMNumber = txtMBMPrevMEMNo
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

