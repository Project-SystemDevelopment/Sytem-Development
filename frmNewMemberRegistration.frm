VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewMemberRegistration 
   Caption         =   "Member Registration"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5565
      Left            =   120
      TabIndex        =   74
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
      TabPicture(0)   =   "frmNewMemberRegistration.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Supplementary"
      TabPicture(1)   =   "frmNewMemberRegistration.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstCovAdd"
      Tab(1).Control(1)=   "Frame6"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame6 
         Height          =   3195
         Left            =   -74880
         TabIndex        =   114
         Top             =   480
         Width           =   11460
         Begin VB.TextBox txtCovAllergic 
            Height          =   795
            Left            =   4680
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   68
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
            TabIndex        =   70
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox txtCovExclusion 
            Height          =   795
            Left            =   6360
            MaxLength       =   2500
            MultiLine       =   -1  'True
            TabIndex        =   69
            Top             =   2160
            Width           =   3135
         End
         Begin VB.ComboBox cboCovSalutation 
            Height          =   315
            Left            =   1320
            TabIndex        =   35
            Top             =   195
            Width           =   975
         End
         Begin VB.TextBox txtCovName 
            Height          =   330
            Left            =   3480
            MaxLength       =   60
            TabIndex        =   36
            Top             =   180
            Width           =   3660
         End
         Begin VB.Frame Frame9 
            Height          =   1695
            Left            =   10320
            TabIndex        =   115
            Top             =   120
            Width           =   1065
            Begin VB.CommandButton cmdRemove 
               Caption         =   "&Remove"
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   118
               Top             =   600
               Width           =   840
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "&Add"
               Height          =   315
               Left            =   120
               TabIndex        =   71
               Top             =   240
               Width           =   840
            End
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "Update"
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   117
               Top             =   960
               Width           =   840
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Clear"
               Height          =   315
               Left            =   120
               TabIndex        =   116
               Top             =   1320
               Width           =   840
            End
         End
         Begin VB.TextBox txtCovIcBcPp 
            Height          =   285
            Left            =   7800
            MaxLength       =   30
            TabIndex        =   37
            Top             =   180
            Width           =   2055
         End
         Begin VB.ComboBox cboCovRelationship 
            Height          =   315
            Left            =   1320
            TabIndex        =   38
            Text            =   "W -Wife"
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox cboCovAdultChild 
            Height          =   315
            Left            =   1320
            TabIndex        =   42
            Text            =   "A"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtPrevPLNCode 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5760
            TabIndex        =   40
            Top             =   480
            Width           =   1380
         End
         Begin VB.TextBox txtAvailableLimit 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   45
            Top             =   1010
            Width           =   975
         End
         Begin VB.TextBox txtPrevINSCom 
            Height          =   330
            Left            =   7800
            MaxLength       =   60
            TabIndex        =   41
            Top             =   440
            Width           =   2055
         End
         Begin VB.TextBox txtRBAmt 
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   49
            Top             =   1290
            Width           =   975
         End
         Begin VB.TextBox txtSuppPayPremGross 
            Height          =   285
            Left            =   1320
            TabIndex        =   53
            Top             =   1545
            Width           =   975
         End
         Begin VB.TextBox txtCOVPayRemarks 
            Height          =   285
            Left            =   5760
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   44
            Top             =   720
            Width           =   4095
         End
         Begin VB.TextBox txtSuppGenWP 
            Height          =   285
            Left            =   5760
            MaxLength       =   9
            TabIndex        =   47
            Top             =   960
            Width           =   945
         End
         Begin VB.TextBox txtSuppPREWP 
            Height          =   285
            Left            =   5760
            MaxLength       =   9
            TabIndex        =   51
            Top             =   1200
            Width           =   945
         End
         Begin VB.TextBox TxtSuppSpeWP 
            Height          =   300
            Left            =   5760
            MaxLength       =   9
            TabIndex        =   55
            Top             =   1455
            Width           =   945
         End
         Begin VB.TextBox txtSuppPayMCOGross 
            Height          =   285
            Left            =   1320
            TabIndex        =   59
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox txtRenewalDisc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   61
            Top             =   2040
            Width           =   975
         End
         Begin MSMask.MaskEdBox txtCovBirthdate 
            Height          =   285
            Left            =   3480
            TabIndex        =   39
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
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
            TabIndex        =   57
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
            TabIndex        =   58
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
            TabIndex        =   48
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
            TabIndex        =   52
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
            TabIndex        =   56
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
            TabIndex        =   43
            Text            =   "M"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtCovAnnualLimit 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   46
            Top             =   1005
            Width           =   1095
         End
         Begin VB.ComboBox cboSuppStaffDisc 
            Height          =   315
            Left            =   3480
            Sorted          =   -1  'True
            TabIndex        =   50
            Top             =   1260
            Width           =   1095
         End
         Begin VB.TextBox txtSuppPayPremNet 
            Height          =   285
            Left            =   3480
            TabIndex        =   54
            Top             =   1550
            Width           =   1095
         End
         Begin VB.TextBox txtSuppPayMCONet 
            Height          =   285
            Left            =   3480
            TabIndex        =   60
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox txtPolicyDisc 
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   62
            Top             =   2040
            Width           =   1095
         End
         Begin VB.TextBox txtPolicyDiscAmt 
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   63
            Top             =   2295
            Width           =   975
         End
         Begin VB.TextBox txtLoadingPrem 
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   65
            Top             =   2550
            Width           =   975
         End
         Begin VB.TextBox txtRenewalDiscAmt 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3480
            MaxLength       =   8
            TabIndex        =   64
            Top             =   2280
            Width           =   1095
         End
         Begin VB.TextBox txtSuppPLNCode 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3480
            MaxLength       =   4
            TabIndex        =   66
            Top             =   2520
            Width           =   1095
         End
         Begin VB.TextBox txtCovUndExcess 
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   67
            Top             =   2780
            Width           =   975
         End
         Begin VB.Label Label22 
            Caption         =   "Und Excess"
            Height          =   255
            Left            =   120
            TabIndex        =   183
            Top             =   2800
            Width           =   1365
         End
         Begin VB.Label Label62 
            Caption         =   "PLN Code"
            Height          =   255
            Left            =   2355
            TabIndex        =   180
            Top             =   2520
            Width           =   1485
         End
         Begin VB.Label Label61 
            Caption         =   "Ren. Disc. Amt"
            Height          =   255
            Left            =   2355
            TabIndex        =   178
            Top             =   2295
            Width           =   1485
         End
         Begin VB.Label Label60 
            Caption         =   "Pol. Disc. Amt"
            Height          =   255
            Left            =   120
            TabIndex        =   177
            Top             =   2325
            Width           =   1365
         End
         Begin VB.Label Label58 
            Caption         =   "Loading Prem"
            Height          =   255
            Left            =   120
            TabIndex        =   176
            Top             =   2550
            Width           =   1365
         End
         Begin VB.Label Label48 
            Caption         =   "Allergic"
            Height          =   255
            Left            =   4680
            TabIndex        =   150
            Top             =   1920
            Width           =   600
         End
         Begin VB.Label Label108 
            Caption         =   "Supp. Clinic Warranty"
            Height          =   255
            Left            =   9600
            TabIndex        =   149
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label32 
            Caption         =   "Policy Disc."
            Height          =   255
            Left            =   2355
            TabIndex        =   148
            Top             =   2040
            Width           =   1365
         End
         Begin VB.Label Label28 
            Caption         =   "Renewal Disc."
            Height          =   255
            Left            =   120
            TabIndex        =   147
            Top             =   2055
            Width           =   1485
         End
         Begin VB.Label Label23 
            Caption         =   "Pay MCO Net"
            Height          =   255
            Left            =   2355
            TabIndex        =   146
            Top             =   1815
            Width           =   1365
         End
         Begin VB.Label Label13 
            Caption         =   "Pay MCO Gross"
            Height          =   255
            Left            =   120
            TabIndex        =   145
            Top             =   1815
            Width           =   1365
         End
         Begin VB.Label Label31 
            Caption         =   "Pay Prem Net"
            Height          =   255
            Left            =   2355
            TabIndex        =   144
            Top             =   1560
            Width           =   1365
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            Caption         =   "To"
            Height          =   255
            Left            =   8400
            TabIndex        =   143
            Top             =   1440
            Width           =   285
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            Caption         =   "From"
            Height          =   255
            Left            =   8400
            TabIndex        =   142
            Top             =   1245
            Width           =   405
         End
         Begin VB.Label Label111 
            Caption         =   "Pay Prem Gross"
            Height          =   255
            Left            =   120
            TabIndex        =   141
            Top             =   1560
            Width           =   1365
         End
         Begin VB.Label Label37 
            Caption         =   "Clm Non-Pay Date"
            Height          =   375
            Left            =   8450
            TabIndex        =   140
            Top             =   1000
            Width           =   1365
         End
         Begin VB.Label Label53 
            Caption         =   "EOP"
            Height          =   255
            Left            =   6840
            TabIndex        =   139
            Top             =   1275
            Width           =   1005
         End
         Begin VB.Label Label104 
            Caption         =   "Spe. Ill. W. P."
            Height          =   255
            Left            =   4680
            TabIndex        =   138
            Top             =   1530
            Width           =   1065
         End
         Begin VB.Label Label103 
            Caption         =   "Pre W. P."
            Height          =   255
            Left            =   4680
            TabIndex        =   137
            Top             =   1245
            Width           =   780
         End
         Begin VB.Label Label78 
            Caption         =   "Gen W. P."
            Height          =   255
            Left            =   4680
            TabIndex        =   136
            Top             =   990
            Width           =   900
         End
         Begin VB.Label Label33 
            Caption         =   "EOP"
            Height          =   255
            Left            =   6840
            TabIndex        =   135
            Top             =   1030
            Width           =   1005
         End
         Begin VB.Label Label45 
            Caption         =   "EOP"
            Height          =   255
            Left            =   6840
            TabIndex        =   134
            Top             =   1560
            Width           =   1005
         End
         Begin VB.Label Label36 
            Caption         =   "Disc Req"
            Height          =   255
            Left            =   2355
            TabIndex        =   133
            Top             =   1300
            Width           =   1065
         End
         Begin VB.Label Label92 
            Caption         =   "Pay Remarks"
            Height          =   255
            Left            =   4680
            TabIndex        =   132
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label91 
            Caption         =   "T/O Ins "
            Height          =   255
            Left            =   7200
            TabIndex        =   131
            Top             =   480
            Width           =   1140
         End
         Begin VB.Label Label90 
            Caption         =   "R&&B Amt"
            Height          =   255
            Left            =   120
            TabIndex        =   130
            Top             =   1300
            Width           =   1140
         End
         Begin VB.Label Label82 
            Caption         =   "P. Pln Code"
            Height          =   255
            Left            =   4680
            TabIndex        =   129
            Top             =   525
            Width           =   1125
         End
         Begin VB.Label Label38 
            Caption         =   "Name"
            Height          =   255
            Left            =   2355
            TabIndex        =   128
            Top             =   195
            Width           =   1140
         End
         Begin VB.Label Label51 
            Caption         =   "Adult/Child"
            Height          =   255
            Left            =   120
            TabIndex        =   127
            Top             =   795
            Width           =   825
         End
         Begin VB.Label Label50 
            Caption         =   "Sex "
            Height          =   255
            Left            =   2355
            TabIndex        =   126
            Top             =   750
            Width           =   825
         End
         Begin VB.Label Label39 
            Caption         =   "Salutation"
            Height          =   255
            Left            =   120
            TabIndex        =   125
            Top             =   225
            Width           =   1140
         End
         Begin VB.Label Label41 
            Caption         =   "IC Num"
            Height          =   255
            Left            =   7200
            TabIndex        =   124
            Top             =   165
            Width           =   1125
         End
         Begin VB.Label Label49 
            Caption         =   "DOB"
            Height          =   255
            Left            =   2355
            TabIndex        =   123
            Top             =   495
            Width           =   1020
         End
         Begin VB.Label Label46 
            Caption         =   "Supp Lmt"
            Height          =   255
            Left            =   2355
            TabIndex        =   122
            Top             =   1020
            Width           =   855
         End
         Begin VB.Label Label52 
            Caption         =   "R'Ship"
            Height          =   255
            Left            =   120
            TabIndex        =   121
            Top             =   525
            Width           =   1005
         End
         Begin VB.Label Label47 
            Caption         =   "Supplementary Warranty "
            Height          =   255
            Left            =   6360
            TabIndex        =   120
            Top             =   1920
            Width           =   2100
         End
         Begin VB.Label Label56 
            Caption         =   "Bal. Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   119
            Top             =   1020
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5085
         Left            =   120
         TabIndex        =   75
         Top             =   360
         Width           =   11295
         Begin VB.TextBox txtLifeTimeStatus 
            Height          =   285
            Left            =   10920
            TabIndex        =   181
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Frame Frame4 
            Height          =   945
            Left            =   120
            TabIndex        =   76
            Top             =   4080
            Width           =   11055
            Begin VB.Label Label75 
               Caption         =   "Cancer Treat. Lmt"
               Height          =   255
               Left            =   5640
               TabIndex        =   189
               Top             =   600
               Width           =   1305
            End
            Begin VB.Label Label74 
               Caption         =   "Home Nursing"
               Height          =   255
               Left            =   8520
               TabIndex        =   188
               Top             =   240
               Width           =   1065
            End
            Begin VB.Label Label73 
               Caption         =   "Kidney Dialysis Lmt"
               Height          =   255
               Left            =   5640
               TabIndex        =   187
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lblKidneyLmt 
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
               Left            =   7080
               TabIndex        =   186
               Top             =   240
               Width           =   1290
            End
            Begin VB.Label lblCancerLmt 
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
               Left            =   7080
               TabIndex        =   185
               Top             =   600
               Width           =   1290
            End
            Begin VB.Label lblHomeNurseLmt 
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
               Left            =   9675
               TabIndex        =   184
               Top             =   240
               Width           =   1290
            End
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
               Left            =   3720
               TabIndex        =   84
               Top             =   600
               Width           =   1320
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
               Left            =   1080
               TabIndex        =   83
               Top             =   240
               Width           =   1320
            End
            Begin VB.Label Label25 
               Caption         =   "Premium"
               Height          =   255
               Left            =   120
               TabIndex        =   82
               Top             =   240
               Width           =   600
            End
            Begin VB.Label Label26 
               Caption         =   "MCO Fees"
               Height          =   240
               Left            =   2760
               TabIndex        =   81
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label27 
               Caption         =   "Annual Limit"
               Height          =   255
               Left            =   2760
               TabIndex        =   80
               Top             =   600
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
               Left            =   3720
               TabIndex        =   79
               Top             =   240
               Width           =   1320
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
               Left            =   1080
               TabIndex        =   78
               Top             =   600
               Width           =   1320
            End
            Begin VB.Label Label30 
               Caption         =   "Access Fee"
               Height          =   255
               Left            =   120
               TabIndex        =   77
               Top             =   600
               Width           =   1140
            End
         End
         Begin VB.TextBox txtMBMExclusion 
            Height          =   810
            Left            =   240
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   3240
            Width           =   4890
         End
         Begin VB.TextBox txtRemarks 
            Height          =   855
            Left            =   5880
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   3240
            Width           =   5130
         End
         Begin VB.TextBox txtMBMAdd1 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   8
            Top             =   150
            Width           =   4065
         End
         Begin VB.ComboBox cboPLNCode 
            Height          =   315
            Left            =   9720
            Sorted          =   -1  'True
            TabIndex        =   23
            Top             =   120
            Width           =   1140
         End
         Begin VB.ComboBox cboINSCode 
            Height          =   315
            ItemData        =   "frmNewMemberRegistration.frx":0038
            Left            =   7080
            List            =   "frmNewMemberRegistration.frx":003A
            Sorted          =   -1  'True
            TabIndex        =   22
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox txtMBMAdd2 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   9
            Top             =   400
            Width           =   4065
         End
         Begin VB.TextBox txtMBMAdd3 
            Height          =   285
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   10
            Top             =   660
            Width           =   4065
         End
         Begin VB.TextBox txtMBMPostCode 
            Height          =   315
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   11
            Top             =   910
            Width           =   1215
         End
         Begin VB.ComboBox cboCTYCode 
            Height          =   315
            ItemData        =   "frmNewMemberRegistration.frx":003C
            Left            =   3480
            List            =   "frmNewMemberRegistration.frx":003E
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   910
            Width           =   1680
         End
         Begin VB.ComboBox cboSTACode 
            Height          =   315
            Left            =   1080
            TabIndex        =   13
            Top             =   1200
            Width           =   4095
         End
         Begin VB.ComboBox cboMBMTakeover 
            Height          =   315
            Left            =   3525
            TabIndex        =   15
            Text            =   "N"
            Top             =   1480
            Width           =   1650
         End
         Begin VB.ComboBox cboMBMMaritalStatus 
            Height          =   315
            Left            =   3525
            TabIndex        =   17
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
            TabIndex        =   14
            Tag             =   "M"
            Text            =   "M"
            Top             =   1480
            Width           =   1215
         End
         Begin VB.ComboBox cboNATCode 
            Height          =   315
            Left            =   3525
            Sorted          =   -1  'True
            TabIndex        =   19
            Text            =   "MY-Malaysia"
            Top             =   2040
            Width           =   1650
         End
         Begin MSMask.MaskEdBox txtMBMPayorExpDate 
            Height          =   315
            Left            =   9720
            TabIndex        =   25
            Top             =   410
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtMBMBirthdate 
            Height          =   315
            Left            =   1080
            TabIndex        =   16
            Top             =   1760
            Width           =   1215
            _ExtentX        =   2143
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
            Height          =   315
            Left            =   7080
            TabIndex        =   24
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
            TabIndex        =   26
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
            TabIndex        =   27
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
            TabIndex        =   28
            Top             =   1230
            Width           =   3825
         End
         Begin VB.TextBox txtMBMDepartment 
            Height          =   315
            Left            =   7080
            MaxLength       =   50
            TabIndex        =   29
            Top             =   1520
            Width           =   3825
         End
         Begin VB.TextBox txtMBMGroupCompany 
            Height          =   315
            Left            =   7080
            MaxLength       =   150
            TabIndex        =   30
            Top             =   1800
            Width           =   3825
         End
         Begin VB.ComboBox cboRACCode 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   18
            Text            =   "C - Chinese"
            Top             =   2040
            Width           =   1215
         End
         Begin VB.ComboBox cboBRNCode 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            TabIndex        =   20
            Top             =   2320
            Width           =   4095
         End
         Begin VB.TextBox txtMBMAgentCode 
            Height          =   285
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   21
            Top             =   2600
            Width           =   4065
         End
         Begin VB.ComboBox CboInsMode 
            Height          =   315
            Left            =   9720
            TabIndex        =   31
            Text            =   "A - Annual"
            Top             =   2080
            Width           =   1185
         End
         Begin VB.ComboBox CboGuaRenewal 
            Height          =   315
            ItemData        =   "frmNewMemberRegistration.frx":0040
            Left            =   7080
            List            =   "frmNewMemberRegistration.frx":0042
            TabIndex        =   32
            Top             =   2370
            Width           =   3855
         End
         Begin VB.Label Label4 
            Caption         =   "Guaranty Renew"
            Height          =   255
            Left            =   5880
            TabIndex        =   182
            Top             =   2415
            Width           =   1260
         End
         Begin VB.Label Label40 
            Caption         =   "Principal Warranty "
            Height          =   255
            Left            =   240
            TabIndex        =   113
            Top             =   3000
            Width           =   1590
         End
         Begin VB.Label Label76 
            Caption         =   "Remarks"
            Height          =   255
            Left            =   5880
            TabIndex        =   112
            Top             =   3000
            Width           =   1590
         End
         Begin VB.Label Label70 
            Caption         =   "Instl mode"
            Height          =   255
            Left            =   8920
            TabIndex        =   111
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label txtMBMPolicyYear 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "1"
            Height          =   315
            Left            =   9735
            TabIndex        =   110
            Top             =   990
            Width           =   1140
         End
         Begin VB.Label Label59 
            Caption         =   "Branch"
            Height          =   255
            Left            =   180
            TabIndex        =   108
            Top             =   2320
            Width           =   600
         End
         Begin VB.Label Label10 
            Caption         =   "Race"
            Height          =   255
            Left            =   180
            TabIndex        =   107
            Top             =   2040
            Width           =   600
         End
         Begin VB.Label Label57 
            Caption         =   "Dept. Name"
            Height          =   255
            Left            =   5880
            TabIndex        =   106
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "E'yee Num"
            Height          =   255
            Left            =   5880
            TabIndex        =   105
            Top             =   1260
            Width           =   1545
         End
         Begin VB.Label txtMBMBatchNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   9735
            TabIndex        =   73
            Top             =   720
            Width           =   1140
         End
         Begin VB.Label Label17 
            Caption         =   "Bordx. Date"
            Height          =   255
            Left            =   5880
            TabIndex        =   104
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label44 
            Caption         =   "Batch Num"
            Height          =   255
            Left            =   8640
            TabIndex        =   103
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
            TabIndex        =   102
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
            TabIndex        =   101
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label Label63 
            Caption         =   "Take Over"
            Height          =   255
            Left            =   2520
            TabIndex        =   100
            Top             =   1480
            Width           =   825
         End
         Begin VB.Label Label24 
            Caption         =   "Marital Status"
            Height          =   255
            Left            =   2520
            TabIndex        =   99
            Top             =   1760
            Width           =   1095
         End
         Begin VB.Label Label42 
            Caption         =   "Date Join"
            Height          =   255
            Left            =   5880
            TabIndex        =   98
            Top             =   990
            Width           =   780
         End
         Begin VB.Label Label7 
            Caption         =   "State"
            Height          =   255
            Left            =   180
            TabIndex        =   97
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
            TabIndex        =   96
            Top             =   150
            Width           =   960
         End
         Begin VB.Label Label12 
            Caption         =   "PostCode"
            Height          =   255
            Left            =   180
            TabIndex        =   95
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
            TabIndex        =   94
            Top             =   1760
            Width           =   765
         End
         Begin VB.Label Label5 
            Caption         =   "Sex (M/F)"
            Height          =   255
            Left            =   180
            TabIndex        =   93
            Top             =   1480
            Width           =   825
         End
         Begin VB.Label Label2 
            Caption         =   "Address"
            Height          =   255
            Left            =   180
            TabIndex        =   92
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
            TabIndex        =   91
            Top             =   440
            Width           =   1275
         End
         Begin VB.Label Label6 
            Caption         =   "City/Town"
            Height          =   255
            Left            =   2520
            TabIndex        =   90
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label11 
            Caption         =   "Nationality"
            Height          =   255
            Left            =   2520
            TabIndex        =   89
            Top             =   2040
            Width           =   960
         End
         Begin VB.Label Label14 
            Caption         =   "Agent "
            Height          =   255
            Left            =   180
            TabIndex        =   88
            Top             =   2600
            Width           =   660
         End
         Begin VB.Label Label43 
            Caption         =   "Pol Year"
            Height          =   255
            Left            =   8640
            TabIndex        =   87
            Top             =   1020
            Width           =   1035
         End
         Begin VB.Label Label18 
            Caption         =   "Grp Company"
            Height          =   255
            Left            =   5880
            TabIndex        =   86
            Top             =   1840
            Width           =   1545
         End
         Begin VB.Label txtAGECode 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7080
            TabIndex        =   109
            Top             =   2115
            Width           =   1785
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
         TabIndex        =   151
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
         NumItems        =   44
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
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   30
            Text            =   "Clm Non-Payment To"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Text            =   "Clinic Warranty"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   32
            Text            =   "PayorPrem"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   33
            Text            =   "PayorPremNet"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   34
            Text            =   "PayorMCOGross"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   35
            Text            =   "PayorMCONet"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   36
            Text            =   "RenewalDisc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   37
            Text            =   "PolDisc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   38
            Text            =   "PolDiscAmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   39
            Text            =   "RenDiscAmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   40
            Text            =   "LoadingAmt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   41
            Text            =   "PLNCode"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   42
            Text            =   "AvailLimit"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   43
            Text            =   "UndExcess"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   80
      TabIndex        =   152
      Top             =   240
      Width           =   11775
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Enabled         =   0   'False
         Height          =   300
         Left            =   6480
         TabIndex        =   154
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Go"
         Enabled         =   0   'False
         Height          =   300
         Left            =   5400
         TabIndex        =   153
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
         ItemData        =   "frmNewMemberRegistration.frx":0044
         Left            =   1320
         List            =   "frmNewMemberRegistration.frx":0046
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
         TabIndex        =   158
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label68 
         Caption         =   "Prev. Pol. Num"
         Height          =   255
         Left            =   7440
         TabIndex        =   157
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label67 
         Caption         =   "Prev Mem No"
         Height          =   240
         Left            =   2040
         TabIndex        =   156
         Top             =   225
         Width           =   1290
      End
      Begin VB.Label Label66 
         Caption         =   "Renewal (Y/N)"
         Height          =   240
         Left            =   120
         TabIndex        =   155
         Top             =   225
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   80
      TabIndex        =   168
      Top             =   840
      Width           =   11775
      Begin VB.CommandButton cmdClinic 
         BackColor       =   &H00808080&
         Caption         =   "Clinical >"
         Height          =   255
         Left            =   9960
         MaskColor       =   &H00808080&
         TabIndex        =   169
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtMBMIcBcPP 
         Height          =   285
         Left            =   8400
         MaxLength       =   30
         TabIndex        =   4
         Top             =   160
         Width           =   3015
      End
      Begin VB.TextBox txtMBMName 
         Height          =   330
         Left            =   3120
         MaxLength       =   60
         TabIndex        =   3
         Top             =   160
         Width           =   4245
      End
      Begin VB.ComboBox cboMBMSalutation 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Text            =   "MR"
         Top             =   160
         Width           =   690
      End
      Begin VB.TextBox txtPAYCode 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   440
         Width           =   690
      End
      Begin VB.TextBox txtMBMPolicyNo 
         Height          =   330
         Left            =   3120
         MaxLength       =   25
         TabIndex        =   6
         Top             =   440
         Width           =   4245
      End
      Begin VB.TextBox txtHLTCode 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8400
         TabIndex        =   7
         Top             =   420
         Width           =   1335
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
         TabIndex        =   175
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label55 
         Caption         =   "Payor"
         Height          =   255
         Left            =   120
         TabIndex        =   174
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
         TabIndex        =   173
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label Label64 
         Caption         =   "Health Type"
         Height          =   255
         Left            =   7440
         TabIndex        =   172
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
         TabIndex        =   171
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Salutation"
         Height          =   255
         Left            =   120
         TabIndex        =   170
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   120
      TabIndex        =   159
      Top             =   7320
      Width           =   11715
      Begin VB.CommandButton CmdMBMOthers 
         Caption         =   "Others Details"
         Height          =   315
         Left            =   7200
         TabIndex        =   167
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox txtcount 
         Height          =   285
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   166
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtINSFlag 
         Height          =   285
         Left            =   3120
         TabIndex        =   165
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtMBMAdultChild 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   164
         Text            =   "A"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtTempBordx 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   163
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtTempBatch 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   162
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Sa&ve"
         Height          =   315
         Left            =   9435
         TabIndex        =   72
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   315
         Left            =   10440
         TabIndex        =   161
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   8445
         TabIndex        =   160
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Member Registration"
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
      TabIndex        =   179
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmNewMemberRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecordSaved As Boolean
Dim startTrans As Boolean
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
Dim ADOrstMBM  As New ADODB.Recordset
Dim sql As String
Dim MembershipNo As String
Dim completeMBM, completeMBMOthers, completeMBMC, completeMBMCLNOthers As Boolean
Dim completeMBMTwo, completeMBMOthersTwo, completeMBMCRef As Boolean
Dim TempPayEffDate As String
Dim TempPayExpDate As String
Dim WLifeStatus As Boolean
Dim PrevLifeMBMNo As String
Dim TempLifeTimeStatus As String
Dim TmpTopupNo As String
Dim CISDCStatus As String
Dim CISDCPayor As String
Dim CISPDentStatus As String
Dim CISSDentStatus As String
Dim CISPSpecStatus As String
Dim CISSSpecStatus As String
Dim CISPMepStatus As String
Dim CISSMepStatus As String

Private Sub cboMBMSalutation_LostFocus()
    cboMBMSalutation = ChkStr(cboMBMSalutation)
End Sub

Private Sub cboMBMSex_LostFocus()
    cboMBMSex.Text = ChkStr(cboMBMSex)
End Sub

Private Sub cboMBMTakeover_LostFocus()
    cboMBMTakeover.Text = ChkStr(cboMBMTakeover.Text)
End Sub

Private Sub cboPLNCode_LostFocus()
Dim strsqlPLN As String
Dim PLNAS As New ADODB.Recordset
Dim medixmasterconnPLN As New ADODB.Connection

    medixmasterconnPLN.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    TempLifeTimeStatus = ""
    strsqlPLN = "select PLNLifeTimeStatus from PLN where PLNCode = '" & Trim(Mid(cboPLNCode, 1, 4)) & "'"
    PLN.Open strsqlPLN, medixmasterconnPLN, adOpenKeyset, adLockOptimistic
    If PLN.recordCount Then
        TempLifeTimeStatus = IIf(Len(PLN!PLNLifeTimeStatus), PLN!PLNLifeTimeStatus, "")
    End If
    
    PLN.Close
    Set PLN = Nothing
    medixmasterconnPLN.Close
    Set medixmasterconnPLN = Nothing
End Sub

Private Sub cmdClinic_Click()
    frmMBMClinical.Show
End Sub

Private Sub CmdMBMOthers_Click()
    frmNewMBMOthers.Source = 3
    
    If frmNewMBMOthers.Visible = False Then
        Load frmNewMBMOthers
    End If
    frmNewMBMOthers.Show
    
    frmNewMBMOthers.Frame5.Enabled = True
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

Private Sub cboPLNCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 0
        Shift = 0
        Clipboard.Clear
    End If
End Sub

Private Sub Form_Load()
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    Call ComboListing
    Call LoadInsuredType
    Call LoadPlanType("I")
    cboMBMRenewal.SelStart = 0
    bExit = False
    tabIndexDone = False
    txtTempProvider = 0
    txtTempPrem = 0
    txtTempMCOFees = 0
    PolDics = 0
    RenDisc = 0
    If frmNewMBMOthers.Visible = False Then
        frmNewMBMOthers.Source = 2
        Load frmNewMBMOthers
    End If
    
    Call ClearForm(frmNewMBMOthers)
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub ComboListing()
Dim RAC As New ADODB.Recordset
Dim NAT As New ADODB.Recordset
Dim REL As New ADODB.Recordset
Dim CTY As New ADODB.Recordset
Dim STA As New ADODB.Recordset
Dim Work As New ADODB.Recordset
Dim cmd As New ADODB.Command

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchRace"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    RAC.CursorLocation = adUseClient
    RAC.CursorType = adOpenForwardOnly
    RAC.CacheSize = 100
    Set RAC = cmd.Execute
        
    Do Until RAC.EOF
        With RAC
            cboRACCode.AddItem !RACCode & " - " & !RACName
            RAC.MoveNext
        End With
    Loop
    RAC.Close
    Set RAC = Nothing

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchNAT"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    NAT.CursorLocation = adUseClient
    NAT.CursorType = adOpenForwardOnly
    NAT.CacheSize = 100
    Set NAT = cmd.Execute

    Do Until NAT.EOF
        With NAT
            cboNATCode.AddItem !NATCode & " - " & !NATName
            NAT.MoveNext
        End With
    Loop
    NAT.Close
    Set NAT = Nothing


    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchREL"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    REL.CursorLocation = adUseClient
    REL.CursorType = adOpenForwardOnly
    REL.CacheSize = 100
    Set REL = cmd.Execute

    Do Until REL.EOF
        With REL
            cboCovRelationship.AddItem !RELCODE & " - " & !RELDescription
            REL.MoveNext
        End With
    Loop
    REL.Close
    Set REL = Nothing


    Set cmd.ActiveConnection = medixmasterconnMaint
    cmd.CommandText = "SearchCity"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    CTY.CursorLocation = adUseClient
    CTY.CursorType = adOpenForwardOnly
    CTY.CacheSize = 100
    Set CTY = cmd.Execute

    Do Until CTY.EOF
        With CTY
            cboBRNCode.AddItem !CTYDescription
            cboCTYCode.AddItem !CTYDescription
            CTY.MoveNext
        End With
    Loop
    CTY.Close
    Set CTY = Nothing

    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchState"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    STA.CursorLocation = adUseClient
    STA.CursorType = adOpenForwardOnly
    STA.CacheSize = 100
    Set STA = cmd.Execute

    Do Until STA.EOF
        With STA
            cboSTACode.AddItem !STADescription
            STA.MoveNext
        End With
    Loop
    STA.Close
    Set STA = Nothing
    
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
    
    CboGuaRenewal.Clear
    CboGuaRenewal.AddItem "  "
    CboGuaRenewal.AddItem "1 - No Life Time Limit"
    CboGuaRenewal.AddItem "3 - With Life Time Limit"
End Sub

Private Sub LoadInsuredType()
    Dim strsql As String
    Dim INS As New ADODB.Recordset
   
    cboINSCode.Clear
    cboINSCode.AddItem " "
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
Dim MBMCovTwo As New ADODB.Recordset
Dim sqlstrTwo As String
Dim CovTwo As Boolean

    lstCovAdd.listitems.Clear
        Dim MemberConveredPerson As ListItem
        Do While Not ADOrstMBM.EOF
            
            With ADOrstMBM
                
                Set MemberConveredPerson = lstCovAdd.listitems.Add(, , IIf(!MBMCName <> "", !MBMCName, ""))
                If Len(Trim(!MBMCICBCPPNo)) <> 0 Then
                    MemberConveredPerson.SubItems(2) = !MBMCICBCPPNo
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
                
                If Len(!MBMCCoverID) Then
                    MemberConveredPerson.SubItems(10) = !MBMCCoverID
                End If
                
                If Len(Trim(!MBMCAnnualLimit)) Then
                    MemberConveredPerson.SubItems(11) = !MBMCAnnualLimit
                End If
                
                
                If Len(!MBMCRenewalDisc) Then
                    MemberConveredPerson.SubItems(36) = !MBMCRenewalDisc
                End If
                
                If Len(!MBMCPolicyDisc) Then
                    MemberConveredPerson.SubItems(37) = !MBMCPolicyDisc
                End If
                                                                
                If Len(!MBMCPLNCode) Then
                    MemberConveredPerson.SubItems(41) = !MBMCPLNCode
                End If
                
                If Len(!MBMCAvailableLimit) Then
                    MemberConveredPerson.SubItems(42) = !MBMCAvailableLimit
                End If
                    
                    If Len(Trim(!MBMCSalutation)) <> 0 Then
                        MemberConveredPerson.SubItems(1) = ChkStr(!MBMCSalutation)
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
            End With
            ADOrstMBM.MoveNext
        Loop
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
    cboCovRelationship.Text = "W - Wife"
    cboCovSalutation.Text = "MR"
    cboCovSex.Text = "M"
    cboCovAdultChild.Text = "A"
    txtCovAnnualLimit = 0
    cboMBMSalutation.Text = "MR"
    cboMBMRenewal = "N"
    cboMBMSex = "M"
    cboMBMTakeover = "N - None"
    cboMBMMaritalStatus = "M"
    cboRACCode.Text = "C - Chinese"
    cboNATCode.Text = "MY - Malaysia"
    cboINSCode = " "
    lstCovAdd.listitems.Clear
    txtAGECode = ""
    txtMBMPremium = ""
    txtMCOFees = ""
    txtMBMProviderCharges = ""
    txtAnnualLimit = ""
    lblKidneyLmt = ""
    lblHomeNurseLmt = ""
    lblCancerLmt = ""
    frmNewMBMOthers.txtGENWPEnd.mask = ""
    frmNewMBMOthers.txtGENWPEnd.Text = ""
    frmNewMBMOthers.txtGENWPEnd.mask = "##/##/####"
    frmNewMBMOthers.txtPREWPEnd.mask = ""
    frmNewMBMOthers.txtPREWPEnd.Text = ""
    frmNewMBMOthers.txtPREWPEnd.mask = "##/##/####"
    frmNewMBMOthers.txtSPEWPEnd.mask = ""
    frmNewMBMOthers.txtSPEWPEnd.Text = ""
    frmNewMBMOthers.txtSPEWPEnd.mask = "##/##/####"
    frmNewMBMOthers.TxtClmNonPayFr.Text = ""
    frmNewMBMOthers.TxtClmNonPayFr.mask = "##/##/####"
    frmNewMBMOthers.TxtClmNonPayTo.Text = ""
    frmNewMBMOthers.TxtClmNonPayTo.mask = "##/##/####"
    frmNewMBMOthers.cboBTBG = "N - No"
    frmNewMBMOthers.txtUNDExcess = ""
    frmMBMClinical.cboINSCode.Text = ""
    frmMBMClinical.cboPLNCode.Text = ""
    frmMBMClinical.CboCardType.Text = ""
    frmMBMClinical.txtMBMPayorEffDate = "__/__/____"
    frmMBMClinical.txtMBMPayorExpDate = "__/__/____"
    frmMBMClinical.txtAGECode = ""
    frmMBMClinical.lblAnnualLimit = ""
    frmMBMClinical.lblMBMPremium = ""
    frmMBMClinical.LblVisitLimit = ""
    frmMBMClinical.lblMCOFees = ""
    TempLifeTimeStatus = ""
    MembershipNo = ""
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
        txtPolicyDiscAmt = lstCovAdd.SelectedItem.SubItems(38)
        txtRenewalDiscAmt = lstCovAdd.SelectedItem.SubItems(39)
        txtLoadingPrem = lstCovAdd.SelectedItem.SubItems(40)
        txtSuppPLNCode = lstCovAdd.SelectedItem.SubItems(41)
        txtAvailableLimit = Format(lstCovAdd.SelectedItem.SubItems(42), "###,##0.00")

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
        txtPolicyDiscAmt = lstCovAdd.SelectedItem.SubItems(38)
        txtRenewalDiscAmt = lstCovAdd.SelectedItem.SubItems(39)
        txtLoadingPrem = lstCovAdd.SelectedItem.SubItems(40)
        txtSuppPLNCode = lstCovAdd.SelectedItem.SubItems(41)
        txtAvailableLimit = Format(lstCovAdd.SelectedItem.SubItems(42), "###,##0.00")

        Call EnableCovered(False)
        cmdUpdate.Enabled = True
        cmdRemove.Enabled = True
        cmdAdd.Enabled = False
    End If
    Call EnableCovered(True)
End Sub

'  Contorl of Tab index  *********************************88

Private Sub SSTab1_Click(PreviousTab As Integer)
    txtCLNCovExclusion.Enabled = False
    If SSTab1.Tab = 1 Then
'        If Trim(Mid(cboINSCode, 1, 1)) = "I" Or Trim(Mid(cboINSCode, 1, 1)) = "G" Then
        If Trim(Mid(cboINSCode, 1, 1)) = "F" Or Trim(Mid(cboINSCode, 1, 1)) = "H" Or _
            Trim(Mid(frmMBMClinical.cboINSCode, 1, 1)) = "F" Or Trim(Mid(frmMBMClinical.cboINSCode, 1, 1)) = "H" Then
            cboCovSalutation.Enabled = True
            If Trim(Mid(frmMBMClinical.cboINSCode, 1, 1)) = "F" Or Trim(Mid(frmMBMClinical.cboINSCode, 1, 1)) = "H" Then
                txtCLNCovExclusion.Enabled = True
            End If
            cboCovSalutation.SetFocus
        Else
            MsgBox "No Covered Person Needed For This Insured Type", vbOKOnly + vbCritical, DialogBoxTitle
            cboINSCode.SetFocus
            SSTab1.Tab = 0
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

Private Sub SSTab1_GotFocus()
    If SSTab1.Tab = 0 Then
        txtMBMAdd1.SetFocus
    End If
End Sub

Private Sub SSTab1_LostFocus()
    If SSTab1.Tab = 0 Then
        txtMBMAdd1.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        cboCovSalutation.SetFocus
    End If
End Sub

Private Sub txtCLNCovExclusion_LostFocus()
    txtCLNCovExclusion = ChkStr(txtCLNCovExclusion)
End Sub

Private Sub txtCovAllergic_Click()
    KeyPreview = False
End Sub

Private Sub txtCovAllergic_LostFocus()
    KeyPreview = True
    txtCovAllergic = ChkStr(txtCovAllergic)
End Sub

Private Sub txtCovBirthdate_LostFocus()
    Dim MBMValidDate As New ADODB.Recordset
    Dim strsql As String
    If txtCovBirthdate <> "__/__/____" Then
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
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
        medixmasterconn.Close
        Set medixmasterconn = Nothing
    End If
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

Private Sub txtCOVPayRemarks_LostFocus()
    txtCOVPayRemarks = ChkStr(txtCOVPayRemarks)
End Sub

Private Sub txtMBMAdd1_LostFocus()
    txtMBMAdd1 = ChkStr(txtMBMAdd1)
End Sub

Private Sub txtMBMAdd2_LostFocus()
    txtMBMAdd2 = ChkStr(txtMBMAdd2)
End Sub

Private Sub txtMBMAdd3_LostFocus()
    txtMBMAdd3 = ChkStr(txtMBMAdd3)
End Sub

Private Sub txtMBMAgentCode_LostFocus()
    txtMBMAgentCode = ChkStr(txtMBMAgentCode)
End Sub

Private Sub txtMBMDepartment_LostFocus()
    txtMBMDepartment = ChkStr(txtMBMDepartment)
End Sub

Private Sub txtMBMEmployeeNo_LostFocus()
    txtMBMEmployeeNo = ChkStr(txtMBMEmployeeNo)
End Sub

Private Sub txtMBMGroupCompany_LostFocus()
    txtMBMGroupCompany = ChkStr(txtMBMGroupCompany)
End Sub

Private Sub txtMBMName_LostFocus()
    txtMBMName = ChkStr(txtMBMName)
End Sub

Private Sub txtMBMPolicyNo_LostFocus()
    txtMBMPolicyNo = ChkStr(txtMBMPolicyNo)
End Sub

Private Sub txtMBMPrevMEMNo_LostFocus()
    txtMBMPrevMEMNo = ChkStr(txtMBMPrevMEMNo)
End Sub

Private Sub txtPrevINSCom_LostFocus()
    txtPrevINSCom = ChkStr(txtPrevINSCom)
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
    txtMBMExclusion = ChkStr(txtMBMExclusion)
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
    txtMBMIcBcPP = ChkStr(txtMBMIcBcPP)
    
    

    If Len(Trim(txtMBMIcBcPP)) Then
        Screen.MousePointer = vbHourglass
        If CheckFlag = False Then
            strsql = "Select MBMNumber From MBMCrossReference Where MBMICBCPP ='" & Trim(txtMBMIcBcPP) & "'"
            
            medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
            
            MBMIC.Open strsql, medixmasterconnMaint, adOpenForwardOnly, adLockBatchOptimistic
            
           If MBMIC.EOF = False Then
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
            medixmasterconnMaint.Close
            Set medixmasterconnMaint = Nothing
            
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
Dim TmpEffDate As String
Dim TmpDOB As String
Dim tmpHLTCode As String
Dim tmpEffStartDate As String
Dim tmpEffEndDate As String

    If IsDate(txtMBMBirthdate) = False Then
        MsgBox "Invalid date format ! ", vbCritical + vbOKOnly
        txtMBMBirthdate.SetFocus
    Else
        TmpEffDate = txtMBMPayorEffDate
        TmpDOB = txtMBMBirthdate
        tmpHLTCode = Trim(txtHLTCode)
        tmpEffStartDate = ""
        tmpEffEndDate = ""
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        strsql = "Select ValidID ,MBMDOBStart , MBMDOBEnd  From MBMValidDateRange "
        MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
        If MBMValidDate.recordCount Then
            tmpEffStartDate = IIf(Len(MBMValidDate!MBMDOBStart), MBMValidDate!MBMDOBStart, "")
            tmpEffEndDate = IIf(Len(MBMValidDate!MBMDOBEnd), MBMValidDate!MBMDOBEnd, "")
        End If
        MBMValidDate.Close
        Set MBMValidDate = Nothing
        medixmasterconn.Close
        Set medixmasterconn = Nothing
        If tmpEffStartDate <> "" And tmpEffEndDate <> "" Then
            If DateDiff("d", tmpEffStartDate, txtMBMBirthdate) >= 0 And DateDiff("d", tmpEffEndDate, txtMBMBirthdate) <= 0 Then
                If Trim(cboPLNCode) <> "" Then
                    If txtMBMPayorEffDate <> "__/__/____" Then
                        TmpEffDate = txtMBMPayorEffDate
                    End If
                ElseIf Trim(frmMBMClinical.cboPLNCode) <> "" Then
                    If frmMBMClinical.txtMBMPayorEffDate <> "__/__/____" Then
                        TmpEffDate = frmMBMClinical.txtMBMPayorEffDate
                    End If
                End If
            Else
                DateExit = MsgBox("Uncertain Birthdate, please select OK to continue " & vbNewLine & _
                            " Otherwise press Cancel to make changes! ", vbExclamation + vbOKCancel + vbDefaultButton2)
                If DateExit = vbCancel Then
                    txtMBMBirthdate.SetFocus
                End If
            End If
            If IsDate(TmpEffDate) Then
                medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
                medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
                Call MemberCallAge(TmpEffDate, TmpDOB, tmpHLTCode)
                Call CallLimit
                Call MemberCallFees
                medixmasterconn.Close
                Set medixmasterconn = Nothing
                medixmasterconnMaint.Close
                Set medixmasterconnMaint = Nothing
            End If
        End If
        frmMBMClinical.txtMBMBirthdate = txtMBMBirthdate
    End If
End Sub

Private Sub cboPLNCode_Click()
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
    If Len(cboPLNCode) >= 2 Then
        Call showGroup
        If txtMBMPayorEffDate <> "__/__/____" Then
            Call CallLimit
            Call MemberCallFees
            Call showGroup
        End If
    End If
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
End Sub



Private Sub cboINSCode_Click()
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"

    If Len(Trim(cboINSCode)) Then
        If Trim(cboINSCode) = "F" Or Trim(cboINSCode) = "H" Or Trim(frmMBMClinical.cboINSCode) = "F" Or Trim(frmMBMClinical.cboINSCode) = "H" Then
            cboMBMMaritalStatus = "M"
                Call EnableSupplementary(True)
        Else
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
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    
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
    frmSearchMBM.Source = 1
    frmSearchMBM.Show vbModal
End Sub


Private Sub txtMBMPayorEffDate_LostFocus()
Dim MBMValidDate As New ADODB.Recordset
Dim strsql As String
Dim tempdate As Date
Dim answer
Dim TmpEffDate As String
Dim TmpDOB As String
Dim TmpValidEffdate As String

    TmpDOB = txtMBMBirthdate

    If (txtMBMPayorEffDate) = "__/__/____" Then
        Exit Sub
    End If
    If IsDate(txtMBMPayorEffDate) = False Then
        answer = MsgBox("Invalid Date Format!", vbCritical + vbOKCancel)
        If answer = vbOK Then
            txtMBMPayorEffDate.SetFocus
        Else
            txtMBMPayorExpDate.SetFocus
        End If
    Else
        If ChkStr(cboMBMRenewal) = "N" Then
            txtMBMDateJoin = txtMBMPayorEffDate
        Else
            If txtMBMDateJoin = "__/__/____" Then
                txtMBMDateJoin = txtMBMPayorEffDate
            End If
        End If
        medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
        
        strsql = "Select ValidID , MBMEffStart, MBMEffEnd From MBMValidDateRange "
        MBMValidDate.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
                        
        If MBMValidDate.recordCount Then
             If DateDiff("d", MBMValidDate!MBMEffStart, txtMBMPayorEffDate) >= 0 And DateDiff("d", MBMValidDate!MBMEffEnd, txtMBMPayorEffDate) <= 0 Then
                 tempdate = Format(DateAdd("yyyy", 1, txtMBMPayorEffDate), "dd/mm/yyyy")
                 txtMBMPayorExpDate = Format(DateAdd("d", -1, tempdate), "dd/mm/yyyy")
                 'txtMBMEmployeeNo.SetFocus
                 TmpEffDate = txtMBMPayorExpDate
             Else
                 MsgBox "Policy Effective Date is out of acceptable range! " & vbNewLine & _
                         "It must be in between " & MBMValidDate!MBMEffStart & " and " & MBMValidDate!MBMEffEnd & "! ", vbCritical + vbOKOnly
                 'txtMBMPayorEffDate.SetFocus
             End If
        Else
            tempdate = Format(DateAdd("yyyy", 1, txtMBMPayorEffDate), "dd/mm/yyyy")
            txtMBMPayorExpDate = Format(DateAdd("d", -1, tempdate), "dd/mm/yyyy")
            TmpEffDate = txtMBMPayorExpDate
        End If
        If txtMBMBirthdate <> "__/__/____" Then
            Call MemberCallAge(CDate(TmpEffDate), CDate(TmpDOB), txtHLTCode)
            Call MemberCallFees
            Call CallLimit
        End If
        MBMValidDate.Close
        Set MBMValidDate = Nothing
        medixmasterconn.Close
        Set medixmasterconn = Nothing
        medixmasterconnMaint.Close
        Set medixmasterconnMaint = Nothing
    
        txtTempEffDate = txtMBMPayorEffDate
        txtTempExpDate = txtMBMPayorExpDate
    End If
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
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

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
                 If CDate(txtMBMPayorEffDate) = CDate(txtTempEffDate) And CDate(txtTempExpDate) = CDate(txtMBMPayorExpDate) Then
                     txtMCOFees = txtTempMCOFees
                     txtMBMPremium = txtTempPrem
                     txtMBMProviderCharges = txtTempProvider
                 End If
                End If
            End If
        Else
            MsgBox "Invalid Date Format!", vbCritical + vbOKOnly
             txtMBMPayorExpDate.SetFocus
        End If
    End If
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    medixmasterconnMaint.Close
    Set medixmasterconnMaint = Nothing
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

Private Sub MemberCallAge(TmpEDate As String, TmpDOB As String, tmpHLTCode As String)
Dim SelectAge As String
Dim tmpPlncode As String
Dim tmpInsType As String

    tmpPlncode = Trim(cboPLNCode)
    tmpInsType = Trim(cboINSCode)
    If Trim(tmpPlncode) = "" Then
        tmpPlncode = Mid(frmMBMClinical.cboPLNCode, 1, 4)
        tmpInsType = Mid(frmMBMClinical.cboINSCode, 1, 1)
    End If
    If IsDate(TmpEDate) And IsDate(TmpDOB) And Len(Trim(tmpHLTCode)) _
        And Trim(tmpPlncode) <> "" And Trim(tmpInsType) <> "" <> 0 Then
        SelectAge = GetAge(CDate(TmpEDate), CDate(TmpDOB))
        txtAGECode = GetAgeBand(Trim(tmpHLTCode), SelectAge, tmpPlncode, tmpInsType)
        If Len(Trim(txtAGECode)) <> 0 Then
            If Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)) = "01" Then
                txtMBMAdultChild = "C"
            Else
                txtMBMAdultChild = "A"
            End If
        End If
    End If
End Sub

Public Sub MemberCallFees()
Dim PLNRec As New ADODB.Recordset
Dim SqlPln, RoundUpStatus As String
Dim tempMCOFee As Double
    If Trim(txtPAYCode) <> "" And Trim(txtAGECode) <> "" And Trim(cboPLNCode) <> "" _
        And Trim(txtHLTCode) <> "" Then
        txtMBMPremium = "0"
        txtMCOFees = "0"
'        If Len(Trim(txtPAYCode)) <> 0 Then
            txtMBMPremium = Format(GetPremium(Trim(cboINSCode), Trim(txtPAYCode), Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)), _
            Trim(cboPLNCode), Trim(txtHLTCode), txtMBMPayorEffDate), "#,##0.00")
            
            txtTempPrem = txtMBMPremium

                If PolDics > 0 Then
                    txtMBMPremium = txtMBMPremium * (100 - PolDics) / 100
                End If
                If RenDisc > 0 Then
                    txtMBMPremium = txtMBMPremium * (100 - RenDisc) / 100
                End If
'        End If
        
        txtMCOFees = Format(GetMCO(Trim(Mid(cboINSCode, 1, 1)), Trim(txtHLTCode), txtMBMAdultChild, txtMBMPayorEffDate, Trim(cboPLNCode)), "#,##0.00")
        txtTempMCOFees = txtMCOFees
        tempMCOFee = txtTempMCOFees
        SqlPln = "SELECT PLNRoundUpStatus from PLN where PlnCode='" & Trim(cboPLNCode) & _
                "' and HLTCode = '" & Trim(txtHLTCode) & _
                "' AND PLNEffDate <='" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "'"
        If Trim(txtHLTCode) = "N" Then
            SqlPln = SqlPln & " and PAYCode='" & Trim(txtPAYCode) & "'"
        End If
        PLNRec.Open SqlPln, medixmasterconn, adOpenKeyset, adLockOptimistic
        RoundUpStatus = "Y"
        If PLNRec.recordCount Then
            RoundUpStatus = IIf(Len(Trim(PLNRec!PLNRoundUpStatus)), Trim(PLNRec!PLNRoundUpStatus), "Y")
        End If
        PLNRec.Close
        Set PLNRec = Nothing
        If RoundUpStatus <> "N" Then
            tempMCOFee = Round(txtTempMCOFees)
        End If
        txtMCOFees = tempMCOFee
    End If
    Call MemberCLNCallFees
End Sub
Private Sub MemberCLNCallFees()
Dim ChkPremium As New ADODB.Recordset
Dim strsql As String
Dim tmpPlncode As String
Dim tmpInsType As String
    
    tmpPlncode = frmMBMClinical.cboPLNCode
    tmpInsType = Mid(frmMBMClinical.cboINSCode, 1, 1)
    If Trim(tmpPlncode) <> "" Then
        tmpPlncode = Trim(Mid(tmpPlncode, 1, IIf(InStr(1, tmpPlncode, "-") = 0, 4, InStr(1, tmpPlncode, "-") - 1)))
        tmpInsType = Mid(frmMBMClinical.cboINSCode, 1, 1)
    End If

    If Trim(txtAGECode) <> "" And Trim(tmpPlncode) <> "" Then
    
        frmMBMClinical.lblMBMPremium = "0"
        frmMBMClinical.lblMCOFees = "0"
        frmMBMClinical.txtTempPrem = "0.00"
        frmMBMClinical.txtTempMCOFees = "0.00"
        MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"

        strsql = " Select PLNCode, MCOInd, PRMInd From PLN Where PLNCode = '" & Trim(tmpPlncode) & "'"
        ChkPremium.Open strsql, MediConnCISDB, adOpenKeyset, adLockOptimistic

        If ChkPremium.recordCount Then

            If ChkPremium!PRMInd = "Y" Then
                frmMBMClinical.lblMBMPremium = Format(GetCLNPremium(tmpInsType, Trim(txtPAYCode), Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)), _
                Trim(tmpPlncode), Trim(txtHLTCode), CDate(frmMBMClinical.txtMBMPayorEffDate)), "#,##0.00")
                frmMBMClinical.txtTempPrem = Format(GetCLNPremium(tmpInsType, Trim(txtPAYCode), Trim(Mid(txtAGECode, 1, InStr(1, txtAGECode, "-") - 1)), _
                Trim(tmpPlncode), Trim(txtHLTCode), CDate(frmMBMClinical.txtMBMPayorEffDate)), "#,##0.00")
            End If
            
            If ChkPremium!MCOInd = "Y" Then
                frmMBMClinical.txtTempMCOFees = Format(GetCLNMCO(tmpInsType, Trim(txtHLTCode), txtMBMAdultChild, CDate(frmMBMClinical.txtMBMPayorEffDate), Trim(tmpPlncode)), "#,##0.00")
                frmMBMClinical.lblMCOFees = Format(GetCLNMCO(tmpInsType, Trim(txtHLTCode), txtMBMAdultChild, CDate(frmMBMClinical.txtMBMPayorEffDate), Trim(tmpPlncode)), "#,##0.00")
                frmMBMClinical.lblMCOFees = Round(InsertMCOFees(CDate(frmMBMClinical.txtMBMPayorExpDate), CDate(frmMBMClinical.txtMBMPayorEffDate), Format(frmMBMClinical.lblMCOFees, "#,##0.00")), 0)

            End If
        End If
        ChkPremium.Close
        Set ChkPremium = Nothing
        MediConnCISDB.Close
        Set MediConnCISDB = Nothing
            
    End If
        
End Sub
Private Sub ProviderCallFees()
    Dim GetProString As String
    
    If Trim(cboINSCode) <> "" And Trim(txtHLTCode) <> "" Then
        txtMBMProviderCharges = ""
        GetProString = GetProviderCharges("ME", Trim(cboINSCode), Trim(txtHLTCode), txtMBMAdultChild)

        txtMBMProviderCharges = Format(GetProString, "#,##0.00")
    Else
        txtMBMProviderCharges = 0
    End If
      txtTempProvider = txtMBMProviderCharges
End Sub

Private Sub CallLimit()
Dim tmpKidneyLmt As Double
Dim tmpCancerLmt As Double
Dim tmpHomeNurseLmt As Double
Dim tmpSuppBNFStatus As String
    tmpKidneyLmt = 0
    tmpCancerLmt = 0
    tmpHomeNurseLmt = 0
    txtAnnualLimit = ""
    If Len(Trim(cboPLNCode)) And IsDate(txtMBMPayorEffDate) Then
        txtAnnualLimit = Format(GetAnnualLimit(Trim(cboINSCode), Trim(cboPLNCode), Trim(txtHLTCode), CDate(txtMBMPayorEffDate)), "#,###,##0.00")
'        If Trim(CboGuaRenewal) <> "" And TempLifeTimeStatus = "Y" Then
        If TempLifeTimeStatus = "Y" Then
            Call GetBenefitLimit(Trim(cboINSCode), Trim(cboPLNCode), Trim(txtHLTCode), CDate(txtMBMPayorEffDate), tmpKidneyLmt, tmpCancerLmt, tmpHomeNurseLmt, tmpSuppBNFStatus)
            lblKidneyLmt = Format(tmpKidneyLmt, "#,##0.00")
            lblCancerLmt = Format(tmpCancerLmt, "#,##0.00")
            lblHomeNurseLmt = Format(tmpHomeNurseLmt, "#,##0.00")
        End If
        Call PrincipalLifeTimeChecking
    End If

End Sub

' ######################### END CALCULATIOn###########

' ######################### BUTTON  CLICKING ###########
Public Sub cmdShow_Click()
    'On Error Resume Next
    Dim sqlstr As String
    Dim MBM As New ADODB.Recordset
    Dim sqlstrMBMTwo As String
    Dim MBMTwo As New ADODB.Recordset
    Dim TmpEffDate As String
    
    Screen.MousePointer = vbHourglass
    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    
    sql = "Exec SearchMBMCovPersonsProc '" & Trim(txtMBMPrevMEMNo) & "'"
    Set ADOrstMBM = New ADODB.Recordset
    ADOrstMBM.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If Not ADOrstMBM.EOF Then
    
    With ADOrstMBM
        txtMBMName = Trim(!MBMName)
        txtMBMIcBcPP = Trim(!MBMICBCPP)
        txtMBMPrevPolNo = Trim(!MBMPolicyNo)
        If Len(!INSCode) Then
            cboINSCode = Trim(!INSCode)
        End If
        If Len(Trim(!PLNCode)) Then
            cboPLNCode = Trim(!PLNCode)
        End If
        If Len(Trim(!MBMAddress1)) Then
            txtMBMAdd1 = Trim(!MBMAddress1)
        End If
        If Len(Trim(!MBMAddress2)) Then
            txtMBMAdd2 = Trim(!MBMAddress2)
        End If
        If Len(Trim(!MBMAddress3)) Then
            txtMBMAdd3 = Trim(!MBMAddress3)
        End If
        If Len(Trim(!MBMPostCode)) Then
            txtMBMPostCode = Trim(!MBMPostCode)
        End If
        If Len(Trim(!CTYCode)) Then
            cboCTYCode = Trim(!CTYCode)
        End If
        If Len(Trim(!MBMSex)) Then
            cboMBMSex = Trim(!MBMSex)
        End If
        
        If Len(Trim(!MBMBranch)) Then
            cboBRNCode = Trim(!MBMBranch)
        End If
        If Len(Trim(!MBMDOB)) Then
            txtMBMBirthdate.mask = ""
            txtMBMBirthdate = Trim(!MBMDOB)
            txtMBMBirthdate.mask = "##/##/####"
        End If
        If Len(Trim(!RACCode)) Then
            cboRACCode = Trim(!RACCode)
        End If
        If Len(Trim(!STACode)) Then
            cboSTACode = Trim(!STACode)
        End If
        
        If Len(Trim(!MBMTakeOver)) Then
            cboMBMTakeover.Enabled = True
            cboMBMTakeover = Trim(!MBMTakeOver)
        End If
        
        If !MBMPolicyYear <> "" Then
            txtMBMPolicyYear = IIf(Trim(!MBMPolicyYear) = "", 1, !MBMPolicyYear + 1)
        Else
            txtMBMPolicyYear = 1
        End If
        If Trim(!HSClnInd) = "C" Or Trim(!HSClnInd) = "M" Or Trim(!HSClnInd) = "B" Then
            Call showClinicalInfo
        End If
    End With
        Call showMBMTwo
        Call showMBMOthers
        Call showMBMOthersTwo
        If Trim(Mid(cboINSCode, 1, 1)) = "F" Or Trim(Mid(cboINSCode, 1, 1)) = "H" Then
            If Len(Trim(ADOrstMBM!MBMCCoverID)) And Trim(ADOrstMBM!MBMCCoverID) <> "00" Then
                Call showMBMCov
            End If
        End If
        Call CallLimit
        Call MemberCallFees
        Call cboPLNCode_LostFocus
    Else
        MsgBox "No record found!", vbCritical + vbOKOnly
    End If
    ADOrstMBM.Close
    Set ADOrstMBM = Nothing
    medixmasterconn.Close
    Set medixmasterconn = Nothing
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

Private Sub cmdClear_Click()
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
    If frmNewMBMOthers.Visible = False Then
        Load frmNewMBMOthers
    End If
    If frmMBMClinical.Visible = True Then
        Unload frmMBMClinical
    End If
    Call ClearForm(frmNewMBMOthers)
    Unload frmNewMemberRegistration
    Set frmNewMemberRegistration = Nothing

End Sub


Private Sub cmdSave_Click()
'Dim MBMLifeRec As New ADODB.Recordset
'Dim LifeTimePss As Boolean
'Dim StrSqlLife As String

    Dim RecordSaved

    If cboMBMRenewal = "R" And _
       Len(Trim(txtMBMPrevMEMNo)) = 0 And _
       Len(Trim(txtMBMPrevPolNo)) = 0 Then
        MsgBox "Please Enter Previous Membership No or Previous Policy No", vbOKOnly + vbCritical, DialogBoxTitle
        cmdShow.SetFocus
    Else
'        LifeTimePss = True
'        StrSqlLife = "SELECT MBMNumber, MBMLifeTimeLmt from View_MBMLifeTime where MBMPolicyNo='" & Trim(txtMBMPolicyNo) & "' AND MBMCovID='00'"
'        medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"
'
'        MBMLifeRec.Open StrSqlLife, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
'        If MBMLifeRec.recordCount < 1 Then
'            If Mid(CboGuaRenewal, 1, 1) <> "3" Then
'                LifeTimePss = False
'            End If
'        End If
'        MBMLifeRec.Close
'        Set MBMLifeRec = Nothing
'        medixmasterconnMaint.Close
'        Set medixmasterconnMaint = Nothing
        If Len(Trim(txtMBMName)) = 0 Then
            MsgBox "Please Enter Membership Name", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMName.SetFocus
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
        ElseIf IsDate(txtMBMBirthdate) = False Then
            MsgBox "Invalid BirthDate", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMBirthdate = "__/__/____"
            txtMBMBirthdate.SetFocus
        ElseIf Len(Trim(txtAGECode)) = 0 Then
            MsgBox "Invalid Age Group", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf Trim(cboPLNCode.Text) = "" And Trim(frmMBMClinical.cboPLNCode) = "" Then
            MsgBox "Please Enter Plan Code or Clinical Plan", vbOKOnly + vbCritical, DialogBoxTitle
            cboPLNCode.SetFocus
        ElseIf Trim(TempLifeTimeStatus) = "Y" And Trim(CboGuaRenewal) = "" Then
            MsgBox "Please Enter Guarantee Renewal for 1st year Policy of Life Time ", vbOKOnly + vbCritical, DialogBoxTitle
' ------------ HIS Verification
        ElseIf Trim(cboPLNCode.Text) <> "" And Len(Trim(txtMBMPolicyNo)) = 0 Then
            MsgBox "Please Enter Policy No", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMPolicyNo.SetFocus
        ElseIf Trim(cboPLNCode.Text) <> "" And IsDate(txtMBMPayorEffDate) = False Then
            MsgBox "Please Enter Payor Effective Date", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMPayorEffDate.SetFocus
        ElseIf Trim(cboPLNCode.Text) <> "" And IsDate(txtMBMPayorExpDate) = False Then
            MsgBox "Please Enter Payor Expiry Date", vbOKOnly + vbCritical, DialogBoxTitle
            txtMBMPayorExpDate.SetFocus
        ElseIf Trim(cboPLNCode.Text) <> "" And Trim(cboINSCode.Text) = "" Then
            MsgBox "Please Enter Insured Code", vbOKOnly + vbCritical, DialogBoxTitle
            cboINSCode.SetFocus
        ElseIf Trim(cboINSCode.Text) = "" And Trim(frmMBMClinical.cboINSCode.Text) = "" Then
            MsgBox "Please Enter Insured Code OR Clinical Insured Code ", vbOKOnly + vbCritical, DialogBoxTitle
            cboINSCode.SetFocus
        Else
            RecordSaved = MsgBox("Are you sure?", vbYesNo + vbExclamation)
              If RecordSaved = vbYes Then
                  Screen.MousePointer = vbHourglass
                    medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
                    medixmasterconnMaint.Open "File Name=" & BOSPath & "Conn\MediConnHISMaint.UDL"

                    Call NewRegistrationSub
                    
                    medixmasterconn.Close
                    Set medixmasterconn = Nothing
                    medixmasterconnMaint.Close
                    Set medixmasterconnMaint = Nothing
                  txtMBMPolicyYear = 1
                  Screen.MousePointer = vbDefault
              End If
        End If
    End If
    

End Sub

' ######### To save the membership data ---> cater for new registration ############3
Private Sub NewRegistrationSub()
On Error GoTo ErrorSave
Dim StrSqlLife As String
Dim MBM As New ADODB.Recordset
Dim MBM1 As New ADODB.Recordset
Dim MBMTwo As New ADODB.Recordset
Dim MBMCRef As New ADODB.Recordset
Dim MBMLifeRec As New ADODB.Recordset
Dim CLNPln As New ADODB.Recordset
Dim AnnLmt As New ADODB.Recordset

Dim strsqlMBM As String
Dim strsqlMBMTwo As String
Dim strsqlMBMCRef As String
Dim CLNPlnSql As String
Dim strsqlannlmt As String

Dim MBMAvailLmt As String
Dim TempDOB As String
Dim TempBordxDate As String
Dim TempValueDate As String
Dim TempMBMAdd2 As String
Dim TempMBMName As String
Dim SuppLmtStatus As String
Dim HisCisEffDate, HISCISExpDate As String
Dim HISCisPLN, HISCiSIns As String
Dim TmpMemberType, TmpHSClnInd, TmpDCStatus As String
Dim tmpKidneyAnnLmt As Double
Dim tmpCancerAnnLmt As Double
Dim tmpHomeNursAnnLmt As Double
Dim tmpSuppBNFStatus  As String


    strsqlPLN = "select * from DiscPlan where PLNCode = '" & Trim(Mid(cboPLNCode, 1, 4)) & "'"
    PLN.Open strsqlPLN, medixmasterconnMaint, adOpenKeyset, adLockOptimistic
    If PLN.recordCount Then
        Disc = True
    End If
    
    PLN.Close
    Set PLN = Nothing
    
    ' INsert into Membership table
    ' Begin transaction
    MBMAvailLmt = CDbl(IIf(IsNumeric(txtAnnualLimit) = True, txtAnnualLimit, 0))
    TempDOB = Format(txtMBMBirthdate, "YYYY/MM/DD")
    TempBordxDate = Format(txtMBMBordxDate, "YYYY/MM/DD")
    TempValueDate = Format(Date, "YYYY/MM/DD")
    TempPayEffDate = Format(txtMBMPayorEffDate, "YYYY/MM/DD")
    TempPayExpDate = Format(txtMBMPayorExpDate, "YYYY/MM/DD")
    TempMBMAdd2 = Replace(txtMBMAdd2, "'", "''")
    TempMBMName = Replace(txtMBMName, "'", "''")
    TmpDCStatus = ""
    CISDCStatus = ""
    CISDCPayor = Mid(txtPAYCode, 1, 2)
    MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"

    If Trim(frmMBMClinical.cboPLNCode) <> "" Then

        CLNPlnSql = "SELECT ANNLmtInd, DCStatus, PAYCode, PDentalStatus, SDentalStatus," & _
                    "PSpecStatus, SSpecStatus, PMepStatus, SMepStatus " & _
                    " from PLN where PLNCode='" & Mid(frmMBMClinical.cboPLNCode, 1, 4) & _
                    "' AND HLTCode='" & Trim(txtHLTCode) & "'"
        CLNPln.Open CLNPlnSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
        If CLNPln.recordCount Then
            TmpDCStatus = IIf(Len(Trim(CLNPln!DCStatus)), Trim(CLNPln!DCStatus), "")
            CISDCStatus = IIf(Len(Trim(CLNPln!DCStatus)), Trim(CLNPln!DCStatus), "")
            CISDCPayor = IIf(Len(Trim(CLNPln!PAYCode)), Trim(CLNPln!PAYCode), "")
            CISPDentStatus = IIf(Len(Trim(CLNPln!PDentalStatus)), Trim(CLNPln!PDentalStatus), "")
            CISSDentStatus = IIf(Len(Trim(CLNPln!SDentalStatus)), Trim(CLNPln!SDentalStatus), "")
            CISPSpecStatus = IIf(Len(Trim(CLNPln!PSpecStatus)), Trim(CLNPln!PSpecStatus), "")
            CISSSpecStatus = IIf(Len(Trim(CLNPln!SSpecStatus)), Trim(CLNPln!SSpecStatus), "")
            CISPMepStatus = IIf(Len(Trim(CLNPln!PMepStatus)), Trim(CLNPln!PMepStatus), "")
            CISSMepStatus = IIf(Len(Trim(CLNPln!SMepStatus)), Trim(CLNPln!SMepStatus), "")

        End If
        CLNPln.Close
        Set CLNPln = Nothing
    End If
    
    TmpHSClnInd = "H"
    If Trim(cboPLNCode.Text) <> "" And Trim(frmMBMClinical.cboPLNCode) = "" Then
        TmpHSClnInd = "H"
    End If
    If Trim(cboPLNCode.Text) = "" And Trim(frmMBMClinical.cboPLNCode) <> "" Then
        TmpHSClnInd = "C"
    End If
    If Trim(cboPLNCode.Text) <> "" And Trim(frmMBMClinical.cboPLNCode) <> "" Then
        If Trim(TmpDCStatus) = "Y" Then
            TmpHSClnInd = "M"
        Else
            TmpHSClnInd = "B"
        End If
    End If
    CISDCStatus = TmpDCStatus
    HISCisPLN = Trim(cboPLNCode.Text)
    HISCiSIns = Trim(cboINSCode)
    HisCisEffDate = txtMBMPayorEffDate
    HISCISExpDate = txtMBMPayorExpDate
        
    If Trim(cboPLNCode.Text) = "" Then
        HisCisEffDate = frmMBMClinical.txtMBMPayorEffDate
        HISCISExpDate = frmMBMClinical.txtMBMPayorExpDate
        HISCisPLN = Trim(frmMBMClinical.cboPLNCode)
        HISCiSIns = Trim(frmMBMClinical.cboINSCode)
    End If

    SuppLmtStatus = "A"
    TmpMemberType = "P"
    If Trim(cboPLNCode.Text) <> "" Then
        strsqlannlmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(HISCisPLN) & _
                        "' AND HLTCode='" & Trim(txtHLTCode) & "' AND INSCode='" & Trim(HISCiSIns) & "'"
        AnnLmt.Open strsqlannlmt, medixmasterconn, adOpenKeyset, adLockOptimistic
    Else
        strsqlannlmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(HISCisPLN) & _
                       "' AND HLTCode='" & Trim(txtHLTCode) & "' AND INSCode='" & Trim(HISCiSIns) & "'"
        AnnLmt.Open strsqlannlmt, MediConnCISDB, adOpenKeyset, adLockOptimistic
    End If
    If AnnLmt.recordCount > 0 Then
        If Trim(AnnLmt!SuppLimitStatus) <> "" Then
            SuppLmtStatus = Trim(AnnLmt!SuppLimitStatus)
        End If
    End If
    AnnLmt.Close
    Set AnnLmt = Nothing
    If Trim(SuppLmtStatus) <> "A" Then
        TmpMemberType = "Q"
    End If
    MediConnCISDB.Close
    Set MediConnCISDB = Nothing
    startTrans = True
    medixmasterconn.beginTrans
    medixmasterconnMaint.beginTrans
    If Trim(cboPLNCode) = "" And Trim(frmMBMClinical.cboPLNCode) <> "" Then
        MembershipNo = Trim(txtPAYCode) & Trim(Mid(frmMBMClinical.cboPLNCode, 1, IIf(InStr(1, frmMBMClinical.cboPLNCode, "-") = 0, 4, InStr(1, frmMBMClinical.cboPLNCode, "-") - 1))) & Trim(Mid(frmMBMClinical.cboINSCode, 1, 1)) & GenerateMembershipNo("DC", Trim(Mid(frmMBMClinical.cboPLNCode, 1, IIf(InStr(1, frmMBMClinical.cboPLNCode, "-") = 0, 4, InStr(1, frmMBMClinical.cboPLNCode, "-") - 1))), Trim(Mid(frmMBMClinical.cboINSCode, 1, 1)))
    Else
        MembershipNo = Trim(txtPAYCode) & Trim(cboPLNCode) & Trim(cboINSCode) & GenerateMembershipNo(Trim(txtPAYCode), Trim(cboPLNCode), Trim(cboINSCode))
    End If
    MembershipNo = ChkStr(MembershipNo)
    TmpTopupNo = ""
    WLifeStatus = False
    PrevLifeMBMNo = ""
 
    tmpKidneyAnnLmt = 0
    tmpCancerAnnLmt = 0
    tmpHomeNursAnnLmt = 0
    If Trim(TempLifeTimeStatus) = "Y" Then
        Call GetBenefitLimit(Trim(cboINSCode), Trim(cboPLNCode), Trim(txtHLTCode), CDate(txtMBMPayorEffDate), tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt, tmpSuppBNFStatus)
        Call SaveMBMBnfAnnLmt(MembershipNo, "00", "A", tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt)
    End If
    Call PrincipalLifeTimeChecking 'LifeTimeLimit & BenefitLimit
    If WLifeStatus = True Then
        Call SaveMBMLifeTimeLmt(TmpTopupNo, "00", "A")
    End If

    If Trim(cboPLNCode) <> "" Then
        strsqlMBM = "insert into MBM (MBMNumber, MBMCOVID, MBMPolicyNo, " & _
                    "MBMIcBcPp, MBMName, HLTCode, PAYCode, MBMAddress1, " & _
                    "MBMAddress2, MBMAddress3, CTYCode, MBMPostCode, MBMDOB, " & _
                    "RACCode, MBMSex, AgeCode, MBMAdultChild, STACode, MBMBordxDate, " & _
                    "MBMTakeOver, MBMRenewFlag,  MBMStatus, MBMValueDate, MBMPolicyYear, " & _
                    "MBMExpiredStatus, MBMBatchNo, MBMDataStatus,MBMPayorEffDate, MBMPayorEXPDate, " & _
                    "INSCode, PLNCode,MBMAvailableLimit,MBMMemberType, HSCLNInd,MBMTopupNumber) " & _
                    "Values('" & MembershipNo & "', '00', '" & txtMBMPolicyNo & "', " & _
                    "'" & txtMBMIcBcPP & "', '" & TempMBMName & "', '" & Trim(Mid(Trim(txtHLTCode), 1, 1)) & "', '" & Trim(Mid(txtPAYCode, 1, 2)) & "', '" & txtMBMAdd1 & "', " & _
                    "'" & TempMBMAdd2 & "', '" & txtMBMAdd3 & "', '" & cboCTYCode & "', '" & txtMBMPostCode & "', '" & TempDOB & "', " & _
                    "'" & Trim(Mid(cboRACCode, 1, 1)) & "', '" & cboMBMSex & "', '" & Trim(Mid(txtAGECode, 1, 2)) & "', '" & txtMBMAdultChild & "', '" & Trim(cboSTACode) & "', '" & TempBordxDate & "', " & _
                    "'" & Mid(cboMBMTakeover, 1, 1) & "', '" & cboMBMRenewal & "', 'A', '" & TempValueDate & "', '" & txtMBMPolicyYear & "', " & _
                    "'N', '" & txtMBMBatchNo & "', 'A', '" & TempPayEffDate & "', '" & TempPayExpDate & "', " & _
                    "'" & Trim(Mid(cboINSCode, 1, 1)) & "', '" & Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") > 0, InStr(1, cboPLNCode, "-") - 1, 4))) & "' , " & MBMAvailLmt & "," & _
                    "'" & TmpMemberType & "','" & TmpHSClnInd & "','" & TmpTopupNo & "')"
    Else
        strsqlMBM = "insert into MBM (MBMNumber, MBMCOVID, MBMPolicyNo, " & _
                    "MBMIcBcPp, MBMName, HLTCode, PAYCode, MBMAddress1, " & _
                    "MBMAddress2, MBMAddress3, CTYCode, MBMPostCode, MBMDOB, " & _
                    "RACCode, MBMSex, AgeCode, MBMAdultChild, STACode, MBMBordxDate, " & _
                    "MBMTakeOver, MBMRenewFlag,  MBMStatus, MBMValueDate, MBMPolicyYear, " & _
                    "MBMExpiredStatus, MBMBatchNo, MBMDataStatus,MBMAvailableLimit,MBMMemberType, HSCLNInd,MBMTopupNumber) " & _
                    "Values('" & MembershipNo & "', '00', '" & txtMBMPolicyNo & "', " & _
                    "'" & txtMBMIcBcPP & "', '" & TempMBMName & "', '" & Mid(txtHLTCode, 1, 1) & "', '" & Mid(txtPAYCode, 1, 2) & "', '" & txtMBMAdd1 & "', " & _
                    "'" & txtMBMAdd2 & "', '" & txtMBMAdd3 & "', '" & cboCTYCode & "', '" & txtMBMPostCode & "', '" & TempDOB & "', " & _
                    "'" & Trim(Mid(cboRACCode, 1, 1)) & "', '" & cboMBMSex & "', '" & Trim(Mid(txtAGECode, 1, 2)) & "', '" & txtMBMAdultChild & "', '" & Trim(cboSTACode) & "', '" & TempBordxDate & "', " & _
                    "'" & Mid(cboMBMTakeover, 1, 1) & "', '" & cboMBMRenewal & "', 'A', '" & TempValueDate & "', '" & txtMBMPolicyYear & "', " & _
                    "'N', '" & txtMBMBatchNo & "', 'A'," & MBMAvailLmt & ",'" & TmpMemberType & "','" & TmpHSClnInd & "','" & TmpTopupNo & "')"
    End If

    medixmasterconn.Execute strsqlMBM

    completeMBM = True
    
    strsqlMBMTwo = "insert into MBMTwo (MBMNumber, MBMSalutation, MBMTelnoH, " & _
                "MBMTelNoM, MBMTelNoO, MBMEmail, " & _
                "NATCode, SRVCodeList, LGNCode, MBMRegUser, MBMLastUpdateUser, MBMUploadStatus) " & _
                "Values('" & MembershipNo & "', '" & cboMBMSalutation & "' , '" & frmNewMBMOthers.txtMBMTelNoH & "', " & _
                "'" & frmNewMBMOthers.txtMBMTelNoM & "', '" & frmNewMBMOthers.txtMBMTelNoO & "', '" & frmNewMBMOthers.txtMBMEmail & "', " & _
                "'" & Trim(Mid(cboNATCode, 1, 2)) & "', 'ME', '" & Trim(Mid(frmNewMBMOthers.cboLGNCode, 1, 1)) & "', '" & Logon & "', '" & Logon & "','M')"
    medixmasterconn.Execute strsqlMBMTwo
    completeMBMTwo = True
    
    If Trim(cboPLNCode.Text) <> "" Then
        strsqlMBMCRef = "insert into MBMCrossReference (MBMNumber, MBMCOVID, MBMPolicyNo, " & _
                    "MBMIcBcPp, MBMName, HLTCode, PAYCode, MBMDOB," & _
                    "MBMStatus, MBMDataStatus,MBMBordxDate,AgeCode," & _
                    "MBMPayorEffDate, MBMPayorEXPDate, MBMBatchNo, " & _
                    "INSCode, PLNCode, MBMTopupNumber) " & _
                    "Values('" & MembershipNo & "', '00', '" & txtMBMPolicyNo & "', " & _
                    "'" & txtMBMIcBcPP & "', '" & TempMBMName & "', '" & txtHLTCode & "', '" & txtPAYCode & "','" & TempDOB & "', " & _
                    "'A', 'A', '" & TempBordxDate & "', '" & Trim(Mid(txtAGECode, 1, 2)) & "', " & _
                    "'" & TempPayEffDate & "', '" & TempPayExpDate & "', '" & txtMBMBatchNo & "', " & _
                    "'" & Trim(Mid(cboINSCode, 1, 1)) & "', '" & Trim(Mid(cboPLNCode, 1, IIf(InStr(1, cboPLNCode, "-") > 0, InStr(1, cboPLNCode, "-") - 1, 4))) & "'," & _
                    "'" & Trim(TmpTopupNo) & "')"
                                
    Else
        strsqlMBMCRef = "insert into MBMCrossReference (MBMNumber, MBMCOVID, MBMPolicyNo, MBMBatchNo," & _
                    "MBMIcBcPp, MBMName, HLTCode, PAYCode, MBMDOB," & _
                    "MBMStatus, MBMDataStatus,MBMBordxDate,AgeCode, MBMTopupNumber)" & _
                    "Values('" & MembershipNo & "', '00', '" & txtMBMPolicyNo & "', '" & txtMBMBatchNo & "', " & _
                    "'" & txtMBMIcBcPP & "', '" & TempMBMName & "', '" & txtHLTCode & "', '" & txtPAYCode & "','" & TempDOB & "', " & _
                    "'A', 'A', '" & TempBordxDate & "', '" & Trim(Mid(txtAGECode, 1, 2)) & "'," & _
                    "'" & Trim(TmpTopupNo) & "')"
    End If
    medixmasterconnMaint.Execute strsqlMBMCRef
    completeMBMCRef = True
    
    Call SaveMBMOthers(MembershipNo)
    
    Call SaveMBMOthersTwo(MembershipNo)
    
    If cboINSCode = "F" Or cboINSCode = "H" Or Mid(frmMBMClinical.cboINSCode, 1, 1) = "F" _
        Or Mid(frmMBMClinical.cboINSCode, 1, 1) = "H" Then
            Call saveCoverPersons(MembershipNo)
        End If
    If Trim(frmMBMClinical.cboPLNCode) <> "" Then
        Call SaveMBMCLNOthers(MembershipNo)
    End If
    
    '========Clinical=======================
    
    Call SaveBatch
    
    medixmasterconn.CommitTrans
    medixmasterconnMaint.CommitTrans
    startTrans = False
     
     cboMBMRenewal.Enabled = True
     MsgBox "Record Saved, Your Membership Number is : " & MembershipNo, vbOKOnly + vbInformation, DialogBoxTitle
     ' Clering previous record
     txtMBMPayorExpDate.Enabled = True
     Call ClearEntries
   
     SSTab1.Tab = 0
     cboMBMRenewal.SetFocus
    Exit Sub
ErrorSave:
    Screen.MousePointer = vbDefault
    If startTrans = True Then
        medixmasterconn.RollbackTrans
        medixmasterconnMaint.RollbackTrans
    End If
    MsgBox Err.Description
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

Private Sub saveCoverPersons(MembershipNo As String)   ' , MBMtype As String)
Dim MBMCoveredPersons As New ADODB.Recordset
Dim MBMCoveredPersonsTwo As New ADODB.Recordset
Dim listrecord As Integer
Dim listloop As Integer
Dim SelectAge As String
Dim strsqlannlmt As String
Dim strsqlPrem As String
Dim strsqlMBMCov As String
Dim strsqlMBMCovTwo As String
Dim Prem As New ADODB.Recordset
Dim AnnLmt As New ADODB.Recordset
Dim MemberDOB As String
Dim FindAge As String
Dim AgeValue As String
Dim tempSPremium As String
Dim tempDPremium As String
Dim tempSAnnLmt As String
Dim CoverID As String
Dim CovMBMType As String
Dim MBMCovAnnLmt As Double
Dim MBMCovAvaLmt As Double
Dim TmpMBMCRBAmt As Double
Dim TmpMBMCPayPremGross As Double
Dim TmpMBMCPayPremNet As Double
Dim TmpMBMCPayMCOGross As Double
Dim TmpMBMCPayMCONet As Double
Dim TmpMBMCFamDiscAmt As Double
Dim TmpMBMCRenewDiscAmt As Double
Dim TmpMBMCLoadingPrem As Double
Dim tempMBMCPolicyDisc As Double
Dim tempMBMCRenewalDisc  As Double
Dim TempMBMCDOB As String
Dim TempMBMCBordxDate As String
Dim TempMBMCExlcusion As String
Dim TempMBMCName As String
Dim TempSuppExpDate As String
Dim SuppPlanCode As String
Dim SptLmtInd As String
Dim TmpPlnRec As New ADODB.Recordset
Dim TmpPlnSql As String
Dim tmpPlncode As String
Dim tmpPLNLifeTimeStatus, TmpSuppPrmStatus As String
Dim TmpSuppLmt, TmpSuppPrmAmt As String
Dim TmpSql2 As String
Dim TmpMBM As New ADODB.Recordset
Dim strSptLmt As String
Dim SptPlnLmt As New ADODB.Recordset
Dim strSptMCOSql As String
Dim SptMCOInd As String
Dim McoCovIDChrg As String
Dim SptPlnMCO As New ADODB.Recordset
Dim tmpUndExcess As String
Dim TmpBasicMCO, tmpMCO As String
Dim tmpInsType As String
Dim TmpAdultChild As String
Dim tmpKidneyAnnLmt As Double
Dim tmpCancerAnnLmt As Double
Dim tmpHomeNursAnnLmt As Double
Dim tmpSuppBNFStatus  As String
Dim tmpPLNRoundUpStatus As String
    listrecord = lstCovAdd.listitems.count
    ' Insert into Membership Covered Person's table

    tmpPlncode = cboPLNCode
    tmpInsType = cboINSCode
    If Trim(tmpPlncode) = "" Then
        tmpPlncode = frmMBMClinical.cboPLNCode
        tmpInsType = Mid(frmMBMClinical.cboINSCode, 1, 1)
    End If

     If listrecord <> 0 Then
         For listloop = 1 To listrecord
            If Trim(cboPLNCode) <> "" Then
                SelectAge = GetAge(txtMBMPayorEffDate, lstCovAdd.listitems(listloop).SubItems(3))
                AgeValue = Mid(GetAgeBand(Trim(txtHLTCode), SelectAge, tmpPlncode, tmpInsType), 1, 2)
            ElseIf Trim(frmMBMClinical.cboPLNCode) <> "" Then
                SelectAge = GetAge(frmMBMClinical.txtMBMPayorEffDate, lstCovAdd.listitems(listloop).SubItems(3))
                AgeValue = Mid(GetAgeBand(Trim(txtHLTCode), SelectAge, tmpPlncode, tmpInsType), 1, 2)
            End If
            TmpAdultChild = "A"
            If AgeValue = "01" Then
                TmpAdultChild = "C"
            End If

            CoverID = Format(listloop, "00")
            CovMBMType = "C"
            SptLmtInd = "A"
            tmpPLNLifeTimeStatus = "N"
            TempMBMCName = Replace(lstCovAdd.listitems(listloop).Text, "'", "''")
            If Trim(cboPLNCode) <> "" Then
                tmpPlncode = Trim(cboPLNCode)
                If Len(Trim(lstCovAdd.listitems(listloop).SubItems(41))) Then
                    tmpPlncode = Trim(lstCovAdd.listitems(listloop).SubItems(41))
                End If
                If Trim(txtHLTCode) = "S" Then
                    TmpPlnSql = "Select PLNLifeTimeStatus, PLNRoundUpStatus from PLN where PLNCode='" & Trim(tmpPlncode) & _
                                "' AND HLTCode='" & Trim(txtHLTCode) & _
                                "' AND PLNEffDate <= '" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' order by PLNEffDate desc"
                Else
                    TmpPlnSql = "Select PLNLifeTimeStatus, PLNRoundUpStatus from PLN where PLNCode='" & Trim(tmpPlncode) & _
                                "' AND HLTCode='" & Trim(txtHLTCode) & "' AND PAYCode='" & Trim(txtPAYCode) & _
                                "' AND PLNEffDate <= '" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' order by PLNEffDate desc"
                End If
                tmpPLNLifeTimeStatus = "N"
                tmpPLNRoundUpStatus = "Y"
                TmpPlnRec.Open TmpPlnSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                If TmpPlnRec.recordCount Then
                    tmpPLNLifeTimeStatus = IIf(Len(Trim(TmpPlnRec!PLNLifeTimeStatus)), Trim(TmpPlnRec!PLNLifeTimeStatus), "N")
                    tmpPLNRoundUpStatus = IIf(Len(Trim(TmpPlnRec!PLNRoundUpStatus)), Trim(TmpPlnRec!PLNRoundUpStatus), "Y")
                End If
                TmpPlnRec.Close
                Set TmpPlnRec = Nothing
                
                SptLmtInd = "A"
                strSptLmt = "Select SuppLimitStatus, SuppLimit from AnnualLimit where PLNCode = '" & Trim(tmpPlncode) & _
                    "' AND HLTCode='" & Trim(txtHLTCode) & "' AND INSCode='" & Trim(cboINSCode) & _
                    "' AND AnnualEffDate <= '" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' order by AnnualEffDate desc"
                SptPlnLmt.Open strSptLmt, medixmasterconn, adOpenKeyset, adLockOptimistic
                TmpSuppLmt = 0
                If SptPlnLmt.recordCount > 0 Then
                    TmpSuppLmt = IIf(IsNumeric(SptPlnLmt!SuppLimit), SptPlnLmt!SuppLimit, 0)
                    If Trim(SptPlnLmt!SuppLimitStatus) <> "" Then
                        SptLmtInd = Trim(SptPlnLmt!SuppLimitStatus)
                    End If
                End If
                SptPlnLmt.Close
                Set SptPlnLmt = Nothing
                If Trim(SptLmtInd) = "B" And Trim(CoverID) = "01" Or Trim(SptLmtInd) = "C" Then
                    MBMCovAnnLmt = CDbl(TmpSuppLmt)
                    MBMCovAvaLmt = CDbl(TmpSuppLmt)
                End If  ' AnnLmimit Calculation
'------------   Life Time Limit
                If Trim(CboGuaRenewal) <> "" Then
                    If Trim(TempLifeTimeStatus) = "Y" Then
                        tmpKidneyAnnLmt = 0
                        tmpCancerAnnLmt = 0
                        tmpHomeNursAnnLmt = 0
                        Call GetBenefitLimit(Trim(cboINSCode), Trim(cboPLNCode), Trim(txtHLTCode), CDate(txtMBMPayorEffDate), tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt, tmpSuppBNFStatus)
                        If (tmpSuppBNFStatus = "B" And CoverID = "01") Or tmpSuppBNFStatus = "C" Then
                            Call SaveMBMBnfAnnLmt(MembershipNo, CoverID, tmpSuppBNFStatus, tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt)
                        End If
                        If WLifeStatus Then
                            Call SaveMBMLifeTimeLmt(PrevLifeMBMNo, CoverID, SptLmtInd)
                        Else
                            TmpSql2 = "Select top 1 MBMCNumber,MBMCAvailableLimit From MBMCoveredPersons where MBMCNumber='" & Trim(PrevLifeMBMNo) & _
                                     "' AND MBMCCoverID='" & Trim(CoverID) & "'"
                            TmpMBM.Open TmpSql2, medixmasterconn, adOpenForwardOnly, adLockReadOnly
                            If TmpMBM.EOF = False Then
                                If Mid(CboGuaRenewal, 1, 1) <> "3" Then
                                    MBMCovAnnLmt = IIf(IsNumeric(TmpMBM!MBMCAvailableLimit), TmpMBM!MBMCAvailableLimit, 0)
                                    MBMCovAvaLmt = IIf(IsNumeric(TmpMBM!MBMCAvailableLimit), TmpMBM!MBMCAvailableLimit, 0)
                                End If
                            End If
                            TmpMBM.Close
                            Set TmpMBM = Nothing
                        End If
                    End If
                End If
' ----- End of Life Time LImit
                
                SptMCOInd = "A"
                McoCovIDChrg = ""
                strSptMCOSql = "Select SuppMCOStatus, CovIDChrg from MCO where PLNCode = '" & Trim(tmpPlncode) & _
                       "' AND HLTCode='" & Trim(txtHLTCode) & "' AND INSCode='" & Trim(cboINSCode) & _
                       "' AND MCOEffdate <= '" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"

                SptPlnMCO.Open strSptMCOSql, medixmasterconn, adOpenKeyset, adLockOptimistic
        
                If SptPlnMCO.recordCount > 0 Then
                    If Trim(SptPlnMCO!SuppMCOStatus) <> "" Then
                        SptMCOInd = Trim(SptPlnMCO!SuppMCOStatus)
                    End If
                    If Trim(SptPlnMCO!CovIDChrg) <> "" Then
                        McoCovIDChrg = Trim(SptPlnMCO!CovIDChrg)
                    End If
                Else
                    SptPlnMCO.Close
                    strSptMCOSql = "Select SuppMCOStatus, CovIDChrg from MCO where PLNCode is null " & _
                                   " AND HLTCode='" & Trim(txtHLTCode) & "' AND INSCode='" & Trim(cboINSCode) & _
                                   "' AND MCOEffdate <= '" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' order by MCOEffDate desc"
                    SptPlnMCO.Open strSptMCOSql, medixmasterconn, adOpenKeyset, adLockOptimistic
                    If Trim(SptPlnMCO!SuppMCOStatus) <> "" Then
                        SptMCOInd = Trim(SptPlnMCO!SuppMCOStatus)
                    End If
                    If Trim(SptPlnMCO!CovIDChrg) <> "" Then
                        McoCovIDChrg = Trim(SptPlnMCO!CovIDChrg)
                    End If
                End If
                SptPlnMCO.Close
                Set SptPlnMCO = Nothing

                If (Trim(SptMCOInd) = "B" And Trim(CoverID) = "01") _
                Or Trim(SptMCOInd) = "C" _
                Or (Trim(SptMCOInd) = "D" And Trim(CoverID) >= Trim(McoCovIDChrg)) Then
                    TmpBasicMCO = Format(GetSuppMCO(Trim(cboINSCode), Trim(txtHLTCode), TmpAdultChild, CDate(txtMBMPayorEffDate), Trim(tmpPlncode)), "#,##0.00")
                    tmpMCO = InsertMCOFees(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(TmpBasicMCO, "#,##0.00"))
                    If tmpPLNRoundUpStatus <> "N" Then
                        tmpMCO = Round(tmpMCO)
                    End If
                End If

                If Trim(txtHLTCode) = "S" Then
                strsqlPrem = "Select SuppPrmStatus, PLNCode, SuppPrmAmt from PRM where " & _
                             " PLNCode = '" & Trim(tmpPlncode) & "' And AgeCode = '" & Trim(AgeValue) & _
                             "' And INSCode = '" & Mid(cboINSCode, 1, 1) & _
                             "' AND HLTCode='" & Trim(txtHLTCode) & _
                             "' AND PRMEffDate <= '" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' order by PRMEffDate desc"
                Else
                strsqlPrem = "Select SuppPrmStatus, PLNCode, SuppPrmAmt from PRM where " & _
                             " PLNCode = '" & Trim(tmpPlncode) & "' And AgeCode = '" & Trim(AgeValue) & _
                             "' And INSCode = '" & Mid(cboINSCode, 1, 1) & _
                             "' AND HLTCode='" & Trim(txtHLTCode) & "' AND PAYCode='" & Trim(txtPAYCode) & _
                             "' AND PRMEffDate <= '" & Format(txtMBMPayorEffDate, "mm/dd/yyyy") & "' order by PRMEffDate desc"
                End If
                Prem.Open strsqlPrem, medixmasterconn, adOpenKeyset, adLockOptimistic
                TmpSuppPrmStatus = "A"
                TmpSuppPrmAmt = 0
                If Prem.recordCount Then
                    TmpSuppPrmAmt = IIf(IsNumeric(Prem!SuppPrmAmt), Prem!SuppPrmAmt, 0)
                    If Trim(Prem!SuppPrmStatus) <> "" Then
                        TmpSuppPrmStatus = Prem!SuppPrmStatus
                    End If
                End If
                Prem.Close
                Set Prem = Nothing
                
                If (Trim(TmpSuppPrmStatus) = "B" And Trim(CoverID) = "01") Or Trim(TmpSuppPrmStatus) = "C" Then
                    tempSPremium = TmpSuppPrmAmt
                    tempDPremium = Format(GetSPremium(Trim(cboINSCode), Trim(txtPAYCode), AgeValue, Trim(cboPLNCode), Trim(txtHLTCode), CDate(txtMBMPayorEffDate)), "#,##0.00")
                    tempDPremium = InsertPremium(CDate(txtMBMPayorExpDate), CDate(txtMBMPayorEffDate), Format(tempSPremium, "#,##0.00"))
                         
                    If Disc = True Then
                        If lstCovAdd.listitems(listloop).SubItems(36) <> "" Then
                             tempDPremium = tempDPremium * (100 - CDbl(lstCovAdd.listitems(listloop).SubItems(36))) / 100
                        End If
                        If lstCovAdd.listitems(listloop).SubItems(37) <> "" Then
                             tempDPremium = tempDPremium * (100 - CDbl(lstCovAdd.listitems(listloop).SubItems(37))) / 100
                        End If
                    End If
                End If
            End If  ' IF PLNCode <> ""
        MBMCovAnnLmt = CDbl(IIf(IsNumeric(MBMCovAnnLmt) = True, MBMCovAnnLmt, 0))
        MBMCovAvaLmt = CDbl(IIf(IsNumeric(MBMCovAvaLmt) = True, MBMCovAvaLmt, 0))
        tempSPremium = CDbl(IIf(IsNumeric(tempSPremium) = True, tempSPremium, 0))
        tempDPremium = CDbl(IIf(IsNumeric(tempDPremium) = True, tempDPremium, 0))
        TmpBasicMCO = CDbl(IIf(IsNumeric(TmpBasicMCO) = True, TmpBasicMCO, 0))
        tmpMCO = CDbl(IIf(IsNumeric(tmpMCO) = True, tmpMCO, 0))
        tempMBMCPolicyDisc = CDbl(IIf(IsNumeric(lstCovAdd.listitems(listloop).SubItems(36)) = True, lstCovAdd.listitems(listloop).SubItems(36), 0))
        tempMBMCRenewalDisc = CDbl(IIf(IsNumeric(lstCovAdd.listitems(listloop).SubItems(37)) = True, lstCovAdd.listitems(listloop).SubItems(37), 0))
        TempMBMCDOB = Format(lstCovAdd.listitems(listloop).SubItems(3), "YYYY/MM/DD")
        TempMBMCBordxDate = Format(txtMBMBordxDate, "YYYY/MM/DD")
        SuppPlanCode = IIf(Len(lstCovAdd.listitems(listloop).SubItems(41)), lstCovAdd.listitems(listloop).SubItems(41), cboPLNCode)
        MembershipNo = MembershipNo
        If Trim(SuppPlanCode) <> "" Then
            TempSuppExpDate = GetSuppExpDate(TempMBMCDOB, SuppPlanCode, TempPayEffDate, TempPayExpDate)
        End If
        If IsDate(TempSuppExpDate) Then
        strsqlMBMCov = "insert into MBMCOveredPersons (MBMCNumber, MBMCName, MBMCIcBcPpNo, MBMCStatus, " & _
                    "MBMCDOB, MBMCAGEBand, MBMCSex, MBMCAdultChild, " & _
                    "MBMCRELCode, MBMCCoverID, MBMCMemberType, MBMCBatchNo, " & _
                    "MBMCBordxDate, MBMCPLNCode, MBMCExpDate,  " & _
                    "MBMCAnnualLimit, MBMCAvailableLimit, MBMCBasicPremium, MBMCPremium,MBMCPolicyDisc, MBMCRenewalDisc, " & _
                    "MBMCBasicMCO, MBMCMCO) " & _
                    "Values('" & MembershipNo & "', '" & TempMBMCName & "', '" & lstCovAdd.listitems(listloop).SubItems(2) & "','A'," & _
                    "'" & TempMBMCDOB & "', '" & AgeValue & "', '" & lstCovAdd.listitems(listloop).SubItems(4) & "', '" & lstCovAdd.listitems(listloop).SubItems(5) & "', " & _
                    "'" & Mid(lstCovAdd.listitems(listloop).SubItems(6), 1, 1) & "' , '" & Trim(CoverID) & "', '" & Trim(CovMBMType) & "', '" & txtMBMBatchNo & "', " & _
                    "'" & TempMBMCBordxDate & "', '" & SuppPlanCode & "', '" & Format(TempSuppExpDate, "YYYY/MM/DD") & "', " & _
                    "" & MBMCovAnnLmt & ", " & MBMCovAvaLmt & ", " & tempSPremium & ", " & tempDPremium & " , " & tempMBMCPolicyDisc & ", " & tempMBMCRenewalDisc & ", " & _
                    "" & TmpBasicMCO & "," & tmpMCO & ")"
        Else
        
        strsqlMBMCov = "insert into MBMCOveredPersons (MBMCNumber, MBMCName, MBMCIcBcPpNo, " & _
                    "MBMCDOB, MBMCAGEBand, MBMCSex, MBMCAdultChild, " & _
                    "MBMCRELCode, MBMCCoverID, MBMCMemberType, MBMCBatchNo, " & _
                    "MBMCBordxDate, MBMCPLNCode, " & _
                    "MBMCAnnualLimit, MBMCAvailableLimit, MBMCBasicPremium, MBMCPremium,MBMCPolicyDisc, MBMCRenewalDisc, " & _
                    "MBMCBasicMCO, MBMCMCO) " & _
                    "Values('" & MembershipNo & "', '" & TempMBMCName & "', '" & lstCovAdd.listitems(listloop).SubItems(2) & "', " & _
                    "'" & TempMBMCDOB & "', '" & AgeValue & "', '" & lstCovAdd.listitems(listloop).SubItems(4) & "', '" & lstCovAdd.listitems(listloop).SubItems(5) & "', " & _
                    "'" & Mid(lstCovAdd.listitems(listloop).SubItems(6), 1, 1) & "' , '" & Trim(CoverID) & "', '" & Trim(CovMBMType) & "', '" & txtMBMBatchNo & "', " & _
                    "'" & TempMBMCBordxDate & "', '" & SuppPlanCode & "', " & _
                    "" & MBMCovAnnLmt & ", " & MBMCovAvaLmt & ", " & tempSPremium & ", " & tempDPremium & " , " & tempMBMCPolicyDisc & ", " & tempMBMCRenewalDisc & "," & _
                    "" & TmpBasicMCO & "," & tmpMCO & ")"
        End If
 
        medixmasterconn.Execute strsqlMBMCov
        '---MBMCovPerTwo---'
        'MBMCoveredPersonsTwo.AddNew
        tmpUndExcess = lstCovAdd.listitems(listloop).SubItems(43)
        tmpUndExcess = CDbl(IIf(IsNumeric(tmpUndExcess) = True, tmpUndExcess, 0))

        TmpMBMCRBAmt = CDbl(IIf(IsNumeric(lstCovAdd.listitems(listloop).SubItems(19)) = True, lstCovAdd.listitems(listloop).SubItems(19), 0))
        TmpMBMCPayPremGross = CDbl(IIf(IsNumeric(lstCovAdd.listitems(listloop).SubItems(32)) = True, lstCovAdd.listitems(listloop).SubItems(32), 0))
        TmpMBMCPayPremNet = CDbl(IIf(IsNumeric(lstCovAdd.listitems(listloop).SubItems(33)) = True, lstCovAdd.listitems(listloop).SubItems(33), 0))
        TmpMBMCPayMCOGross = CDbl(IIf(IsNumeric(lstCovAdd.listitems(listloop).SubItems(34)) = True, lstCovAdd.listitems(listloop).SubItems(34), 0))
        TmpMBMCPayMCONet = CDbl(IIf(IsNumeric(lstCovAdd.listitems(listloop).SubItems(35)) = True, lstCovAdd.listitems(listloop).SubItems(35), 0))
        TmpMBMCFamDiscAmt = CDbl(IIf(IsNumeric(lstCovAdd.listitems(listloop).SubItems(38)) = True, lstCovAdd.listitems(listloop).SubItems(38), 0))
        TmpMBMCRenewDiscAmt = CDbl(IIf(IsNumeric(lstCovAdd.listitems(listloop).SubItems(39)) = True, lstCovAdd.listitems(listloop).SubItems(39), 0))
        TmpMBMCLoadingPrem = CDbl(IIf(IsNumeric(lstCovAdd.listitems(listloop).SubItems(40)) = True, lstCovAdd.listitems(listloop).SubItems(40), 0))
        TempMBMCExlcusion = Replace(lstCovAdd.listitems(listloop).SubItems(9), "'", "''")
        strsqlMBMCovTwo = "insert into MBMCOveredPersonsTwo (MBMCNumber, MBMCCoverID, MBMCSalutation, MBMCBloodType, " & _
                    "MBMCAllergic, MBMCExclusion, MBMCOccupation, MBMCWorkNature, " & _
                    "MBMCSmoking, MBMCAlcohol, MBMCPrevPLNCode, MBMCPAYRemarks, MBMCPrevInsCom,MBMCStaffDec," & _
                    "MBMCRBAmt,MBMCPayorPremGross, MBMCPayorPremNet, MBMCPayorMCOGross, MBMCPayorMCONet,MBMCFamDiscAmt,MBMCLoadingPrem,MBMCRenewDiscAmt, " & _
                    "MBMCUndExcess)" & _
                    "Values('" & MembershipNo & "', '" & CoverID & "', '" & lstCovAdd.listitems(listloop).SubItems(1) & "', '" & lstCovAdd.listitems(listloop).SubItems(7) & "', " & _
                    "'" & lstCovAdd.listitems(listloop).SubItems(8) & "', '" & TempMBMCExlcusion & "', '" & lstCovAdd.listitems(listloop).SubItems(12) & "', '" & lstCovAdd.listitems(listloop).SubItems(13) & "', " & _
                    "'" & lstCovAdd.listitems(listloop).SubItems(14) & "' , '" & lstCovAdd.listitems(listloop).SubItems(15) & "', '" & lstCovAdd.listitems(listloop).SubItems(18) & "', '" & lstCovAdd.listitems(listloop).SubItems(20) & "', '" & ChkStr(lstCovAdd.listitems(listloop).SubItems(21)) & "',  '" & lstCovAdd.listitems(listloop).SubItems(28) & "', " & _
                    "" & TmpMBMCRBAmt & ", " & TmpMBMCPayPremGross & " , " & TmpMBMCPayPremNet & " ,  " & TmpMBMCPayMCOGross & ", " & TmpMBMCPayMCONet & ", " & TmpMBMCFamDiscAmt & "," & TmpMBMCLoadingPrem & " , " & TmpMBMCRenewDiscAmt & ", " & _
                    "" & tmpUndExcess & ") "
        medixmasterconn.Execute strsqlMBMCovTwo
        If Trim(frmMBMClinical.cboPLNCode) <> "" Then
            Call saveCLNCoverPersons(MembershipNo, CoverID, lstCovAdd.listitems(listloop).SubItems(3), lstCovAdd.listitems(listloop).SubItems(31))
        End If
        Next listloop
    End If
End Sub


Private Sub saveCLNCoverPersons(MembershipNo As String, TmpCovID, TmpDOB As String, tmpCisWarranty As String)
Dim FindAge As String
Dim AgeValue As String
Dim CISSuppAnnLmt, CISSuppVisLmt As String
Dim CISSuppBPrem, CISSuppPrem As String
Dim CISSuppBMCO, CISSuppMCO As String
Dim CisDentalLmt As Double
Dim CisSpecLmt As Double
Dim CisMepLmt As Double
Dim CisAnnLmtType As String
Dim CisSuppSql As String
Dim SysDate As String
Dim tmpCisExpDate As String
Dim tmpSuppExpDate As String
'Dim CISPayor As String
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
Dim tmpCardType As String
Dim CLNStatus As String

'    CISPayor = IIf(Trim(CISDCStatus) = "Y", "DC", Trim(txtPAYCode))

    tmpCISPln = Mid(frmMBMClinical.cboPLNCode, 1, 4)
    tmpCisEffDate = frmMBMClinical.txtMBMPayorEffDate
    tmpCisExpDate = frmMBMClinical.txtMBMPayorExpDate
    tmpCisInsType = Mid(frmMBMClinical.cboINSCode, 1, 1)
    MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"

    ' ******* Clinical Supp Limit
        CISstrSptLmt = "Select SuppLimitStatus, PLNCode from AnnualLimit where PLNCode = '" & Trim(tmpCISPln) & _
                       "' AND HLTCode='" & Trim(txtHLTCode) & "' AND INSCode='" & Trim(tmpCisInsType) & _
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
        If Trim(txtHLTCode) = "S" Then
            CISstrSptPrem = "Select SuppPrmStatus, PLNCode from PRM where PLNCode = '" & Trim(tmpCISPln) & _
                             "' AND HLTCode='" & Trim(txtHLTCode) & "' AND INSCode='" & Trim(tmpCisInsType) & _
                             "' AND PRMEffDate <= '" & Format(tmpCisEffDate, "mm/dd/yyyy") & "' order by PRMEffDate desc"
        Else
            CISstrSptPrem = "Select SuppPrmStatus, PLNCode from PRM where PLNCode = '" & Trim(tmpCISPln) & _
                             "' AND HLTCode='" & Trim(txtHLTCode) & "' AND INSCode='" & Trim(tmpCisInsType) & _
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
                         "' AND HLTCode='" & Trim(txtHLTCode) & "' AND INSCode='" & Trim(tmpCisInsType) & _
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
                              " AND HLTCode='" & Trim(txtHLTCode) & "' AND INSCode='" & Trim(tmpCisInsType) & _
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
        CISSuppBPrem = 0
        CISSuppPrem = 0
        CISSuppBMCO = 0
        CISSuppMCO = 0
        CISSuppAnnLmt = 0
        CISSuppVisLmt = 0
        CisDentalLmt = 0
        CisSpecLmt = 0
        CisMepLmt = 0

        FindAge = GetAge(CDate(tmpCisEffDate), CDate(TmpDOB))
        AgeValue = Mid(GetAgeBand(txtHLTCode, FindAge, tmpCISPln, tmpCisInsType), 1, 2)
        TmpCovID = Format(TmpCovID, "00")
    
        If Trim(CISSptPremInd) = "B" And Format(TmpCovID, "00") = "01" Or Trim(CISSptPremInd) = "C" Then
            CISSuppPrem = Format(GetCLNSPremium(tmpCisInsType, CISDCPayor, AgeValue, tmpCISPln, Trim(txtHLTCode), CDate(tmpCisEffDate)), "#,##0.00")
            CISSuppBPrem = CISSuppPrem
            CISSuppPrem = InsertPremium(CDate(tmpCisExpDate), CDate(tmpCisEffDate), Format(CISSuppPrem, "#,##0.00"))
            CISSuppPrem = Format(Round(CISSuppPrem, 0), "#,##0.00")
        End If
        If Trim(CISSptMCOInd) = "B" And Format(TmpCovID, "00") = "01" Or Trim(CISSptMCOInd) = "C" Then
            If AgeValue = "01" Then
                CISSuppMCO = Format(GetCLNSMCO(tmpCisInsType, Trim(txtHLTCode), "C", CDate(tmpCisEffDate), tmpCISPln), "#,##0.00")
            Else
                CISSuppMCO = Format(GetCLNSMCO(tmpCisInsType, Trim(txtHLTCode), "A", CDate(tmpCisEffDate), tmpCISPln), "#,##0.00")
            End If
            CISSuppBMCO = CISSuppMCO
            CISSuppMCO = InsertMCOFees(CDate(tmpCisExpDate), CDate(tmpCisEffDate), Format(CISSuppMCO, "#,##0.00"))
            CISSuppMCO = Round(CISSuppMCO)
        End If
        If Trim(CISSptLmtInd) = "B" And Format(TmpCovID, "00") = "01" Or Trim(CISSptLmtInd) = "C" Then
            CISSuppAnnLmt = Format(GetCLNSuppLimit(tmpCisInsType, tmpCISPln, Trim(txtHLTCode), CDate(tmpCisEffDate)), "#,##0.00")
            CISSuppVisLmt = GetCLNSVisitLimit(tmpCisInsType, tmpCISPln, Trim(txtHLTCode), CDate(tmpCisEffDate))
            If Trim(CISSDentStatus) = "S" Then 'stand alone dental limit
                CisAnnLmtType = "SDentalAnnLimit"
                CisDentalLmt = GetCISLimit(tmpCisInsType, tmpCISPln, Trim(txtHLTCode), CDate(tmpCisEffDate), CisAnnLmtType)
            End If
            If Trim(CISSSpecStatus) = "S" Then 'stand alone Spec limit
                CisAnnLmtType = "SSpecAnnLimit"
                CisSpecLmt = GetCISLimit(tmpCisInsType, tmpCISPln, Trim(txtHLTCode), CDate(tmpCisEffDate), CisAnnLmtType)
            End If
            If Trim(CISSMepStatus) = "S" Then 'stand alone MEP limit
                CisAnnLmtType = "SMepAnnLimit"
                CisMepLmt = GetCISLimit(tmpCisInsType, tmpCISPln, Trim(txtHLTCode), CDate(tmpCisEffDate), CisAnnLmtType)
            End If
        End If
        MediConnCISDB.Close
        Set MediConnCISDB = Nothing
' ****** End of Clinical Supp Limit
        CISSuppBPrem = CDbl(IIf(IsNumeric(CISSuppBPrem) = True, CISSuppBPrem, 0))
        CISSuppPrem = CDbl(IIf(IsNumeric(CISSuppPrem) = True, CISSuppPrem, 0))
        CISSuppBMCO = CDbl(IIf(IsNumeric(CISSuppBMCO) = True, CISSuppBMCO, 0))
        CISSuppMCO = CDbl(IIf(IsNumeric(CISSuppMCO) = True, CISSuppMCO, 0))
        CISSuppAnnLmt = CDbl(IIf(IsNumeric(CISSuppAnnLmt) = True, CISSuppAnnLmt, 0))
        CISSuppVisLmt = CDbl(IIf(IsNumeric(CISSuppVisLmt) = True, CISSuppVisLmt, 0))
        SysDate = Format(Date, "yyyy/mm/dd")
        tmpSuppExpDate = Format(tmpCisExpDate, "yyyy/mm/dd")
        tmpCardType = Replace(frmMBMClinical.CboCardType, "'", "''")
        tmpCisWarranty = Replace(tmpCisWarranty, "'", "''")
        CLNStatus = "A"
        CisSuppSql = "EXEC Upload_MBMCovPersonsCis '" & MembershipNo & "','" & TmpCovID & "', " & _
                    "" & CISSuppMCO & "," & CISSuppBMCO & "," & CISSuppPrem & "," & CISSuppBPrem & ", " & _
                    "" & CISSuppAnnLmt & "," & CISSuppVisLmt & "," & CisDentalLmt & "," & CisSpecLmt & ", " & _
                    "" & CisMepLmt & ",'" & SysDate & "','" & Logon & "'," & _
                    "'" & SysDate & "','" & Logon & "','" & tmpSuppExpDate & "'," & _
                    "'" & tmpCardType & "','" & tmpCisWarranty & "','" & CLNStatus & "','" & Trim(tmpCISPln) & "'"
        medixmasterconn.Execute CisSuppSql
    ' Insert into Membership Clinic Covered Person's table
'    Next listloop
End Sub

Private Sub SaveMBMOthers(MembershipNo As String)
Dim MBMOthers As New ADODB.Recordset
Dim MBMOthers1 As New ADODB.Recordset
Dim strsqlMBMOth As String
Dim FirstJoinDate As String
Dim TempMBMProviderCharges As String
Dim TempMBMPremium As String
Dim TempMCOFees As String
Dim TempMCOFees2 As String
Dim TempPrem As String
Dim TempProvider As String
Dim TempPolicyDisc As String
Dim TempRenewalDisc As String
Dim TempGRPCompany As String
Dim TempUNDExcess As String
Dim TmpGuaRenew As String
Dim tmpBranch As String

     FirstJoinDate = Format(txtMBMDateJoin, "YYYY/MM/DD")
     TempMBMProviderCharges = CDbl(IIf(IsNumeric(txtMBMProviderCharges) = True, txtMBMProviderCharges, 0))
     TempMBMPremium = CDbl(IIf(IsNumeric(txtMBMPremium) = True, txtMBMPremium, 0))
     TempMCOFees = CDbl(IIf(IsNumeric(txtMCOFees) = True, txtMCOFees, 0))
     TempMCOFees2 = CDbl(IIf(IsNumeric(txtTempMCOFees) = True, txtTempMCOFees, 0))
     TempPrem = CDbl(IIf(IsNumeric(txtTempPrem) = True, txtTempPrem, 0))
     TempProvider = CDbl(IIf(IsNumeric(txtTempProvider) = True, txtTempProvider, 0))
     TempPolicyDisc = CInt(IIf(IsNumeric(frmNewMBMOthers.txtPolicyDisc) = True, frmNewMBMOthers.txtPolicyDisc, 0))
     TempRenewalDisc = CInt(IIf(IsNumeric(frmNewMBMOthers.txtRenewalDisc) = True, frmNewMBMOthers.txtRenewalDisc, 0))
     TempGRPCompany = Replace(txtMBMGroupCompany, "'", "''")
     TempUNDExcess = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtUNDExcess) = True, frmNewMBMOthers.txtUNDExcess, 0))
     TmpGuaRenew = IIf(Len(CboGuaRenewal), Mid(CboGuaRenewal, 1, 1), "")
     tmpBranch = Trim(cboBRNCode)
    ' INsert into Membership Other details table
     'MBMOthers.Open "Select MBMNumber, * from MBMOthers", medixmasterconn, adOpenKeyset, adLockOptimistic
     strsqlMBMOth = "insert into MBMOthers (MBMNumber, MBMGroupCompany, MBMEmployeeNo, " & _
                    "MBMAgentCode, MBMPolSubNo1,  MBMBranch, " & _
                    "MBMAccessFee, MBMPremium, MBMMCOFee, MBMBasicMCO, MBMBasicPrem, MBMBasicEA, " & _
                    "MBMDepartment, MBMInstallment, RELCODE, MBMUndExcess," & _
                    "MBMFirstMEMNo, MBMFirstPolNo, MBMPrevPolNo, MBMPrevMemNo,MBMFirstJoinedDate, " & _
                    "MBMPolicyDisc, MBMRenewalDisc,MBMPolCondition,MBMGuaRenewal)" & _
                    "Values('" & MembershipNo & "', '" & TempGRPCompany & "', '" & txtMBMEmployeeNo & "', " & _
                    "'" & txtMBMAgentCode & "', '" & frmNewMBMOthers.TxtPOLSubNo & "', '" & tmpBranch & "', " & _
                    "" & TempMBMProviderCharges & " , " & TempMBMPremium & ", " & TempMCOFees & ", " & TempMCOFees2 & ", " & TempPrem & ", " & TempProvider & ", " & _
                    "'" & txtMBMDepartment & "', '" & Mid(CboInsMode, 1, 1) & "', 'P', " & TempUNDExcess & ", " & _
                    "'" & ChkStr(txtFirstMeMNo) & "', '" & txtFirstPolNo & "', '" & txtMBMPrevPolNo & "', '" & txtMBMPrevMEMNo & "', '" & FirstJoinDate & "', " & _
                    "" & TempPolicyDisc & ", " & TempRenewalDisc & ", '" & Mid(frmNewMBMOthers.cboPOLCondition, 1, 2) & "','" & TmpGuaRenew & "')"

    medixmasterconn.Execute strsqlMBMOth
    completeMBMOthers = True
End Sub

Private Sub SaveBatch()
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
    completeMBMC = True
End Sub
    
Private Sub UpdateBordxPriting()
Dim BordxPrinting As New ADODB.Recordset

    BordxPrinting.Open "BordxPrinting", medixmasterconn, adOpenKeyset, adLockOptimistic
    BordxPrinting.AddNew
    With BordxPrinting
        !MBMBordxDate = txtMBMBordxDate
        !PAYCode = txtPAYCode
        !MBMBatchNO = txtMBMBatchNo
        !MBMBordxType = "N"
    End With
    BordxPrinting.Update
    BordxPrinting.Close
    Set BordxPrinting = Nothing
End Sub

'  ################# Start Suppplementary ######################


Private Sub CmdAdd_Click()
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
            .SubItems(28) = Mid(cboSuppStaffDisc, 1, 1)
            .SubItems(29) = TxtSuppClmNonPayFr
            .SubItems(30) = TxtSuppClmNonPayTo
            .SubItems(31) = txtCLNCovExclusion
            .SubItems(32) = txtSuppPayPremGross
            .SubItems(33) = txtSuppPayPremNet
            .SubItems(34) = txtSuppPayMCOGross
            .SubItems(35) = txtSuppPayMCONet
            .SubItems(36) = txtRenewalDisc
            .SubItems(37) = txtPolicyDisc
            .SubItems(38) = txtPolicyDiscAmt
            .SubItems(39) = txtRenewalDiscAmt
            .SubItems(40) = txtLoadingPrem
            .SubItems(41) = Trim(txtSuppPLNCode)
            .SubItems(42) = txtAvailableLimit
            .SubItems(43) = IIf(IsNumeric(txtCovUndExcess), txtCovUndExcess, 0)
            End With
            Call ClearCoveredEntry
            If cboCovSalutation.Enabled = True Then
                cboCovSalutation.SetFocus
            End If
            cboCovRelationship = "S - SON"
    End If

End Sub

Private Sub CmdRemove_Click()
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


Private Sub CmdUpdate_Click()
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
                lstCovAdd.SelectedItem.SubItems(28) = Mid(cboSuppStaffDisc, 1, 1)
                lstCovAdd.SelectedItem.SubItems(29) = TxtSuppClmNonPayFr
                lstCovAdd.SelectedItem.SubItems(30) = TxtSuppClmNonPayTo
                lstCovAdd.SelectedItem.SubItems(31) = txtCLNCovExclusion
                lstCovAdd.SelectedItem.SubItems(32) = txtSuppPayPremGross
                lstCovAdd.SelectedItem.SubItems(33) = txtSuppPayPremNet
                lstCovAdd.SelectedItem.SubItems(34) = txtSuppPayMCOGross
                lstCovAdd.SelectedItem.SubItems(35) = txtSuppPayMCONet
                lstCovAdd.SelectedItem.SubItems(36) = txtRenewalDisc
                lstCovAdd.SelectedItem.SubItems(37) = txtPolicyDisc
                lstCovAdd.SelectedItem.SubItems(38) = txtPolicyDiscAmt
                lstCovAdd.SelectedItem.SubItems(39) = txtRenewalDiscAmt
                lstCovAdd.SelectedItem.SubItems(40) = txtLoadingPrem
                lstCovAdd.SelectedItem.SubItems(41) = Trim(txtSuppPLNCode)
                lstCovAdd.SelectedItem.SubItems(42) = txtAvailableLimit
                lstCovAdd.SelectedItem.SubItems(43) = txtCovUndExcess
                Call ClearCoveredEntry
                Exit For
            End If
        End If
    Next ForLoop
    
    cboCovSalutation.SetFocus
    cmdRemove.Enabled = True
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = False
End Sub

Private Sub lstCovAdd_Click()
'On Error Resume Next
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
        If Len(lstCovAdd.SelectedItem.SubItems(23)) Then
            txtSuppGENWPEnd = lstCovAdd.SelectedItem.SubItems(23)
        End If
        txtSuppPREWP = lstCovAdd.SelectedItem.SubItems(24)
        If Len(lstCovAdd.SelectedItem.SubItems(25)) Then
            txtSuppPREWPEnd = lstCovAdd.SelectedItem.SubItems(25)
        End If
        TxtSuppSpeWP = lstCovAdd.SelectedItem.SubItems(26)
        If Len(lstCovAdd.SelectedItem.SubItems(27)) Then
            txtSuppSPEWPEnd = lstCovAdd.SelectedItem.SubItems(27)
        End If
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
        txtPolicyDiscAmt = lstCovAdd.SelectedItem.SubItems(38)
        txtRenewalDiscAmt = lstCovAdd.SelectedItem.SubItems(39)
        txtLoadingPrem = lstCovAdd.SelectedItem.SubItems(40)
        txtSuppPLNCode = lstCovAdd.SelectedItem.SubItems(41)
        txtAvailableLimit = Format(lstCovAdd.SelectedItem.SubItems(42), "###,##0.00")
        If IsNumeric(lstCovAdd.SelectedItem.SubItems(43)) Then
            txtCovUndExcess = lstCovAdd.SelectedItem.SubItems(43)
        End If
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


Private Sub EnableCovered(status As Boolean)
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

Private Sub ClearCoveredEntry()
    cboCovSalutation = "MR"
    txtCovName = ""
    txtCovIcBcPp = ""
    txtCovBirthdate.mask = ""
    txtCovBirthdate.Text = ""
    txtCovBirthdate.mask = "##/##/####"
    cboCovSex = "M"
    cboCovAdultChild = "A"
    cboCovRelationship = "S - Son"
    txtCovAllergic = ""
    txtCovExclusion = ""
    txtCLNCovExclusion = ""
    txtcount = ""
    txtPrevPLNCode = ""
    txtRBAmt = ""
    txtPrevINSCom = ""
    txtCOVPayRemarks = ""
    txtCovAnnualLimit = ""
    txtSuppGenWP = ""
    txtSuppPREWP = ""
    TxtSuppSpeWP = ""
    cboSuppStaffDisc = ""
    
    txtSuppGENWPEnd.mask = ""
    txtSuppGENWPEnd.Text = ""
    txtSuppGENWPEnd.mask = "##/##/####"
    txtSuppPREWPEnd.mask = ""
    txtSuppPREWPEnd.Text = ""
    txtSuppPREWPEnd.mask = "##/##/####"
    txtSuppSPEWPEnd.mask = ""
    txtSuppSPEWPEnd.Text = ""
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
    txtSuppPLNCode = ""
    txtAvailableLimit = ""
    txtCovUndExcess = ""
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

Private Sub cboCTYCode_KeyPress(KeyAscii As Integer)
  KeyAscii = KeyCharacter(KeyAscii)
End Sub

Private Sub cboCTYCode_LostFocus()
    cboCTYCode = ChkStr(cboCTYCode)
    If Len(cboCTYCode) > 50 Then
        MsgBox "City/Town Exceeded 50 Characters!", vbCritical
        cboCTYCode.SetFocus
    End If
End Sub

Private Sub cboSTACode_LostFocus()
    cboSTACode = ChkStr(cboSTACode)
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

Private Sub InsertProviderFees()
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
    txtRemarks = ChkStr(txtRemarks)
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

Private Function showMBMOthers() As Integer
Dim sqlstr As String
Dim MBM As New ADODB.Recordset
Dim strCount As String
Dim rstCount As New ADODB.Recordset
Dim MBMNumber As String
    
        MBMNumber = txtMBMPrevMEMNo
        With ADOrstMBM
            
            frmNewMBMOthers.Label4 = Trim(MBMNumber)
            If Trim(!MBMCLMNonPayFr) <> "" Then
                frmNewMBMOthers.TxtClmNonPayFr = Format(!MBMCLMNonPayFr, "dd/mm/yyyy")
            End If
            If Trim(!MBMCLMNonPayTo) <> "" Then
                frmNewMBMOthers.TxtClmNonPayTo = Format(!MBMCLMNonPayTo, "dd/mm/yyyy")
            End If
            If Trim(!LabelStatus) = "Y" Then
                frmNewMBMOthers.Check1.value = 1
            End If
            
            frmNewMBMOthers.TxtPOLSubNo = IIf(Len(Trim(!MBMPolSubNo1)), Trim(!MBMPolSubNo1), "")
            frmNewMBMOthers.txtRenewalDisc = IIf(IsNumeric(!MBMRenewalDisc), !MBMRenewalDisc, 0)
            frmNewMBMOthers.txtPolicyDisc = IIf(IsNumeric(!MBMPolicyDisc), !MBMPolicyDisc, 0)
            frmNewMBMOthers.txtPLNCode = IIf(Len(!MBMNewPlanCode), !MBMNewPlanCode, "")
            
            If Len(!MBMNewPolEffDate) Then
                frmNewMBMOthers.txtNewPolEffDate = Format(!MBMNewPolEffDate, "dd/mm/yyyy")
            End If
            
            txtMBMGroupCompany = IIf(Len(Trim(!MBMGroupCompany)), Trim(!MBMGroupCompany), "")
            txtMBMEmployeeNo = IIf(Len(Trim(!MBMEmployeeNo)), Trim(!MBMEmployeeNo), "")
            txtFirstPolNo = IIf(Len(Trim(!MBMFirstPolNo)), Trim(!MBMFirstPolNo), "")
            txtFirstMeMNo = IIf(Len(Trim(!MBMFirstMEMNo)), Trim(!MBMFirstMEMNo), "")
            txtMBMAgentCode = IIf(Len(Trim(!MBMAgentCode)), Trim(!MBMAgentCode), "")

            If Trim(!MBMFirstJoinedDate) <> "" Then
                txtMBMDateJoin = !MBMFirstJoinedDate
            End If

            txtMBMDepartment = IIf(Len(Trim(!MBMDepartment)), Trim(!MBMDepartment), "")
            frmNewMBMOthers.cboPOLCondition = IIf(Len(!MBMPolCondition), !MBMPolCondition, "")
            frmNewMBMOthers.txtUNDExcess = IIf(IsNumeric(!MBMUndExcess), !MBMUndExcess, 0)
            CboGuaRenewal.Text = IIf(Len(!MBMGuaRenewal), !MBMGuaRenewal, "")
        
        End With
End Function

Private Function showMBMTwo() As Integer
Dim sqlstr As String
Dim MBMTwo As New ADODB.Recordset
Dim strCount As String
Dim rstCount As New ADODB.Recordset
    
    If frmNewMBMOthers.Visible = False Then
        Load frmNewMBMOthers
    End If
    
    Call ClearForm(frmNewMBMOthers)

    If Not ADOrstMBM.EOF Then
        With ADOrstMBM
            If Len(Trim(!MBMSalutation)) Then
               cboMBMSalutation = Trim(!MBMSalutation)
            End If
            
            If Len(Trim(!NATCode)) Then
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
                frmNewMBMOthers.txtMBMEmail = !MBMEmail
            End If
            
            If Len(!LGNCode) Then
               frmNewMBMOthers.cboLGNCode = !LGNCode
            End If
            
        End With
    End If
End Function

Private Function showMBMOthersTwo() As Integer
Dim sqlstr As String
Dim MBMOthersTwo As New ADODB.Recordset
Dim strCount As String
Dim rstCount As New ADODB.Recordset
Dim MBMNumber As String
    
    If frmNewMBMOthers.Visible = False Then
        Load frmNewMBMOthers
    End If
    
    MBMNumber = txtMBMPrevMEMNo
    If Not ADOrstMBM.EOF Then
        With ADOrstMBM
            frmNewMBMOthers.Label4 = Trim(MBMNumber)
            frmNewMBMOthers.txtMBMPrevPlan = IIf(Len(!MBMPrevPLNCode), !MBMPrevPLNCode, "")
            frmNewMBMOthers.txtMBMRBAmt = IIf(IsNumeric(!MBMRBAmt), !MBMRBAmt, 0)
            frmNewMBMOthers.txtMBMPrevINSCom = IIf(Len(!MBMPrevInsCom), !MBMPrevInsCom, "")
            frmNewMBMOthers.cboCOPayRB = IIf(Len(!MBMCoPayRB), !MBMCoPayRB, "")
            frmNewMBMOthers.TxtGenWP = IIf(Len(!MBMGenWP), !MBMGenWP, "")
            frmNewMBMOthers.txtPREWP = IIf(Len(!MBMPreExWP), !MBMPreExWP, "")
            frmNewMBMOthers.TxtSpeWP = IIf(Len(!MBMSpeIllWP), !MBMSpeIllWP, "")
            frmNewMBMOthers.cboStaffDisc = IIf(Len(!MBMStaffDec), !MBMStaffDec, "")
            If Len(!MBMGenEndWP) Then
                frmNewMBMOthers.txtGENWPEnd = !MBMGenEndWP
            End If
            If Len(!MBMPreExEndWP) Then
                frmNewMBMOthers.txtPREWPEnd = !MBMPreExEndWP
            End If
            If Len(!MBMSpeIllEndWP) Then
                frmNewMBMOthers.txtSPEWPEnd = !MBMSpeIllEndWP
            End If
            frmNewMBMOthers.txtMBMAllergic = IIf(Len(Trim(!MBMAllergic)), Trim(!MBMAllergic), "")
            txtRemarks = IIf(Len(!MBMRemarks), !MBMRemarks, "")
            frmNewMBMOthers.cboMBMOccupation = IIf(Len(!MBMOccupation), !MBMOccupation, "")
            frmNewMBMOthers.txtMBMWorkNature = IIf(Len(!MBMWorkNature), !MBMWorkNature, "")
            frmNewMBMOthers.txtMBMWeight = IIf(Len(!MBMWeight), !MBMWeight, "")
            frmNewMBMOthers.txtMBMHeight = IIf(Len(!MBMHeight), !MBMHeight, "")
            frmNewMBMOthers.cboMBMBlood = IIf(Len(!MBMBlood), !MBMBlood, "")
            frmNewMBMOthers.cboMBMSmoking = IIf(Len(!MBMSmoking), !MBMSmoking, "")
            frmNewMBMOthers.cboMBMAlcohol = IIf(Len(!MBMAlcohol), !MBMAlcohol, "")
            frmNewMBMOthers.txtPayPrem = IIf(IsNumeric(!MBMPayorPremGross), !MBMPayorPremGross, 0)
            frmNewMBMOthers.txtPayPremNett = IIf(IsNumeric(!MBMPayorPremNet), !MBMPayorPremNet, 0)
            frmNewMBMOthers.txtPayMCOGross = IIf(IsNumeric(!MBMPayorMCOGross), !MBMPayorMCOGross, 0)
            frmNewMBMOthers.txtPayMCONett = IIf(IsNumeric(!MBMPayorMCONet), !MBMPayorMCONet, 0)
            frmNewMBMOthers.cboBTBG = IIf(Len(!MBMBTBG), !MBMBTBG, "")
            frmNewMBMOthers.txtMBMPAYRemarks = IIf(Len(!MBMPAYRemarks), !MBMPAYRemarks, "")
            frmNewMBMOthers.txtPolicyDiscAmt = IIf(IsNumeric(!MBMFamDiscAmt), !MBMFamDiscAmt, 0)
            frmNewMBMOthers.txtRenewalDiscAmt = IIf(IsNumeric(!MBMRenewDiscAmt), !MBMRenewDiscAmt, 0)
            frmNewMBMOthers.txtLoadingPrem = IIf(IsNumeric(!MBMLoadingPrem), !MBMLoadingPrem, 0)
            txtMBMExclusion = IIf(Len(Trim(!MBMExclusion)), Trim(!MBMExclusion), "")
            frmNewMBMOthers.txtBankACNo = IIf(Len(Trim(!MBMBankACNo)), Trim(!MBMBankACNo), "")

        End With
    End If
End Function

Private Sub SaveMBMOthersTwo(MembershipNo As String)
    Dim strsqlMBMOthTwo As String
    Dim TmpPayPremium As String
    Dim TmpPayPremiumNett As String
    Dim TmpPayMCOGross As String
    Dim TmpPayMCONett As String
    Dim TmpMBMRBAmt As String
    Dim TmpPolicyDiscAmt As String
    Dim TmpRenewalDiscAmt  As String
    Dim TmpLoadingPrem As String
    Dim TempMBMExclusion As String
    
    TmpPayPremium = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtPayPrem) = True, frmNewMBMOthers.txtPayPrem, 0))
    TmpPayPremiumNett = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtPayPremNett) = True, frmNewMBMOthers.txtPayPremNett, 0))
    TmpPayMCOGross = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtPayMCOGross) = True, frmNewMBMOthers.txtPayMCOGross, 0))
    TmpPayMCONett = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtPayMCONett) = True, frmNewMBMOthers.txtPayMCONett, 0))
    TmpMBMRBAmt = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtMBMRBAmt) = True, frmNewMBMOthers.txtMBMRBAmt, 0))
    TmpPolicyDiscAmt = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtPolicyDiscAmt) = True, frmNewMBMOthers.txtPolicyDiscAmt, 0))
    TmpRenewalDiscAmt = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtRenewalDiscAmt) = True, frmNewMBMOthers.txtRenewalDiscAmt, 0))
    TmpLoadingPrem = CDbl(IIf(IsNumeric(frmNewMBMOthers.txtLoadingPrem) = True, frmNewMBMOthers.txtLoadingPrem, 0))
    TempMBMExclusion = Replace(txtMBMExclusion, "'", "''")

    ' INsert into Membership Other details table
     
        strsqlMBMOthTwo = "insert into MBMOthersTwo (MBMNumber, MBMWeight, MBMHeight, " & _
                    "MBMOccupation, MBMWorkNature, MBMBlood, MBMSmoking, " & _
                    "MBMAlcohol, MBMPrevPLNCode, MBMPAYRemarks, " & _
                    "MBMPrevInsCom, MBMMaritalStatus, MBMCoPayRB, MBMStaffDec, MBMPayorPremGross, " & _
                    "MBMPayorPremNet,MBMPayorMCOGross,MBMPayorMCONet," & _
                    "MBMRBAmt, MBMFamDiscAmt, MBMRenewDiscAmt,MBMLoadingPrem,MBMExclusion,MBMBankACNo, MBMAllergic,MBMRemarks)" & _
                    "Values('" & MembershipNo & "', '" & ChkStr(frmNewMBMOthers.txtMBMWeight) & "', '" & ChkStr(frmNewMBMOthers.txtMBMHeight) & "', " & _
                    "'" & Trim(Mid(frmNewMBMOthers.cboMBMOccupation, 1, 2)) & "', '" & Mid(Trim(frmNewMBMOthers.txtMBMWorkNature), 1, 3) & "', '" & ChkStr(frmNewMBMOthers.cboMBMBlood) & "', '" & frmNewMBMOthers.cboMBMSmoking & "', " & _
                    "'" & frmNewMBMOthers.cboMBMAlcohol & "' , '" & frmNewMBMOthers.txtMBMPrevPlan & "', '" & ChkStr(frmNewMBMOthers.txtMBMPAYRemarks) & "', " & _
                    "'" & ChkStr(frmNewMBMOthers.txtMBMPrevINSCom) & "', '" & Mid(cboMBMMaritalStatus, 1, 1) & "', '" & frmNewMBMOthers.cboCOPayRB & "', '" & Mid(frmNewMBMOthers.cboStaffDisc, 1, 1) & "', " & TmpPayPremium & ", " & _
                    "" & TmpPayPremiumNett & ", " & TmpPayMCOGross & ", " & TmpPayMCONett & ", " & _
                    "" & TmpMBMRBAmt & ", " & TmpPolicyDiscAmt & ", " & TmpRenewalDiscAmt & ", " & TmpLoadingPrem & " , '" & ChkStr(TempMBMExclusion) & "', '" & ChkStr(frmNewMBMOthers.txtBankACNo) & "'," & _
                    "'" & frmNewMBMOthers.txtMBMAllergic & "','" & txtRemarks & "') "
        medixmasterconn.Execute strsqlMBMOthTwo
        completeMBMOthersTwo = True
End Sub

Private Sub ClearMBMOthers()
On Error Resume Next

    frmNewMBMOthers.Label4 = ""
    frmNewMBMOthers.TxtClmNonPayFr.mask = ""
    frmNewMBMOthers.TxtClmNonPayFr.Text = ""
    frmNewMBMOthers.TxtClmNonPayFr.mask = "##/##/####"
    frmNewMBMOthers.TxtClmNonPayTo.mask = ""
    frmNewMBMOthers.TxtClmNonPayTo.Text = ""
    frmNewMBMOthers.TxtClmNonPayTo.mask = "##/##/####"
    frmNewMBMOthers.Check1.value = 0
    frmNewMBMOthers.TxtPOLSubNo = ""
    frmNewMBMOthers.txtMBMAllergic = ""
    frmNewMBMOthers.txtMBMPAYRemarks = ""
    frmNewMBMOthers.txtPolicyDisc = ""
    frmNewMBMOthers.txtRenewalDisc = ""
    frmNewMBMOthers.txtPLNCode = ""
    frmNewMBMOthers.txtNewPolEffDate.mask = ""
    frmNewMBMOthers.txtNewPolEffDate.Text = ""
    frmNewMBMOthers.txtNewPolEffDate.mask = "##/##/####"
    txtMBMGroupCompany = ""
    txtMBMEmployeeNo = ""
    txtRemarks = ""
    txtMBMExclusion = ""
    txtFirstPolNo = ""
    txtFirstMeMNo = ""
    txtMBMAgentCode = ""
    txtMBMDateJoin.mask = ""
    txtMBMDateJoin.Text = ""
    txtMBMDateJoin.mask = "##/##/####"
    frmNewMBMOthers.txtGENWPEnd.mask = ""
    frmNewMBMOthers.txtGENWPEnd.Text = ""
    frmNewMBMOthers.txtGENWPEnd.mask = "##/##/####"
    frmNewMBMOthers.txtPREWPEnd.mask = ""
    frmNewMBMOthers.txtPREWPEnd.Text = ""
    frmNewMBMOthers.txtPREWPEnd.mask = "##/##/####"
    frmNewMBMOthers.txtSPEWPEnd.mask = ""
    frmNewMBMOthers.txtSPEWPEnd.Text = ""
    frmNewMBMOthers.txtSPEWPEnd.mask = "##/##/####"
    frmNewMBMOthers.TxtClmNonPayFr.Text = ""
    frmNewMBMOthers.TxtClmNonPayFr.mask = "##/##/####"
    frmNewMBMOthers.TxtClmNonPayTo.Text = ""
    frmNewMBMOthers.TxtClmNonPayTo.mask = "##/##/####"
    frmNewMBMOthers.cboBTBG = "N - No"
End Sub

Private Sub SaveMBMCLNOthers(MembershipNo As String)
Dim ClnOthersSql As String
Dim tmpCisPrem, tmpCisBPrem, tmpCISMCO, tmpCisBMCO As Double
Dim tmpPOlNo, TmpCisWarr, tmpRemark As String
Dim CISPLN As String
Dim CisInsType As String
Dim CLnAnnLmt As New ADODB.Recordset
Dim CLNPln As New ADODB.Recordset
Dim CLNPlnSql, ClnAnnLmtSql As String
Dim CISCovSuppLMtStat, TmpMBMType, TmpCisPayor As String
Dim tmpCisEffDate As String
Dim tmpCisExpDate As String
Dim CisDentalLmt As Double
Dim CisSpecLmt As Double
Dim CisMepLmt As Double
Dim CisAnnLmtType As String
Dim CisVisitLmt As Double
Dim CisBalLmt As Double
Dim tmpCardType As String

    TmpCisPayor = Trim(txtPAYCode)
    CISPLN = Mid(frmMBMClinical.cboPLNCode, 1, 4)
    CisInsType = Mid(frmMBMClinical.cboINSCode, 1, 1)
    
    CISCovSuppLMtStat = "A"
    MediConnCISDB.Open "File Name=" & BOSPath & "Conn\MediConnCISDB.UDL"

    CLNPlnSql = "SELECT ANNLmtInd, DCStatus, PAYCode from PLN where PLNCode='" & Trim(CISPLN) & _
            "' AND HLTCode='" & Trim(txtHLTCode) & "'"
    CLNPln.Open CLNPlnSql, MediConnCISDB, adOpenKeyset, adLockOptimistic
    
    If CLNPln.recordCount Then
        TmpCisPayor = IIf(Trim(CLNPln!DCStatus) = "Y", CLNPln!PAYCode, Trim(txtPAYCode))
        CISDCPayor = Trim(TmpCisPayor)
        If Trim(CLNPln!ANNLmtInd) = "Y" Then
            ClnAnnLmtSql = "Select SuppLimitStatus from AnnualLimit where PLNCode = '" & Trim(CISPLN) & _
                           "' AND HLTCode='" & Trim(txtHLTCode) & "' AND INSCode='" & Trim(CisInsType) & "'"
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
    
    CisDentalLmt = 0
    CisSpecLmt = 0
    CisMepLmt = 0
    
    If Trim(CISPDentStatus) = "S" Then 'stand alone dental limit
        CisAnnLmtType = "PDentalAnnLimit"
        CisDentalLmt = GetCISLimit(CisInsType, CISPLN, Trim(txtHLTCode), CDate(frmMBMClinical.txtMBMPayorEffDate), CisAnnLmtType)
    End If
    If Trim(CISPSpecStatus) = "S" Then 'stand alone Spec limit
        CisAnnLmtType = "PSpecAnnLimit"
        CisSpecLmt = GetCISLimit(CisInsType, CISPLN, Trim(txtHLTCode), CDate(frmMBMClinical.txtMBMPayorEffDate), CisAnnLmtType)
    End If
    If Trim(CISPMepStatus) = "S" Then 'stand alone MEP limit
        CisAnnLmtType = "PMepAnnLimit"
        CisMepLmt = GetCISLimit(CisInsType, CISPLN, Trim(txtHLTCode), CDate(frmMBMClinical.txtMBMPayorEffDate), CisAnnLmtType)
    End If
    If Trim(CISCovSuppLMtStat) = "A" Then
        TmpMBMType = "P"
    Else
        TmpMBMType = "Q"
    End If
    MediConnCISDB.Close
    Set MediConnCISDB = Nothing
    tmpCisPrem = IIf(IsNumeric(frmMBMClinical.lblMBMPremium) = True, frmMBMClinical.lblMBMPremium, 0)
    tmpCisBPrem = IIf(IsNumeric(frmMBMClinical.txtTempPrem) = True, frmMBMClinical.txtTempPrem, 0)
    tmpCISMCO = IIf(IsNumeric(frmMBMClinical.lblMCOFees) = True, frmMBMClinical.lblMCOFees, 0)
    tmpCisBMCO = IIf(IsNumeric(frmMBMClinical.txtTempMCOFees) = True, frmMBMClinical.txtTempMCOFees, 0)
    tmpCisEffDate = Format(frmMBMClinical.txtMBMPayorEffDate, "yyyy/mm/dd")
    tmpCisExpDate = Format(frmMBMClinical.txtMBMPayorExpDate, "yyyy/mm/dd")
    tmpPOlNo = Replace(txtMBMPolicyNo, "'", "''")
    TmpCisWarr = Replace(frmMBMClinical.txtWarranty, "'", "''")
    tmpRemark = Replace(frmMBMClinical.txtRemarks, "'", "''")
    CisBalLmt = IIf(IsNumeric(frmMBMClinical.lblAnnualLimit) = True, frmMBMClinical.lblAnnualLimit, 0)
    CisVisitLmt = IIf(IsNumeric(frmMBMClinical.LblVisitLimit) = True, frmMBMClinical.LblVisitLimit, 0)
    CisDentalLmt = CDbl(IIf(IsNumeric(CisDentalLmt) = True, CisDentalLmt, 0))
    CisSpecLmt = CDbl(IIf(IsNumeric(CisSpecLmt) = True, CisSpecLmt, 0))
    CisMepLmt = CDbl(IIf(IsNumeric(CisMepLmt) = True, CisMepLmt, 0))
    tmpCardType = Replace(frmMBMClinical.CboCardType, "'", "''")

    ClnOthersSql = "Insert Into MBMCLNOthers(MBMNumber, MBMPolNumber, PAYCode, " & _
                 "PLNType, InsType, MBMPolEffDate, MBMPolExpDate, MBMRemarks, MBMCLNWarranty," & _
                 "MBMCLNPremium, MBMCLNBasicPrem, MBMCLNMCOFee, MBMCLNBasicMCO, " & _
                 "MBMCLNMemberType, MBMCLNExpiryStatus, MBMCLNDataStatus, MBMCLNStatus, MBMCardType," & _
                 "MBMCLNVisitLmt, MBMCLNBalLmt, MBMMEPLmt, MBMDentalLmt, MBMSpecLmt) " & _
                 "Values('" & MembershipNo & "','" & Trim(txtMBMPolicyNo) & "', '" & TmpCisPayor & "'," & _
                 "'" & CISPLN & "','" & CisInsType & "','" & tmpCisEffDate & "'," & _
                 "'" & tmpCisExpDate & "','" & Trim(tmpRemark) & "','" & Trim(TmpCisWarr) & "'," & _
                 "" & tmpCisPrem & "," & tmpCisBPrem & "," & tmpCISMCO & "," & tmpCisBMCO & "," & _
                 "'" & TmpMBMType & "','N','A','A'" & ",'" & Trim(tmpCardType) & "'," & _
                 "" & CisVisitLmt & "," & CisBalLmt & "," & CisMepLmt & "," & CisDentalLmt & "," & CisSpecLmt & ")"
                 
    medixmasterconn.Execute ClnOthersSql
    
End Sub

Private Function GetSuppExpDate(DateOfBirth, tmpPlanCode, TmpEffDate, TmpExpDate As String)
Dim CheckYr As Long
Dim PLNRec As New ADODB.Recordset
Dim PLNSql As String
Dim PlnAgeLmt As String
Dim SuppDobExpDate As String

    PlnAgeLmt = ""
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
    If Trim(PlnAgeLmt) <> "" Then
        If CheckYr = PlnAgeLmt Then
            If CDate(SuppDobExpDate) <= CDate(TmpExpDate) Then
                GetSuppExpDate = CDate(SuppDobExpDate)
            End If
        End If
    End If

End Function

Private Sub SaveMBMLifeTimeLmt(LMBMNumber, LMBMCovID, LmtType As String)
Dim TmpSql As String
Dim TmpLifeSql As String
Dim TmpLimit As Double
Dim TmpLmtVar As String
Dim cmd As New ADODB.Command
Dim AccumAmtRec As New ADODB.Recordset
Dim Prm1 As New ADODB.Parameter
Dim Prm2 As New ADODB.Parameter
Dim TmpSql2 As Integer
Dim tmpAccAmt As Double
    TmpLimit = 0
    tmpAccAmt = 0

    TmpSql2 = 0
    If Trim(LMBMCovID) <> "00" Then
        TmpSql2 = 1
    End If
    TmpSql = " And MBMTopupNumber ='" & Trim(LMBMNumber) & "' AND MBMPatientCovID='" & Trim(LMBMCovID) & "'"
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchAccAppAmt"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 15
    Set Prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 255, TmpSql)
    cmd.Parameters.Append Prm1
    Set Prm2 = cmd.CreateParameter("strAppendCase", adVarChar, adParamInput, 255, TmpSql2)
    cmd.Parameters.Append Prm2
    AccumAmtRec.CursorLocation = adUseClient
    AccumAmtRec.CursorType = adOpenForwardOnly
    AccumAmtRec.CacheSize = 100
    Set AccumAmtRec = cmd.Execute
    Do While Not AccumAmtRec.EOF
        With AccumAmtRec
            tmpAccAmt = IIf(IsNumeric(!CLMTotalApproved), !CLMTotalApproved, 0)
            AccumAmtRec.MoveNext
        End With
    Loop
    AccumAmtRec.Close
    Set AccumAmtRec = Nothing

    TmpLmtVar = "SuppLifeTimeLimit"
    If LmtType = "A" Then
        TmpLmtVar = "LifeTimeLimit"
    End If
    TmpLimit = GetLifeTimeLimit(TmpLmtVar, Trim(cboINSCode), Trim(cboPLNCode), Trim(txtHLTCode), CDate(txtMBMPayorEffDate))
    TmpLimit = CDbl(TmpLimit) - tmpAccAmt
    TmpLifeSql = "Insert Into MBMLifeTime(MBMNumber, MBMCOVID, MBMLifeTimeLmt) " & _
          "Values('" & LMBMNumber & "','" & LMBMCovID & "'," & TmpLimit & ")"
    medixmasterconn.Execute TmpLifeSql
End Sub

Private Sub SaveMBMBnfAnnLmt(LMBMNumber, LMBMCovID, LmtType, tmpKidneyAnnLmt, tmpCancerAnnLmt, tmpHomeNursAnnLmt)
Dim TmpBnfLmtSql As String
    
    TmpBnfLmtSql = "Insert Into MBMBnfLmt(MBMNumber, MBMCoveredID, MBMKidneyAnnLmt, MBMKidneyBalLmt, " & _
                 "MBMCancerAnnLmt, MBMCancerBalLmt, MBMHomeNCAnnLmt, MBMHomeNCBalLmt) " & _
                 "Values('" & LMBMNumber & "','" & LMBMCovID & "'," & tmpKidneyAnnLmt & "," & tmpKidneyAnnLmt & _
                 "," & tmpCancerAnnLmt & "," & tmpCancerAnnLmt & "," & tmpHomeNursAnnLmt & "," & tmpHomeNursAnnLmt & ")"
    medixmasterconn.Execute TmpBnfLmtSql

End Sub

Private Sub showClinicalInfo()
Dim tmpMBMNo As String
Dim MBMClnOth As New ADODB.Recordset
Dim ClnOthSql As String

    tmpMBMNo = txtMBMPrevMEMNo
    ClnOthSql = "SELECT PLNType, INSType, MBMPolEffDate, MBMPolExpDate, MBMCLNWarranty, MBMRemarks " & _
              " From MBMCLNOthers where MBMNumber='" & Trim(tmpMBMNo) & "'"
    MBMClnOth.Open ClnOthSql, medixmasterconn, adOpenKeyset, adLockOptimistic
    If MBMClnOth.recordCount Then
        frmMBMClinical.cboPLNCode.Text = MBMClnOth!PLNType
        frmMBMClinical.cboINSCode.Text = MBMClnOth!InsType
        frmMBMClinical.txtMBMPayorEffDate = MBMClnOth!MBMPolEffDate
        frmMBMClinical.txtMBMPayorExpDate = MBMClnOth!MBMPolExpDate
        frmMBMClinical.txtWarranty = IIf(Len(MBMClnOth!MBMCLNWarranty), MBMClnOth!MBMCLNWarranty, "")
        frmMBMClinical.txtRemarks = IIf(Len(MBMClnOth!MBMRemarks), MBMClnOth!MBMRemarks, "")
    End If
    MBMClnOth.Close
    Set MBMClnOth = Nothing
End Sub

Private Sub PrincipalLifeTimeChecking()
Dim tmpSuppBNFStatus  As String
Dim TmpSql2 As String
Dim TmpLMBM As New ADODB.Recordset
Dim MBMAvailLmt As Double
Dim tmpLmt As Double
    If Trim(TempLifeTimeStatus) = "Y" Then
        TmpSql2 = "Select top 1 MBMNumber, MBMTopupNumber, MBMAvailableLimit From MBM where MBMPolicyNo='" & Trim(txtMBMPolicyNo) & _
                  "' And MBMStatus = 'A' AND PayCode ='" & Trim(txtPAYCode) & "' AND INSCode='" & Trim(cboINSCode) & _
                  "' ORDER BY MBMPayorEffDate , MBMNumber "
        TmpLMBM.Open TmpSql2, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        TmpTopupNo = MembershipNo
        PrevLifeMBMNo = MembershipNo

        If TmpLMBM.EOF = False Then
            TmpTopupNo = TmpLMBM!MBMTopupNumber
            PrevLifeMBMNo = TmpLMBM!MBMTopupNumber
            If Mid(CboGuaRenewal, 1, 1) <> "3" Then
                MBMAvailLmt = IIf(IsNumeric(TmpLMBM!MBMAvailableLimit), TmpLMBM!MBMAvailableLimit, 0)
                txtAnnualLimit = MBMAvailLmt
            End If
        End If
        TmpLMBM.Close
        Set TmpLMBM = Nothing
        
        TmpSql2 = "Select MBMNumber, MBMLifeTimeLmt From MBMLifeTime where MBMNumber='" & Trim(PrevLifeMBMNo) & _
                  "' AND MBMCovID='00'"
        TmpLMBM.Open TmpSql2, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        If TmpLMBM.EOF Then
            If Mid(CboGuaRenewal, 1, 1) = "3" Then
                WLifeStatus = True
            End If
        End If
        TmpLMBM.Close
        Set TmpLMBM = Nothing
    End If
End Sub



