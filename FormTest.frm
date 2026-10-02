VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Height          =   3300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11520
      Begin VB.TextBox txtSuppPayPrem 
         Height          =   285
         Left            =   1275
         TabIndex        =   39
         Top             =   2920
         Width           =   1020
      End
      Begin VB.TextBox txtSuppPrem 
         Height          =   285
         Left            =   3405
         TabIndex        =   38
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtCovAnnualLimit 
         Height          =   285
         Left            =   3405
         TabIndex        =   37
         Top             =   2400
         Width           =   975
      End
      Begin VB.ComboBox cboSmoking 
         Height          =   315
         Left            =   3405
         TabIndex        =   36
         Top             =   2130
         Width           =   975
      End
      Begin VB.ComboBox cboAlcohol 
         Height          =   315
         Left            =   3405
         TabIndex        =   35
         Top             =   1845
         Width           =   975
      End
      Begin VB.ComboBox cboWorkNature 
         Height          =   315
         Left            =   1275
         TabIndex        =   34
         Top             =   2640
         Width           =   1020
      End
      Begin VB.ComboBox cboOccupation 
         Height          =   315
         Left            =   1275
         TabIndex        =   33
         Top             =   2400
         Width           =   1020
      End
      Begin VB.ComboBox cboCovBlood 
         Height          =   315
         Left            =   1275
         TabIndex        =   32
         Top             =   2120
         Width           =   1020
      End
      Begin VB.TextBox txtPrevINSCom 
         Height          =   330
         Left            =   1280
         MaxLength       =   60
         TabIndex        =   31
         Top             =   1815
         Width           =   1020
      End
      Begin VB.TextBox txtRBAmt 
         Height          =   330
         Left            =   3405
         MaxLength       =   8
         TabIndex        =   30
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtPrevPLNCode 
         Height          =   285
         Left            =   1280
         MaxLength       =   15
         TabIndex        =   29
         Top             =   1565
         Width           =   1020
      End
      Begin VB.ComboBox cboCovAdultChild 
         Height          =   315
         Left            =   3405
         TabIndex        =   28
         Text            =   "A"
         Top             =   1275
         Width           =   975
      End
      Begin VB.ComboBox cboCovSex 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1280
         TabIndex        =   27
         Text            =   "M"
         Top             =   1275
         Width           =   1020
      End
      Begin VB.ComboBox cboCovRelationship 
         Height          =   315
         Left            =   1280
         TabIndex        =   26
         Text            =   "W- Wife"
         Top             =   990
         Width           =   3105
      End
      Begin VB.TextBox txtCovExclusion 
         Height          =   855
         Left            =   5715
         MaxLength       =   2500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   2390
         Width           =   2700
      End
      Begin VB.TextBox txtCovAllergic 
         Height          =   495
         Left            =   5715
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   1920
         Width           =   2685
      End
      Begin VB.TextBox txtCOVPayRemarks 
         Height          =   405
         Left            =   5715
         MaxLength       =   60
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   1560
         Width           =   2685
      End
      Begin VB.ComboBox cboSuppStaffDisc 
         Height          =   315
         Left            =   5720
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   1280
         Width           =   2700
      End
      Begin VB.TextBox TxtSuppSpeWP 
         Height          =   300
         Left            =   5720
         MaxLength       =   9
         TabIndex        =   17
         Top             =   1010
         Width           =   945
      End
      Begin VB.TextBox txtSuppPREWP 
         Height          =   285
         Left            =   5720
         MaxLength       =   9
         TabIndex        =   16
         Top             =   755
         Width           =   945
      End
      Begin VB.TextBox txtSuppGenWP 
         Height          =   285
         Left            =   5720
         MaxLength       =   9
         TabIndex        =   15
         Top             =   500
         Width           =   945
      End
      Begin VB.Frame Frame9 
         Height          =   1695
         Left            =   10200
         TabIndex        =   10
         Top             =   240
         Width           =   1215
         Begin VB.CommandButton cmdCovClear 
            Caption         =   "Clear"
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   1240
            Width           =   960
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   180
            Width           =   960
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "&Remove"
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   880
            Width           =   960
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "&Update"
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   520
            Width           =   960
         End
      End
      Begin VB.TextBox txtCovIcBcPp 
         Height          =   285
         Left            =   1280
         MaxLength       =   30
         TabIndex        =   9
         Top             =   480
         Width           =   3105
      End
      Begin VB.TextBox txtCovName 
         Height          =   330
         Left            =   3360
         MaxLength       =   60
         TabIndex        =   8
         Top             =   202
         Width           =   5055
      End
      Begin VB.ComboBox cboCovSalutation 
         Height          =   315
         Left            =   1280
         TabIndex        =   7
         Text            =   "MS"
         Top             =   195
         Width           =   915
      End
      Begin VB.Frame Frame12 
         Height          =   795
         Left            =   8520
         TabIndex        =   2
         Top             =   360
         Width           =   1575
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
            Left            =   480
            TabIndex        =   3
            Top             =   150
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
            Left            =   480
            TabIndex        =   4
            Top             =   405
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label92 
            Alignment       =   2  'Center
            Caption         =   "To"
            Height          =   255
            Left            =   80
            TabIndex        =   6
            Top             =   430
            Width           =   285
         End
         Begin VB.Label Label102 
            Alignment       =   2  'Center
            Caption         =   "From"
            Height          =   255
            Left            =   80
            TabIndex        =   5
            Top             =   120
            Width           =   405
         End
      End
      Begin VB.TextBox txtCLNCovExclusion 
         Enabled         =   0   'False
         Height          =   855
         Left            =   8520
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   2400
         Width           =   2895
      End
      Begin MSMask.MaskEdBox txtCovBirthdate 
         Height          =   285
         Left            =   1275
         TabIndex        =   22
         Top             =   735
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   503
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
         TabIndex        =   23
         Top             =   500
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
         Height          =   285
         Left            =   7320
         TabIndex        =   24
         Top             =   780
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
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
         TabIndex        =   25
         Top             =   1035
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   476
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label51 
         Caption         =   "Adult/Child"
         Height          =   240
         Left            =   2400
         TabIndex        =   69
         Top             =   1330
         Width           =   825
      End
      Begin VB.Label Label45 
         Caption         =   "Supp. Limit"
         Height          =   255
         Left            =   2400
         TabIndex        =   68
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label Label53 
         Caption         =   "Blood"
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label Label52 
         Caption         =   "R'Ship"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   1040
         Width           =   1005
      End
      Begin VB.Label Label49 
         Caption         =   "D.O.B"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   800
         Width           =   1005
      End
      Begin VB.Label Label48 
         Caption         =   "IC Num"
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   500
         Width           =   1125
      End
      Begin VB.Label Label47 
         Caption         =   "Salutation"
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label50 
         Caption         =   "Sex (M/F)"
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   1330
         Width           =   825
      End
      Begin VB.Label Label46 
         Caption         =   "Work Nature"
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label58 
         Caption         =   "Smoking"
         Height          =   255
         Left            =   2400
         TabIndex        =   60
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label62 
         Caption         =   "Alcohol"
         Height          =   240
         Left            =   2400
         TabIndex        =   59
         Top             =   1875
         Width           =   1095
      End
      Begin VB.Label Label69 
         Caption         =   "Mem Name"
         Height          =   255
         Left            =   2400
         TabIndex        =   58
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label60 
         Caption         =   "Allergic"
         Height          =   255
         Left            =   4560
         TabIndex        =   57
         Top             =   1920
         Width           =   870
      End
      Begin VB.Label Label61 
         Caption         =   "Supp. Warranty "
         Height          =   255
         Left            =   4560
         TabIndex        =   56
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label28 
         Caption         =   "Occupation"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label71 
         Caption         =   "Supp. Prem"
         Height          =   255
         Left            =   2400
         TabIndex        =   54
         Top             =   2640
         Width           =   1005
      End
      Begin VB.Label Label72 
         Caption         =   "T/O Pln Code"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   1600
         Width           =   1125
      End
      Begin VB.Label Label74 
         Caption         =   "Prev Ins Com"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   1800
         Width           =   1140
      End
      Begin VB.Label Label75 
         Caption         =   "R&&B Amount"
         Height          =   255
         Left            =   2400
         TabIndex        =   51
         Top             =   1600
         Width           =   1140
      End
      Begin VB.Label Label80 
         Caption         =   "Payor Remarks"
         Height          =   255
         Left            =   4560
         TabIndex        =   50
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label83 
         Caption         =   "Spe. Ill. W. P."
         Height          =   255
         Left            =   4560
         TabIndex        =   49
         Top             =   1080
         Width           =   2025
      End
      Begin VB.Label Label84 
         Caption         =   "Pre W. P."
         Height          =   255
         Left            =   4560
         TabIndex        =   48
         Top             =   795
         Width           =   1380
      End
      Begin VB.Label Label85 
         Caption         =   "Gen W. P."
         Height          =   255
         Left            =   4560
         TabIndex        =   47
         Top             =   525
         Width           =   1620
      End
      Begin VB.Label Label86 
         Caption         =   "EOP"
         Height          =   255
         Left            =   6840
         TabIndex        =   46
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label87 
         Caption         =   "EOP"
         Height          =   255
         Left            =   6840
         TabIndex        =   45
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label88 
         Caption         =   "EOP"
         Height          =   255
         Left            =   6840
         TabIndex        =   44
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label89 
         Caption         =   "Staff Disc Req"
         Height          =   255
         Left            =   4560
         TabIndex        =   43
         Top             =   1320
         Width           =   1065
      End
      Begin VB.Label Label93 
         Caption         =   "Clm Non-Pay Date"
         Height          =   375
         Left            =   8520
         TabIndex        =   42
         Top             =   200
         Width           =   1365
      End
      Begin VB.Label Label108 
         Caption         =   "Supp. Clinic Warranty"
         Height          =   255
         Left            =   8520
         TabIndex        =   41
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label111 
         Caption         =   "Payor Prem"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   2880
         Width           =   1005
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
