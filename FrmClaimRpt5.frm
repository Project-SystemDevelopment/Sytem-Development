VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmClaimRpt5 
   Caption         =   "Claim Report 5"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Caption         =   "Reports Selection"
      Height          =   3240
      Left            =   8160
      TabIndex        =   0
      Top             =   3780
      Width           =   3615
      Begin VB.OptionButton Option37 
         Caption         =   "By Plan No"
         Height          =   195
         Left            =   360
         TabIndex        =   123
         Top             =   1320
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton Option35 
         Caption         =   "Detail Listing"
         Height          =   195
         Left            =   360
         TabIndex        =   122
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton Option34 
         Caption         =   "General Summary"
         Height          =   195
         Left            =   360
         TabIndex        =   121
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.OptionButton Option33 
         Caption         =   "Top 20 by Cases"
         Height          =   195
         Left            =   360
         TabIndex        =   120
         Top             =   840
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.OptionButton Option29 
         Caption         =   "Top 20 by Amount"
         Height          =   195
         Left            =   360
         TabIndex        =   119
         Top             =   1080
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.OptionButton Option32 
         Caption         =   "By Nature of Claims"
         Height          =   195
         Left            =   360
         TabIndex        =   118
         Top             =   2040
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton Option38 
         Caption         =   "Ageband"
         Height          =   195
         Left            =   360
         TabIndex        =   117
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton Option30 
         Caption         =   "By Nature of Claims"
         Height          =   195
         Left            =   360
         TabIndex        =   116
         Top             =   1560
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.OptionButton Option31 
         Caption         =   "By Nature of Claims"
         Height          =   195
         Left            =   360
         TabIndex        =   115
         Top             =   1800
         Visible         =   0   'False
         Width           =   3135
      End
   End
   Begin VB.Frame Frame22 
      Caption         =   "Selection Criteria"
      Height          =   6795
      Left            =   0
      TabIndex        =   162
      Top             =   240
      Width           =   8145
      Begin VB.CheckBox ChkCaseStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   3570
         Width           =   255
      End
      Begin VB.CheckBox ChkPolStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   3285
         Width           =   255
      End
      Begin VB.CheckBox ChkExpired 
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
         Left            =   2640
         TabIndex        =   26
         Top             =   3840
         Width           =   495
      End
      Begin VB.CheckBox ChkPreMode 
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
         Left            =   2640
         TabIndex        =   28
         Top             =   4155
         Width           =   495
      End
      Begin VB.CheckBox ChkState 
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
         Left            =   2640
         TabIndex        =   30
         Top             =   4440
         Width           =   255
      End
      Begin VB.CheckBox ChkSex 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox ChkRace 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   1900
         Width           =   255
      End
      Begin VB.CheckBox ChkMarital 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   1620
         Width           =   255
      End
      Begin VB.CheckBox ChkAdultChild 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   1340
         Width           =   255
      End
      Begin VB.CheckBox ChkINS 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   500
         Width           =   255
      End
      Begin VB.CheckBox ChkHLT 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox ChkAgentCode 
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
         Left            =   2640
         TabIndex        =   36
         Top             =   5265
         Width           =   255
      End
      Begin VB.CheckBox ChkTakeOver 
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
         Left            =   2640
         TabIndex        =   34
         Top             =   4995
         Width           =   375
      End
      Begin VB.CheckBox ChkRelation 
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
         Left            =   2640
         TabIndex        =   16
         Top             =   2450
         Width           =   495
      End
      Begin VB.CheckBox ChkDataStatus 
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
         Left            =   2640
         TabIndex        =   18
         Top             =   2745
         Width           =   495
      End
      Begin VB.CheckBox ChkRenewal 
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
         Left            =   2640
         TabIndex        =   32
         Top             =   4710
         Width           =   615
      End
      Begin VB.CheckBox ChkClaimability 
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
         Left            =   2640
         TabIndex        =   20
         Top             =   3000
         Width           =   495
      End
      Begin VB.CheckBox ChkRepCat 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7680
         TabIndex        =   81
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox ChkRep 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7680
         TabIndex        =   78
         Top             =   3120
         Width           =   255
      End
      Begin VB.CheckBox ChkFinal 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7680
         TabIndex        =   75
         Top             =   2910
         Width           =   255
      End
      Begin VB.CheckBox ChkPlan 
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
         Left            =   7680
         TabIndex        =   87
         Top             =   3840
         Width           =   375
      End
      Begin VB.CheckBox ChkEffDate 
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
         Height          =   315
         Left            =   7680
         TabIndex        =   90
         Top             =   4080
         Width           =   255
      End
      Begin VB.CheckBox ChkFirst 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7680
         TabIndex        =   72
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox ChkDoctor 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   40
         Top             =   5760
         Width           =   255
      End
      Begin VB.CheckBox ChkHOSGrp 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7750
         TabIndex        =   54
         Top             =   1040
         Width           =   255
      End
      Begin VB.CheckBox ChkHOS 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7750
         TabIndex        =   52
         Top             =   750
         Width           =   255
      End
      Begin VB.CheckBox ChkPolNo 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2640
         TabIndex        =   38
         Top             =   5550
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpCom 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7750
         TabIndex        =   50
         Top             =   500
         Width           =   255
      End
      Begin VB.CheckBox ChkPayor 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7750
         TabIndex        =   48
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox ChkDOB 
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
         Height          =   375
         Left            =   7680
         TabIndex        =   99
         Top             =   4800
         Width           =   255
      End
      Begin VB.CheckBox ChkJoinDate 
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
         Height          =   315
         Left            =   7680
         TabIndex        =   96
         Top             =   4560
         Width           =   255
      End
      Begin VB.CheckBox chkDOR 
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
         Left            =   7680
         TabIndex        =   66
         Top             =   2100
         Width           =   255
      End
      Begin VB.CheckBox chkExpDate 
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
         Height          =   315
         Left            =   7680
         TabIndex        =   93
         Top             =   4320
         Width           =   255
      End
      Begin VB.CheckBox ChkDOD 
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
         Left            =   7680
         TabIndex        =   63
         Top             =   1815
         Width           =   255
      End
      Begin VB.CheckBox ChkDOA 
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
         Left            =   7680
         TabIndex        =   60
         Top             =   1575
         Width           =   255
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "S - Sihat"
         Top             =   200
         Width           =   1320
      End
      Begin VB.CheckBox ChkClaimNature 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7680
         TabIndex        =   84
         Top             =   3600
         Width           =   255
      End
      Begin VB.CheckBox ChkMBMNo 
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
         Height          =   240
         Left            =   7680
         TabIndex        =   102
         Top             =   5160
         Width           =   255
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
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   47
         Top             =   150
         Width           =   3360
      End
      Begin VB.CheckBox ChkPolYear 
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
         Height          =   240
         Left            =   2640
         TabIndex        =   46
         Top             =   6480
         Width           =   255
      End
      Begin VB.CheckBox ChkICNo 
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
         Left            =   7680
         TabIndex        =   107
         Top             =   5640
         Width           =   375
      End
      Begin VB.CheckBox ChkDOW 
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
         Left            =   7680
         TabIndex        =   69
         Top             =   2380
         Width           =   255
      End
      Begin VB.ComboBox CboInsType 
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
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   1320
      End
      Begin VB.ComboBox CboMemType 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1320
      End
      Begin VB.ComboBox cboSuppType 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1000
         Width           =   1320
      End
      Begin VB.ComboBox CboAdultChild 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   1290
         Width           =   1320
      End
      Begin VB.ComboBox CboMarital 
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   1580
         Width           =   1320
      End
      Begin VB.ComboBox CboRace 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   1860
         Width           =   1320
      End
      Begin VB.ComboBox CboSex 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   2150
         Width           =   1320
      End
      Begin VB.ComboBox CboRelation 
         Height          =   315
         Left            =   1200
         TabIndex        =   15
         Top             =   2430
         Width           =   1320
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
         Left            =   1200
         TabIndex        =   17
         Top             =   2715
         Width           =   1320
      End
      Begin VB.ComboBox CboClaimability 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   3000
         Width           =   1320
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
         Left            =   1200
         TabIndex        =   21
         Top             =   3255
         Width           =   1320
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
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   3540
         Width           =   1320
      End
      Begin VB.ComboBox CboExpired 
         Height          =   315
         Left            =   1200
         TabIndex        =   25
         Top             =   3840
         Width           =   1320
      End
      Begin VB.ComboBox CboPreMode 
         Height          =   315
         Left            =   1200
         TabIndex        =   27
         Top             =   4125
         Width           =   1320
      End
      Begin VB.ComboBox CboState 
         Height          =   315
         Left            =   1200
         TabIndex        =   29
         Top             =   4410
         Width           =   1320
      End
      Begin VB.ComboBox cboRenewalInd 
         Height          =   315
         Left            =   1200
         TabIndex        =   31
         Top             =   4680
         Width           =   1320
      End
      Begin VB.ComboBox cboTakeOver 
         Height          =   315
         Left            =   1200
         TabIndex        =   33
         Top             =   4965
         Width           =   1320
      End
      Begin VB.TextBox txtAgentCode 
         Height          =   285
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   35
         Top             =   5250
         Width           =   1320
      End
      Begin VB.TextBox txtPolNo 
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   37
         Top             =   5505
         Width           =   1320
      End
      Begin VB.TextBox txtDOCName 
         Height          =   285
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   39
         Top             =   5730
         Width           =   1320
      End
      Begin VB.TextBox TxtPatientName 
         Height          =   285
         Left            =   1200
         MaxLength       =   60
         TabIndex        =   41
         Top             =   5985
         Width           =   1320
      End
      Begin VB.TextBox TxtAgeband1 
         Height          =   285
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   42
         Top             =   6240
         Width           =   525
      End
      Begin VB.TextBox TxtAgeband2 
         Height          =   285
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   43
         Top             =   6240
         Width           =   480
      End
      Begin VB.TextBox txtPOLYear1 
         Height          =   285
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   44
         Top             =   6480
         Width           =   525
      End
      Begin VB.TextBox txtPOLYear2 
         Height          =   285
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   45
         Top             =   6480
         Width           =   480
      End
      Begin VB.ComboBox CboGrpCom 
         Height          =   315
         Left            =   4320
         TabIndex        =   49
         Top             =   440
         Width           =   3360
      End
      Begin VB.ComboBox cboHOSName 
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
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   51
         Top             =   720
         Width           =   3360
      End
      Begin VB.ComboBox cboHOSGrp 
         Height          =   315
         Left            =   4320
         TabIndex        =   53
         Top             =   980
         Width           =   3360
      End
      Begin VB.ComboBox CboStayType 
         Height          =   315
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   55
         Top             =   1260
         Width           =   1200
      End
      Begin VB.ComboBox CboCashType 
         Height          =   315
         Left            =   6360
         Sorted          =   -1  'True
         TabIndex        =   56
         Top             =   1260
         Width           =   1320
      End
      Begin VB.TextBox TxtYear 
         Height          =   315
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   58
         Top             =   1515
         Width           =   1200
      End
      Begin MSMask.MaskEdBox TxtDOA2 
         Height          =   315
         Left            =   6240
         TabIndex        =   59
         Top             =   1515
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOA1 
         Height          =   315
         Left            =   4320
         TabIndex        =   57
         Top             =   1515
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOD1 
         Height          =   315
         Left            =   4320
         TabIndex        =   61
         Top             =   1800
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOR1 
         Height          =   315
         Left            =   4320
         TabIndex        =   64
         Top             =   2090
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOW1 
         Height          =   315
         Left            =   4320
         TabIndex        =   67
         Top             =   2370
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPreDiag1 
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   70
         Top             =   2610
         Width           =   1200
      End
      Begin MSMask.MaskEdBox txtDOD2 
         Height          =   315
         Left            =   6240
         TabIndex        =   62
         Top             =   1800
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOR2 
         Height          =   315
         Left            =   6240
         TabIndex        =   65
         Top             =   2090
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOW2 
         Height          =   315
         Left            =   6240
         TabIndex        =   68
         Top             =   2370
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPreDiag2 
         Height          =   285
         Left            =   6240
         MaxLength       =   3
         TabIndex        =   71
         Top             =   2610
         Width           =   1200
      End
      Begin VB.TextBox txtFinalDiag1 
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   73
         Top             =   2865
         Width           =   1200
      End
      Begin VB.TextBox txtRepudCat1 
         Height          =   285
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   76
         Top             =   3120
         Width           =   1200
      End
      Begin VB.TextBox txtRepudCode1 
         Height          =   285
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   79
         Top             =   3375
         Width           =   1200
      End
      Begin VB.TextBox TxtClaimNature1 
         Height          =   285
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   82
         Top             =   3615
         Width           =   1200
      End
      Begin VB.TextBox TxtPlan1 
         Height          =   315
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   85
         Top             =   3840
         Width           =   1200
      End
      Begin MSMask.MaskEdBox TxtPayEffDate1 
         Height          =   315
         Left            =   4320
         TabIndex        =   88
         Top             =   4080
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPayExpDate1 
         Height          =   315
         Left            =   4320
         TabIndex        =   91
         Top             =   4365
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDateJoin1 
         Height          =   315
         Left            =   4320
         TabIndex        =   94
         Top             =   4650
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOB1 
         Height          =   315
         Left            =   4320
         TabIndex        =   97
         Top             =   4920
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtMBMNo1 
         Height          =   285
         Left            =   4320
         MaxLength       =   15
         TabIndex        =   100
         Top             =   5160
         Width           =   1200
      End
      Begin VB.TextBox TxtCaseNo1 
         Height          =   285
         Left            =   4320
         MaxLength       =   15
         TabIndex        =   103
         Top             =   5400
         Width           =   1200
      End
      Begin VB.TextBox TxtICNo1 
         Height          =   285
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   105
         Top             =   5650
         Width           =   1200
      End
      Begin VB.TextBox txtFinalDiag2 
         Height          =   285
         Left            =   6240
         MaxLength       =   3
         TabIndex        =   74
         Top             =   2865
         Width           =   1200
      End
      Begin VB.TextBox txtRepudCat2 
         Height          =   285
         Left            =   6240
         MaxLength       =   4
         TabIndex        =   77
         Top             =   3120
         Width           =   1200
      End
      Begin VB.TextBox txtRepudCode2 
         Height          =   285
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   80
         Top             =   3375
         Width           =   1200
      End
      Begin VB.TextBox TxtClaimNature2 
         Height          =   285
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   83
         Top             =   3615
         Width           =   1200
      End
      Begin VB.TextBox TxtPlan2 
         Height          =   315
         Left            =   6240
         MaxLength       =   4
         TabIndex        =   86
         Top             =   3840
         Width           =   1200
      End
      Begin MSMask.MaskEdBox TxtPayEffDate2 
         Height          =   315
         Left            =   6240
         TabIndex        =   89
         Top             =   4080
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPayExpDate2 
         Height          =   315
         Left            =   6240
         TabIndex        =   92
         Top             =   4365
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDateJoin2 
         Height          =   315
         Left            =   6240
         TabIndex        =   95
         Top             =   4650
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOB2 
         Height          =   315
         Left            =   6240
         TabIndex        =   98
         Top             =   4920
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtMBMNo2 
         Height          =   285
         Left            =   6240
         MaxLength       =   15
         TabIndex        =   101
         Top             =   5160
         Width           =   1200
      End
      Begin VB.TextBox TxtCaseNo2 
         Height          =   285
         Left            =   6240
         MaxLength       =   15
         TabIndex        =   104
         Top             =   5400
         Width           =   1200
      End
      Begin VB.TextBox TxtICNo2 
         Height          =   285
         Left            =   6240
         MaxLength       =   20
         TabIndex        =   106
         Top             =   5650
         Width           =   1200
      End
      Begin VB.TextBox txtActualAmt1 
         Height          =   285
         Left            =   4320
         TabIndex        =   108
         Top             =   5880
         Width           =   1215
      End
      Begin VB.TextBox txtActualAmt2 
         Height          =   285
         Left            =   6240
         TabIndex        =   109
         Top             =   5880
         Width           =   1215
      End
      Begin VB.TextBox txtAppAmt1 
         Height          =   285
         Left            =   4320
         TabIndex        =   110
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox txtAppAmt2 
         Height          =   285
         Left            =   6240
         TabIndex        =   111
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox TxtNoofDay2 
         Height          =   285
         Left            =   6240
         MaxLength       =   4
         TabIndex        =   113
         Top             =   6360
         Width           =   1200
      End
      Begin VB.TextBox TxtNoofDay1 
         Height          =   285
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   112
         Top             =   6360
         Width           =   1200
      End
      Begin VB.Label Label68 
         Caption         =   "No of Day"
         Height          =   255
         Left            =   3240
         TabIndex        =   238
         Top             =   6465
         Width           =   975
      End
      Begin VB.Label Label69 
         Caption         =   "To"
         Height          =   255
         Left            =   5880
         TabIndex        =   237
         Top             =   6465
         Width           =   375
      End
      Begin VB.Label Label80 
         Caption         =   "To"
         Height          =   255
         Left            =   5880
         TabIndex        =   236
         ToolTipText     =   "Approved Amt"
         Top             =   6200
         Width           =   405
      End
      Begin VB.Label Label79 
         Caption         =   "App Amt"
         Height          =   255
         Left            =   3240
         TabIndex        =   235
         ToolTipText     =   "Approved Amt"
         Top             =   6200
         Width           =   645
      End
      Begin VB.Label Label78 
         Caption         =   "To"
         Height          =   255
         Left            =   5880
         TabIndex        =   234
         ToolTipText     =   "Approved Amt"
         Top             =   5950
         Width           =   405
      End
      Begin VB.Label Label77 
         Caption         =   "Act Amt"
         Height          =   255
         Left            =   3240
         TabIndex        =   233
         ToolTipText     =   "Approved Amt"
         Top             =   5950
         Width           =   645
      End
      Begin VB.Label Label62 
         Caption         =   "Patient Name"
         Height          =   255
         Left            =   120
         TabIndex        =   231
         Top             =   6030
         Width           =   1095
      End
      Begin VB.Label Label59 
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
         Height          =   255
         Left            =   2640
         TabIndex        =   230
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label58 
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
         Height          =   255
         Left            =   7800
         TabIndex        =   229
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   1680
         TabIndex        =   228
         Top             =   6570
         Width           =   375
      End
      Begin VB.Label Label56 
         Caption         =   "Pol. Year"
         Height          =   255
         Left            =   120
         TabIndex        =   227
         Top             =   6570
         Width           =   855
      End
      Begin VB.Label Label55 
         Caption         =   "Agent Code"
         Height          =   255
         Left            =   120
         TabIndex        =   226
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label45 
         Caption         =   "Take Over"
         Height          =   255
         Left            =   120
         TabIndex        =   225
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label44 
         Caption         =   "Renewal "
         Height          =   255
         Left            =   120
         TabIndex        =   224
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   223
         Top             =   4920
         Width           =   375
      End
      Begin VB.Label Label42 
         Caption         =   "DOB"
         Height          =   255
         Left            =   3240
         TabIndex        =   222
         ToolTipText     =   "Date of Birth"
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   221
         Top             =   4680
         Width           =   375
      End
      Begin VB.Label Label40 
         Caption         =   "Date Join"
         Height          =   255
         Left            =   3240
         TabIndex        =   220
         ToolTipText     =   "Payor Expiry Date"
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Label39 
         Caption         =   "Doctor"
         Height          =   255
         Left            =   120
         TabIndex        =   219
         Top             =   5760
         Width           =   1095
      End
      Begin VB.Label Label37 
         Caption         =   "Hosp. Grp"
         Height          =   255
         Left            =   3240
         TabIndex        =   218
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   3240
         TabIndex        =   217
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   216
         Top             =   3495
         Width           =   375
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   215
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   214
         Top             =   2970
         Width           =   375
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   213
         Top             =   2655
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Repud Code"
         Height          =   255
         Left            =   3240
         TabIndex        =   212
         Top             =   3495
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Repud Cat"
         Height          =   255
         Left            =   3240
         TabIndex        =   211
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Final Diag"
         Height          =   255
         Left            =   3240
         TabIndex        =   210
         Top             =   2970
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Pre. Diag"
         Height          =   255
         Left            =   3240
         TabIndex        =   209
         Top             =   2655
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Grp Company"
         Height          =   255
         Left            =   3240
         TabIndex        =   208
         Top             =   540
         Width           =   1095
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
         Left            =   120
         TabIndex        =   207
         Top             =   3600
         Width           =   945
      End
      Begin VB.Label Label53 
         Caption         =   "Claimability"
         Height          =   255
         Left            =   120
         TabIndex        =   206
         Top             =   3000
         Width           =   975
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
         Left            =   120
         TabIndex        =   205
         Top             =   3360
         Width           =   1020
      End
      Begin VB.Label Label51 
         Caption         =   "Plan Code"
         Height          =   255
         Left            =   3240
         TabIndex        =   204
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   203
         Top             =   3960
         Width           =   375
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
         Left            =   120
         TabIndex        =   202
         Top             =   2760
         Width           =   1020
      End
      Begin VB.Label Label38 
         Caption         =   "Payor"
         Height          =   225
         Left            =   3240
         TabIndex        =   201
         Top             =   240
         Width           =   975
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
         TabIndex        =   200
         Top             =   480
         Width           =   690
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
         TabIndex        =   199
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label7 
         Caption         =   "Eff Date"
         Height          =   255
         Left            =   3240
         TabIndex        =   198
         ToolTipText     =   "Payor Effective Date"
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label23 
         Caption         =   "Exp Date"
         Height          =   255
         Left            =   3240
         TabIndex        =   197
         ToolTipText     =   "Payor Expiry Date"
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "DOR"
         Height          =   255
         Left            =   3240
         TabIndex        =   196
         ToolTipText     =   "Date of Register"
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   195
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   194
         Top             =   4440
         Width           =   375
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   193
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label47 
         Caption         =   "DOA"
         Height          =   255
         Left            =   3240
         TabIndex        =   192
         ToolTipText     =   "Date of Admission"
         Top             =   1650
         Width           =   975
      End
      Begin VB.Label Label48 
         Caption         =   "DOD"
         Height          =   255
         Left            =   3240
         TabIndex        =   191
         ToolTipText     =   "Date of Discharged"
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   190
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   189
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Sex"
         Height          =   255
         Left            =   120
         TabIndex        =   188
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Mem Type"
         Height          =   255
         Left            =   120
         TabIndex        =   187
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Adult / Child"
         Height          =   255
         Left            =   120
         TabIndex        =   186
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Marital Status"
         Height          =   255
         Left            =   120
         TabIndex        =   185
         Top             =   1725
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Race"
         Height          =   255
         Left            =   120
         TabIndex        =   184
         Top             =   2025
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Relationship"
         Height          =   255
         Left            =   120
         TabIndex        =   183
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label19 
         Height          =   255
         Left            =   2520
         TabIndex        =   182
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label25 
         Caption         =   "Expired Status"
         Height          =   255
         Left            =   120
         TabIndex        =   181
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label27 
         Caption         =   "Prem. Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   180
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "State"
         Height          =   255
         Left            =   120
         TabIndex        =   179
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label30 
         Caption         =   "Policy No."
         Height          =   255
         Left            =   120
         TabIndex        =   178
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Supp. Type"
         Height          =   255
         Left            =   120
         TabIndex        =   177
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Claim Nature"
         Height          =   255
         Left            =   3240
         TabIndex        =   176
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   195
         Left            =   5760
         TabIndex        =   175
         Top             =   3780
         Width           =   375
      End
      Begin VB.Label Label24 
         Caption         =   "Member No"
         Height          =   255
         Left            =   3240
         TabIndex        =   174
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label65 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   173
         Top             =   5160
         Width           =   375
      End
      Begin VB.Label Label66 
         Caption         =   "Claim No"
         Height          =   255
         Left            =   3240
         TabIndex        =   172
         Top             =   5450
         Width           =   1095
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   171
         Top             =   5450
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Cash Type"
         Height          =   240
         Left            =   5520
         TabIndex        =   170
         Top             =   1335
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Stay Type"
         Height          =   255
         Left            =   3240
         TabIndex        =   169
         Top             =   1335
         Width           =   1095
      End
      Begin VB.Label Label85 
         Caption         =   "I/C No"
         Height          =   255
         Left            =   3240
         TabIndex        =   168
         Top             =   5730
         Width           =   975
      End
      Begin VB.Label Label84 
         Caption         =   "To"
         Height          =   255
         Left            =   5880
         TabIndex        =   167
         Top             =   5730
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "To"
         Height          =   255
         Left            =   1800
         TabIndex        =   166
         Top             =   6320
         Width           =   375
      End
      Begin VB.Label Label74 
         Caption         =   "Age Band"
         Height          =   255
         Left            =   120
         TabIndex        =   165
         Top             =   6320
         Width           =   855
      End
      Begin VB.Label Label75 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   164
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label76 
         Caption         =   "DOW"
         Height          =   255
         Left            =   3240
         TabIndex        =   163
         ToolTipText     =   "Worksheet Date"
         Top             =   2400
         Width           =   855
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      TabIndex        =   160
      Top             =   0
      Width           =   11805
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Claims Reports 5"
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
         TabIndex        =   161
         Top             =   0
         Width           =   3285
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Reports Type"
      Height          =   3615
      Left            =   8160
      TabIndex        =   139
      Top             =   240
      Width           =   3615
      Begin VB.OptionButton Option36 
         Caption         =   "Agent No"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   158
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Doctor"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   157
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Hosp. Grp"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   156
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option23 
         Caption         =   "Acc/Illness"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   114
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option21 
         Caption         =   "Discharge"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   155
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton Option20 
         Caption         =   "Admission "
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   154
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option18 
         Caption         =   "Repudiation"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   153
         Top             =   2160
         Width           =   1455
      End
      Begin VB.OptionButton Option16 
         Caption         =   "Duration Stay"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   152
         Top             =   1920
         Width           =   1335
      End
      Begin VB.OptionButton Option15 
         Caption         =   "Diagnosis"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   151
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Hospital"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   150
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Payor"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   149
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option8 
         Caption         =   "State"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   148
         Top             =   1680
         Width           =   735
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Race"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   147
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Gender"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   146
         Top             =   2400
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Age Band"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   145
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Plan No."
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   144
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "General Listing"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   143
         Top             =   2640
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Insured Type"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   142
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option13 
         Caption         =   "I/C"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   141
         Top             =   1920
         Width           =   1575
      End
      Begin VB.OptionButton Option19 
         Caption         =   "Bordx"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   140
         Top             =   2160
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   124
      Top             =   7320
      Width           =   11775
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   6240
         TabIndex        =   134
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   375
         Left            =   7920
         TabIndex        =   133
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   375
         Left            =   7080
         TabIndex        =   132
         Top             =   160
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   131
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   2760
         TabIndex        =   130
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print to File"
         Height          =   375
         Left            =   8760
         TabIndex        =   129
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   375
         Left            =   10560
         TabIndex        =   128
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9720
         TabIndex        =   127
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   5400
         TabIndex        =   126
         Top             =   160
         Width           =   855
      End
      Begin VB.TextBox TxtDay 
         Height          =   285
         Left            =   1440
         TabIndex        =   125
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   255
         Left            =   3600
         TabIndex        =   135
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
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
         NumItems        =   32
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   22
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   24
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   25
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   26
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   27
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   28
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   29
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   30
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   31
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label64 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1560
         TabIndex        =   138
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   240
         TabIndex        =   137
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblTotalRecord 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3960
         TabIndex        =   136
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   255
      Left            =   0
      TabIndex        =   159
      Top             =   7080
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   450
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   63
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   36
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   37
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   38
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   39
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(41) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   40
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(42) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   41
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(43) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   42
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(44) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   43
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(45) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   44
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(46) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   45
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(47) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   46
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(48) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   47
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(49) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   48
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(50) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   49
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(51) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   50
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(52) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   51
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(53) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   52
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(54) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   53
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(55) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   54
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(56) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   55
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(57) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   56
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(58) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   57
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(59) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   58
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(60) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   59
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(61) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   60
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(62) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   61
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(63) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   62
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   232
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmClaimRpt5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DiagCode, PostCode, InsureCode, PlanCode, ClaimNatureCode As String
Private aa As Integer
Private AgeCode, INSCode, AgentCode, DoctorName, Hospital, HosGrp, Hospitals As String
Private Description, StateCode, ClaimType, Description1, Description2 As String
Private SexCode, RaceCode, Payor, Day As String
Private TotalDuration, TotalFollowUp, TotalPreHos, Total As String
Private StrAppend, TotalChkAmt As String
Private AdmissionDate1, AdmissionDate2, AdmissionDate3 As String
Private AdmissionDate4, AdmissionDate5, AdmissionDate6 As String
Private AdmissionDate7, AdmissionDate8, AdmissionDate9 As String
Private AdmissionDate10, AdmissionDate11, AdmissionDate12 As String
Private AdmissionDate13, AdmissionDate14, AdmissionDate15 As String
Private AdmissionDate16, AdmissionDate17, AdmissionDate18 As String
Private AdmissionDate19, AdmissionDate20, AdmissionDate21 As String
Private AdmissionDate22, AdmissionDate23, AdmissionDate24 As String
Private AdmissionDate25, AdmissionDate26, AdmissionDate27 As String
Private AdmissionDate28, AdmissionDate29, AdmissionDate30 As String
Private AdmissionDate31 As String
Private Day1, Day2, Day3, Day4, Day5, Day6, Day7 As String
Private Day8, Day9, Day10, Day11, Day12, Day13, Day14 As String
Private Day15, Day16, Day17, Day18, Day19, Day20, Day21 As String
Private Day22, Day23, Day24, Day25, Day26, Day27, Day28 As String
Private Day29, Day30, Day31 As String
Private Total1, Total2, Total3, Total4, Total5, Total6, Total7 As String
Private Total8, Total9, Total10, Total11, Total12, Total13, Total14 As String
Private Total15, Total16, Total17, Total18, Total19, Total20, Total21 As String
Private Total22, Total23, Total24, Total25, Total26, Total27, Total28 As String
Private Total29, Total30, Total31 As String
Private intExit As Integer
Private Current As String
Private GrpCompany As String
Private m_blnDirection(1 To 63) As Boolean

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


Private Sub CmdBack_Click()
    ListView2.Height = 255
    ListView2.Left = 0
    ListView2.Top = 7080
    ListView2.Width = 11775
    Frame22.Visible = True
    'Frame1.Visible = True
    Frame4.Visible = True
    Frame6.Visible = True
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    cmdExit.Enabled = True
    CmdClear.Enabled = True
    CmdPrint.Enabled = True
End Sub

Private Sub CmdView_Click()
    ListView2.Height = 6720
    ListView2.Left = 0
    ListView2.Top = 480
    ListView2.Width = 11775
    Frame22.Visible = False
    Frame4.Visible = False
    'Frame1.Visible = False
    Frame6.Visible = False
End Sub

Private Sub ListViewHeader1()
    ListView2.ColumnHeaders(1).Text = "Claim Nature"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For aa = 6 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
End Sub

Private Sub ListViewHeader15()

    If Option35.value = True Then
        ListView2.ColumnHeaders(1).Text = "Payor"
        ListView2.ColumnHeaders(2).Text = "Claim No"
        ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(3).Text = "BordxNo"
        ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(4).Text = "Bordx Date"
        ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(5).Text = "Member No"
        ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(6).Text = "Policy No"
        ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(7).Text = "Covered ID"
        ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(8).Text = "Prin. Name"
        ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(9).Text = "Patient Name"
        ListView2.ColumnHeaders(9).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(10).Text = "Prin. IC No"
        ListView2.ColumnHeaders(10).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(11).Text = "Prin. Sex"
        ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(12).Text = "Prin. Ageband"
        ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(13).Text = "Prin. DOB"
        ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(14).Text = "Patient IC No"
        ListView2.ColumnHeaders(14).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(15).Text = "Patient Sex"
        ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(16).Text = "Patient Ageband"
        ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(17).Text = "Patient DOB"
        ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(18).Text = "Insured Type"
        ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(19).Text = "Plan Code"
        ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(20).Text = "Hospital"
        ListView2.ColumnHeaders(20).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(21).Text = "First Diag."
        ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(22).Text = "Description"
        ListView2.ColumnHeaders(22).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(23).Text = "Final Diag."
        ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(24).Text = "Description"
        ListView2.ColumnHeaders(24).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(25).Text = "Bill Amt"
        ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(26).Text = "Approved Amt"
        ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(27).Text = "Reserve Amt"
        ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(28).Text = "Recovery Amt"
        ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(29).Text = "Case Status"
        ListView2.ColumnHeaders(29).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(30).Text = "Claimability"
        ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(31).Text = "Policy Status"
        ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(32).Text = "DOA"
        ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(33).Text = "DOD"
        ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(34).Text = "DOR"
        ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(35).Text = "Duration"
        ListView2.ColumnHeaders(35).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(36).Text = "Worksheet Date"
        ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(37).Text = "Payor Eff. Date"
        ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(38).Text = "Payor Exp. Date"
        ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
        For aa = 39 To 63
            ListView2.ColumnHeaders(aa).Text = ""
        Next aa
    ElseIf Option34.value = True Then
        ListView2.ColumnHeaders(1).Text = "Payor"
        ListView2.ColumnHeaders(2).Text = "Claim No"
        ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(3).Text = "BordxNo"
        ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(4).Text = "Bordx Date"
        ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(5).Text = "Member No"
        ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(6).Text = "Policy No"
        ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(7).Text = "Covered ID"
        ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(8).Text = "Prin. Name"
        ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(9).Text = "Patient Name"
        ListView2.ColumnHeaders(9).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(10).Text = "Prin. IC No"
        ListView2.ColumnHeaders(10).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(11).Text = "Prin. Sex"
        ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(12).Text = "Prin. Ageband"
        ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(13).Text = "Prin. DOB"
        ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(14).Text = "Patient IC No"
        ListView2.ColumnHeaders(14).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(15).Text = "Patient Sex"
        ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(16).Text = "Patient Ageband"
        ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(17).Text = "Patient DOB"
        ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(18).Text = "Insured Type"
        ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(19).Text = "Plan Code"
        ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(20).Text = "Hospital"
        ListView2.ColumnHeaders(20).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(21).Text = "First Diag."
        ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(22).Text = "Description"
        ListView2.ColumnHeaders(22).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(23).Text = "Final Diag."
        ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(24).Text = "Description"
        ListView2.ColumnHeaders(24).Alignment = lvwColumnLeft
        ListView2.ColumnHeaders(25).Text = "Bill Amt"
        ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(26).Text = "Approved Amt"
        ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(27).Text = "Reserve Amt"
        ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(28).Text = "Recovery Amt"
        ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(29).Text = "Case Status"
        ListView2.ColumnHeaders(29).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(30).Text = "Claimability"
        ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(31).Text = "Policy Status"
        ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(32).Text = "DOA"
        ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(33).Text = "DOD"
        ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(34).Text = "DOR"
        ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(35).Text = "Duration"
        ListView2.ColumnHeaders(35).Alignment = lvwColumnRight
        ListView2.ColumnHeaders(36).Text = "Worksheet Date"
        ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(37).Text = "Payor Eff. Date"
        ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(38).Text = "Payor Exp. Date"
        ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
        ListView2.ColumnHeaders(39).Text = "Case Remarks"
        ListView2.ColumnHeaders(39).Alignment = lvwColumnLeft
        For aa = 40 To 63
            ListView2.ColumnHeaders(aa).Text = ""
        Next aa
    End If
End Sub

Private Sub ListViewHeader16()

    ListView2.ColumnHeaders(1).Text = "Hospital Group"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Bill Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For aa = 6 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
End Sub

Private Sub ListViewHeader17()

If Option34.value = True Or Option37.value = True Or Option29.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Hospital"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Bill Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Average Bill"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Average Approved"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option33.value = True Then
    ListView2.ColumnHeaders(1).Text = "Hospital"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Bill Amt"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Approved Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Average Bill"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Average Approved"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    For aa = 7 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa

ElseIf Option30.value = True Then
    ListView2.ColumnHeaders(1).Text = "Hospital"
    ListView2.ColumnHeaders(2).Text = "Final Diag."
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Description"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Bill Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
Else

    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Hospital"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Case No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Member No"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Covered ID"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Patient Name"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Bill Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Approved Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Member Amt"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "DOA"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Text = "DOD"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Case Status"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Text = "Claimability"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Text = "Final Diag."
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    
    For aa = 15 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
End If

End Sub

Private Sub ListViewHeader2()
If Option35.value = True Then
    ListView2.ColumnHeaders(1).Text = "No"
    ListView2.ColumnHeaders(2).Text = "Claim No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Member No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Policy No"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Patient Name"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "DOA"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "Hospital Name"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(8).Text = "First Diag."
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    For aa = 9 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option34.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Date 1"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Total"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Date 2"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Total"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Date 3"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Total"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Date 4"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(9).Text = "Total"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "Date 5"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(11).Text = "Total"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "Date 6"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(13).Text = "Total"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "Date 7"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(15).Text = "Total"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "Date 8"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(17).Text = "Total"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "Date 9"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(19).Text = "Total"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "Date 10"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(21).Text = "Total"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "Date 11"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(23).Text = "Total"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "Date 12"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(25).Text = "Total"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Date 13"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(27).Text = "Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Date 14"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Date 15"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(31).Text = "Total"
    ListView2.ColumnHeaders(31).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(32).Text = "Date 16"
    ListView2.ColumnHeaders(32).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(33).Text = "Total"
    ListView2.ColumnHeaders(33).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(34).Text = "Date 17"
    ListView2.ColumnHeaders(34).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(35).Text = "Total"
    ListView2.ColumnHeaders(35).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(36).Text = "Date 18"
    ListView2.ColumnHeaders(36).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(37).Text = "Total"
    ListView2.ColumnHeaders(37).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(38).Text = "Date 19"
    ListView2.ColumnHeaders(38).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(39).Text = "Total"
    ListView2.ColumnHeaders(39).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(40).Text = "Date 20"
    ListView2.ColumnHeaders(40).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(41).Text = "Total"
    ListView2.ColumnHeaders(41).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(42).Text = "Date 21"
    ListView2.ColumnHeaders(42).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(43).Text = "Total"
    ListView2.ColumnHeaders(43).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(44).Text = "Date 22"
    ListView2.ColumnHeaders(44).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(45).Text = "Total"
    ListView2.ColumnHeaders(45).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(46).Text = "Date 23"
    ListView2.ColumnHeaders(46).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(47).Text = "Total"
    ListView2.ColumnHeaders(47).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(48).Text = "Date 24"
    ListView2.ColumnHeaders(48).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(49).Text = "Total"
    ListView2.ColumnHeaders(49).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(50).Text = "Date 25"
    ListView2.ColumnHeaders(50).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(51).Text = "Total"
    ListView2.ColumnHeaders(51).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(52).Text = "Date 26"
    ListView2.ColumnHeaders(52).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(53).Text = "Total"
    ListView2.ColumnHeaders(53).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(54).Text = "Date 27"
    ListView2.ColumnHeaders(54).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(55).Text = "Total"
    ListView2.ColumnHeaders(55).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(56).Text = "Date 28"
    ListView2.ColumnHeaders(56).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(57).Text = "Total"
    ListView2.ColumnHeaders(57).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(58).Text = "Date 29"
    ListView2.ColumnHeaders(58).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(59).Text = "Total"
    ListView2.ColumnHeaders(59).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(60).Text = "Date 30"
    ListView2.ColumnHeaders(60).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(61).Text = "Total"
    ListView2.ColumnHeaders(61).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(62).Text = "Date 31"
    ListView2.ColumnHeaders(62).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(63).Text = "Total"
    ListView2.ColumnHeaders(63).Alignment = lvwColumnRight

ElseIf Option33.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "January"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "February"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "March"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "April"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "May"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "June"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "July"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "August"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "September"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "October"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "November"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "December"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    
    For aa = 14 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
End If

End Sub

Private Sub ListViewHeader9()
If Option35.value = True Then
    ListView2.ColumnHeaders(1).Text = "No"
    ListView2.ColumnHeaders(2).Text = "Claim No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Member No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Policy No"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Patient Name"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "DOA"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "DOD"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Text = "Hospital Name"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(9).Text = "First Diag."
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    For aa = 10 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option34.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Date 1"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Total"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Date 2"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Total"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Date 3"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Total"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Date 4"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(9).Text = "Total"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "Date 5"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(11).Text = "Total"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "Date 6"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(13).Text = "Total"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "Date 7"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(15).Text = "Total"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(16).Text = "Date 8"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(17).Text = "Total"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(18).Text = "Date 9"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(19).Text = "Total"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(20).Text = "Date 10"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(21).Text = "Total"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(22).Text = "Date 11"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(23).Text = "Total"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(24).Text = "Date 12"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(25).Text = "Total"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Date 13"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(27).Text = "Total"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Date 14"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(29).Text = "Total"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Date 15"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(31).Text = "Total"
    ListView2.ColumnHeaders(31).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(32).Text = "Date 16"
    ListView2.ColumnHeaders(32).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(33).Text = "Total"
    ListView2.ColumnHeaders(33).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(34).Text = "Date 17"
    ListView2.ColumnHeaders(34).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(35).Text = "Total"
    ListView2.ColumnHeaders(35).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(36).Text = "Date 18"
    ListView2.ColumnHeaders(36).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(37).Text = "Total"
    ListView2.ColumnHeaders(37).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(38).Text = "Date 19"
    ListView2.ColumnHeaders(38).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(39).Text = "Total"
    ListView2.ColumnHeaders(39).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(40).Text = "Date 20"
    ListView2.ColumnHeaders(40).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(41).Text = "Total"
    ListView2.ColumnHeaders(41).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(42).Text = "Date 21"
    ListView2.ColumnHeaders(42).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(43).Text = "Total"
    ListView2.ColumnHeaders(43).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(44).Text = "Date 22"
    ListView2.ColumnHeaders(44).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(45).Text = "Total"
    ListView2.ColumnHeaders(45).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(46).Text = "Date 23"
    ListView2.ColumnHeaders(46).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(47).Text = "Total"
    ListView2.ColumnHeaders(47).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(48).Text = "Date 24"
    ListView2.ColumnHeaders(48).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(49).Text = "Total"
    ListView2.ColumnHeaders(49).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(50).Text = "Date 25"
    ListView2.ColumnHeaders(50).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(51).Text = "Total"
    ListView2.ColumnHeaders(51).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(52).Text = "Date 26"
    ListView2.ColumnHeaders(52).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(53).Text = "Total"
    ListView2.ColumnHeaders(53).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(54).Text = "Date 27"
    ListView2.ColumnHeaders(54).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(55).Text = "Total"
    ListView2.ColumnHeaders(55).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(56).Text = "Date 28"
    ListView2.ColumnHeaders(56).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(57).Text = "Total"
    ListView2.ColumnHeaders(57).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(58).Text = "Date 29"
    ListView2.ColumnHeaders(58).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(59).Text = "Total"
    ListView2.ColumnHeaders(59).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(60).Text = "Date 30"
    ListView2.ColumnHeaders(60).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(61).Text = "Total"
    ListView2.ColumnHeaders(61).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(62).Text = "Date 31"
    ListView2.ColumnHeaders(62).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(63).Text = "Total"
    ListView2.ColumnHeaders(63).Alignment = lvwColumnRight

ElseIf Option33.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "January"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "February"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "March"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "April"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "May"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "June"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "July"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "August"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "September"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "October"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "November"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "December"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    
    For aa = 14 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
End If

End Sub

    


Private Sub ListViewHeader4()
    
If Option35.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Agent No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Claim No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Member No"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Covered ID"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Patient Name"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Doctor"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(8).Text = "Bill Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Approved Amt"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "Recovery Amt"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "First Diag."
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Description"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(13).Text = "Final Diag."
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Text = "Description"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(15).Text = "DOA"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Text = "DOD"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Text = "Ins Code"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(18).Text = "Plan Code"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Text = "Case Status"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    For aa = 20 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
ElseIf Option34.value = True Or Option33.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Agent No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For aa = 7 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option29.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Agent No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Doctor Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
End If
End Sub

Private Sub ListViewHeader3()
If Option35.value = True Then
    ListView2.ColumnHeaders(1).Text = "Ageband"
    ListView2.ColumnHeaders(2).Text = "Total Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For aa = 6 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
ElseIf Option34.value = True Then
    ListView2.ColumnHeaders(1).Text = "Gender"
    ListView2.ColumnHeaders(2).Text = "Ageband"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For aa = 7 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
ElseIf Option33.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Ageband"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
ElseIf Option29.value = True Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Ageband"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For aa = 7 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
ElseIf Option37.value = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Plan Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Age Band"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
ElseIf Option30.value = True Then
    ListView2.ColumnHeaders(1).Text = "Nature of Claim"
    ListView2.ColumnHeaders(2).Text = "Age Band"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For aa = 7 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
End If
End Sub

Private Sub ListViewHeader11()
If Option16.value = True And Option35.value = True Then
    ListView2.ColumnHeaders(1).Text = "Duration"
    ListView2.ColumnHeaders(2).Text = "Total Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    
    For aa = 5 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa

ElseIf Option16.value = True And Option34.value = True Then
    ListView2.ColumnHeaders(1).Text = "Claim No"
    ListView2.ColumnHeaders(2).Text = "DOA"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "DOD"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Patient Name"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "Doctor Name"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "First Diag"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Text = "Description"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(9).Text = "Final Diag"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(10).Text = "Description"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnLeft
        
    For aa = 11 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
End If
End Sub

Private Sub ListViewHeader14()
If Option35.value = True Then
    ListView2.ColumnHeaders(1).Text = "Sex"
    ListView2.ColumnHeaders(2).Text = "Total Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    For aa = 6 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
ElseIf Option34.value = True Then
    ListView2.ColumnHeaders(1).Text = "Gender"
    ListView2.ColumnHeaders(2).Text = "Age Band"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For aa = 7 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
ElseIf Option33.value = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Gender"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
ElseIf Option29.value = True Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Gender"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For aa = 7 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
ElseIf Option37.value = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Plan Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Gender"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
ElseIf Option30.value = True Then
    ListView2.ColumnHeaders(1).Text = "Nature of Claim"
    ListView2.ColumnHeaders(2).Text = "Gender"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For aa = 7 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
End If
End Sub

Private Sub ListViewHeader23()
If Option35.value = True Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Total Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
        
    For aa = 6 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa

ElseIf Option34.value = True Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Age Band"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For aa = 7 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
ElseIf Option33.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Race Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option29.value = True Then
    ListView2.ColumnHeaders(1).Text = "Race Code"
    ListView2.ColumnHeaders(2).Text = "Sex"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For aa = 7 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option37.value = True Then
    ListView2.ColumnHeaders(1).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Text = "Plan No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Race Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option30.value = True Then
    ListView2.ColumnHeaders(1).Text = "Nature of Claim"
    ListView2.ColumnHeaders(2).Text = "Race Code"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
        
    For aa = 7 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
End If
End Sub

Private Sub ListViewHeader18()
If Option35.value = True Then
    
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Ins Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    For aa = 7 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa

ElseIf Option34.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Plan Type"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration "
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa

ElseIf Option33.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Ageband"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
ElseIf Option29.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Gender"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option37.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Race Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Total Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
        
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
End If

End Sub

Private Sub ListViewHeader21()
If Option35.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For aa = 7 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa

ElseIf Option31.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Basic Premium"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Gross Premium"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    
    For aa = 9 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option34.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Hospital Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option33.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Final Diag. Code"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Descriptions"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "No of Claims"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Duration "
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    
    For aa = 9 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option29.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Plan Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Ageband"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration "
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option37.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Plan Type"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "No of Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration "
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa

ElseIf Option30.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Insured Type"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "Plan Type"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Ageband"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "No of Claims"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Duration "
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    
    For aa = 9 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa

End If
End Sub

Private Sub ListViewHeader27()

    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "State"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Total Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For aa = 7 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
End Sub

Private Sub ListViewHeader25()

If Option35.value = True Then
    ListView2.ColumnHeaders(1).Text = "Repudiation Code"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    
    For aa = 4 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option34.value = True Then
    ListView2.ColumnHeaders(1).Text = "RepCategory"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight

    For aa = 4 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option33.value = True Then
    ListView2.ColumnHeaders(1).Text = "First Diag"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    
    For aa = 4 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option29.value = True Then
    ListView2.ColumnHeaders(1).Text = "Final Diag"
    ListView2.ColumnHeaders(2).Text = "Description"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight

    For aa = 4 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option37.value = True Then
    ListView2.ColumnHeaders(1).Text = "Case No"
    ListView2.ColumnHeaders(2).Text = "Bordx No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Bordx Date"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "Claimabity"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "Patient"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "Repudiate Cat"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "Description"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(8).Text = "Repudiate Code"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Text = "Description"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(10).Text = "Diag1"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Text = "Diag2"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Hospital"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(13).Text = "DOA"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Text = "DOD"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    
    For aa = 15 To 63
       ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option30.value = True Then
    ListView2.ColumnHeaders(1).Text = "Hospital"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    
    For aa = 3 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option31.value = True Then
    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight

    For aa = 3 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa

ElseIf Option32.value = True Then
    ListView2.ColumnHeaders(1).Text = "Gender"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight

    For aa = 3 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
ElseIf Option38.value = True Then
    ListView2.ColumnHeaders(1).Text = "Age Band"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight

    For aa = 3 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
End If

End Sub

'Private Sub ListViewHeader26()

'If Option17.value = True Then
    'ListView2.ColumnHeaders(1).Text = "Claim No"
    'ListView2.ColumnHeaders(2).Text = "Date"
    'ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    'ListView2.ColumnHeaders(3).Text = "MDCN"
    'ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    'ListView2.ColumnHeaders(4).Text = "HDN"
    'ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    'ListView2.ColumnHeaders(5).Text = "HCN"
    'ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    'ListView2.ColumnHeaders(6).Text = "MDCNAmt"
    'ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    'ListView2.ColumnHeaders(7).Text = "HDNAmt"
    'ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    'ListView2.ColumnHeaders(8).Text = "HCNAmt"
    'ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    'ListView2.ColumnHeaders(9).Text = "Reimbursement Amt"
    'ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    'ListView2.ColumnHeaders(10).Text = "Check Total"
    'ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
        
    'For AA = 11 To 63
       'ListView2.ColumnHeaders(AA).Text = ""
    'Next AA
'End If

'End Sub

Private Sub ListViewHeader20()

    ListView2.ColumnHeaders(1).Text = "Payor"
    ListView2.ColumnHeaders(2).Text = "Total Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For aa = 6 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
End Sub

Private Sub Form_Load()
    intExit = 0
    Call ComboListing
    Call ListViewHeader15
    ListView2.Visible = True
    Option35.value = True
    Option35.Caption = "Without Remarks"
    Option35.Visible = True
    Option34.Caption = "With Remarks"
    Option34.Visible = True
End Sub

' =================== Command Button Click =============================

Private Sub CmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub cmdprint_Click()
    Screen.MousePointer = vbHourglass
    
    If Option1.value = True And Option35.value = True Then 'General Listing
        Call ExportToExcelAnalysis15(ListView2)
    ElseIf Option1.value = True And Option34.value = True Then 'General Listing 2
        Call ExportListViewtoExcel(ListView2)
    ElseIf Option3.value = True And Option35.value = True Then 'Plan No 1
        Call ExportToExcelAnalysis211(ListView2)
    ElseIf Option3.value = True And Option31.value = True Then 'Plan No 7
        Call ExportToExcelAnalysis211(ListView2)
    ElseIf Option3.value = True And Option34.value = True Then 'Plan No 2
        Call ExportToExcelAnalysis212(ListView2)
    ElseIf Option3.value = True And Option33.value = True Then 'Plan No 3
        Call ExportToExcelAnalysis213(ListView2)
    ElseIf Option3.value = True And Option29.value = True Then 'Plan No 4
        Call ExportToExcelAnalysis214(ListView2)
    ElseIf Option3.value = True And Option37.value = True Then ' Plan No 5
        Call ExportToExcelAnalysis215(ListView2)
    ElseIf Option3.value = True And Option30.value = True Then ' Plan No 6
        Call ExportToExcelAnalysis215(ListView2)
    ElseIf Option5.value = True And Option35.value = True Then 'Ageband 1
        Call ExportToExcelAnalysis31(ListView2)
    ElseIf Option5.value = True And Option34.value = True Then 'Ageband 2
        Call ExportToExcelAnalysis32(ListView2)
    ElseIf Option5.value = True And Option33.value = True Then 'Ageband 3
        Call ExportToExcelAnalysis33(ListView2)
    ElseIf Option5.value = True And Option29.value = True Then 'Ageband 4
        Call ExportToExcelAnalysis34(ListView2)
    ElseIf Option5.value = True And Option37.value = True Then 'Ageband 5
        Call ExportToExcelAnalysis35(ListView2)
    ElseIf Option5.value = True And Option30.value = True Then 'Ageband 6
        Call ExportToExcelAnalysis36(ListView2)
    ElseIf Option6.value = True And Option35.value = True Then 'Gender 1
        Call ExportToExcelAnalysis61(ListView2)
    ElseIf Option6.value = True And Option34.value = True Then 'Gender 2
        Call ExportToExcelAnalysis62(ListView2)
    ElseIf Option6.value = True And Option33.value = True Then 'Gender 3
        Call ExportToExcelAnalysis63(ListView2)
    ElseIf Option6.value = True And Option29.value = True Then 'Gender 4
        Call ExportToExcelAnalysis64(ListView2)
    ElseIf Option6.value = True And Option37.value = True Then 'Gender 5
        Call ExportToExcelAnalysis65(ListView2)
    ElseIf Option6.value = True And Option30.value = True Then 'Gender 6
        Call ExportToExcelAnalysis66(ListView2)
    ElseIf Option8.value = True Then 'State
        Call ExportToExcelAnalysis8(ListView2)
    ElseIf Option12.value = True Then 'State
        Call ExportToExcelAnalysis12(ListView2)
    ElseIf Option15.value = True And Option35.value = True Then 'Diagnosis 1
        Call ExportToExcelAnalysis151(ListView2)
    ElseIf Option15.value = True And Option34.value = True Then 'Diagnosis 2
        Call ExportToExcelAnalysis152(ListView2)
    ElseIf Option15.value = True And Option33.value = True Then 'Diagnosis 2
        Call ExportToExcelAnalysis153(ListView2)
    ElseIf Option15.value = True And Option29.value = True Then 'Diagnosis 2
        Call ExportToExcelAnalysis153(ListView2)
    ElseIf Option16.value = True And Option35.value = True Then 'Diagnosis 1
        Call ExportToExcelAnalysis111(ListView2)
    ElseIf Option16.value = True And Option34.value = True Then 'Diagnosis 2
        Call ExportToExcelAnalysis112(ListView2)
    'ElseIf Option17.value = True Then 'Claim Check Total
        'Call ExportToExcelAnalysis26(ListView2)
    ElseIf Option23.value = True Then 'Acc / Illness
        Call ExportToExcelAnalysis1(ListView2)
    ElseIf Option36.value = True And Option35.value = True Then 'Agent No 1
        Call ExportToExcelAnalysis41(ListView2)
    ElseIf Option36.value = True And Option34.value = True Then 'Agent No 2
        Call ExportToExcelAnalysis42(ListView2)
    ElseIf Option36.value = True And Option33.value = True Then 'Agent No 3
        Call ExportToExcelAnalysis43(ListView2)
    ElseIf Option36.value = True And Option29.value = True Then 'Agent No 4
        Call ExportToExcelAnalysis44(ListView2)
    Else
        Call ExportToExcelAnalysis1(ListView2)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
Dim StrAppend As String
    StrAppend = ""
    CmdPrint.Enabled = False
    CmdClear.Enabled = False
    cmdExit.Enabled = False
    Screen.MousePointer = vbHourglass
    Call SearchRecords
    Screen.MousePointer = vbDefault
End Sub

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

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
End Sub
' =================== Command Button Click / Change ==========================


'============== ComboBox Listing =================================
Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim INS As New ADODB.Recordset
'Dim REL As New ADODB.Recordset
Dim PolStatus As New ADODB.Recordset
Dim state As New ADODB.Recordset
Dim HOS As New ADODB.Recordset
Dim HosGrp As New ADODB.Recordset

    INS.Open "Select INSCode, INSDescription from INS", medixmasterconn, adOpenKeyset, adLockOptimistic
    PAY.Open "Select PAYCode, PAYCompanyName from PAY", medixmasterconn, adOpenKeyset, adLockOptimistic
    PolStatus.Open "Select pstatuscode ,pstatusDesc from PolicyStatus", medixmasterconn, adOpenKeyset, adLockOptimistic
    'REL.Open "Select RELCode, RELDescription from REL", medixmasterconn, adOpenKeyset, adLockOptimistic
    state.Open "Select STADescription from STA", medixmasterconn, adOpenKeyset, adLockOptimistic
    HOS.Open "Select HOSCode, HOSName from HOS", medixmasterconn, adOpenKeyset, adLockOptimistic
    HosGrp.Open "Select HOSGRPCode, HOSGRPName from HOSGRP", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    
    cboHLTCode.AddItem "S" & " - " & "Sihat"
    cboHLTCode.AddItem "N" & " - " & "Non-Sihat"
    
    Do Until INS.EOF
        CboInsType.AddItem ChkStr(INS!INSCode) & " - " & ChkStr(INS!INSDescription)
        INS.MoveNext
    Loop
    INS.Close
    Set INS = Nothing
    
    CboMemType.AddItem "Both"
    CboMemType.AddItem "P" & " - " & "Principal"
    CboMemType.AddItem "S" & " - " & "Supplementary"
    CboMemType.Text = "Both"
    
    CboAdultChild.AddItem "A - Adult"
    CboAdultChild.AddItem "C - Child"
    
    CboMarital.AddItem "S - Single"
    CboMarital.AddItem "M - Married"
    
    CboRace.AddItem "C - Chinese"
    CboRace.AddItem "I - Indian"
    CboRace.AddItem "M - Malay"
    CboRace.AddItem "O - Other"
            
    CboSex.AddItem "M" & " - " & "Male"
    CboSex.AddItem "F" & " - " & "Female"
    
    'Do Until REL.EOF
        'CboRelation.AddItem ChkStr(REL!RelCode) & " - " & ChkStr(REL!RELDescription)
        'REL.MoveNext
    'Loop
    'REL.Close
    'Set REL = Nothing
    
    CboRelation.AddItem "D - Daughter"
    CboRelation.AddItem "H - Husband"
    CboRelation.AddItem "S - Son"
    CboRelation.AddItem "W - Wife"
    
    cboDataStatus.AddItem "A - Active"
    cboDataStatus.AddItem "D - Delete"
    
    With CboClaimability
        .AddItem "A - Accepted"
        .AddItem "R - Rejected"
    End With
    
    Do Until PolStatus.EOF
        CboStatus.AddItem (PolStatus!pstatuscode) & " - " & (PolStatus!pstatusDesc)
        PolStatus.MoveNext
    Loop
    PolStatus.Close
    Set PolStatus = Nothing
    
    With CboCaseStatus
        .AddItem "O - Open"
        .AddItem "D - Delete"
        .AddItem "A - All"
    End With
    
    CboExpired.AddItem "N - No"
    CboExpired.AddItem "Y - Yes"
        
    CboPreMode.AddItem "A - Annually"
    CboPreMode.AddItem "Q - Quarterly"
    CboPreMode.AddItem "S - Semi Annual"
    CboPreMode.AddItem "M - Monthly"
    
    With CboStayType
        .AddItem "I - IN-PATIENT"
        .AddItem "O - OUT-PATIENT"
    End With
    
    With CboCashType
        .AddItem "C - CASHLESS"
        .AddItem "R - REIMBURSEMENT"
    End With
        
    Do Until state.EOF
        CboState.AddItem ChkStr(state!STADescription)
        state.MoveNext
    Loop
    state.Close
    Set state = Nothing
        
    cboRenewalInd.AddItem "N - No"
    cboRenewalInd.AddItem "Y - Yes"
    
    cboTakeOver.AddItem "N - No"
    cboTakeOver.AddItem "Y - Yes"
           
    Do Until PAY.EOF
        cboPayorCode.AddItem PAY!PAYCode & " - " & Trim(PAY!PAYCompanyName)
        PAY.MoveNext
    Loop
    PAY.Close
    Set PAY = Nothing
    
    cboSuppType.AddItem "Both"
    cboSuppType.AddItem "C" & " - " & "Combined Limit"
    cboSuppType.AddItem "S" & " - " & "Separated Limit"
    cboSuppType.Text = "C" & " - " & "Combined Limit"
    
    Do Until HOS.EOF
        cboHOSName.AddItem Trim(HOS!HOSName)
        HOS.MoveNext
    Loop
    HOS.Close
    Set HOS = Nothing
    
    Do Until HosGrp.EOF
        cboHOSGrp.AddItem Trim(HosGrp!HOSGRPName)
        HosGrp.MoveNext
    Loop
    HosGrp.Close
    Set HosGrp = Nothing
    
End Sub
'============== ComboBox Listing =================================


'==================== List Final Diag. Listing Resullts ==========================
Private Sub FinalDiagListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim rst3 As New ADODB.Recordset
    
    DiagCode = ""
    lblCurrent = 0
    ListView2.listitems.Clear
            
    If Option35.value = True Then
        Sql = "SELECT Top 20 COUNT(FinalDiag) AS Total, FinalDiag, Sum(CLMTotalApproved) As TotalApproved, " & _
              " Sum(CLMTotalActual) As TotalActual, " & _
              " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 " & StrAppend & _
              " GROUP BY FinalDiag ORDER BY Total DESC"
    
    ElseIf Option34.value = True Then
        Sql = "SELECT Top 20 Sum(CLMTotalApproved) As TotalApproved, COUNT(FinalDiag) AS Total, FinalDiag, " & _
              " Sum(CLMTotalActual) As TotalActual, " & _
              " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 " & StrAppend & _
              " Group By FinalDiag Order By TotalApproved Desc "
    
    Else
        Sql = "SELECT COUNT(FinalDiag) AS Total, FinalDiag, Sum(CLMTotalApproved) As TotalApproved, " & _
              " Sum(CLMTotalActual) As TotalActual, " & _
              " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 " & StrAppend & _
              " GROUP BY FinalDiag ORDER BY Total DESC"
    End If
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
       
        Do Until rst2.EOF
    
            With rst2
                'If Len(rst2!PAYCode) Then
                    'Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                'Else
                    'Set Record = ListView2.listitems.Add(, , "Empty")
                'End If
                
                If Len(rst2!FinalDiag) Then
                    Set Record = ListView2.listitems.Add(, , rst2!FinalDiag)
                    'Record.SubItems(1) = Trim(rst2!FinalDiag)
                    DiagCode = rst2!FinalDiag
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                    'Record.SubItems(1) = "Empty"
                    DiagCode = ""
                End If
                
                sql2 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(DiagCode) & "'"
                rst3.Open sql2, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                If rst3.recordCount = 0 Then
                    rst3.Close
                    Set rst3 = Nothing
                    Record.SubItems(1) = ""
                Else
                    Description = Trim(rst3!ICDDescription)
                    Record.SubItems(1) = Trim(Description)
                    rst3.Close
                    Set rst3 = Nothing
                End If
                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(4) = Format(rst2!TotalActual, "##,#0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "##,#0.00")
                End If
            
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub FinalDiagbyHos(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim rst3 As New ADODB.Recordset
    
    DiagCode = ""
    lblCurrent = 0
    ListView2.listitems.Clear
    
    Sql = " SELECT COUNT(FinalDiag) AS Total, FinalDiag, Sum(CLMTotalApproved) As TotalApproved, " & _
          " Sum(CLMTotalActual) As TotalActual, HOSName, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 " & StrAppend & _
          " GROUP BY FinalDiag, HOSName ORDER BY FinalDiag "
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
       
        Do Until rst2.EOF
    
            With rst2
                               
                If Len(rst2!FinalDiag) Then
                    Set Record = ListView2.listitems.Add(, , rst2!FinalDiag)
                    'Record.SubItems(1) = Trim(rst2!FinalDiag)
                    DiagCode = rst2!FinalDiag
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                    'Record.SubItems(1) = "Empty"
                    DiagCode = ""
                End If
                
                sql2 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(DiagCode) & "'"
                rst3.Open sql2, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                If rst3.recordCount = 0 Then
                    rst3.Close
                    Set rst3 = Nothing
                    Record.SubItems(1) = ""
                Else
                    Description = Trim(rst3!ICDDescription)
                    Record.SubItems(1) = Trim(Description)
                    rst3.Close
                    Set rst3 = Nothing
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(2) = Trim(rst2!HOSName)
                End If
                
                If Len(rst2!Total) Then
                    Record.SubItems(3) = rst2!Total
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(4) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(4) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(5) = Format(rst2!TotalActual, "##,#0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(6) = Format(rst2!TotalApproved, "##,#0.00")
                End If
            
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub


'================ Daily Admission Listing ===========================
Private Sub DailyAdmissionListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim intloop As Integer
       
    
    ListView2.listitems.Clear
                
    Sql = " SELECT CASNumber, claimversion, MBMNumber, MBMPolicyNo, " & _
          " PatientName, CASDateAdmitted, HOSName, FirstDiag " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    'If rst2.recordCount = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'cmdExit.Enabled = True
        'CmdClear.Enabled = True
        'CmdPrint.Enabled = True
    'Else
    
        Do Until rst2.EOF
                    
            With rst2
            
                intloop = intloop + 1
                
                Set Record = ListView2.listitems.Add(, , intloop)
                                                
                If Len(rst2!CASNumber & "-" & rst2!claimversion) Then
                       Record.SubItems(1) = rst2!CASNumber & "-" & rst2!claimversion
                End If
                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(2) = rst2!MBMNumber
                End If
                
                If Len(rst2!MBMPolicyNo) Then
                    Record.SubItems(3) = rst2!MBMPolicyNo
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(4) = rst2!PatientName
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(5) = rst2!CASDateAdmitted
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(6) = rst2!HOSName
                End If
                
                If Len(rst2!FirstDiag) Then
                    Record.SubItems(7) = rst2!FirstDiag
                End If
            End With
            rst2.MoveNext
        Loop
    'End If
    
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    
    rst2.Close
    Set rst2 = Nothing
    cmdExit.Enabled = True
    CmdClear.Enabled = True
    CmdPrint.Enabled = True
End Sub
'================ Daily Admission Listing ===========================

'================ Daily Discharge Listing ===========================
Private Sub DailyDischargeListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim intloop As Integer
       
    
    ListView2.listitems.Clear
    
    'sql = " SELECT * From VIEW_AnalysisReportNew Where 0 = 0 "
    
    Sql = " SELECT CASNumber, claimversion, MBMNumber, MBMPolicyNo, " & _
          " PatientName, CASDateAdmitted, CASDischargeDate, HOSName, FinalDiag " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    'If rst2.recordCount = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'cmdExit.Enabled = True
        'CmdClear.Enabled = True
        'CmdPrint.Enabled = True
    'Else
    
        Do Until rst2.EOF
                    
            With rst2
            
                intloop = intloop + 1
                
                Set Record = ListView2.listitems.Add(, , intloop)
                
                                
                If Len(rst2!CASNumber & "-" & rst2!claimversion) Then
                       Record.SubItems(1) = rst2!CASNumber & "-" & rst2!claimversion
                End If
                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(2) = rst2!MBMNumber
                End If
                
                If Len(rst2!MBMPolicyNo) Then
                    Record.SubItems(3) = rst2!MBMPolicyNo
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(4) = rst2!PatientName
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(5) = rst2!CASDateAdmitted
                End If
                
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(6) = rst2!CASDischargeDate
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(7) = rst2!HOSName
                End If
                
                If Len(rst2!FinalDiag) Then
                    Record.SubItems(8) = rst2!FinalDiag
                End If
                
            End With
            rst2.MoveNext
        Loop
    'End If
    
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    
    rst2.Close
    Set rst2 = Nothing
    cmdExit.Enabled = True
    CmdClear.Enabled = True
    CmdPrint.Enabled = True
End Sub
'================ Daily Discharge Listing ===========================


'=========== Option Button, Combo Box Change/KeyPress ===================

Private Sub cboPayorCode_Click()
Dim strsql As String
Dim PAYGRP As New ADODB.Recordset

If Len(Mid(cboPayorCode, 1, 2)) Then
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
End If
End Sub

Private Sub cboPayorCode_Change()

    If cboPayorCode.Text = "" Then
        If InStr(cboPayorCode.Text, "null") Then
           cboPayorCode.Enabled = False
        End If
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



Private Sub cboHLTcode_KeyPress(KeyAscii As Integer)
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

Private Sub CboInsType_Change()

    If CboInsType.Text = "" Then
        If InStr(CboInsType.Text, "null") Then
           CboInsType.Enabled = False
        End If
    End If
End Sub

Private Sub CboInsType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboInsType.Text, CboInsType.SelStart) + Chr(KeyAscii) + Mid(CboInsType.Text, CboInsType.SelStart + CboInsType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboInsType.ListCount - 1
 
        If NewText = LCase(Left(CboInsType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboInsType.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboInsType.Text = ValidValue
                CboInsType.SelStart = 0
                CboInsType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboMemType_Change()

    If CboMemType.Text = "" Then
        If InStr(CboMemType.Text, "null") Then
           CboMemType.Enabled = False
        End If
    End If
End Sub

Private Sub cboSuppType_Change()

    If cboSuppType.Text = "" Then
        If InStr(cboSuppType.Text, "null") Then
           cboSuppType.Enabled = False
        End If
    End If
End Sub



Private Sub CboAdultChild_Change()

    If CboAdultChild.Text = "" Then
        If InStr(CboAdultChild.Text, "null") Then
           CboAdultChild.Enabled = False
        End If
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

Private Sub CboMarital_Change()

    If CboMarital.Text = "" Then
        If InStr(CboMarital.Text, "null") Then
           CboMarital.Enabled = False
        End If
    End If
End Sub

Private Sub CboMarital_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboMarital.Text, CboMarital.SelStart) + Chr(KeyAscii) + Mid(CboMarital.Text, CboMarital.SelStart + CboMarital.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboMarital.ListCount - 1
 
        If NewText = LCase(Left(CboMarital.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboMarital.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboMarital.Text = ValidValue
                CboMarital.SelStart = 0
                CboMarital.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboRace_Change()

    If CboRace.Text = "" Then
        If InStr(CboRace.Text, "null") Then
           CboRace.Enabled = False
        End If
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

    If CboSex.Text = "" Then
        If InStr(CboSex.Text, "null") Then
           CboSex.Enabled = False
        End If
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

Private Sub CboRelation_Change()

    If CboRelation.Text = "" Then
        If InStr(CboRelation.Text, "null") Then
           CboRelation.Enabled = False
        End If
    End If
End Sub

Private Sub CboRelation_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboRelation.Text, CboRelation.SelStart) + Chr(KeyAscii) + Mid(CboRelation.Text, CboRelation.SelStart + CboRelation.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboRelation.ListCount - 1
 
        If NewText = LCase(Left(CboRelation.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboRelation.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboRelation.Text = ValidValue
                CboRelation.SelStart = 0
                CboRelation.SelLength = Len(ValidValue)
            End If
        End If
End Sub



Private Sub cboDataStatus_Change()

    If cboDataStatus.Text = "" Then
        If InStr(cboDataStatus.Text, "null") Then
           cboDataStatus.Enabled = False
        End If
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

Private Sub CboClaimability_Change()

    If CboClaimability.Text = "" Then
        If InStr(CboClaimability.Text, "null") Then
           CboClaimability.Enabled = False
        End If
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

Private Sub CboStatus_Change()

    If CboStatus.Text = "" Then
        If InStr(CboStatus.Text, "null") Then
           CboStatus.Enabled = False
        End If
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

Private Sub CboCaseStatus_Change()

    If CboCaseStatus.Text = "" Then
        If InStr(CboCaseStatus.Text, "null") Then
           CboCaseStatus.Enabled = False
        End If
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



Private Sub CboExpired_Change()

    If CboExpired.Text = "" Then
        If InStr(CboExpired.Text, "null") Then
           CboExpired.Enabled = False
        End If
    End If
End Sub

Private Sub CboExpired_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboExpired.Text, CboExpired.SelStart) + Chr(KeyAscii) + Mid(CboExpired.Text, CboExpired.SelStart + CboExpired.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboExpired.ListCount - 1
 
        If NewText = LCase(Left(CboExpired.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboExpired.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboExpired.Text = ValidValue
                CboExpired.SelStart = 0
                CboExpired.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboPreMode_Change()

    If CboPreMode.Text = "" Then
        If InStr(CboPreMode.Text, "null") Then
           CboPreMode.Enabled = False
        End If
    End If
End Sub

Private Sub CboPreMode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboPreMode.Text, CboPreMode.SelStart) + Chr(KeyAscii) + Mid(CboPreMode.Text, CboPreMode.SelStart + CboPreMode.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPreMode.ListCount - 1
 
        If NewText = LCase(Left(CboPreMode.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPreMode.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPreMode.Text = ValidValue
                CboPreMode.SelStart = 0
                CboPreMode.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboState_Change()

    If CboState.Text = "" Then
        If InStr(CboState.Text, "null") Then
           CboState.Enabled = False
        End If
    End If
End Sub

Private Sub CboState_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboState.Text, CboState.SelStart) + Chr(KeyAscii) + Mid(CboState.Text, CboState.SelStart + CboState.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboState.ListCount - 1
 
        If NewText = LCase(Left(CboState.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboState.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboState.Text = ValidValue
                CboState.SelStart = 0
                CboState.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboRenewalInd_Change()

    If cboRenewalInd.Text = "" Then
        If InStr(cboRenewalInd.Text, "null") Then
           cboRenewalInd.Enabled = False
        End If
    End If
End Sub

Private Sub cboRenewalInd_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboRenewalInd.Text, cboRenewalInd.SelStart) + Chr(KeyAscii) + Mid(cboRenewalInd.Text, cboRenewalInd.SelStart + cboRenewalInd.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboRenewalInd.ListCount - 1
 
        If NewText = LCase(Left(cboRenewalInd.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboRenewalInd.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboRenewalInd.Text = ValidValue
                cboRenewalInd.SelStart = 0
                cboRenewalInd.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboTakeOver_Change()

    If cboTakeOver.Text = "" Then
        If InStr(cboTakeOver.Text, "null") Then
           cboTakeOver.Enabled = False
        End If
    End If
End Sub

Private Sub cboTakeOver_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboTakeOver.Text, cboTakeOver.SelStart) + Chr(KeyAscii) + Mid(cboTakeOver.Text, cboTakeOver.SelStart + cboTakeOver.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboTakeOver.ListCount - 1
 
        If NewText = LCase(Left(cboTakeOver.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboTakeOver.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboTakeOver.Text = ValidValue
                cboTakeOver.SelStart = 0
                cboTakeOver.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboGrpCom_Change()

    If CboGrpCom.Text = "" Then
        If InStr(CboGrpCom.Text, "null") Then
           CboGrpCom.Enabled = False
        End If
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

Private Sub cboHOSName_Change()

    If cboHOSName.Text = "" Then
        If InStr(cboHOSName.Text, "null") Then
           cboHOSName.Enabled = False
        End If
    End If
End Sub

Private Sub cboHOSName_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboHOSName.Text, cboHOSName.SelStart) + Chr(KeyAscii) + Mid(cboHOSName.Text, cboHOSName.SelStart + cboHOSName.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboHOSName.ListCount - 1
 
        If NewText = LCase(Left(cboHOSName.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboHOSName.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboHOSName.Text = ValidValue
                cboHOSName.SelStart = 0
                cboHOSName.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboHOSGrp_Change()

    If cboHOSGrp.Text = "" Then
        If InStr(cboHOSGrp.Text, "null") Then
           cboHOSGrp.Enabled = False
        End If
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

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    
If Option23.value = True Or _
   (Option11.value = True And Option35.value = True) Or (Option2.value = True And Option35.value = True) Or _
   Option9.value = True Or _
   (Option7.value = True And Option35.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
ElseIf (Option3.value = True And Option35.value = True) Or _
       (Option36.value = True And Option33.value = True) Or _
       (Option36.value = True And Option34.value = True) Or Option8.value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
        
ElseIf Option3.value = True And Option31.value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        End Select

ElseIf (Option5.value = True And Option34.value = True) Or _
   (Option5.value = True And Option33.value = True) Or (Option5.value = True And Option29.value = True) Or _
   (Option5.value = True And Option30.value = True) Or (Option6.value = True And Option34.value = True) Or _
   (Option2.value = True And Option33.value = True) Or (Option7.value = True And Option34.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
ElseIf (Option3.value = True And Option29.value = True) Or _
    (Option10.value = True And Option34.value = True) Or (Option10.value = True And Option37.value = True) Or _
    (Option10.value = True And Option29.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
ElseIf Option10.value = True And Option33.value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
ElseIf (Option6.value = True And Option33.value = True) Or (Option6.value = True And Option29.value = True) Or _
   (Option6.value = True And Option30.value = True) Or (Option2.value = True And Option34.value = True) Or _
   (Option2.value = True And Option29.value = True) Or (Option2.value = True And Option37.value = True) Or _
   (Option7.value = True And Option33.value = True) Or (Option7.value = True And Option29.value = True) Or _
   (Option7.value = True And Option30.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
ElseIf (Option15.value = True And Option35.value = True) Or (Option15.value = True And Option34.value = True) Or _
   (Option15.value = True And Option33.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
ElseIf Option15.value = True And Option29.value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
ElseIf (Option36.value = True And Option29.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
ElseIf (Option3.value = True And Option34.value = True) Or (Option3.value = True And Option37.value = True) Or _
       (Option7.value = True And Option30.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
ElseIf Option3.value = True And Option30.value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        End Select
ElseIf Option3.value = True And Option33.value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        End Select
ElseIf (Option7.value = True And Option37.value = True) Or _
   (Option6.value = True And Option37.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
ElseIf (Option20.value = True And Option35.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(8)
        End Select
ElseIf (Option5.value = True And Option35.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
ElseIf (Option16.value = True And Option35.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        End Select
ElseIf (Option16.value = True And Option34.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(10)
        End Select
'ElseIf (Option17.value = True) Then
        'Select Case ColumnHeader.Index
        'Case 1
            'SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        'Case 2
            'SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
        'Case 3
            'SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        'Case 4
            'SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        'Case 5
            'SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        'Case 6
            'SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        'Case 7
            'SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        'Case 8
            'SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        'Case 9
            'SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        'Case 10
            'SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        'End Select
ElseIf (Option5.value = True And Option37.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
ElseIf Option19.value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select
ElseIf (Option6.value = True And Option35.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        End Select
ElseIf (Option10.value = True And Option35.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(14)
        End Select
ElseIf (Option10.value = True And Option30.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        End Select
ElseIf Option1.value = True And Option35.value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
        Case 27
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
        Case 28
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
        Case 29
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(29)
        Case 30
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(30)
        Case 31
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(31)
        Case 32
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(32)
        Case 33
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(33)
        Case 34
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(34)
        Case 35
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(35)
        Case 36
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(36)
        Case 37
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(37)
        Case 38
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(38)
        End Select
ElseIf Option1.value = True And Option34.value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(26)
        Case 27
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
        Case 28
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(28)
        Case 29
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(29)
        Case 30
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(30)
        Case 31
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(31)
        Case 32
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(32)
        Case 33
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(33)
        Case 34
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(34)
        Case 35
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(35)
        Case 36
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(36)
        Case 37
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(37)
        Case 38
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(38)
        Case 39
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(39)
        End Select
ElseIf (Option20.value = True And Option33.value = True) Or _
       (Option21.value = True And Option33.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
        End Select
ElseIf (Option21.value = True And Option35.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(9)
        End Select
ElseIf Option12.value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(10)
        End Select
        
ElseIf Option13.value = True Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(6)
        End Select

ElseIf (Option36.value = True And Option35.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(19)
        End Select
ElseIf (Option20.value = True And Option34.value = True) Or _
       (Option21.value = True And Option34.value = True) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(26)
        Case 27
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(27)
        Case 28
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(28)
        Case 29
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(29)
        Case 30
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(30)
        Case 31
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(31)
        Case 32
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(32)
        Case 33
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(33)
        Case 34
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(34)
        Case 35
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(35)
        Case 36
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(36)
        Case 37
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(37)
        Case 38
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(38)
        Case 39
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(39)
        Case 40
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(40)
        Case 41
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(41)
        Case 42
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(42)
        Case 43
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(43)
        Case 44
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(44)
        Case 45
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(45)
        Case 46
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(46)
        Case 47
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(47)
        Case 48
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(48)
        Case 49
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(49)
        Case 50
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(50)
        Case 51
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(51)
        Case 52
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(52)
        Case 53
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(53)
        Case 54
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(54)
        Case 55
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(55)
        Case 56
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(56)
        Case 57
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(57)
        Case 58
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(58)
        Case 59
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(59)
        Case 60
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(60)
        Case 61
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(61)
        Case 62
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(62)
        Case 63
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(63)
        End Select
End If

End Sub

'=========== Combo Box Change/KeyPress ===================

Private Sub Option1_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    Option35.value = True
    Option35.Caption = "Without Remarks"
    Option35.Visible = True
    Option34.Caption = "With Remarks"
    Option34.Visible = True
    ChkDOA.Enabled = True
    ChkDOA.Visible = True
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    TxtDOA1.Visible = True
    TxtDOA2.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label49.Visible = True
    Label50.Visible = True
    Label47.Caption = "DOA"
    Label48.Caption = "DOD"
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
    Option29.Visible = False
    Option33.Visible = False
    Option37.Visible = False
    Option30.Visible = False
    Call ListViewHeader15
End Sub

Private Sub Option10_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    ChkDOA.Visible = True
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    TxtDOA1.Visible = True
    TxtDOA2.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label49.Visible = True
    Label50.Visible = True
    Label47.Caption = "DOA"
    Label48.Caption = "DOD"
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
        
    Option35.value = True
    Option35.Caption = "Detail Listing"
    Option34.Caption = "General Summary (By Payor)"
    Option33.Caption = "General Summary (All)"
    Option29.Caption = "Top 20 by Cases"
    Option37.Caption = "Top 20 by Amount"
    Option30.Caption = "By Final Diag."
    Option30.Visible = True
    Option37.Visible = True
    Option35.Visible = True
    Option34.Visible = True
    Option33.Visible = True
    Option29.Visible = True
           
    Call ListViewHeader17

End Sub

Private Sub Option11_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    ChkDOA.Visible = True
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    TxtDOA1.Visible = True
    TxtDOA2.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label49.Visible = True
    Label50.Visible = True
    Label47.Caption = "DOA"
    Label48.Caption = "DOD"
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.value = True
    Option35.Visible = True
    Option35.Caption = "General Summary"
    Option37.Visible = False
    Option30.Visible = False
    Call ListViewHeader16
End Sub

Private Sub Option12_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    ChkDOA.Visible = True
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    TxtDOA1.Visible = True
    TxtDOA2.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label49.Visible = True
    Label50.Visible = True
    Label47.Caption = "DOA"
    Label48.Caption = "DOD"
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
    Option37.Visible = False
    Option30.Visible = False
    
    ListView2.ColumnHeaders(1).Text = "Case No"
    ListView2.ColumnHeaders(2).Text = "Doctor"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Patient Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Agent Code"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Hospital"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "DOA"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "DOD"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Text = "Inv Amt"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "Final Diag"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(10).Text = "Description"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnLeft
    For aa = 11 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
End Sub

Private Sub Option13_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    ChkDOA.Visible = True
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    TxtDOA1.Visible = True
    TxtDOA2.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label49.Visible = True
    Label50.Visible = True
    Label47.Caption = "DOA"
    Label48.Caption = "DOD"
        
    ListView2.ColumnHeaders(1).Text = "I/C No"
    ListView2.ColumnHeaders(2).Text = "Claim No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Patient Name"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "O/S Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
            
    For aa = 7 To 63
       ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
    Option37.Visible = False
    Option30.Visible = False
End Sub

'Private Sub Option14_Click()
    'lblCurrent = ""
    'lblTotalRecord = ""
    'ListView2.listitems.Clear
    'CboClaimability.Text = ""
    'ChkDOA.Enabled = True
    'ChkDOA.Visible = True
    'ChkDOD.Enabled = True
    'ChkDOD.Visible = True
    'TxtDOA1.Visible = True
    'TxtDOA2.Visible = True
    'txtDOD1.Visible = True
    'txtDOD2.Visible = True
    'Label49.Visible = True
    'Label50.Visible = True
    'Label47.Caption = "DOA"
    'Label48.Caption = "DOD"
    'Option31.Visible = False
    'Option32.Visible = False
    'Option38.Visible = False
    'Option29.Visible = False
    'Option33.Visible = False
    'Option34.Visible = False
    'Option35.Visible = False
    'Option37.Visible = False
    'Option30.Visible = False
'End Sub

Private Sub Option15_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    ChkDOA.Visible = True
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    TxtDOA1.Visible = True
    TxtDOA2.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label49.Visible = True
    Label50.Visible = True
    Label47.Caption = "DOA"
    Label48.Caption = "DOD"
    
    Option35.value = True
    Call ListViewHeader8
    Option35.Caption = "Top 20 by Cases"
    Option34.Caption = "Top 20 by Amount"
    Option33.Caption = "General Summary"
    Option29.Caption = "By Hospital"
    
    Option35.Visible = True
    Option34.Visible = True
    Option33.Visible = True
    Option29.Visible = True
    Option37.Visible = False
    Option30.Visible = False
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
  
End Sub

Private Sub Option16_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    Option35.value = True
    Option35.Visible = True
    Option35.Caption = "Summary"
    Option34.Visible = True
    Option34.Caption = "Details"
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
    Option33.Visible = False
    Option29.Visible = False
    Option37.Visible = False
    Option30.Visible = False
    Call ListViewHeader11
End Sub

'Private Sub Option17_Click()
    'lblCurrent = ""
    'lblTotalRecord = ""
    'ListView2.listitems.Clear
    'CboClaimability.Text = ""
    'ChkDOA.Enabled = True
    'ChkDOA.Visible = True
    'ChkDOD.Enabled = True
    'ChkDOD.Visible = True
    'TxtDOA1.Visible = True
    'TxtDOA2.Visible = True
    'txtDOD1.Visible = True
    'txtDOD2.Visible = True
    'Label49.Visible = True
    'Label50.Visible = True
    'Label47.Caption = "DOA"
    'Label48.Caption = "DOD"
    'Call ListViewHeader26
    'Option31.Visible = False
    'Option32.Visible = False
    'Option38.Visible = False
    'Option29.Visible = False
    'Option33.Visible = False
    'Option34.Visible = False
    'Option35.Visible = False
    'Option37.Visible = False
    'Option30.Visible = False
'End Sub

Private Sub Option18_Click()
    llblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = "R - Rejected"
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label50.Visible = True
    Label48.Caption = "DOD"
    Option35.value = True
    Call ListViewHeader25
    Option35.Caption = "Repudiation Code"
    Option34.Caption = "Repudiation Category"
    Option33.Caption = "Diagnosis (First Diagnosis)"
    Option29.Caption = "Diagnosis (Final Diagnosis)"
    Option37.Caption = "Listing"
    Option30.Caption = "Hospital"
    Option31.Caption = "Payor"
    Option32.Caption = "Gender"
    Option38.Caption = "Ageband"
    Option35.Visible = True
    Option34.Visible = True
    Option33.Visible = True
    Option29.Visible = True
    Option37.Visible = True
    Option30.Visible = True
    Option31.Visible = True
    Option32.Visible = True
    Option38.Visible = True
    
End Sub

Private Sub Option19_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    ChkDOA.Visible = True
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    TxtDOA1.Visible = True
    TxtDOA2.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label49.Visible = True
    Label50.Visible = True
    Label47.Caption = "DOA"
    Label48.Caption = "DOD"
        
    ListView2.ColumnHeaders(1).Text = "Bordx No"
    ListView2.ColumnHeaders(2).Text = "Bordx Date"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "No of Claim"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Bordx Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
            
    For aa = 7 To 63
       ListView2.ColumnHeaders(aa).Text = ""
    Next aa
    
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
    Option37.Visible = False
    Option30.Visible = False
End Sub

Private Sub Option2_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    ChkDOA.Visible = True
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    TxtDOA1.Visible = True
    TxtDOA2.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label49.Visible = True
    Label50.Visible = True
    Label47.Caption = "DOA"
    Label48.Caption = "DOD"
    Option35.value = True
    Option35.Visible = True
    Option35.Caption = "Summary"
    Option34.Visible = True
    Option34.Caption = "By Plan No"
    Option33.Visible = True
    Option33.Caption = "By Ageband"
    Option29.Visible = True
    Option29.Caption = "By Gender"
    Option37.Visible = True
    Option37.Caption = "By Race"
    Option30.Visible = False
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
    Call ListViewHeader18
End Sub

Private Sub Option20_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label50.Visible = True
    Label48.Caption = "DOD"
    Option35.value = True
    Call ListViewHeader2
    Option35.Caption = "Daily Report"
    Option34.Caption = "Monthly Report"
    Option33.Caption = "YTD Report"
    Option35.Visible = True
    Option34.Visible = True
    Option33.Visible = True
    Option29.Visible = False
    Option37.Visible = False
    Option30.Visible = False
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
End Sub

Private Sub Option21_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    ChkDOA.Visible = True
    TxtDOA1.Visible = True
    TxtDOA2.Visible = True
    Label49.Visible = True
    Label47.Caption = "DOA"
    ChkDOA.Enabled = True
    ChkDOD.Enabled = False
    ChkDOD.Visible = True
    Option35.value = True
    Call ListViewHeader9
    Option35.Caption = "Daily Report"
    Option34.Caption = "Monthly Report"
    Option33.Caption = "YTD Report"
    Option35.Visible = True
    Option34.Visible = True
    Option33.Visible = True
    Option29.Visible = False
    Option37.Visible = False
    Option30.Visible = False
End Sub

'Private Sub Option22_Click()
    'lblCurrent = ""
    'lblTotalRecord = ""
    'ListView2.listitems.Clear
    'CboClaimability.Text = ""
    'ChkDOA.Enabled = True
    'ChkDOA.Visible = True
    'ChkDOD.Enabled = True
    'ChkDOD.Visible = True
    'TxtDOA1.Visible = True
    'TxtDOA2.Visible = True
    'txtDOD1.Visible = True
    'txtDOD2.Visible = True
    'Label49.Visible = True
    'Label50.Visible = True
    'Label47.Caption = "DOA"
    'Label48.Caption = "DOD"
    'Option31.Visible = False
    'Option32.Visible = False
    'Option38.Visible = False
    'Option29.Visible = False
    'Option33.Visible = False
    'Option34.Visible = True
    'Option35.value = True
    'Option35.Visible = True
    'Option35.Caption = "General Summary"
    'Option34.Caption = "By Follow-Up"
    'Option37.Visible = False
    'Option30.Visible = False
    'Call ListViewHeader6
'End Sub

Private Sub Option23_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    ChkDOA.Visible = True
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    TxtDOA1.Visible = True
    TxtDOA2.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label49.Visible = True
    Label50.Visible = True
    Label47.Caption = "DOA"
    Label48.Caption = "DOD"
    Option29.Visible = False
    Option30.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
    Option37.Visible = False
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
    Call ListViewHeader1
End Sub

'Private Sub Option24_Click()
    'lblCurrent = ""
    'lblTotalRecord = ""
    'ListView2.listitems.Clear
    'CboClaimability.Text = ""
    'ChkDOA.Enabled = True
    'ChkDOA.Visible = True
    'ChkDOD.Enabled = True
    'ChkDOD.Visible = True
    'TxtDOA1.Visible = True
    'TxtDOA2.Visible = True
    'txtDOD1.Visible = True
    'txtDOD2.Visible = True
    'Label49.Visible = True
    'Label50.Visible = True
    'Label47.Caption = "DOA"
    'Label48.Caption = "DOD"
    'Option29.Visible = False
    'Option33.Visible = False
    'Option34.Visible = False
    'Option35.Visible = False
    'Option37.Visible = False
    'Option30.Visible = False
    'Option31.Visible = False
    'Option32.Visible = False
    'Option38.Visible = False
'End Sub

Private Sub Option29_Click()
    
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    
    If Option2.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader18
    ElseIf Option3.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader21
    ElseIf Option5.value = True Then
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader3
    ElseIf Option6.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader14
    ElseIf Option7.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader23
    ElseIf Option15.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader8
    ElseIf Option18.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader25
    ElseIf Option36.value = True Then
        Call ListViewHeader4
    ElseIf Option10.value = True Then
        cboHOSName.Enabled = False
        cboHOSGrp.Enabled = False
        Call ListViewHeader17
    End If
End Sub

Private Sub Option3_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    ChkDOA.Visible = True
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    TxtDOA1.Visible = True
    TxtDOA2.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label49.Visible = True
    Label50.Visible = True
    Label47.Caption = "DOA"
    Label48.Caption = "DOD"
    Option37.Visible = True
    Option29.Visible = True
    Option30.Visible = True
    Option33.Visible = True
    Option34.Visible = True
    Option35.value = True
    Option35.Visible = True
    Option31.Visible = True
    Option35.Caption = "General Summary"
    Option34.Caption = "Top 10 Hospitals"
    Option33.Caption = "Top 10 Diagnosis"
    Option29.Caption = "By Ageband"
    Option37.Caption = "By Insured Type"
    Option30.Caption = "By Insured Type && Ageband"
    Option31.Caption = "General Summary with Premium"
    Call ListViewHeader21
End Sub

Private Sub Option30_Click()
    
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    
    If Option5.value = True Then
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader3
    ElseIf Option3.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader21
    ElseIf Option6.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader14
    ElseIf Option7.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader23
    ElseIf Option18.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader25
    ElseIf Option10.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader17
    End If
End Sub

Private Sub Option31_Click()

    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear

If Option18.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader25
ElseIf Option3.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader21
End If
End Sub

Private Sub Option32_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear

If Option18.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader25
End If
End Sub

Private Sub Option33_Click()
    
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    
    If Option2.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader18
    ElseIf Option3.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader21
    ElseIf Option5.value = True Then
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader3
    ElseIf Option6.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader14
    ElseIf Option7.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader23
    ElseIf Option18.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader25
    ElseIf Option20.value = True Then
        Call ListViewHeader2
        Label47.Caption = "DOA Year"
        TxtYear.Top = "1560"
        TxtDOA1.Visible = False
        TxtDOA2.Visible = False
        Label49.Visible = False
        ChkDOA.Visible = False
    ElseIf Option21.value = True Then
        Call ListViewHeader9
        Label48.Caption = "DOD Year"
        Label47.Caption = "DOA"
        TxtYear.Top = "1830"
        txtDOD1.Visible = False
        txtDOD2.Visible = False
        Label50.Visible = False
        ChkDOD.Visible = False
    ElseIf Option10.value = True Then
        cboHOSName.Enabled = True
        cboHOSGrp.Enabled = True
        Call ListViewHeader17
    ElseIf Option36.value = True Then
        Call ListViewHeader4
    End If
End Sub

Private Sub Option34_Click()
    
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    
    If Option1.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader15
    ElseIf Option2.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader18
    ElseIf Option3.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader21
    ElseIf Option5.value = True Then
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader3
    ElseIf Option6.value = True Then
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader14
    ElseIf Option7.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader23
    ElseIf Option16.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader11
    ElseIf Option18.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader25
    ElseIf Option20.value = True Then
        Call ListViewHeader2
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
    ElseIf Option21.value = True Then
        ChkDOA.Enabled = True
        ChkDOD.Enabled = False
        txtDOD1.Visible = True
        txtDOD2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label50.Visible = True
        ChkDOD.Visible = True
        Call ListViewHeader9
    ElseIf Option10.value = True Then
        cboHOSName.Enabled = True
        cboHOSGrp.Enabled = True
        Call ListViewHeader17
    'ElseIf Option22.value = True Then
        'Call ListViewHeader6
    ElseIf Option36.value = True Then
        Call ListViewHeader4
    End If
End Sub

Private Sub Option35_Click()
    
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    
    If Option1.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader15
    ElseIf Option2.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader18
    ElseIf Option3.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader21
    ElseIf Option5.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader3
    ElseIf Option6.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader14
    ElseIf Option7.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader23
    ElseIf Option16.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader11
    ElseIf Option18.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader25
    ElseIf Option20.value = True Then
        ChkDOA.Enabled = False
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader2
    ElseIf Option21.value = True Then
        ChkDOA.Enabled = True
        ChkDOD.Enabled = False
        txtDOD1.Visible = True
        txtDOD2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label50.Visible = True
        ChkDOD.Visible = True
        Call ListViewHeader9
    ElseIf Option10.value = True Then
        cboHOSName.Enabled = True
        cboHOSGrp.Enabled = True
        Call ListViewHeader17
    'ElseIf Option22.value = True Then
        'Call ListViewHeader6
    ElseIf Option36.value = True Then
        Call ListViewHeader4
    End If
End Sub

Private Sub Option36_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    ChkDOA.Visible = True
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    TxtDOA1.Visible = True
    TxtDOA2.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label49.Visible = True
    Label50.Visible = True
    Label47.Caption = "DOA"
    Label48.Caption = "DOD"
    Option37.Visible = False
    Option30.Visible = False
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
    Option29.Visible = True
    Option33.Visible = True
    Option34.Visible = True
    Option35.value = True
    Option35.Visible = True
    Option35.Caption = "Details Listing"
    Option34.Caption = "General Summary"
    Option33.Caption = "By Top 50 Cases"
    Option29.Caption = "By Same Doctor"
    Call ListViewHeader4
End Sub

Private Sub Option37_Click()
    
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    
    If Option2.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader18
    ElseIf Option3.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        Call ListViewHeader21
    ElseIf Option5.value = True Then
        Call ListViewHeader3
    ElseIf Option6.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader14
    ElseIf Option7.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader23
    ElseIf Option18.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader25
    ElseIf Option10.value = True Then
        cboHOSName.Enabled = False
        cboHOSGrp.Enabled = False
        Call ListViewHeader17
    End If
End Sub

Private Sub Option38_Click()
lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear

If Option18.value = True Then
        ChkDOA.Enabled = True
        TxtDOA1.Visible = True
        TxtDOA2.Visible = True
        Label47.Caption = "DOA"
        Label48.Caption = "DOD"
        Label49.Visible = True
        ChkDOA.Visible = True
        ListViewHeader25
End If
End Sub

'Private Sub Option4_Click()
    'lblCurrent = ""
    'lblTotalRecord = ""
    'ListView2.listitems.Clear
    'CboClaimability.Text = ""
    'ChkDOA.Enabled = True
    'ChkDOA.Visible = True
    'ChkDOD.Enabled = True
    'ChkDOD.Visible = True
    'TxtDOA1.Visible = True
    'TxtDOA2.Visible = True
    'txtDOD1.Visible = True
    'txtDOD2.Visible = True
    'Label49.Visible = True
    'Label50.Visible = True
    'Label47.Caption = "DOA"
    'Label48.Caption = "DOD"
    'Option31.Visible = False
    'Option32.Visible = False
    'Option38.Visible = False
    'Option29.Visible = False
    'Option33.Visible = False
    'Option34.Visible = False
    'Option35.Visible = False
    'Option37.Visible = False
    'Option30.Visible = False
'End Sub

Private Sub Option5_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    Option35.value = True
    Option35.Visible = True
    Option35.Caption = "Summary"
    Option34.Visible = True
    Option34.Caption = "By Gender"
    Option33.Visible = True
    Option33.Caption = "By Insured Type"
    Option29.Visible = True
    Option29.Caption = "By Race"
    Option37.Visible = True
    Option37.Caption = "By Plan No"
    Option30.Visible = True
    Option30.Caption = "By Nature of Claims"
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
    Call ListViewHeader3
End Sub

Private Sub Option6_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    ChkDOA.Visible = True
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    TxtDOA1.Visible = True
    TxtDOA2.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label49.Visible = True
    Label50.Visible = True
    Label47.Caption = "DOA"
    Label48.Caption = "DOD"
    Option35.value = True
    Option35.Visible = True
    Option35.Caption = "Summary"
    Option34.Visible = True
    Option34.Caption = "By Ageband"
    Option33.Visible = True
    Option33.Caption = "By Insured Type"
    Option29.Visible = True
    Option29.Caption = "By Race"
    Option37.Visible = True
    Option37.Caption = "By Plan No"
    Option30.Visible = True
    Option30.Caption = "By Nature of Claims"
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
    Call ListViewHeader14
End Sub

Private Sub Option7_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    ChkDOA.Visible = True
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    TxtDOA1.Visible = True
    TxtDOA2.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label49.Visible = True
    Label50.Visible = True
    Label47.Caption = "DOA"
    Label48.Caption = "DOD"
    Option35.value = True
    Option35.Visible = True
    Option35.Caption = "Summary"
    Option34.Visible = True
    Option34.Caption = "By Ageband"
    Option33.Visible = True
    Option33.Caption = "By Insured Type"
    Option29.Visible = True
    Option29.Caption = "By Gender"
    Option37.Visible = True
    Option37.Caption = "By Plan No"
    Option30.Visible = True
    Option30.Caption = "By Nature of Claims"
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
    Call ListViewHeader23
End Sub

Private Sub Option8_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    ChkDOA.Visible = True
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    TxtDOA1.Visible = True
    TxtDOA2.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label49.Visible = True
    Label50.Visible = True
    Label47.Caption = "DOA"
    Label48.Caption = "DOD"
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
    Option37.Visible = False
    Option30.Visible = False
    Call ListViewHeader27
End Sub

Private Sub Option9_Click()
    lblCurrent = ""
    lblTotalRecord = ""
    ListView2.listitems.Clear
    CboClaimability.Text = ""
    ChkDOA.Enabled = True
    ChkDOA.Visible = True
    ChkDOD.Enabled = True
    ChkDOD.Visible = True
    TxtDOA1.Visible = True
    TxtDOA2.Visible = True
    txtDOD1.Visible = True
    txtDOD2.Visible = True
    Label49.Visible = True
    Label50.Visible = True
    Label47.Caption = "DOA"
    Label48.Caption = "DOD"
    Option31.Visible = False
    Option32.Visible = False
    Option38.Visible = False
    Option29.Visible = False
    Option33.Visible = False
    Option34.Visible = False
    Option35.Visible = False
    Option37.Visible = False
    Option30.Visible = False
    Call ListViewHeader20
End Sub

'=========== Check Date Format ===========================

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

Private Sub TxtDOW1_LostFocus()
    If IsDate(txtDOW1) Then
    Else
        If txtDOW1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOW1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDOW2_LostFocus()
    If IsDate(txtDOW2) Then
    Else
        If txtDOW2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOW2.SetFocus
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

Private Sub TxtPayExpDate1_LostFocus()
    If IsDate(TxtPayExpDate1) Then
    Else
        If TxtPayExpDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPayExpDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtPayExpDate2_LostFocus()
    If IsDate(TxtPayExpDate2) Then
    Else
        If TxtPayExpDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtPayExpDate2.SetFocus
        End If
    End If
End Sub

Private Sub TxtDateJoin1_LostFocus()
    If IsDate(txtDateJoin1) Then
    Else
        If txtDateJoin1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateJoin1.SetFocus
        End If
    End If
End Sub

Private Sub TxtDateJoin2_LostFocus()
    If IsDate(txtDateJoin2) Then
    Else
        If txtDateJoin2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDateJoin2.SetFocus
        End If
    End If
End Sub

Private Sub txtDOB1_LostFocus()
    If IsDate(txtDOB1) Then
    Else
        If txtDOB1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOB1.SetFocus
        End If
    End If
End Sub

Private Sub txtDOB2_LostFocus()
    If IsDate(txtDOB2) Then
    Else
        If txtDOB2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            txtDOB2.SetFocus
        End If
    End If
End Sub
'=========== Check Date Format ===========================

'================ Monthy Reports Listing ===========================
Private Sub MonthyReportListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
       
    ListView2.listitems.Clear
    
            
    Sql = " SELECT Distinct PAYCode From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Order by PAYCode "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                    
            With rst2
                If Len(rst2!PAYCode) Then
                
                Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Payor = rst2!PAYCode
                
                If Option20.value = True And Option34.value = True Then
                    Call SearchAdmissionRecords(StrAppend)
                ElseIf Option21.value = True And Option34.value = True Then
                    Call SearchDischargeRecords(StrAppend)
                End If
                
                If Len(AdmissionDate1) And Len(Day1) Then
                    Record.SubItems(1) = AdmissionDate1 & "-" & Day1
                End If
                
                If Len(Total1) Then
                    Record.SubItems(2) = Total1
                End If
                
                If Len(AdmissionDate2) And Len(Day2) Then
                    Record.SubItems(3) = AdmissionDate2 & "-" & Day2
                End If
                
                If Len(Total2) Then
                    Record.SubItems(4) = Total2
                End If
                
                If Len(AdmissionDate3) And Len(Day3) Then
                    Record.SubItems(5) = AdmissionDate3 & "-" & Day3
                End If
                
                If Len(Total3) Then
                    Record.SubItems(6) = Total3
                End If
                
                If Len(AdmissionDate4) And Len(Day4) Then
                    Record.SubItems(7) = AdmissionDate4 & "-" & Day4
                End If
                
                If Len(Total4) Then
                    Record.SubItems(8) = Total4
                End If
                
                If Len(AdmissionDate5) And Len(Day5) Then
                    Record.SubItems(9) = AdmissionDate5 & "-" & Day5
                End If
                
                If Len(Total5) Then
                    Record.SubItems(10) = Total5
                End If
                
                If Len(AdmissionDate6) And Len(Day6) Then
                    Record.SubItems(11) = AdmissionDate6 & "-" & Day6
                End If
                
                If Len(Total6) Then
                    Record.SubItems(12) = Total6
                End If
                
                If Len(AdmissionDate7) And Len(Day7) Then
                    Record.SubItems(13) = AdmissionDate7 & "-" & Day7
                End If
                
                If Len(Total7) Then
                    Record.SubItems(14) = Total7
                End If
                
                If Len(AdmissionDate8) And Len(Day8) Then
                    Record.SubItems(15) = AdmissionDate8 & "-" & Day8
                End If
                
                If Len(Total8) Then
                    Record.SubItems(16) = Total8
                End If
                
                If Len(AdmissionDate9) And Len(Day9) Then
                    Record.SubItems(17) = AdmissionDate9 & "-" & Day9
                End If
                
                If Len(Total9) Then
                    Record.SubItems(18) = Total9
                End If
                
                If Len(AdmissionDate10) And Len(Day10) Then
                    Record.SubItems(19) = AdmissionDate10 & "-" & Day10
                End If
                
                If Len(Total10) Then
                    Record.SubItems(20) = Total10
                End If
                
                If Len(AdmissionDate11) And Len(Day11) Then
                    Record.SubItems(21) = AdmissionDate11 & "-" & Day11
                End If
                
                If Len(Total11) Then
                    Record.SubItems(22) = Total11
                End If
                
                If Len(AdmissionDate12) And Len(Day12) Then
                    Record.SubItems(23) = AdmissionDate12 & "-" & Day12
                End If
                
                If Len(Total12) Then
                    Record.SubItems(24) = Total12
                End If
                
                If Len(AdmissionDate13) And Len(Day13) Then
                    Record.SubItems(25) = AdmissionDate13 & "-" & Day13
                End If
                
                If Len(Total13) Then
                    Record.SubItems(26) = Total13
                End If
                
                If Len(AdmissionDate14) And Len(Day14) Then
                    Record.SubItems(27) = AdmissionDate14 & "-" & Day14
                End If
                
                If Len(Total14) Then
                    Record.SubItems(28) = Total14
                End If
                
                If Len(AdmissionDate15) And Len(Day15) Then
                    Record.SubItems(29) = AdmissionDate15 & "-" & Day15
                End If
                
                If Len(Total15) Then
                    Record.SubItems(30) = Total15
                End If
                
                If Len(AdmissionDate16) And Len(Day16) Then
                    Record.SubItems(31) = AdmissionDate16 & "-" & Day16
                End If
                
                If Len(Total16) Then
                    Record.SubItems(32) = Total6
                End If
                
                If Len(AdmissionDate17) And Len(Day17) Then
                    Record.SubItems(33) = AdmissionDate17 & "-" & Day17
                End If
                
                If Len(Total17) Then
                    Record.SubItems(34) = Total17
                End If
                
                If Len(AdmissionDate18) And Len(Day18) Then
                    Record.SubItems(35) = AdmissionDate18 & "-" & Day18
                End If
                
                If Len(Total18) Then
                    Record.SubItems(36) = Total8
                End If
                
                If Len(AdmissionDate19) And Len(Day19) Then
                    Record.SubItems(37) = AdmissionDate19 & "-" & Day19
                End If
                
                If Len(Total19) Then
                    Record.SubItems(38) = Total19
                End If
                
                If Len(AdmissionDate20) And Len(Day20) Then
                    Record.SubItems(39) = AdmissionDate20 & "-" & Day20
                End If
                
                If Len(Total20) Then
                    Record.SubItems(40) = Total20
                End If
                
                If Len(AdmissionDate21) And Len(Day21) Then
                    Record.SubItems(41) = AdmissionDate21 & "-" & Day21
                End If
                
                If Len(Total21) Then
                    Record.SubItems(42) = Total21
                End If
                
                If Len(AdmissionDate22) And Len(Day22) Then
                    Record.SubItems(43) = AdmissionDate22 & "-" & Day22
                End If
                
                If Len(Total22) Then
                    Record.SubItems(44) = Total22
                End If
                
                If Len(AdmissionDate23) And Len(Day23) Then
                    Record.SubItems(45) = AdmissionDate23 & "-" & Day23
                End If
                
                If Len(Total23) Then
                    Record.SubItems(46) = Total23
                End If
                
                If Len(AdmissionDate24) And Len(Day24) Then
                    Record.SubItems(47) = AdmissionDate24 & "-" & Day24
                End If
                
                If Len(Total24) Then
                    Record.SubItems(48) = Total24
                End If
                
                If Len(AdmissionDate25) And Len(Day24) Then
                    Record.SubItems(49) = AdmissionDate25 & "-" & Day25
                End If
                
                If Len(Total25) Then
                    Record.SubItems(50) = Total25
                End If
                
                If Len(AdmissionDate26) And Len(Day26) Then
                    Record.SubItems(51) = AdmissionDate26 & "-" & Day26
                End If
                
                If Len(Total26) Then
                    Record.SubItems(52) = Total26
                End If
                
                If Len(AdmissionDate27) And Len(Day27) Then
                    Record.SubItems(53) = AdmissionDate27 & "-" & Day27
                End If
                
                If Len(Total27) Then
                    Record.SubItems(54) = Total27
                End If
                
                If Len(AdmissionDate28) And Len(Day28) Then
                    Record.SubItems(55) = AdmissionDate28 & "-" & Day28
                End If
                
                If Len(Total28) Then
                    Record.SubItems(56) = Total28
                End If
                
                If Len(AdmissionDate29) And Len(Day29) Then
                    Record.SubItems(57) = AdmissionDate29 & "-" & Day29
                End If
                
                If Len(Total29) Then
                    Record.SubItems(58) = Total29
                End If
                
                If Len(AdmissionDate30) And Len(Day30) Then
                    Record.SubItems(59) = AdmissionDate30 & "-" & Day30
                End If
                
                If Len(Total30) Then
                    Record.SubItems(60) = Total30
                End If
                
                If Len(AdmissionDate31) And Len(Day31) Then
                    Record.SubItems(61) = AdmissionDate31 & "-" & Day31
                End If
                
                If Len(Total31) Then
                    Record.SubItems(62) = Total31
                End If
                End If
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub
'================ Admission Listing ===========================

'==================== Search One Month Admission Records ===========================
Private Sub SearchAdmissionRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String

    ListView3.listitems.Clear
    Call ClearBoolean
    
    Sql = " Select PAYCode, CASDateAdmitted, Count(PAYCode) AS Total From VIEW_CLMReport Where PAYCode = '" & Trim(Payor) & "'"
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + (" Group By PAYCode, CASDateAdmitted Order by CASDateAdmitted")
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
                
        Do Until rst2.EOF
        
            With rst2
            
            
                TxtDay.Text = ""
                txtDate.Text = ""
                
                Set Record = ListView3.listitems.Add(, , rst2!PAYCode)

                
                If one = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate1 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day1 = Day
                        Record.SubItems(1) = AdmissionDate1 & "-" & Day1
                        Record.SubItems(2) = rst2!Total
                        Total1 = rst2!Total
                        one = True
                    End If
                
                ElseIf two = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate2 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day2 = Day
                        Record.SubItems(1) = AdmissionDate2 & "-" & Day2
                        Record.SubItems(2) = rst2!Total
                        Total2 = rst2!Total
                        two = True
                    End If
                
                ElseIf three = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate3 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day3 = Day
                        Record.SubItems(1) = AdmissionDate3 & "-" & Day3
                        Record.SubItems(2) = rst2!Total
                        Total3 = rst2!Total
                        three = True
                    End If
                
                ElseIf four = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate4 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day4 = Day
                        Record.SubItems(1) = AdmissionDate4 & "-" & Day4
                        Record.SubItems(2) = rst2!Total
                        Total4 = rst2!Total
                        four = True
                    End If
                ElseIf five = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate5 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day5 = Day
                        Record.SubItems(1) = AdmissionDate5 & "-" & Day5
                        Record.SubItems(2) = rst2!Total
                        Total5 = rst2!Total
                        five = True
                    End If
                ElseIf six = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate6 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day6 = Day
                        Record.SubItems(1) = AdmissionDate6 & "-" & Day6
                        Record.SubItems(2) = rst2!Total
                        Total6 = rst2!Total
                        six = True
                    End If
                ElseIf seven = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate7 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day7 = Day
                        Record.SubItems(1) = AdmissionDate7 & "-" & Day7
                        Record.SubItems(2) = rst2!Total
                        Total7 = rst2!Total
                        seven = True
                    End If
                ElseIf eight = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate8 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day8 = Day
                        Record.SubItems(1) = AdmissionDate8 & "-" & Day8
                        Record.SubItems(2) = rst2!Total
                        Total8 = rst2!Total
                        eight = True
                    End If
                ElseIf nine = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate9 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day9 = Day
                        Record.SubItems(1) = AdmissionDate9 & "-" & Day9
                        Record.SubItems(2) = rst2!Total
                        Total9 = rst2!Total
                        nine = True
                    End If
                ElseIf ten = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate10 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day10 = Day
                        Record.SubItems(1) = AdmissionDate10 & "-" & Day10
                        Record.SubItems(2) = rst2!Total
                        Total10 = rst2!Total
                        ten = True
                    End If
                ElseIf eleven = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate11 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day11 = Day
                        Record.SubItems(1) = AdmissionDate11 & "-" & Day11
                        Record.SubItems(2) = rst2!Total
                        Total11 = rst2!Total
                        eleven = True
                    End If
                ElseIf twelve = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate12 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day12 = Day
                        Record.SubItems(1) = AdmissionDate12 & "-" & Day12
                        Record.SubItems(2) = rst2!Total
                        Total12 = rst2!Total
                        twelve = True
                    End If
                ElseIf thirteen = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate13 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day13 = Day
                        Record.SubItems(1) = AdmissionDate13 & "-" & Day13
                        Record.SubItems(2) = rst2!Total
                        Total13 = rst2!Total
                        thirteen = True
                    End If
                ElseIf fourteen = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate14 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day14 = Day
                        Record.SubItems(1) = AdmissionDate14 & "-" & Day14
                        Record.SubItems(2) = rst2!Total
                        Total14 = rst2!Total
                        fourteen = True
                    End If
                ElseIf fiftteen = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate15 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day15 = Day
                        Record.SubItems(1) = AdmissionDate15 & "-" & Day15
                        Record.SubItems(2) = rst2!Total
                        Total15 = rst2!Total
                        fiftteen = True
                    End If
                ElseIf sixteen = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate16 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day16 = Day
                        Record.SubItems(1) = AdmissionDate16 & "-" & Day16
                        Record.SubItems(2) = rst2!Total
                        Total16 = rst2!Total
                        sixteen = True
                    End If
                ElseIf seventeen = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate17 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day17 = Day
                        Record.SubItems(1) = AdmissionDate17 & "-" & Day17
                        Record.SubItems(2) = rst2!Total
                        Total17 = rst2!Total
                        seventeen = True
                    End If
                ElseIf eightteen = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate18 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day18 = Day
                        Record.SubItems(1) = AdmissionDate18 & "-" & Day18
                        Record.SubItems(2) = rst2!Total
                        Total18 = rst2!Total
                        eightteen = True
                    End If
                ElseIf nineteen = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate19 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day19 = Day
                        Record.SubItems(1) = AdmissionDate19 & "-" & Day19
                        Record.SubItems(2) = rst2!Total
                        Total19 = rst2!Total
                        nineteen = True
                    End If
                ElseIf twenty = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate20 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day20 = Day
                        Record.SubItems(1) = AdmissionDate20 & "-" & Day20
                        Record.SubItems(2) = rst2!Total
                        Total20 = rst2!Total
                        twenty = True
                    End If
                ElseIf twentyone = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate21 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day21 = Day
                        Record.SubItems(1) = AdmissionDate21 & "-" & Day21
                        Record.SubItems(2) = rst2!Total
                        Total21 = rst2!Total
                        twentyone = True
                    End If
                ElseIf twentytwo = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate22 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day22 = Day
                        Record.SubItems(1) = AdmissionDate22 & "-" & Day22
                        Record.SubItems(2) = rst2!Total
                        Total22 = rst2!Total
                        twentytwo = True
                    End If
                ElseIf twentythree = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate23 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day20 = Day
                        Record.SubItems(1) = AdmissionDate23 & "-" & Day23
                        Record.SubItems(2) = rst2!Total
                        Total23 = rst2!Total
                        twentythree = True
                    End If
                ElseIf twentyfour = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate24 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day24 = Day
                        Record.SubItems(1) = AdmissionDate24 & "-" & Day24
                        Record.SubItems(2) = rst2!Total
                        Total24 = rst2!Total
                        twentyfour = True
                    End If
                ElseIf twentyfive = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate25 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day25 = Day
                        Record.SubItems(1) = AdmissionDate25 & "-" & Day25
                        Record.SubItems(2) = rst2!Total
                        Total25 = rst2!Total
                        twentyfive = True
                    End If
                ElseIf twentysix = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate26 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day26 = Day
                        Record.SubItems(1) = AdmissionDate26 & "-" & Day26
                        Record.SubItems(2) = rst2!Total
                        Total26 = rst2!Total
                        twentysix = True
                    End If
                ElseIf twentyseven = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate27 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day27 = Day
                        Record.SubItems(1) = AdmissionDate27 & "-" & Day27
                        Record.SubItems(2) = rst2!Total
                        Total27 = rst2!Total
                        twentyseven = True
                    End If
                ElseIf twentyeight = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate28 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day28 = Day
                        Record.SubItems(1) = AdmissionDate28 & "-" & Day28
                        Record.SubItems(2) = rst2!Total
                        Total28 = rst2!Total
                        twentyeight = True
                    End If
                ElseIf twentynine = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate29 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day29 = Day
                        Record.SubItems(1) = AdmissionDate29 & "-" & Day29
                        Record.SubItems(2) = rst2!Total
                        Total29 = rst2!Total
                        twentynine = True
                    End If
                ElseIf thirty = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate30 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day30 = Day
                        Record.SubItems(1) = AdmissionDate30 & "-" & Day30
                        Record.SubItems(2) = rst2!Total
                        Total30 = rst2!Total
                        thirty = True
                    End If
                ElseIf thirtyone = False Then
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(1) = rst2!CASDateAdmitted
                        AdmissionDate31 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day31 = Day
                        Record.SubItems(1) = AdmissionDate31 & "-" & Day31
                        Record.SubItems(2) = rst2!Total
                        Total31 = rst2!Total
                        thirtyone = True
                    End If
                End If
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'==================== Search One Month Admission Records ===========================

'==================== Search One Month Discharge Records ===========================
Private Sub SearchDischargeRecords(Optional strAppend3 As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String

    ListView3.listitems.Clear
    Call ClearBoolean
    
    Sql = " Select PAYCode, CASDischargeDate, Count(PAYCode) AS Total From VIEW_CLMReport Where PAYCode = '" & Trim(Payor) & "'"

    
    If Len(Trim(strAppend3)) Then
        Sql = Sql + strAppend3 + (" Group By PAYCode, CASDischargeDate ")
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
        
            With rst2
            
                TxtDay.Text = ""
                txtDate.Text = ""
                
                Set Record = ListView3.listitems.Add(, , rst2!PAYCode)

                
                If one = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate1 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day1 = Day
                        Record.SubItems(1) = AdmissionDate1 & "-" & Day1
                        Record.SubItems(2) = rst2!Total
                        Total1 = rst2!Total
                        one = True
                    End If
                
                ElseIf two = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate2 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day2 = Day
                        Record.SubItems(1) = AdmissionDate2 & "-" & Day2
                        Record.SubItems(2) = rst2!Total
                        Total2 = rst2!Total
                        two = True
                    End If
                
                ElseIf three = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate3 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day3 = Day
                        Record.SubItems(1) = AdmissionDate3 & "-" & Day3
                        Record.SubItems(2) = rst2!Total
                        Total3 = rst2!Total
                        three = True
                    End If
                
                ElseIf four = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate4 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day4 = Day
                        Record.SubItems(1) = AdmissionDate4 & "-" & Day4
                        Record.SubItems(2) = rst2!Total
                        Total4 = rst2!Total
                        four = True
                    End If
                ElseIf five = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate5 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day5 = Day
                        Record.SubItems(1) = AdmissionDate5 & "-" & Day5
                        Record.SubItems(2) = rst2!Total
                        Total5 = rst2!Total
                        five = True
                    End If
                ElseIf six = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate6 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day6 = Day
                        Record.SubItems(1) = AdmissionDate6 & "-" & Day6
                        Record.SubItems(2) = rst2!Total
                        Total6 = rst2!Total
                        six = True
                    End If
                ElseIf seven = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate7 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day7 = Day
                        Record.SubItems(1) = AdmissionDate7 & "-" & Day7
                        Record.SubItems(2) = rst2!Total
                        Total7 = rst2!Total
                        seven = True
                    End If
                ElseIf eight = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate8 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day8 = Day
                        Record.SubItems(1) = AdmissionDate8 & "-" & Day8
                        Record.SubItems(2) = rst2!Total
                        Total8 = rst2!Total
                        eight = True
                    End If
                ElseIf nine = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate9 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day9 = Day
                        Record.SubItems(1) = AdmissionDate9 & "-" & Day9
                        Record.SubItems(2) = rst2!Total
                        Total9 = rst2!Total
                        nine = True
                    End If
                ElseIf ten = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate10 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day10 = Day
                        Record.SubItems(1) = AdmissionDate10 & "-" & Day10
                        Record.SubItems(2) = rst2!Total
                        Total10 = rst2!Total
                        ten = True
                    End If
                ElseIf eleven = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate11 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day11 = Day
                        Record.SubItems(1) = AdmissionDate11 & "-" & Day11
                        Record.SubItems(2) = rst2!Total
                        Total11 = rst2!Total
                        eleven = True
                    End If
                ElseIf twelve = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate12 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day12 = Day
                        Record.SubItems(1) = AdmissionDate12 & "-" & Day12
                        Record.SubItems(2) = rst2!Total
                        Total12 = rst2!Total
                        twelve = True
                    End If
                ElseIf thirteen = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate13 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day13 = Day
                        Record.SubItems(1) = AdmissionDate13 & "-" & Day13
                        Record.SubItems(2) = rst2!Total
                        Total13 = rst2!Total
                        thirteen = True
                    End If
                ElseIf fourteen = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate14 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day14 = Day
                        Record.SubItems(1) = AdmissionDate14 & "-" & Day14
                        Record.SubItems(2) = rst2!Total
                        Total14 = rst2!Total
                        fourteen = True
                    End If
                ElseIf fiftteen = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate15 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day15 = Day
                        Record.SubItems(1) = AdmissionDate15 & "-" & Day15
                        Record.SubItems(2) = rst2!Total
                        Total15 = rst2!Total
                        fiftteen = True
                    End If
                ElseIf sixteen = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate16 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day16 = Day
                        Record.SubItems(1) = AdmissionDate16 & "-" & Day16
                        Record.SubItems(2) = rst2!Total
                        Total16 = rst2!Total
                        sixteen = True
                    End If
                ElseIf seventeen = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate17 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day17 = Day
                        Record.SubItems(1) = AdmissionDate17 & "-" & Day17
                        Record.SubItems(2) = rst2!Total
                        Total17 = rst2!Total
                        seventeen = True
                    End If
                ElseIf eightteen = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate18 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day18 = Day
                        Record.SubItems(1) = AdmissionDate18 & "-" & Day18
                        Record.SubItems(2) = rst2!Total
                        Total18 = rst2!Total
                        eightteen = True
                    End If
                ElseIf nineteen = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate19 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day19 = Day
                        Record.SubItems(1) = AdmissionDate19 & "-" & Day19
                        Record.SubItems(2) = rst2!Total
                        Total19 = rst2!Total
                        nineteen = True
                    End If
                ElseIf twenty = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate20 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day20 = Day
                        Record.SubItems(1) = AdmissionDate20 & "-" & Day20
                        Record.SubItems(2) = rst2!Total
                        Total20 = rst2!Total
                        twenty = True
                    End If
                ElseIf twentyone = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate21 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day21 = Day
                        Record.SubItems(1) = AdmissionDate21 & "-" & Day21
                        Record.SubItems(2) = rst2!Total
                        Total21 = rst2!Total
                        twentyone = True
                    End If
                ElseIf twentytwo = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate22 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day22 = Day
                        Record.SubItems(1) = AdmissionDate22 & "-" & Day22
                        Record.SubItems(2) = rst2!Total
                        Total22 = rst2!Total
                        twentytwo = True
                    End If
                ElseIf twentythree = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate23 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day20 = Day
                        Record.SubItems(1) = AdmissionDate23 & "-" & Day23
                        Record.SubItems(2) = rst2!Total
                        Total23 = rst2!Total
                        twentythree = True
                    End If
                ElseIf twentyfour = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate24 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day24 = Day
                        Record.SubItems(1) = AdmissionDate24 & "-" & Day24
                        Record.SubItems(2) = rst2!Total
                        Total24 = rst2!Total
                        twentyfour = True
                    End If
                ElseIf twentyfive = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate25 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day25 = Day
                        Record.SubItems(1) = AdmissionDate25 & "-" & Day25
                        Record.SubItems(2) = rst2!Total
                        Total25 = rst2!Total
                        twentyfive = True
                    End If
                ElseIf twentysix = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate26 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day26 = Day
                        Record.SubItems(1) = AdmissionDate26 & "-" & Day26
                        Record.SubItems(2) = rst2!Total
                        Total26 = rst2!Total
                        twentysix = True
                    End If
                ElseIf twentyseven = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate27 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day27 = Day
                        Record.SubItems(1) = AdmissionDate27 & "-" & Day27
                        Record.SubItems(2) = rst2!Total
                        Total27 = rst2!Total
                        twentyseven = True
                    End If
                ElseIf twentyeight = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate28 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day28 = Day
                        Record.SubItems(1) = AdmissionDate28 & "-" & Day28
                        Record.SubItems(2) = rst2!Total
                        Total28 = rst2!Total
                        twentyeight = True
                    End If
                ElseIf twentynine = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate29 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day29 = Day
                        Record.SubItems(1) = AdmissionDate29 & "-" & Day29
                        Record.SubItems(2) = rst2!Total
                        Total29 = rst2!Total
                        twentynine = True
                    End If
                ElseIf thirty = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate30 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day30 = Day
                        Record.SubItems(1) = AdmissionDate30 & "-" & Day30
                        Record.SubItems(2) = rst2!Total
                        Total30 = rst2!Total
                        thirty = True
                    End If
                ElseIf thirtyone = False Then
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(1) = rst2!CASDischargeDate
                        AdmissionDate31 = Record.SubItems(1)
                        txtDate = Record.SubItems(1)
                        Call CheckDay
                        Day31 = Day
                        Record.SubItems(1) = AdmissionDate31 & "-" & Day31
                        Record.SubItems(2) = rst2!Total
                        Total31 = rst2!Total
                        thirtyone = True
                    End If
                End If
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'==================== Search One Month Discharge Records ===========================


'==================== Check Day ===========================
Private Sub CheckDay()
Dim intYears As Integer 'Used as a counter
Dim intNumDaysToCheck As Long 'Used to limit search
Dim tmpDate As Date 'Temporary date variable
Dim tmpMonth As Date 'Temporary month variable
Dim intPaydays As Integer 'Counts number of paydays Each month
Dim intPayDayOfWeek As Integer 'Pay Day of Week
Dim intX As Long 'Counter
    'computes number of days between pay date and 10 years from paydate
    intNumDaysToCheck = DateDiff("d", CDate(txtDate), DateAdd("yyyy", 100, CDate(txtDate)))
    intPayDayOfWeek = Weekday(CDate(txtDate))
    intPaydays = 0
    tmpMonth = DatePart("m", CDate(txtDate))
    'get paid every two weeks
    'if you get paid every friday and want to check for
    '5 paycheck months, change this to step 7 and change
    'commented value below


    For intX = 0 To intNumDaysToCheck Step 14
        'set the day we are working with
        tmpDate = CDate(txtDate) + intX
        'make sure the month hasn't changed


        If tmpMonth = DatePart("m", tmpDate) Then
            'if it is a payday


            If Weekday(tmpDate) <> intPayDayOfWeek Then
                'there is a problem
                MsgBox "An Error was encountered. Calculation cannot continue"
            Else
                'increment the paydays
                intPaydays = intPaydays + 1
                'if you have arrived at a 3 payday month, add it to the list
                'if you get weekly paychecks, change this value from 3 to 5
            End If
        Else
            'change the month
            tmpMonth = DatePart("m", tmpDate)
            'set the number of paydays to 1
            intPaydays = 1
        End If
    Next
    
    Call DateDay
End Sub
'==================== Check Day ===========================

' ================= Call Day ==============================
Private Sub DateDay()

    Select Case Weekday(CDate(txtDate))
        Case vbMonday
        TxtDay = "Monday"
        Case vbTuesday
        TxtDay = "Tuesday"
        Case vbWednesday
        TxtDay = "Wednesday"
        Case vbThursday
        TxtDay = "Thursday"
        Case vbFriday
        TxtDay = "Friday"
        Case vbSaturday
        TxtDay = "Saturday"
        Case vbSunday
        TxtDay = "Sunday"
        Case Else
        TxtDay = "Error With date"
    End Select
    
    Day = TxtDay
End Sub
' ================= Call Day ==============================

'=================== Clear String Records ========================
Public Function ClearBoolean()
    one = False
    two = False
    three = False
    four = False
    five = False
    six = False
    seven = False
    eight = False
    nine = False
    ten = False
    eleven = False
    twelve = False
    thirteen = False
    fourteen = False
    fifteen = False
    sixteen = False
    seventeen = False
    eightteen = False
    nineteen = False
    twenty = False
    twentyone = False
    twentytwo = False
    twentythree = False
    twentyfour = False
    twentyfive = False
    twentysix = False
    twentyseven = False
    twentyeight = False
    twentynine = False
    thirty = False
    thirtyone = False
    
    AdmissionDate1 = ""
    Day1 = ""
    Total1 = ""
    
    AdmissionDate2 = ""
    Day2 = ""
    Total2 = ""
    
        
    AdmissionDate3 = ""
    Day3 = ""
    Total3 = ""
    
    AdmissionDate4 = ""
    Day4 = ""
    Total4 = ""
    
    AdmissionDate5 = ""
    Day5 = ""
    Total5 = ""
    
    AdmissionDate6 = ""
    Day6 = ""
    Total6 = ""
    
    AdmissionDate7 = ""
    Day7 = ""
    Total7 = ""
    
    AdmissionDate8 = ""
    Day8 = ""
    Total8 = ""
    
    AdmissionDate9 = ""
    Day9 = ""
    Total9 = ""
    
    AdmissionDate10 = ""
    Day10 = ""
    Total10 = ""
    
    AdmissionDate11 = ""
    Day11 = ""
    Total11 = ""
    
    AdmissionDate12 = ""
    Day12 = ""
    Total12 = ""
    
    AdmissionDate13 = ""
    Day13 = ""
    Total13 = ""
    
    AdmissionDate14 = ""
    Day14 = ""
    Total14 = ""
    
    AdmissionDate15 = ""
    Day15 = ""
    Total15 = ""
    
    AdmissionDate16 = ""
    Day16 = ""
    Total16 = ""
    
    AdmissionDate17 = ""
    Day17 = ""
    Total17 = ""
    
    AdmissionDate18 = ""
    Day18 = ""
    Total18 = ""
    
    AdmissionDate19 = ""
    Day19 = ""
    Total19 = ""
    
    AdmissionDate20 = ""
    Day20 = ""
    Total20 = ""
    
    AdmissionDate21 = ""
    Day21 = ""
    Total21 = ""
    
    AdmissionDate22 = ""
    Day22 = ""
    Total22 = ""
    
    AdmissionDate23 = ""
    Day23 = ""
    Total23 = ""
    
    AdmissionDate24 = ""
    Day24 = ""
    Total24 = ""
    
    AdmissionDate25 = ""
    Day25 = ""
    Total25 = ""
    
    AdmissionDate26 = ""
    Day26 = ""
    Total26 = ""
    
    AdmissionDate27 = ""
    Day27 = ""
    Total27 = ""
    
    AdmissionDate28 = ""
    Day28 = ""
    Total28 = ""
    
    AdmissionDate29 = ""
    Day29 = ""
    Total29 = ""
    
    AdmissionDate30 = ""
    Day30 = ""
    Total30 = ""
    
    AdmissionDate31 = ""
    Day31 = ""
    Total31 = ""
    
End Function
'=================== Clear String Records ========================

' =================== Search Year Admission / Discharge Results ======================
Private Function SearchYearReports()

    If TxtYear = "" Then
        MsgBox "Please Enter Year", vbOKOnly + vbCritical, DialogBoxTitle
        TxtYear.SetFocus
    Else
        StrAppend = ""
    
        If Len(Trim(TxtYear)) Then
            StrAppend = StrAppend + " AND Year = '" & Trim(TxtYear) & "'"
        End If
        
        If Option33.value = True Then
            Sql = ""
            Sql = "AND CASStatus = 'O' "
            StrAppend = StrAppend + Sql
        End If
        
        If Option34.value = True Then
            Sql = ""
            Sql = "AND CASStatus = 'C'"
            StrAppend = StrAppend + Sql
        End If
        
        If Option35.value = True Then
            Sql = ""
            Sql = "AND CASStatus = 'D'"
            StrAppend = StrAppend + Sql
        End If
        
        If Option36.value = True Then
            Sql = ""
            Sql = "AND (CASStatus = 'O' OR CASStatus = 'C')"
            StrAppend = StrAppend + Sql
        End If
        
        If Option37.value = True Then
            Sql = ""
            Sql = "AND CASClaimability = 'A'"
            StrAppend = StrAppend + Sql
        End If
        
        If Option38.value = True Then
            Sql = ""
            Sql = "AND CASClaimability = 'R'"
            StrAppend = StrAppend + Sql
        End If
    
        'Call YearListing(strAppend)
    End If
End Function
' =================== Search Year Admission / Discharge Results ======================

'================ Admission / Discharge Year Listing ===========================
Private Sub YearListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
       
    Payor = ""
    ListView2.listitems.Clear
    
    Sql = " SELECT Distinct PAYCode From VIEW_CLMReport Where 0 = 0 Order by PAYCode"
       
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                    
            With rst2
                If Len(rst2!PAYCode) Then
                
                Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Payor = rst2!PAYCode
                
                If Option20.value = True And Option33.value = True Then
                    Call SearchAdmissionYearRecords(StrAppend)
                ElseIf Option21.value = True And Option33.value = True Then
                    Call SearchDischargeYearRecords(StrAppend)
                End If
                
                If Len(Total1) Then
                    Record.SubItems(1) = Total1
                End If
                                
                If Len(Total2) Then
                    Record.SubItems(2) = Total2
                End If
                
                If Len(Total3) Then
                    Record.SubItems(3) = Total3
                End If
                
                If Len(Total4) Then
                    Record.SubItems(4) = Total4
                End If
                                
                If Len(Total5) Then
                    Record.SubItems(5) = Total5
                End If
                
                If Len(Total6) Then
                    Record.SubItems(6) = Total6
                End If
                
                If Len(Total7) Then
                    Record.SubItems(7) = Total7
                End If
                
                If Len(Total8) Then
                    Record.SubItems(8) = Total8
                End If
                
                If Len(Total9) Then
                    Record.SubItems(9) = Total9
                End If
                
                If Len(Total10) Then
                    Record.SubItems(10) = Total10
                End If
                
                If Len(Total11) Then
                    Record.SubItems(11) = Total11
                End If
                
                If Len(Total12) Then
                    Record.SubItems(12) = Total12
                End If
                End If
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub
'================ Admission / Discharge Year Listing ===========================

'==================== Search One Year Admission Records ===========================
Private Sub SearchAdmissionYearRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String

    ListView3.listitems.Clear
    Call ClearBoolean
    
    Sql = " SELECT PAYCode, Month(CASDateAdmitted) As Month, " & _
          " COUNT(CASNumber) AS TotalCases " & _
          " FROM VIEW_CLMReport Where PAYCode = '" & Trim(Payor) & "'"
          
    'sql = " Select * From VIEW_DOAYearReport Where PAYCode = '" & Trim(Payor) & "'"
    
    sql2 = " GROUP BY PAYCode, MONTH(CASDateAdmitted), YEAR(CASDateAdmitted) ORDER BY MONTH(CASDateAdmitted) "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
         
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
        
            With rst2
                               
                Set Record = ListView3.listitems.Add(, , rst2!PAYCode)
    
                    If rst2!Month = "1" And one = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total1 = rst2!TotalCases
                            one = True
                        End If
                
                    ElseIf rst2!Month = "2" And two = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total2 = rst2!TotalCases
                            two = True
                        End If
                
                    ElseIf rst2!Month = "3" And three = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total3 = rst2!TotalCases
                            three = True
                        End If
                
                    ElseIf rst2!Month = "4" And four = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total4 = rst2!TotalCases
                            four = True
                        End If
                
                    ElseIf rst2!Month = "5" And five = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total5 = rst2!TotalCases
                            five = True
                        End If
                
                    ElseIf rst2!Month = "6" And six = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total6 = rst2!TotalCases
                            six = True
                        End If
                
                    ElseIf rst2!Month = "7" And seven = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total7 = rst2!TotalCases
                            seven = True
                        End If
                
                    ElseIf rst2!Month = "8" And eight = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total8 = rst2!TotalCases
                            eight = True
                        End If
                
                    ElseIf rst2!Month = "9" And nine = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total9 = rst2!TotalCases
                            nine = True
                        End If
                
                    ElseIf rst2!Month = "10" And ten = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total10 = rst2!TotalCases
                            ten = True
                        End If
                
                    ElseIf rst2!Month = "11" And eleven = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total11 = rst2!TotalCases
                            eleven = True
                        End If
                
                    ElseIf rst2!Month = "12" And twelve = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total12 = rst2!TotalCases
                            twelve = True
                        End If
                    End If
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'==================== Search One Year Admission Records ===========================

'==================== Search One Year Discharge Records ===========================
Private Sub SearchDischargeYearRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String

    ListView3.listitems.Clear
    Call ClearBoolean
    
    Sql = " SELECT PAYCode, Month(CASDischargeDate) As Month, " & _
          " YEAR(CASDischargeDate) AS Year, COUNT(CASNumber) AS TotalCases " & _
          " FROM VIEW_CLMReport Where PAYCode = '" & Trim(Payor) & "'"
          
       
    sql2 = " GROUP BY PAYCode, MONTH(CASDischargeDate), YEAR(CASDischargeDate) ORDER BY MONTH(CASDischargeDate) "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
        
            With rst2
                               
                Set Record = ListView3.listitems.Add(, , rst2!PAYCode)
    
                    If rst2!Month = "1" And one = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total1 = rst2!TotalCases
                            one = True
                        End If
                
                    ElseIf rst2!Month = "2" And two = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total2 = rst2!TotalCases
                            two = True
                        End If
                
                    ElseIf rst2!Month = "3" And three = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total3 = rst2!TotalCases
                            three = True
                        End If
                
                    ElseIf rst2!Month = "4" And four = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total4 = rst2!TotalCases
                            four = True
                        End If
                
                    ElseIf rst2!Month = "5" And five = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total5 = rst2!TotalCases
                            five = True
                        End If
                
                    ElseIf rst2!Month = "6" And six = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total6 = rst2!TotalCases
                            six = True
                        End If
                
                    ElseIf rst2!Month = "7" And seven = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total7 = rst2!TotalCases
                            seven = True
                        End If
                
                    ElseIf rst2!Month = "8" And eight = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total8 = rst2!TotalCases
                            eight = True
                        End If
                
                    ElseIf rst2!Month = "9" And nine = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total9 = rst2!TotalCases
                            nine = True
                        End If
                
                    ElseIf rst2!Month = "10" And ten = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total10 = rst2!TotalCases
                            ten = True
                        End If
                
                    ElseIf rst2!Month = "11" And eleven = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total11 = rst2!TotalCases
                            eleven = True
                        End If
                
                    ElseIf rst2!Month = "12" And twelve = False Then
                        If Len(rst2!Month) Then
                            Record.SubItems(1) = rst2!TotalCases
                            Total12 = rst2!TotalCases
                            twelve = True
                        End If
                    End If
                
            End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'==================== Search One Year Discharge Records ===========================

Private Function SearchRecords()
Dim Sql As String
Dim StrAppend As String

    StrAppend = ""
    GrpCompany = ""
    Hospitals = ""

    
    If Len(Mid(cboHLTCode, 1, 1)) Then
        StrAppend = StrAppend + " AND HLTCode ='" & Trim(Mid(cboHLTCode, 1, 1)) & "'"
    End If

    If Len(Mid(CboInsType, 1, 1)) Then
        StrAppend = StrAppend & " AND INSCode = '" & Trim(Mid(CboInsType, 1, 1)) & "'"
    End If
    
    If Len(Trim(TxtAgeband1)) Then
        StrAppend = StrAppend + " AND AGEBand >= '" & Trim(TxtAgeband1) & "'"
    End If
    
    If Len(Trim(TxtAgeband2)) Then
        StrAppend = StrAppend + " AND AGEBand <= '" & Trim(TxtAgeband2) & "'"
    End If
    
    If Len(Trim(TxtICNo1)) Then
        StrAppend = StrAppend + " AND ICNo >= '" & Trim(TxtICNo1) & "'"
    End If
    If Len(Trim(TxtICNo2)) Then
        StrAppend = StrAppend + " AND ICNo <= '" & Trim(TxtICNo2) & "'"
    End If
    If ChkICNo.value = 1 Then
        StrAppend = StrAppend + " AND (ICNo Is null)"
    End If
       
     If Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(cboSuppType) = "Both" Then
         StrAppend = StrAppend & " AND (MBMMemberType = 'P' Or MBMMemberType = 'Q') And MBMPatientCovID = '00' "
      ElseIf Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
         StrAppend = StrAppend & " AND MBMMemberType = 'P'  And MBMPatientCovID = '00'"
      ElseIf Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
         StrAppend = StrAppend & " AND MBMMemberType = 'Q' And MBMPatientCovID = '00'"
      ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(cboSuppType) = "Both" Then
         StrAppend = StrAppend & " AND (MBMCMemberType = 'S' OR MBMCMemberType = 'C')"
      ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
         StrAppend = StrAppend & " AND MBMCMemberType = 'C'"
      ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
         StrAppend = StrAppend & " AND MBMCMemberType = 'S'"
      ElseIf Trim(CboMemType) = "Both" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
         StrAppend = StrAppend & " AND (MBMMemberType = 'P' OR MBMCMemberType = 'C')"
      ElseIf Trim(CboMemType) = "Both" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
         StrAppend = StrAppend & " AND (MBMMemberType = 'Q' OR MBMCMemberType = 'S')"
      End If
    
    If Len(Mid(CboAdultChild, 1, 1)) Then
        StrAppend = StrAppend & " AND AdultChild = '" & Trim(Mid(CboAdultChild, 1, 1)) & "'"
    End If
    
    If Len(Mid(CboMarital, 1, 1)) Then
        StrAppend = StrAppend & " AND MBMMaritalStatus = '" & Trim(Mid(CboMarital, 1, 1)) & "'"
    End If
    
    If Len(Mid(CboRace, 1, 1)) Then
        StrAppend = StrAppend & " AND RACCode = '" & Trim(Mid(CboRace, 1, 1)) & "'"
    End If
    
    If Len(Mid(CboSex, 1, 1)) Then
        StrAppend = StrAppend & " AND Sex = '" & Trim(Mid(CboSex, 1, 1)) & "'"
    End If
    
    If Len(Mid(CboRelation, 1, 1)) Then
        StrAppend = StrAppend & " AND Relation = '" & Trim(Mid(CboRelation, 1, 1)) & "'"
    End If
    
    If Len(Mid(cboDataStatus, 1, 1)) Then
        StrAppend = StrAppend & " AND MBMDataStatus = '" & Trim(Mid(cboDataStatus, 1, 1)) & "'"
    End If
    
    If Len(Mid(CboClaimability, 1, 1)) Then
        StrAppend = StrAppend & " AND CASClaimability = '" & Trim(Mid(CboClaimability, 1, 1)) & "'"
    End If

    If Len(Mid(CboStatus, 1, 1)) Then
        StrAppend = StrAppend & " AND MBMStatus = '" & Trim(Mid(CboStatus, 1, 1)) & "'"
    End If

    If Len(Mid(CboCaseStatus, 1, 1)) Then
        StrAppend = StrAppend & " AND CASStatus = '" & Trim(Mid(CboCaseStatus, 1, 1)) & "'"
    End If
    
    If Len(Trim(Mid(CboExpired, 1, 1))) Then
          StrAppend = StrAppend & "AND MBMExpiredStatus = '" & Trim(Mid(CboExpired, 1, 1)) & "' "
    End If
    
    If Len(Mid(CboPreMode, 1, 1)) Then
        StrAppend = StrAppend & " AND MBMInstallment = '" & Trim(Mid(CboPreMode, 1, 1)) & "'"
    End If
    
    If Len(Trim(CboState)) Then
          StrAppend = StrAppend & "AND STACode= '" & Trim(CboState) & "' "
    End If
    
    If Len(Mid(cboRenewalInd, 1, 1)) Then
        StrAppend = StrAppend & " AND MBMRenewFlag = '" & Trim(Mid(cboRenewalInd, 1, 1)) & "'"
    End If
    
    If Len(Mid(cboTakeOver, 1, 1)) Then
        StrAppend = StrAppend & " AND MBMTakeOver = '" & Trim(Mid(cboTakeOver, 1, 1)) & "'"
    End If
    
    If Len(Trim(txtAgentCode)) Then
       StrAppend = StrAppend + " And  MBMAgentCode like '%" & ChkStr(txtAgentCode) & "%' "
    End If
    
    If Len(Trim(txtPolNo)) Then
       StrAppend = StrAppend + " And  MBMPolicyNo like '%" & ChkStr(txtPolNo) & "%' "
    End If
    
    If Len(Mid(cboPayorCode, 1, 2)) Then
        StrAppend = StrAppend & " AND PAYCode = '" & Trim(Mid(cboPayorCode, 1, 2)) & "'"
    End If
           
    If Len(Trim(CboGrpCom)) Then
         GrpCompany = Replace(CboGrpCom, "'", "''")
         StrAppend = StrAppend + " AND MBMGroupCompany = '" & Trim(GrpCompany) & "'"
      End If
    
    'If Len(Trim(cboHOSName)) Then
        'strAppend = strAppend + " And HOSName = '" & Trim(cboHOSName) & "'"
    'End If
    If Len(Trim(cboHOSName)) Then
        Hospitals = Replace(cboHOSName, "'", "''")
        StrAppend = StrAppend + " AND HOSName = '" & Trim(Hospitals) & "'"
    End If
    
    If Len(Trim(cboHOSGrp)) Then
        StrAppend = StrAppend + " And HOSGRPName = '" & Trim(cboHOSGrp) & "'"
    End If
    
    If Len(Trim(txtDOCName)) Then
       StrAppend = StrAppend + " And  AttendDoctor like '%" & ChkStr(txtDOCName) & "%' "
    End If
    
    If Len(Trim(TxtPatientName)) Then
       StrAppend = StrAppend + " And  PatientName like '%" & ChkStr(TxtPatientName) & "%' "
    End If
    
    If Len(Trim(Mid(CboStayType, 1, 1))) Then
        StrAppend = StrAppend + " And CASStayType = '" & Trim(Mid(CboStayType, 1, 1)) & "'"
    End If
    
    If Len(Trim(Mid(CboCashType, 1, 1))) Then
        StrAppend = StrAppend + " And CASPayType = '" & Trim(Mid(CboCashType, 1, 1)) & "'"
    End If
    
    If Len(Trim(TxtDOA1)) > 0 And IsDate(TxtDOA1) = True Then
        StrAppend = StrAppend + " And CasDateAdmitted >= '" & Format(TxtDOA1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOA2)) > 0 And IsDate(TxtDOA2) = True Then
        StrAppend = StrAppend + " And CasDateAdmitted <= '" & Format(TxtDOA2, "mm/dd/yyyy") & "'"
    End If
       
    If (txtDOD1) <> "__/__/____" And IsDate(txtDOD1) = True _
        And (txtDOD2) <> "__/__/____" And IsDate(txtDOD2) = True Then
        Sql = ""
        Sql = " AND CasDischargeDate >= '" & Format(Trim(txtDOD1), "mm/dd/yyyy") & _
              "' And CasDischargeDate <= '" & Format(Trim(txtDOD2), "mm/dd/yyyy") & "'"
        StrAppend = StrAppend + Sql
    End If
    
    If Len(Trim(txtDOD1)) > 0 And IsDate(txtDOD1) = True Then
        StrAppend = StrAppend + " And CasDischargeDate >= '" & Format(txtDOD1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOD2)) > 0 And IsDate(txtDOD2) = True Then
        StrAppend = StrAppend + " And CasDischargeDate <= '" & Format(txtDOD2, "mm/dd/yyyy") & "'"
    End If
    
    If (TxtDOR1) <> "__/__/____" And IsDate(TxtDOR1) = True _
        And (TxtDOR2) <> "__/__/____" And IsDate(TxtDOR2) = True Then
        Sql = ""
        Sql = " AND CaseRegDate >= '" & Format(Trim(TxtDOR1), "mm/dd/yyyy") & _
              "' And CaseRegDate <= '" & Format(Trim(TxtDOR2), "mm/dd/yyyy") & "'"
        StrAppend = StrAppend + Sql
    End If
    
    If Len(Trim(TxtDOR1)) > 0 And IsDate(TxtDOR1) = True Then
        StrAppend = StrAppend + " And CaseRegDate >= '" & Format(TxtDOR1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOR2)) > 0 And IsDate(TxtDOR2) = True Then
        StrAppend = StrAppend + " And CaseRegDate <= '" & Format(TxtDOR2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtPreDiag1)) Then
        StrAppend = StrAppend + " And FirstDiag >= '" & Trim(txtPreDiag1) & "'"
    End If
    
    If Len(Trim(txtPreDiag2)) Then
        StrAppend = StrAppend + " And FirstDiag <= '" & Trim(txtPreDiag2) & "'"
    End If
    
    If Len(Trim(txtFinalDiag1)) Then
        StrAppend = StrAppend + " And FinalDiag >= '" & Trim(txtFinalDiag1) & "'"
    End If
    
    If Len(Trim(txtFinalDiag2)) Then
        StrAppend = StrAppend + " And FinalDiag <= '" & Trim(txtFinalDiag2) & "'"
    End If
    
    If Len(Trim(txtRepudCat1)) Then
        StrAppend = StrAppend + " And CASREPCode1 >= '" & Trim(txtRepudCat1) & "'"
    End If
    
    If Len(Trim(txtRepudCat2)) Then
        StrAppend = StrAppend + " And CASREPCode1 <= '" & Trim(txtRepudCat2) & "'"
    End If
    
    If Len(Trim(txtRepudCode1)) Then
        StrAppend = StrAppend + " And RepCatCode >= '" & Trim(txtRepudCode1) & "'"
    End If
    
    If Len(Trim(txtRepudCode2)) Then
        StrAppend = StrAppend + " And RepCatCode <= '" & Trim(txtRepudCode2) & "'"
    End If
    
    If Len(Trim(TxtClaimNature1)) Then
        StrAppend = StrAppend & " AND ClaimNatureCode >= '" & Trim(TxtClaimNature1) & "'"
    End If
    
    If Len(Trim(TxtClaimNature2)) Then
        StrAppend = StrAppend & " AND ClaimNatureCode <= '" & Trim(TxtClaimNature2) & "'"
    End If
    
    If Len(Trim(TxtPlan1)) Then
        StrAppend = StrAppend & " AND PLNCode >= '" & Trim(TxtPlan1) & "'"
    End If
    
    If Len(Trim(TxtPlan2)) Then
        StrAppend = StrAppend & " AND PLNCode <= '" & Trim(TxtPlan2) & "'"
    End If
    
    If Len(Trim(TxtPolicyNo1)) Then
        StrAppend = StrAppend & " AND MBMPolicyNo >= '" & Trim(TxtPolicyNo1) & "'"
    End If
    
    If Len(Trim(TxtPolicyNo2)) Then
        StrAppend = StrAppend & " AND MBMPolicyNo <= '" & Trim(TxtPolicyNo2) & "'"
    End If
    
    If Len(Trim(TxtBordxNo1)) Then
        StrAppend = StrAppend & " AND BordxRefNo >= '" & Trim(TxtBordxNo1) & "'"
    End If
    
    If Len(Trim(TxtBordxNo2)) Then
        StrAppend = StrAppend & " AND BordxRefNo <= '" & Trim(TxtBordxNo2) & "'"
    End If
    
    If Option19.value = True Then
    
    Else
    
        If Len(Trim(txtDOW1)) > 0 And IsDate(txtDOW1) = True Then
            StrAppend = StrAppend + " And ClaimOpenDate >= '" & Format(txtDOW1, "mm/dd/yyyy") & "'"
        End If
    
        If Len(Trim(txtDOW2)) > 0 And IsDate(txtDOW2) = True Then
            StrAppend = StrAppend + " And ClaimOpenDate <= '" & Format(txtDOW2, "mm/dd/yyyy") & "'"
        End If
    
        If ChkDOW.value = 1 Then
            StrAppend = StrAppend & " AND (ClaimOpenDate Is null OR ClaimOpenDate = '')"
        End If
        
    End If
    
    
    If Len(Trim(TxtPayEffDate1)) > 0 And IsDate(TxtPayEffDate1) = True Then
        StrAppend = StrAppend + " And MBMPayorEffDate >= '" & Format(TxtPayEffDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtPayEffDate2)) > 0 And IsDate(TxtPayEffDate2) = True Then
        StrAppend = StrAppend + " And MBMPayorEffDate <= '" & Format(TxtPayEffDate2, "mm/dd/yyyy") & "'"
    End If
       
    If (TxtPayEffDate1) <> "__/__/____" And IsDate(TxtPayEffDate1) = True _
        And (TxtPayEffDate2) <> "__/__/____" And IsDate(TxtPayEffDate2) = True Then
        Sql = ""
        Sql = " AND MBMPayorEffDate >= '" & Format(Trim(TxtPayEffDate1), "mm/dd/yyyy") & _
              "' And MBMPayorEffDate <= '" & Format(Trim(TxtPayEffDate2), "mm/dd/yyyy") & "'"
        StrAppend = StrAppend + Sql
    End If
    
    If (TxtPayExpDate1) <> "__/__/____" And IsDate(TxtPayExpDate1) = True _
        And (TxtPayExpDate2) <> "__/__/____" And IsDate(TxtPayExpDate2) = True Then
        Sql = ""
        Sql = " AND MBMPayorExpDate >= '" & Format(Trim(TxtPayExpDate1), "mm/dd/yyyy") & _
              "' And MBMPayorExpDate <= '" & Format(Trim(TxtPayExpDate2), "mm/dd/yyyy") & "'"
        StrAppend = StrAppend + Sql
    End If
    
    If Len(Trim(TxtPayExpDate1)) And IsDate(TxtPayExpDate1) = True Then
        StrAppend = StrAppend & " AND MBMPayorExpDate >= '" & Format(TxtPayExpDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtPayExpDate2)) And IsDate(TxtPayExpDate2) = True Then
        StrAppend = StrAppend & " AND MBMPayorExpDate <= '" & Format(TxtPayExpDate2, "mm/dd/yyyy") & "'"
    End If
    
    If (txtDateJoin1) <> "__/__/____" And IsDate(txtDateJoin1) = True _
        And (txtDateJoin2) <> "__/__/____" And IsDate(txtDateJoin2) = True Then
        Sql = ""
        Sql = " AND MBMFirstJoinedDate >= '" & Format(Trim(txtExpDate1), "mm/dd/yyyy") & _
              "' And MBMFirstJoinedDate <= '" & Format(Trim(txtExpDate2), "mm/dd/yyyy") & "'"
        StrAppend = StrAppend + Sql
    End If
    
    If Len(Trim(txtDateJoin1)) And IsDate(txtDateJoin1) = True Then
        StrAppend = StrAppend & " AND MBMFirstJoinedDate >= '" & Format(txtDateJoin1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDateJoin2)) And IsDate(txtDateJoin2) = True Then
        StrAppend = StrAppend & " AND MBMFirstJoinedDate <= '" & Format(txtDateJoin2, "mm/dd/yyyy") & "'"
    End If
    
    If (txtDOB1) <> "__/__/____" And IsDate(txtDOB1) = True _
        And (txtDOB2) <> "__/__/____" And IsDate(txtDOB2) = True Then
        Sql = ""
        Sql = " AND DOB >= '" & Format(Trim(txtDOB1), "mm/dd/yyyy") & _
              "' And DOB <= '" & Format(Trim(txtDOB2), "mm/dd/yyyy") & "'"
        StrAppend = StrAppend + Sql
    End If
    
    If Len(Trim(txtDOB1)) And IsDate(txtDOB1) = True Then
        StrAppend = StrAppend & " AND DOB >= '" & Format(txtDOB1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOB2)) And IsDate(txtDOB2) = True Then
        StrAppend = StrAppend & " AND DOB <= '" & Format(txtDOB2, "mm/dd/yyyy") & "'"
    End If
    
    If (TxtBordxDate1) <> "__/__/____" And IsDate(TxtBordxDate1) = True _
        And (TxtBordxDate2) <> "__/__/____" And IsDate(TxtBordxDate2) = True Then
        Sql = ""
        Sql = " AND BordxDate >= '" & Format(Trim(TxtBordxDate1), "mm/dd/yyyy") & _
              "' And BordxDate <= '" & Format(Trim(TxtBordxDate2), "mm/dd/yyyy") & "'"
        StrAppend = StrAppend + Sql
    End If
    
    If Len(Trim(TxtBordxDate1)) And IsDate(TxtBordxDate1) = True Then
        StrAppend = StrAppend & " AND BordxDate >= '" & Format(TxtBordxDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtBordxDate2)) And IsDate(TxtBordxDate2) = True Then
        StrAppend = StrAppend & " AND BordxDate <= '" & Format(TxtBordxDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtPOLYear1)) Then
        StrAppend = StrAppend & " AND MBMPolicyYear >= '" & Trim(txtPOLYear1) & "'"
    End If
    
    If Len(Trim(txtPOLYear2)) Then
        StrAppend = StrAppend & " AND MBMPolicyYear <= '" & Trim(txtPOLYear2) & "'"
    End If
    
    If Len(Trim(TxtCaseNo1)) Then
       StrAppend = StrAppend + " AND CASNumber >= '" & RemStr(Trim(TxtCaseNo1)) & "'"
    End If
    
    If Len(Trim(TxtCaseNo2)) Then
       StrAppend = StrAppend + " AND CASNumber <= '" & RemStr(Trim(TxtCaseNo2)) & "'"
    End If
    
    If Len(Trim(TxtVersion1)) Then
       StrAppend = StrAppend + " AND ClaimVersion >= '" & RemStr(Trim(TxtVersion1)) & "'"
    End If
    
    If Len(Trim(TxtVersion2)) Then
       StrAppend = StrAppend + " AND ClaimVersion <= '" & RemStr(Trim(TxtVersion2)) & "'"
    End If
    
    If Len(Trim(TxtMBMNo1)) Then
       StrAppend = StrAppend + " AND MBMNumber >= '" & RemStr(Trim(TxtMBMNo1)) & "'"
    End If
    
    If Len(Trim(TxtMBMNo2)) Then
       StrAppend = StrAppend + " AND MBMNumber <= '" & RemStr(Trim(TxtMBMNo2)) & "'"
    End If
    
    If ChkHLT.value = 1 Then
        StrAppend = StrAppend & " AND (HLTCode Is null OR HLTCode = '')"
    End If
    
    If ChkINS.value = 1 Then
        StrAppend = StrAppend & " AND (INSCode Is null OR INSCode = '')"
    End If
    
    'If ChkMemType.value = 1 Then
        'strAppend = strAppend & " AND (MemberType Is null OR MemberType = '')"
    'End If
    
    If ChkAdultChild.value = 1 Then
        StrAppend = StrAppend & " AND (AdultChild Is null OR AdultChild = '')"
    End If
    
    If ChkMarital.value = 1 Then
        StrAppend = StrAppend & " AND (MBMMaritalStatus Is null OR MBMMaritalStatus = '')"
    End If
    
    If ChkRace.value = 1 Then
        StrAppend = StrAppend & " AND (RACCode Is null OR RACCode = '')"
    End If
    
    If ChkSex.value = 1 Then
        StrAppend = StrAppend & " AND (Sex Is null OR Sex = '')"
    End If
    
    If ChkRelation.value = 1 Then
        StrAppend = StrAppend & " AND (Relation Is null OR Relation = '')"
    End If
    
    If ChkDataStatus.value = 1 Then
        StrAppend = StrAppend & " AND (MBMDataStatus Is null OR MBMDataStatus = '')"
    End If
    
    If ChkClaimability.value = 1 Then
        StrAppend = StrAppend & " AND (CASClaimability Is null OR CASClaimability = '')"
    End If
    
    If ChkPolStatus.value = 1 Then
        StrAppend = StrAppend & " AND (MBMStatus Is null OR MBMStatus = '')"
    End If
    
    If ChkCaseStatus.value = 1 Then
        StrAppend = StrAppend & " AND (CASStatus Is null OR CASStatus = '')"
    End If
    
    
    If ChkExpired.value = 1 Then
        StrAppend = StrAppend & " AND (MBMExpiredStatus Is null OR MBMExpiredStatus = '')"
    End If
    
    If ChkPreMode.value = 1 Then
        StrAppend = StrAppend & " AND (MBMInstallment Is null OR MBMInstallment = '')"
    End If
    
    If ChkState.value = 1 Then
        StrAppend = StrAppend & " AND (STACode Is null OR STACode = '')"
    End If
    
    If ChkRenewal.value = 1 Then
        StrAppend = StrAppend & " AND (MBMRenewFlag Is null OR MBMRenewFlag = '')"
    End If
    
    If ChkTakeOver.value = 1 Then
        StrAppend = StrAppend & " AND (MBMTakeOver Is null OR MBMTakeOver = '')"
    End If
    
    If ChkAgentCode.value = 1 Then
        StrAppend = StrAppend & " AND (MBMAgentCode Is null OR MBMAgentCode = '')"
    End If
    
    If ChkPolNo.value = 1 Then
        StrAppend = StrAppend & " AND (MBMPolicyNo Is null OR MBMPolicyNo = '')"
    End If
    
    If ChkPayor.value = 1 Then
        StrAppend = StrAppend & " AND (PAYCode Is null OR PAYCode = '')"
    End If
    
    If ChkGrpCom.value = 1 Then
        StrAppend = StrAppend & " AND (MBMGroupCompany Is null OR MBMGroupCompany = '')"
    End If
    
    If ChkHOS.value = 1 Then
        StrAppend = StrAppend & " AND (HOSName Is null OR HOSName = '')"
    End If
    
    If ChkHOSGrp.value = 1 Then
        StrAppend = StrAppend & " AND (HOSGRPName Is null OR HOSGRPName = '')"
    End If
    
    If ChkDoctor.value = 1 Then
        StrAppend = StrAppend & " AND (AttendDoctor Is null OR AttendDoctor = '')"
    End If
            
    If ChkDOA.value = 1 Then
        StrAppend = StrAppend & " AND (CASDateAdmitted Is null OR CASDateAdmitted = '')"
    End If
    
    If ChkDOD.value = 1 Then
        StrAppend = StrAppend & " AND (CASDischargeDate Is null OR CASDischargeDate = '')"
    End If
    
    If chkDOR.value = 1 Then
        StrAppend = StrAppend & " AND (CaseRegDate Is null OR CaseRegDate = '')"
    End If
    
    If ChkFirst.value = 1 Then
        StrAppend = StrAppend & " AND (FirstDiag Is null OR FirstDiag = '')"
    End If
    
    If ChkFinal.value = 1 Then
        StrAppend = StrAppend & " AND (FinalDiag Is null OR FinalDiag = '')"
    End If
    
    If ChkRep.value = 1 Then
        StrAppend = StrAppend & " AND (CASRepCode1 Is null OR CASRepCode1 = '')"
    End If
    
    If ChkRepCat.value = 1 Then
        StrAppend = StrAppend & " AND (RepCatCode Is null OR RepCatCode = '')"
    End If
    
    If ChkPlan.value = 1 Then
        StrAppend = StrAppend & " AND (PLNCode Is null OR PLNCode = '')"
    End If
    
    If ChkBordx.value = 1 Then
        StrAppend = StrAppend & " AND (BordxRefNo Is null OR BordxRefNo = '')"
    End If
    
    If ChkEffDate.value = 1 Then
        StrAppend = StrAppend & " AND (MBMPayorEffDate Is null OR MBMPayorEffDate = '')"
    End If
    
    If chkExpDate.value = 1 Then
        StrAppend = StrAppend & " AND (MBMPayorExpDate Is null OR MBMPayorExpDate = '')"
    End If
    
    If ChkJoinDate.value = 1 Then
        StrAppend = StrAppend & " AND (MBMFirstJoinedDate Is null OR MBMFirstJoinedDate = '')"
    End If
    
    If ChkDOB.value = 1 Then
        StrAppend = StrAppend & " AND (DOB Is null OR DOB = '')"
    End If
    
    If ChkPolYear.value = 1 Then
        StrAppend = StrAppend & " AND (MBMPolicyYear Is null OR MBMPolicyYear = '')"
    End If
    
    If Option20.value = True And Option33.value = True Then
        If Len(Trim(TxtYear)) Then
            StrAppend = StrAppend + " AND DOAYear = '" & Trim(TxtYear) & "'"
        End If
    ElseIf Option21.value = True And Option33.value = True Then
        If Len(Trim(TxtYear)) Then
            StrAppend = StrAppend + " AND DODYear = '" & Trim(TxtYear) & "'"
        End If
    End If
    
    If ChkClaimNature.value = 1 Then
        StrAppend = StrAppend & " AND (ClaimNatureCode Is null OR ClaimNatureCode = '')"
    End If
    
    If ChkBordxDate.value = 1 Then
        StrAppend = StrAppend & " AND (BordxDate Is null OR BordxDate = '')"
    End If
    
    If Len(Trim(TxtNoofDay1)) Then
        StrAppend = StrAppend + " AND NoofDay >= '" & Trim(TxtNoofDay1) & "'"
      End If
      If Len(Trim(TxtNoofDay2)) Then
        StrAppend = StrAppend + " AND NoofDay <= '" & Trim(TxtNoofDay2) & "'"
      End If
      'If ChkNoofDay.value = 1 Then
        'strAppend = strAppend + " AND (NoofDay Is null OR NoofDay = '')"
      'End If
      
    If Len(txtAppAmt1) Then
        StrAppend = StrAppend + " AND CLMTotalApproved >= " & Trim(txtAppAmt1) & ""
    End If
        
    If Len(txtAppAmt2) Then
        StrAppend = StrAppend + " AND CLMTotalApproved <= " & Trim(txtAppAmt2) & ""
    End If
    
    If Len(txtActualAmt1) Then
        StrAppend = StrAppend + " AND CLMTotalActual >= " & Trim(txtActualAmt1) & ""
    End If
        
    If Len(txtActualAmt2) Then
        StrAppend = StrAppend + " AND CLMTotalActual <= " & Trim(txtActualAmt2) & ""
    End If
      
    
    If Option23.value = True Then
        Call IllnessListing(StrAppend)
    ElseIf Option1.value = True And Option35.value = True Then
        Call ResultListing(StrAppend)
    ElseIf Option1.value = True And Option34.value = True Then
        Call ResultListing2(StrAppend)
    ElseIf Option2.value = True And Option35.value = True Then
        Call INSListing(StrAppend)
    ElseIf Option2.value = True And Option34.value = True Then
        Call PlanInsuredListing(StrAppend)
    ElseIf Option2.value = True And Option33.value = True Then
        Call AgeInsuredListing(StrAppend)
    ElseIf Option2.value = True And Option29.value = True Then
        Call SexInsuredListing(StrAppend)
    ElseIf Option2.value = True And Option37.value = True Then
        Call RaceInsuredListing(StrAppend)
    ElseIf Option3.value = True And Option31.value = True Then
        Call PlanPremiumListing(StrAppend)
    ElseIf Option3.value = True And Option35.value = True Then
        Call PlanListing(StrAppend)
    ElseIf Option3.value = True And Option34.value = True Then
        Call PlanTopHosListing(StrAppend)
    ElseIf Option3.value = True And Option33.value = True Then
        Call PlanTopDiagListing(StrAppend)
    ElseIf Option3.value = True And Option29.value = True Then
        Call PlanAgebandListing(StrAppend)
    ElseIf Option3.value = True And Option37.value = True Then
        Call PlanInsuredListing(StrAppend)
    ElseIf Option3.value = True And Option30.value = True Then
        Call PlanInsuredListing2(StrAppend)
    ElseIf Option5.value = True And Option35.value = True Then
        Call AgeGroupListing(StrAppend)
    ElseIf Option5.value = True And Option34.value = True Then
        Call AgeGenderListing(StrAppend)
    ElseIf Option5.value = True And Option33.value = True Then
        Call AgeInsuredListing(StrAppend)
    ElseIf Option5.value = True And Option29.value = True Then
        Call AgeRaceListing(StrAppend)
    ElseIf Option5.value = True And Option37.value = True Then
        Call AgePlanListing(StrAppend)
    ElseIf Option5.value = True And Option30.value = True Then
        Call AgeIllnessListing(StrAppend)
    ElseIf Option6.value = True And Option35.value = True Then
        Call SexListing(StrAppend)
    ElseIf Option6.value = True And Option34.value = True Then
        Call SexAgebandListing(StrAppend)
    ElseIf Option6.value = True And Option33.value = True Then
        Call SexInsuredListing(StrAppend)
    ElseIf Option6.value = True And Option29.value = True Then
        Call SexRaceListing(StrAppend)
    ElseIf Option6.value = True And Option37.value = True Then
        Call SexPlanListing(StrAppend)
    ElseIf Option6.value = True And Option30.value = True Then
        Call SexIllnessListing(StrAppend)
    ElseIf Option7.value = True And Option35.value = True Then
        Call RaceListing(StrAppend)
    ElseIf Option7.value = True And Option34.value = True Then
        Call RaceAgebandListing(StrAppend)
    ElseIf Option7.value = True And Option33.value = True Then
        Call RaceInsuredListing(StrAppend)
    ElseIf Option7.value = True And Option29.value = True Then
        Call RaceSexListing(StrAppend)
    ElseIf Option7.value = True And Option37.value = True Then
        Call RacePlanListing(StrAppend)
    ElseIf Option7.value = True And Option30.value = True Then
        Call RaceIllnessListing(StrAppend)
    ElseIf Option8.value = True Then
        Call StateListing(StrAppend)
    ElseIf Option9.value = True Then
        Call PayorListing(StrAppend)
    ElseIf Option10.value = True And Option29.value = True Or Option10.value = True And Option37.value = True Then
        Call HOSTop20(StrAppend)
    ElseIf Option10.value = True And Option34.value = True Then
        Call HOSSumListing(StrAppend)
    ElseIf Option10.value = True And Option33.value = True Then
        Call HOSSumListing2(StrAppend)
    ElseIf Option10.value = True And Option35.value = True Then
        Call HOSDetailListing(StrAppend)
    ElseIf Option10.value = True And Option30.value = True Then
        Call HOSByDiag(StrAppend)
    ElseIf Option11.value = True And Option35.value = True Then
        Call HOSGrpSumListing(StrAppend)
    ElseIf Option12.value = True Then
        Call DoctorListing(StrAppend)
    ElseIf Option13.value = True Then
        Call ICSummary(StrAppend)
    ElseIf Option15.value = True And Option35.value = True Then
        Call FinalDiagListing(StrAppend)
    ElseIf Option15.value = True And Option34.value = True Then
        Call FinalDiagListing(StrAppend)
    ElseIf Option15.value = True And Option33.value = True Then
        Call FinalDiagListing(StrAppend)
    ElseIf Option15.value = True And Option29.value = True Then
        Call FinalDiagbyHos(StrAppend)
    ElseIf Option16.value = True And Option35.value = True Then
        Call DurationSum(StrAppend)
    ElseIf Option16.value = True And Option34.value = True Then
        Call DurationDetail(StrAppend)
    'ElseIf Option17.value = True Then
        'Call ClmCheckTotalListing(strAppend)
    ElseIf Option18.value = True And Option35.value = True Then
        Call RepcodeSummary(StrAppend)
    ElseIf Option18.value = True And Option34.value = True Then
        Call RepCatcodeSummary(StrAppend)
    ElseIf Option18.value = True And Option37.value = True Then
        Call ListingSummary(StrAppend)
    ElseIf Option18.value = True And Option33.value = True Then
        Call FirstDiagSummary(StrAppend)
    ElseIf Option18.value = True And Option29.value = True Then
        Call FinalDiagSummary(StrAppend)
    ElseIf Option18.value = True And Option30.value = True Then
        Call HospitalSummary(StrAppend)
    ElseIf Option18.value = True And Option31.value = True Then
        Call PayorSummary(StrAppend)
    ElseIf Option18.value = True And Option32.value = True Then
        Call GenderSummary(StrAppend)
    ElseIf Option18.value = True And Option38.value = True Then
        Call AgeBandSummary(StrAppend)
    ElseIf Option19.value = True Then
        Call BordxNoListing(StrAppend)
    ElseIf Option20.value = True And Option35.value = True Then
        If TxtDOA2 <> TxtDOA1 Then
            MsgBox "Must key in same Date of Admitted", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf TxtDOA1 = "__/__/____" Or TxtDOA2 = "__/__/____" Then
            MsgBox "Date of Admitted can't be empty", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            Call DailyAdmissionListing(StrAppend)
        End If
   ElseIf Option20.value = True And Option34.value = True Then
        If Mid(TxtDOA1, 4, 2) <> Mid(TxtDOA2, 4, 2) Or Mid(TxtDOA1, 7, 4) <> Mid(TxtDOA2, 7, 4) Then
            MsgBox "Must key in same Month", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf TxtDOA1 = "__/__/____" Or TxtDOA2 = "__/__/____" Then
            MsgBox "Date of Admitted can't be empty", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            Call MonthyReportListing(StrAppend)
        End If
   ElseIf Option20.value = True And Option33.value = True Then
        If TxtYear = "" Then
            MsgBox "Please enter Year", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            Call YearListing(StrAppend)
        End If
   ElseIf Option21.value = True And Option35.value = True Then
        If txtDOD2 <> txtDOD1 Then
            MsgBox "Must key in same Date of Discharge", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf txtDOD1 = "__/__/____" Or txtDOD2 = "__/__/____" Then
            MsgBox "Date of Discharge can't be empty", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            Call DailyDischargeListing(StrAppend)
        End If
   ElseIf Option21.value = True And Option34.value = True Then
        If Mid(txtDOD1, 4, 2) <> Mid(txtDOD2, 4, 2) Or Mid(txtDOD1, 7, 4) <> Mid(txtDOD2, 7, 4) Then
            MsgBox "Must key in same Month", vbOKOnly + vbCritical, DialogBoxTitle
        ElseIf txtDOD1 = "__/__/____" Or txtDOD2 = "__/__/____" Then
            MsgBox "Date of Discharge can't be empty", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            Call MonthyReportListing(StrAppend)
        End If
   ElseIf Option21.value = True And Option33.value = True Then
        If TxtYear = "" Then
            MsgBox "Please enter Year", vbOKOnly + vbCritical, DialogBoxTitle
        Else
            Call YearListing(StrAppend)
        End If
   'ElseIf Option22.value = True And Option35.value = True Then
        'Call ClaimTypeListing(strAppend)
   'ElseIf Option22.value = True And Option34.value = True Then
        'Call ClaimTypeFollowUpListing(strAppend)
   ElseIf Option36.value = True And Option35.value = True Then
        Call AgentListing(StrAppend)
   ElseIf Option36.value = True And Option34.value = True Then
        Call AgentSumListing(StrAppend)
   ElseIf Option36.value = True And Option33.value = True Then
        Call AgentSumListing(StrAppend)
   ElseIf Option36.value = True And Option29.value = True Then
        Call AgentDoctorListing(StrAppend)
   End If
End Function

Private Sub HOSTop20(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
       
    lblCurrent = 0
    ListView2.listitems.Clear
        
    Sql = " SELECT TOP 20 PAYCode, HOSName, COUNT(HOSName) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved" & _
          " From VIEW_CLMReport Where  0 = 0"
    
    If Option33.value = True Then
        sql2 = " GROUP BY PAYCode, HOSName ORDER BY Total DESC "
    ElseIf Option29.value = True Then
        sql2 = " GROUP BY PAYCode, HOSName ORDER BY TotalActual DESC "
    End If
    
    Sql = Sql + StrAppend + sql2
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
    
            With rst2
                
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(1) = Trim(rst2!HOSName)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                            
                If Len(rst2!Total) <> 0 Then
                    Record.SubItems(5) = Format(Round(rst2!TotalActual / rst2!Total, 0), "#,##0.00")
                    Record.SubItems(6) = Format(Round(rst2!TotalApproved / rst2!Total, 0), "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub HOSSumListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    
    lblCurrent = 0
    ListView2.listitems.Clear
        
    Sql = " SELECT PAYCode, HOSName, COUNT(HOSName) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " GROUP BY PAYCode, HOSName "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
        
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(1) = Trim(rst2!HOSName)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                                
                If Len(rst2!Total) <> 0 Then
                    Record.SubItems(5) = Format(Round(rst2!TotalActual / rst2!Total, 0), "#,##0.00")
                    Record.SubItems(6) = Format(Round(rst2!TotalApproved / rst2!Total, 0), "#,##0.00")
                End If
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub HOSSumListing2(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    
    lblCurrent = 0
    ListView2.listitems.Clear
        
    Sql = " SELECT HOSName, COUNT(HOSName) AS Total, " & _
          " SUM(CLMTotalActual) AS TotalActual, SUM(CLMTotalApproved) AS TotalApproved " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " GROUP BY HOSName "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
        
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                
            With rst2
                If Len(rst2!HOSName) Then
                    Set Record = ListView2.listitems.Add(, , Trim(rst2!HOSName))
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                               
                                
                If Len(rst2!Total) Then
                    Record.SubItems(1) = rst2!Total
                End If
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(2) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(3) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                                
                If Len(rst2!Total) <> 0 Then
                    Record.SubItems(4) = Format(Round(rst2!TotalActual / rst2!Total, 0), "#,##0.00")
                    Record.SubItems(5) = Format(Round(rst2!TotalApproved / rst2!Total, 0), "#,##0.00")
                End If
                           
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub HOSByDiag(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim rst3 As New ADODB.Recordset
    
    DiagCode = ""
    lblCurrent = 0
    ListView2.listitems.Clear
    
    Sql = " SELECT COUNT(FinalDiag) AS Total, FinalDiag, Sum(CLMTotalApproved) As TotalApproved, " & _
          " Sum(CLMTotalActual) As TotalActual, HOSName, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 " & StrAppend & _
          " GROUP BY FinalDiag, HOSName ORDER BY FinalDiag "
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
       
        Do Until rst2.EOF
    
            With rst2
                               
                If Len(rst2!HOSName) Then
                    Set Record = ListView2.listitems.Add(, , Trim(rst2!HOSName))
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                If Len(rst2!FinalDiag) Then
                    Record.SubItems(1) = Trim(rst2!FinalDiag)
                    DiagCode = rst2!FinalDiag
                Else
                    Record.SubItems(1) = "Empty"
                    DiagCode = ""
                End If
                
                sql2 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(DiagCode) & "'"
                rst3.Open sql2, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                If rst3.recordCount = 0 Then
                    rst3.Close
                    Set rst3 = Nothing
                    Record.SubItems(2) = ""
                Else
                    Description = Trim(rst3!ICDDescription)
                    Record.SubItems(2) = Trim(Description)
                    rst3.Close
                    Set rst3 = Nothing
                End If
                
                If Len(rst2!Total) Then
                    Record.SubItems(3) = rst2!Total
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(4) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(4) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                   Record.SubItems(5) = Format(rst2!TotalActual, "##,#0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(6) = Format(rst2!TotalApproved, "##,#0.00")
                End If
            
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub HOSDetailListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    ListView2.listitems.Clear
    Current = 0
    Total = 0
            
    Sql = " SELECT PAYCode, HOSName,CASNumber, MBMNumber, MBMPatientCovID, " & _
          " PatientName, CLMTotalActual, CLMTotalApproved, " & _
          " CLMRecoveryPaidbyM, CASDateAdmitted, CASDischargeDate, " & _
          " CASStatus, CASClaimability, FinalDiag From VIEW_CLMReport Where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    'If rst2.recordCount = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
    
    Do
    
            Do Until rst2.EOF
            
            Total = rst2.recordCount
            lblTotalRecord = Total
                
            With rst2
                
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(1) = Trim(rst2!HOSName)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                                
                If Len(rst2!CASNumber) Then
                    Record.SubItems(2) = rst2!CASNumber
                End If
                                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(3) = rst2!MBMNumber
                End If
                
                If Len(rst2!MBMPatientCovID) Then
                    Record.SubItems(4) = rst2!MBMPatientCovID
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(5) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!CLMTotalActual) Then
                    Record.SubItems(6) = Format(rst2!CLMTotalActual, "#,##0.00")
                End If
                
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(7) = Format(rst2!CLMTotalApproved, "#,##0.00")
                End If
                
                If Len(rst2!CLMRecoveryPaidByM) Then
                    Record.SubItems(8) = Format(rst2!CLMRecoveryPaidByM, "#,##0.00")
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(9) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(10) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASStatus) Then
                   Record.SubItems(11) = rst2!CASStatus
                End If
                
                If Len(rst2!CASClaimability) Then
                    Record.SubItems(12) = rst2!CASClaimability
                End If
                                
                If Len(rst2!FinalDiag) Then
                    Record.SubItems(13) = rst2!FinalDiag
                End If
                
                Current = Current + 1
                lblCurrent = Current
                           
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    'End If
    intExit = 0
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub DoctorListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim sql3 As String
    
    ListView2.listitems.Clear
    Current = 0
    Total = 0
            
    Sql = " SELECT CASNumber, AttendDoctor, HOSName, " & _
          " PatientName, MBMAgentCode, SUM(INVAmount) AS INVAmount, " & _
          " CASDateAdmitted, CASDischargeDate, " & _
          " FinalDiag From VIEW_CLMReport2 Where 0 = 0 "
          
    sql2 = " Group by CASNumber, AttendDoctor, HOSName, PatientName, MBMAgentCode, " & _
           "CASDateAdmitted, CASDischargeDate, FinalDiag "
 
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    End If
 
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
            With rst2
   
                Set Record = ListView2.listitems.Add(, , !CASNumber)
   
                If Len(rst2!AttendDoctor) Then
                    Record.SubItems(1) = Trim(!AttendDoctor)
                End If
                                
                If Len(rst2!PatientName) Then
                    Record.SubItems(2) = Trim(!PatientName)
                End If
   
                If Len(rst2!MBMAgentCode) Then
                    Record.SubItems(3) = Trim(!MBMAgentCode)
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(4) = Trim(!HOSName)
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(5) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
   
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(6) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
   
                If Len(rst2!INVAmount) Then
                    Record.SubItems(7) = Format(rst2!INVAmount, "#,##0.00")
                End If
   
                If Len(rst2!FinalDiag) Then
                   Record.SubItems(8) = rst2!FinalDiag
                   FinalICDCode = Record.SubItems(8)
                   sql3 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(FinalICDCode) & "'"
                   rst3.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
   
                   If rst3.recordCount = 0 Then
                        Record.SubItems(9) = ""
                        rst3.Close
                        Set rst3 = Nothing
                   Else
                        Record.SubItems(9) = Trim(rst3!ICDDescription)
                        rst3.Close
                        Set rst3 = Nothing
                   End If
                End If
   
                Current = Current + 1
                lblCurrent = Current
    
            End With
            rst2.MoveNext
        Loop
    
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
   
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ClmCheckTotalListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim CLM As New ADODB.Recordset
Dim Pay2MBM As Double
Dim strsql As String
    
    ListView2.listitems.Clear
    Current = 0
    Total = 0
    TotalChkAmt = 0
            
    Sql = " SELECT CASNumber, ClaimVersion, DebitCreditDate, CLMPayableByMedix2H, CLMTotalApproved, " & _
          " DebitCreditRefNo, DebitCreditHRefNo, CLMPayableByMedix2M, DebitCreditHAmt " & _
          " From VIEW_CLMReport Where INVPayTo = 'H' And ((Substring(DebitCreditRefNo,1,1)= 'D' or Substring(DebitCreditRefNo,1,1) = 'C') or (Substring(DebitCreditHRefNo,2,1)= 'D' or Substring(DebitCreditHRefNo,2,1) = 'C'))"
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            'sql = ""
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    'If rst2.recordCount = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
    
    Do
    
            Do Until rst2.EOF
            
            'Total = rst2.recordCount
            'lblTotalRecord = Total
                
            With rst2
                
                If Len(rst2!CASNumber) And Len(rst2!claimversion) Then
                    Set Record = ListView2.listitems.Add(, , rst2!CASNumber & "-" & rst2!claimversion)
                End If
                                
                If Len(rst2!DebitCreditDate) Then
                    Record.SubItems(1) = Format(rst2!DebitCreditDate, "dd/mm/yyyy")
                End If
                                
                If Len(rst2!DebitCreditRefNo) Then
                    Record.SubItems(2) = Trim(rst2!DebitCreditRefNo)
                End If
                
                If Len(rst2!DebitCreditHRefNo) Then
                    If Mid(rst2!DebitCreditHRefNo, 1, 2) = "HD" Then
                        Record.SubItems(3) = Trim(rst2!DebitCreditHRefNo)
                        If Len(rst2!DebitCreditHAmt) Then
                            Record.SubItems(6) = Format(rst2!DebitCreditHAmt, "#,##0.00")
                            Record.SubItems(7) = "0.00"
                        End If
                    ElseIf Mid(rst2!DebitCreditHRefNo, 1, 2) = "HC" Then
                        Record.SubItems(4) = Trim(rst2!DebitCreditHRefNo)
                        If Len(rst2!DebitCreditHAmt) Then
                            Record.SubItems(7) = Format(rst2!DebitCreditHAmt, "#,##0.00")
                            Record.SubItems(6) = "0.00"
                        End If
                    End If
                Else
                    Record.SubItems(6) = "0.00"
                    Record.SubItems(7) = "0.00"
                End If
                

                strsql = " SELECT CLMPayableByMedix2H, CLMTotalApproved, CLMPayableByMedix2M, CLMReimburse " & _
                " From Claim Where CASNumber = '" & rst2!CASNumber & "' And claimversion = '" & rst2!claimversion & "'"
                
                CLM.Open strsql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
                If CLM.recordCount Then
                    If Len(CLM!CLMPayableByMedix2H) Then
                        Record.SubItems(5) = Format(CLM!CLMPayableByMedix2M, "#,##0.00")
                    Else
                        Record.SubItems(5) = "0.00"
                    End If
                    
                If Len(CLM!CLMReimburse) Then
                    Record.SubItems(8) = Format(CLM!CLMReimburse, "#,##0.00")
                Else
                    Record.SubItems(8) = "0.00"
                End If
    
                CLM.Close
                Set CLM = Nothing

                End If
                TotalChkAmt = CDbl(Record.SubItems(5)) - CDbl(Record.SubItems(6)) - CDbl(Record.SubItems(7)) - CDbl(Record.SubItems(8))
                'LblMCO = Format(TotalMCO, "#,##0.00")
                Record.SubItems(9) = Format(TotalChkAmt, "#,##0.00")
                
                Current = Current + 1
                lblCurrent = Current
                
                'End If
          
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    'End If
    intExit = 0
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ResultListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim sql3 As String
Dim FirstICDCode As String
Dim FinalICDCode As String
Dim FirstDescription As String
Dim FinalDescription As String

        
    ListView2.listitems.Clear
    Current = 0
    Total = 0
          
    Sql = " SELECT PAYCode, CASNumber, claimversion, MBMNumber, MBMPolicyNo, " & _
          " MBMPatientCovID, MBMName, MBMDOB, DOB, PatientName, MBMSex, AgeCode, Sex, Ageband, " & _
          " InsCode, PLNCode, HOSName, FirstDiag, FinalDiag, " & _
          " CLMTotalActual, CLMTotalApproved, CASReserveAmount, " & _
          " CLMRecoveryPaidbyM, BordxRefNo, BordxDate, ClaimOpenDate, " & _
          " CASClaimability, CASStatus, MBMStatus, CASDateAdmitted, " & _
          " CASDischargeDate, MBMPayorEffDate, MBMPayorExpDate, " & _
          " NoofDay, CaseRegDate, MBMIcBcPp, ICNo From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Order By CASNumber, ClaimVersion "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend '+ sql2
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    'If rst2.recordCount = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
    
    Do
              
       Do Until rst2.EOF
       
       Total = rst2.recordCount
       lblTotalRecord = Total
        
        FirstICDCode = ""
        FinalICDCode = ""
        FirstDescription = ""
        FinalDescription = ""
        'Text1 = ""
        
        'If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
            'Text1 = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
        'End If
        
        'If Text1 = "0" Then
            'Text1 = Text1 + 1
        'End If
        
            With rst2
                    
                    If Len(!PAYCode) Then
                        Set Record = ListView2.listitems.Add(, , !PAYCode)
                    Else
                        Set Record = ListView2.listitems.Add(, , "Empty")
                    End If
                    
                    'Set Record = ListView2.listitems.Add(, , rst2!CASNumber & "-" & rst2!claimversion)
                    
                    If Len(!CASNumber & "-" & !claimversion) Then
                       Record.SubItems(1) = (!CASNumber & "-" & !claimversion)
                    End If
                    
                    If Len(!BordxRefNo) Then
                       Record.SubItems(2) = !BordxRefNo
                    End If
                    
                    If Len(!BordxDate) Then
                       Record.SubItems(3) = Format(!BordxDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMNumber) Then
                       Record.SubItems(4) = !MBMNumber
                    End If
                    
                    If Len(!MBMPolicyNo) Then
                       Record.SubItems(5) = Trim(!MBMPolicyNo)
                    End If
                                    
                    If Len(!MBMPatientCovID) Then
                       Record.SubItems(6) = !MBMPatientCovID
                    End If
                    
                    If Len(!MBMName) Then
                       Record.SubItems(7) = Trim(!MBMName)
                    End If
                    
                    If Len(!PatientName) Then
                       Record.SubItems(8) = Trim(!PatientName)
                    End If
                    
                    If Len(!MBMIcBcPp) Then
                       Record.SubItems(9) = !MBMIcBcPp
                    End If
                    
                    If Len(!MBMSex) Then
                       Record.SubItems(10) = !MBMSex
                    End If
                    
                    If Len(!AgeCode) Then
                       Record.SubItems(11) = !AgeCode
                    End If
                    
                    If Len(!MBMDOB) Then
                        Record.SubItems(12) = Format(!MBMDOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(!ICNo) Then
                       Record.SubItems(13) = !ICNo
                    End If
                    
                    If Len(!Sex) Then
                       Record.SubItems(14) = !Sex
                    End If
                    
                    If Len(!AgeBand) Then
                       Record.SubItems(15) = !AgeBand
                    End If
                    
                    If Len(!DOB) Then
                        Record.SubItems(16) = Format(!DOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(!INSCode) Then
                       Record.SubItems(17) = !INSCode
                    End If
                    
                    If Len(!PLNCode) Then
                       Record.SubItems(18) = !PLNCode
                    End If
                    
                    If Len(!HOSName) Then
                       Record.SubItems(19) = Trim(!HOSName)
                    End If
                    
                    If Len(!FirstDiag) Then
                       Record.SubItems(20) = !FirstDiag
                       FirstICDCode = Record.SubItems(20)
                       sql3 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(FirstICDCode) & "'"
                       rst3.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                       If rst3.recordCount = 0 Then
                            Record.SubItems(21) = ""
                            rst3.Close
                            Set rst3 = Nothing
                       Else
                            Record.SubItems(21) = Trim(rst3!ICDDescription)
                            rst3.Close
                            Set rst3 = Nothing
                       End If
                    End If
                                                            
                    If Len(!FinalDiag) Then
                       Record.SubItems(22) = !FinalDiag
                       FinalICDCode = Record.SubItems(22)
                       sql3 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(FinalICDCode) & "'"
                       rst3.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                       If rst3.recordCount = 0 Then
                            Record.SubItems(23) = ""
                            rst3.Close
                            Set rst3 = Nothing
                       Else
                            Record.SubItems(23) = Trim(rst3!ICDDescription)
                            rst3.Close
                            Set rst3 = Nothing
                       End If
                    End If
                    
                    If Len(!CLMTotalActual) Then
                       Record.SubItems(24) = Format(!CLMTotalActual, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CLMTotalApproved) Then
                       Record.SubItems(25) = Format(!CLMTotalApproved, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CASReserveAmount) Then
                       Record.SubItems(26) = Format(!CASReserveAmount, "#,##0.00;(#,##0.00)")
                    End If
                                        
                    If Len(!CLMRecoveryPaidByM) Then
                       Record.SubItems(27) = Format(!CLMRecoveryPaidByM, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CASStatus) Then
                       Record.SubItems(28) = !CASStatus
                    End If
                    
                    If Len(!CASClaimability) Then
                       Record.SubItems(29) = !CASClaimability
                    End If
                    
                    If Len(!MBMStatus) Then
                       Record.SubItems(30) = !MBMStatus
                    End If
                    
                    If Len(!CASDateAdmitted) Then
                        Record.SubItems(31) = Format(!CASDateAdmitted, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CASDischargeDate) Then
                        Record.SubItems(32) = Format(!CASDischargeDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CaseRegDate) Then
                        Record.SubItems(33) = Format(!CaseRegDate, "dd/mm/yyyy")
                    End If
                    
                    If !NoofDay = "0" Then
                        Record.SubItems(34) = "1"
                    ElseIf Len(!NoofDay) Then
                        Record.SubItems(34) = !NoofDay
                    End If
                    
                    If Len(!ClaimOpendate) Then
                        Record.SubItems(35) = Format(!ClaimOpendate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEffDate) Then
                        Record.SubItems(36) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEXPDate) Then
                        Record.SubItems(37) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    
                    Current = Current + 1
                    lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ResultListing2(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim sql3 As String
Dim FirstICDCode As String
Dim FinalICDCode As String
Dim FirstDescription As String
Dim FinalDescription As String

        
    ListView2.listitems.Clear
    Current = 0
    Total = 0
          
    Sql = " SELECT PAYCode, CASNumber, claimversion, MBMNumber, MBMPolicyNo, " & _
          " MBMPatientCovID, MBMName, MBMDOB, DOB, PatientName, MBMSex, AgeCode, Sex, Ageband, " & _
          " InsCode, PLNCode, HOSName, FirstDiag, FinalDiag, " & _
          " CLMTotalActual, CLMTotalApproved, CASReserveAmount, " & _
          " CLMRecoveryPaidbyM, BordxRefNo, BordxDate, ClaimOpenDate, " & _
          " CASClaimability, CASStatus, MBMStatus, CASDateAdmitted, " & _
          " CASDischargeDate, MBMPayorEffDate, MBMPayorExpDate, " & _
          " NoofDay, CaseRegDate, CASRemark, MBMIcBcPp, ICNo From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Order By CASNumber, ClaimVersion "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend '+ sql2
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    'If rst2.recordCount = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
    
    Do
              
       Do Until rst2.EOF
       
       Total = rst2.recordCount
       lblTotalRecord = Total
        
        FirstICDCode = ""
        FinalICDCode = ""
        FirstDescription = ""
        FinalDescription = ""
        'Text1 = ""
        
        'If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
            'Text1 = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
        'End If
        
        'If Text1 = "0" Then
            'Text1 = Text1 + 1
        'End If
        
            With rst2
                    
                    If Len(!PAYCode) Then
                        Set Record = ListView2.listitems.Add(, , !PAYCode)
                    Else
                        Set Record = ListView2.listitems.Add(, , "Empty")
                    End If
                    
                    'Set Record = ListView2.listitems.Add(, , rst2!CASNumber & "-" & rst2!claimversion)
                    
                    If Len(rst2!CASNumber & "-" & rst2!claimversion) Then
                       Record.SubItems(1) = (rst2!CASNumber & "-" & rst2!claimversion)
                    End If
                    
                    If Len(rst2!BordxRefNo) Then
                       Record.SubItems(2) = rst2!BordxRefNo
                    End If
                    
                    If Len(rst2!BordxDate) Then
                       Record.SubItems(3) = Format(rst2!BordxDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!MBMNumber) Then
                       Record.SubItems(4) = rst2!MBMNumber
                    End If
                    
                    If Len(rst2!MBMPolicyNo) Then
                       Record.SubItems(5) = Trim(rst2!MBMPolicyNo)
                    End If
                                    
                    If Len(rst2!MBMPatientCovID) Then
                       Record.SubItems(6) = rst2!MBMPatientCovID
                    End If
                    
                    If Len(rst2!MBMName) Then
                       Record.SubItems(7) = Trim(rst2!MBMName)
                    End If
                    
                    If Len(rst2!PatientName) Then
                       Record.SubItems(8) = Trim(rst2!PatientName)
                    End If
                    
                    If Len(!MBMIcBcPp) Then
                       Record.SubItems(9) = !MBMIcBcPp
                    End If
                    
                    If Len(!MBMSex) Then
                       Record.SubItems(10) = !MBMSex
                    End If
                    
                    If Len(!AgeCode) Then
                       Record.SubItems(11) = !AgeCode
                    End If
                    
                    If Len(!MBMDOB) Then
                        Record.SubItems(12) = Format(!MBMDOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(!ICNo) Then
                       Record.SubItems(13) = !ICNo
                    End If
                    
                    If Len(!Sex) Then
                       Record.SubItems(14) = !Sex
                    End If
                    
                    If Len(!AgeBand) Then
                       Record.SubItems(15) = !AgeBand
                    End If
                    
                    If Len(!DOB) Then
                        Record.SubItems(16) = Format(!DOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(!INSCode) Then
                       Record.SubItems(17) = !INSCode
                    End If
                    
                    If Len(!PLNCode) Then
                       Record.SubItems(18) = !PLNCode
                    End If
                    
                    If Len(!HOSName) Then
                       Record.SubItems(19) = Trim(!HOSName)
                    End If
                    
                    If Len(!FirstDiag) Then
                       Record.SubItems(20) = !FirstDiag
                       FirstICDCode = Record.SubItems(20)
                       sql3 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(FirstICDCode) & "'"
                       rst3.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                       If rst3.recordCount = 0 Then
                            Record.SubItems(21) = ""
                            rst3.Close
                            Set rst3 = Nothing
                       Else
                            Record.SubItems(21) = Trim(rst3!ICDDescription)
                            rst3.Close
                            Set rst3 = Nothing
                       End If
                    End If
                                                            
                    If Len(!FinalDiag) Then
                       Record.SubItems(22) = !FinalDiag
                       FinalICDCode = Record.SubItems(22)
                       sql3 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(FinalICDCode) & "'"
                       rst3.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                       If rst3.recordCount = 0 Then
                            Record.SubItems(23) = ""
                            rst3.Close
                            Set rst3 = Nothing
                       Else
                            Record.SubItems(23) = Trim(rst3!ICDDescription)
                            rst3.Close
                            Set rst3 = Nothing
                       End If
                    End If
                    
                    If Len(!CLMTotalActual) Then
                       Record.SubItems(24) = Format(!CLMTotalActual, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CLMTotalApproved) Then
                       Record.SubItems(25) = Format(!CLMTotalApproved, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CASReserveAmount) Then
                       Record.SubItems(26) = Format(!CASReserveAmount, "#,##0.00;(#,##0.00)")
                    End If
                                        
                    If Len(!CLMRecoveryPaidByM) Then
                       Record.SubItems(27) = Format(!CLMRecoveryPaidByM, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(!CASStatus) Then
                       Record.SubItems(28) = !CASStatus
                    End If
                    
                    If Len(!CASClaimability) Then
                       Record.SubItems(29) = !CASClaimability
                    End If
                    
                    If Len(!MBMStatus) Then
                       Record.SubItems(30) = !MBMStatus
                    End If
                    
                    If Len(!CASDateAdmitted) Then
                        Record.SubItems(31) = Format(!CASDateAdmitted, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CASDischargeDate) Then
                        Record.SubItems(32) = Format(!CASDischargeDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CaseRegDate) Then
                        Record.SubItems(33) = Format(!CaseRegDate, "dd/mm/yyyy")
                    End If
                    
                    If !NoofDay = "0" Then
                        Record.SubItems(34) = "1"
                    ElseIf Len(!NoofDay) Then
                        Record.SubItems(34) = !NoofDay
                    End If
                    
                    If Len(!ClaimOpendate) Then
                        Record.SubItems(35) = Format(!ClaimOpendate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEffDate) Then
                        Record.SubItems(36) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEXPDate) Then
                        Record.SubItems(37) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!CASRemark) Then
                        Record.SubItems(38) = Trim(rst2!CASRemark)
                    End If
                    
                    Current = Current + 1
                    lblCurrent = Current
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub INSListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    lblCurrent = 0
    ListView2.listitems.Clear
        
    Sql = " SELECT PAYCode, INSCode, COUNT(*) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved, SUM(NoofDay) AS NoofDay " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By PAYCode, INSCode "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!INSCode) Then
                    Record.SubItems(1) = rst2!INSCode
                End If
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(2) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountINSRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
        
    
    TotalDuration = 0
        
    ListView3.listitems.Clear
                
    If Len(INSCode) Or INSCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where INSCode  = '" & Trim(INSCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where INSCode IS NULL "
    End If
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
                
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub



Private Sub StateListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
    
    lblCurrent = 0
    ListView2.listitems.Clear
        
    Sql = " SELECT PAYCode, STACode, COUNT(*) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By PAYCode, STACode "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!STACode) Then
                    Record.SubItems(1) = Trim(rst2!STACode)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(2) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub PayorListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    lblCurrent = 0
    ListView2.listitems.Clear
        
    Sql = " SELECT PAYCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved, SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By PAYCode Order By PAYCode"
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
     
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(1) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(2) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ListViewHeader6()
If Option35.value = True Then
    ListView2.ColumnHeaders(1).Text = "Claim Type"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Total Duration"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    
    For aa = 6 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
ElseIf Option34.value = True Then
    ListView2.ColumnHeaders(1).Text = "Claim Type"
    ListView2.ColumnHeaders(2).Text = "No of Claims"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "No of Follow-Up"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "No of Pre-Hosp."
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
        
    For aa = 5 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
End If
End Sub

Private Sub ListViewHeader8()
    
If Option29.value = True Then
    ListView2.ColumnHeaders(1).Text = "Diag. Code"
    ListView2.ColumnHeaders(2).Text = "Diag. Desc"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Hospital"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "No of Claims"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Duration"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    
    For aa = 8 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
Else
    ListView2.ColumnHeaders(1).Text = "Diag. Code"
    ListView2.ColumnHeaders(2).Text = "Diag. Desc"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "No of Claims"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "Total Duration"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "Total Actual Amt"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Total Approved Amt"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    
    For aa = 7 To 63
        ListView2.ColumnHeaders(aa).Text = ""
    Next aa
End If
End Sub

'********************** List Ageband Results *******************************
Private Sub AgeGroupListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount
    
    lblCurrent = 0
    
    ListView2.listitems.Clear
            
    Sql = " SELECT Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY Ageband ORDER BY Ageband "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
        
        Do Until rst2.EOF
        
            'With rst2
                If Len(rst2!AgeBand) Then
                    Set Record = ListView2.listitems.Add(, , rst2!AgeBand)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!TotalCases) Then
                   Record.SubItems(1) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(2) = rst2!NoofDay
                End If
                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(3) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(4) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                
            'End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountAgeGroupRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
        
    
    TotalDuration = 0
        
    ListView3.listitems.Clear
                
    If Len(AgeCode) Or AgeCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where Ageband  = '" & Trim(AgeCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where Ageband IS NULL "
    End If
    
        
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
        
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
                        
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Ageband Results *******************************

'********************** List Ageband By Gender Results *******************************
Private Sub AgeGenderListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount

    
    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT Sex, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY Sex, Ageband ORDER BY Sex, Ageband "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
        
        Do Until rst2.EOF
        
            'With rst2
                If Len(rst2!Sex) Then
                    Set Record = ListView2.listitems.Add(, , rst2!Sex)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                                              
                If Len(rst2!AgeBand) Then
                   Record.SubItems(1) = rst2!AgeBand
                End If
                                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(2) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(4) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(5) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                
            'End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountAgeGenderRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
        
    
    TotalDuration = 0
            
    ListView3.listitems.Clear
                
    If Len(SexCode) Or SexCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where Sex  = '" & Trim(SexCode) & "' AND Ageband = '" & Trim(AgeCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where Sex IS NULL AND Ageband = '" & Trim(AgeCode) & "'"
    End If
    
        
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
        
            If rst2!CASDischargeDate <> "" And rst2!CASDateAdmitted <> "" Then
            
                If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                    TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
                End If
                
                If TxtDay = "0" Then
                    TxtDay = TxtDay + 1
                End If
            Else
                TxtDay = "0"
            End If
            
            'With rst2
                'If Len(rst2!CASDischargeDate) Then
                    'Set Record = ListView3.listitems.Add(, , rst2!CASDischargeDate)
                'End If
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Ageband By Gender Results *******************************

'********************** List Ageband By Insured Type Results *******************************
Private Sub AgeInsuredListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount
   
    lblCurrent = 0
    
    ListView2.listitems.Clear
            
    Sql = " SELECT PAYCode, INSCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY PAYCode, INSCode, Ageband "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
        
        Do Until rst2.EOF
        
            'With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!INSCode) Then
                   Record.SubItems(1) = rst2!INSCode
                End If
                                              
                If Len(rst2!AgeBand) Then
                   Record.SubItems(2) = rst2!AgeBand
                End If
                         
                If Len(rst2!TotalCases) Then
                   Record.SubItems(3) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(4) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(4) = rst2!NoofDay
                End If
                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(5) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(6) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                
            'End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountAgeInsuredRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
        
    
    TotalDuration = 0
        
    ListView3.listitems.Clear
                
    If Len(INSCode) Or INSCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where INSCode  = '" & Trim(INSCode) & "' AND Ageband = '" & Trim(AgeCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where INSCode IS NULL AND Ageband = '" & Trim(AgeCode) & "'"
    End If
    
        
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Ageband By Insured Type Results *******************************

'********************** List Ageband By Race Type Results *******************************
Private Sub AgeRaceListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount

    
    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT RACCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY RACCode, Ageband ORDER BY RACCode, Ageband "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
        
        Do Until rst2.EOF
        
            'With rst2
                If Len(rst2!RACCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!RACCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                                              
                If Len(rst2!AgeBand) Then
                   Record.SubItems(1) = rst2!AgeBand
                End If
                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(2) = rst2!TotalCases
                End If
                    
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(4) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(5) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                
            'End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountAgeRaceRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
        
    
    TotalDuration = 0
        
    ListView3.listitems.Clear
                
    If Len(RaceCode) Or RaceCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where RACCode  = '" & Trim(RaceCode) & "' AND Ageband = '" & Trim(AgeCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where RACCode IS NULL AND Ageband = '" & Trim(AgeCode) & "'"
    End If
    
        
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Ageband By Race Type Results *******************************

'********************** List Ageband By Plan Type Results *******************************
Private Sub AgePlanListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount

    lblCurrent = 0
    
    ListView2.listitems.Clear
            
    Sql = " SELECT INSCode, PLNCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY INSCode, PLNCode, Ageband ORDER BY INSCode "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
        
        Do Until rst2.EOF
        
            'With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!INSCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                                              
                If Len(rst2!PLNCode) Then
                   Record.SubItems(1) = rst2!PLNCode
                End If
                    
                If Len(rst2!AgeBand) Then
                   Record.SubItems(2) = rst2!AgeBand
                End If
                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(3) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(4) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(4) = rst2!NoofDay
                End If
                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(5) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(6) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                
            'End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountAgePlanRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
        
    
    TotalDuration = 0
        
    ListView3.listitems.Clear
                
    If Len(INSCode) Or INSCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where INSCode  = '" & Trim(INSCode) & "'" & _
              " AND PLNCode = '" & Trim(PlanCode) & "' AND Ageband = '" & Trim(AgeCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where INSCode IS NULL " & _
              " AND PLNCode = '" & Trim(PlanCode) & "' AND Ageband = '" & Trim(AgeCode) & "'"
    End If
    
        
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Ageband By Plan Type Results *******************************

'********************** List Ageband By Illness Results *******************************
Private Sub AgeIllnessListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount

    lblCurrent = 0
    
    ListView2.listitems.Clear
            
    Sql = " SELECT ClaimNatureCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY ClaimNatureCode, Ageband ORDER BY ClaimNatureCode "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
        
        Do Until rst2.EOF
        
            'With rst2
                If Len(rst2!ClaimNatureCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!ClaimNatureCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                                        
                If Len(rst2!AgeBand) Then
                   Record.SubItems(1) = rst2!AgeBand
                End If
                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(2) = rst2!TotalCases
                End If
                    
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                                    
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(4) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(5) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                    
            'End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
       
    If ListView2.listitems.count = 0 Then
     MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
     CmdPrint.Enabled = True
     CmdClear.Enabled = True
     cmdExit.Enabled = True
    End If
    'End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountAgeIllnessRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
        
    
    TotalDuration = 0
        
    ListView3.listitems.Clear
                
    If Len(ClaimNatureCode) Or ClaimNatureCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where ClaimNatureCode  = '" & Trim(ClaimNatureCode) & "'" & _
              " AND Ageband = '" & Trim(AgeCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where ClaimNatureCode IS NULL " & _
              " AND Ageband = '" & Trim(AgeCode) & "'"
    End If
    
        
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
        
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Ageband By Illness Results *******************************

'********************** List Gender Results *******************************
Private Sub SexListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT SEX, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By SEX ORDER BY SEX "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
      
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!Sex) Then
                    Set Record = ListView2.listitems.Add(, , rst2!Sex)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                               
                If Len(rst2!TotalCases) Then
                    Record.SubItems(1) = rst2!TotalCases
                End If
                    
                If rst2!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(2) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountSexRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(SexCode) Or SexCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where Sex = '" & Trim(SexCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where Sex IS NULL "
    End If
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Gender Results *******************************

'********************** List Gender By Ageband Results *******************************
Private Sub SexAgebandListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT SEX, Ageband, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By SEX, Ageband ORDER BY SEX "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
      
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!Sex) Then
                    Set Record = ListView2.listitems.Add(, , rst2!Sex)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!AgeBand) Then
                    Record.SubItems(1) = rst2!AgeBand
                End If
                    
                If Len(rst2!TotalCases) Then
                    Record.SubItems(2) = rst2!TotalCases
                End If
                    
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                    
                If Len(rst2!TotalActual) Then
                    Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountSexAgebandRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(SexCode) Or SexCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where Sex  = '" & Trim(SexCode) & "'" & _
              " AND Ageband = '" & Trim(AgeCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where Sex IS NULL " & _
              " AND Ageband = '" & Trim(AgeCode) & "'"
    End If
    
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Gender By Ageband Results *******************************

'********************** List Gender By Insured Type Results *******************************
Private Sub SexInsuredListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT PAYCode, INSCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By PAYCode, INSCode, Sex "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
      
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!INSCode) Then
                    Record.SubItems(1) = rst2!INSCode
                End If
                
                If Len(rst2!Sex) Then
                    Record.SubItems(2) = rst2!Sex
                End If

                If Len(rst2!TotalCases) Then
                    Record.SubItems(3) = rst2!TotalCases
                End If
                    
                If rst2!NoofDay = "0" Then
                    Record.SubItems(4) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(4) = rst2!NoofDay
                End If
                    
                If Len(rst2!TotalActual) Then
                    Record.SubItems(5) = Format(rst2!TotalActual, "#,##0.00")
                End If
                    
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(6) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountSexInsuredRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(INSCode) Or INSCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where INSCode  = '" & Trim(INSCode) & "'" & _
              " AND Sex = '" & Trim(SexCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where INSCode IS NULL " & _
              " AND Sex = '" & Trim(SexCode) & "'"
    End If
    
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Gender By Insured Type Results *******************************

'********************** List Gender By Race Results *******************************
Private Sub SexRaceListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT RACCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By RACCode, Sex ORDER BY RACCode "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
      
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!RACCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!RACCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!Sex) Then
                    Record.SubItems(1) = rst2!Sex
                End If
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(2) = rst2!TotalCases
                End If
                    
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                    
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountSexRaceRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(RaceCode) Or RaceCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where RACCode  = '" & Trim(RaceCode) & "'" & _
              " AND Sex = '" & Trim(SexCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where RACCode IS NULL " & _
              " AND Sex = '" & Trim(SexCode) & "'"
    End If
    
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Gender By Race Results *******************************

'********************** List Gender By Plan Results *******************************
Private Sub SexPlanListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT INSCode, PLNCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, SUM(NoofDay) AS NoofDay  " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By INSCode, PLNCode, Sex ORDER BY INSCode "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
      
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!INSCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!PLNCode) Then
                    Record.SubItems(1) = rst2!PLNCode
                End If
                    
                If Len(rst2!Sex) Then
                    Record.SubItems(2) = rst2!Sex
                End If
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(3) = rst2!TotalCases
                End If
                    
                If rst2!NoofDay = "0" Then
                    Record.SubItems(4) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(4) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(5) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(6) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountSexPlanRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(INSCode) Or INSCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where INSCode  = '" & Trim(INSCode) & "'" & _
              " AND PLNCode = '" & Trim(PlanCode) & "' AND Sex  = '" & Trim(SexCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where INSCode IS NULL " & _
              " AND PLNCode = '" & Trim(PlanCode) & "' AND Sex  = '" & Trim(SexCode) & "'"
    End If
                
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Gender By Plan Results *******************************

'********************** List Gender By Illness Results *******************************
Private Sub SexIllnessListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    
    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT ClaimNatureCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By ClaimNatureCode, Sex ORDER BY ClaimNatureCode "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
      
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!ClaimNatureCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!ClaimNatureCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                    
                If Len(rst2!Sex) Then
                    Record.SubItems(1) = rst2!Sex
                End If
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(2) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
     MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
     CmdPrint.Enabled = True
     CmdClear.Enabled = True
     cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountSexIllnessRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(ClaimNatureCode) Or ClaimNatureCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where ClaimNatureCode  = '" & Trim(ClaimNatureCode) & "'" & _
              " AND Sex = '" & Trim(SexCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where ClaimNatureCode IS NULL " & _
              " AND Sex = '" & Trim(SexCode) & "'"
    End If
    
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Gender By Illness Results *******************************

'********************** List Race Results *******************************
Private Sub RaceListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT RACCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By RACCode ORDER BY RACCode "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
      
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!RACCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!RACCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                                               
                If Len(rst2!TotalCases) Then
                    Record.SubItems(1) = rst2!TotalCases
                End If
                    
                If rst2!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(2) = rst2!NoofDay
                End If
                    
                If Len(rst2!TotalActual) Then
                    Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountRaceRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(RaceCode) Or RaceCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where RACCode  = '" & Trim(RaceCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where RACCode IS NULL "
    End If
    
        
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'Else
                    'Set Record = ListView3.listitems.Add(, , "Empty")
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Race Results *******************************

'********************** List Race By Ageband Results *******************************
Private Sub RaceAgebandListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String


    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT RACCode, Ageband, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By RACCode, Ageband ORDER BY RACCode "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
      
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!RACCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!RACCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!AgeBand) Then
                    Record.SubItems(1) = rst2!AgeBand
                End If
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(2) = rst2!TotalCases
                End If
                    
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                    
                If Len(rst2!TotalActual) Then
                    Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountRaceAgebandRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(RaceCode) Or RaceCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where RACCode  = '" & Trim(RaceCode) & "'" & _
              " AND Ageband  = '" & Trim(AgeCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where RACCode IS NULL " & _
              " AND Ageband  = '" & Trim(AgeCode) & "'"
    End If
    
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Race By Ageband Results *******************************

'********************** List Race By Insured Type Results *******************************
Private Sub RaceInsuredListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT PAYCode, INSCode, RACCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By PAYCode, INSCode, RACCode "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!INSCode) Then
                    Record.SubItems(1) = rst2!INSCode
                End If
                
                If Len(rst2!RACCode) Then
                    Record.SubItems(2) = rst2!RACCode
                Else
                    Record.SubItems(2) = "Empty"
                End If
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(3) = rst2!TotalCases
                End If
                    
                If rst2!NoofDay = "0" Then
                    Record.SubItems(4) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(4) = rst2!NoofDay
                End If
                    
                If Len(rst2!TotalActual) Then
                    Record.SubItems(5) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(6) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountRaceInsuredRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(INSCode) Or INSCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where INSCode  = '" & Trim(INSCode) & "'" & _
              " AND RACCode = '" & Trim(RaceCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where INSCode IS NULL " & _
              " AND RACCode = '" & Trim(RaceCode) & "'"
    End If
    
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Race By Insured Type Results *******************************

'********************** List Race By Gender Results *******************************
Private Sub RaceSexListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT RACCode, Sex, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By RACCode, Sex ORDER BY RACCode "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!RACCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!RACCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!Sex) Then
                    Record.SubItems(1) = rst2!Sex
                End If
                              
                If Len(rst2!TotalCases) Then
                    Record.SubItems(2) = rst2!TotalCases
                End If
                    
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                    
                If Len(rst2!TotalActual) Then
                    Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountRaceSexRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(RaceCode) Or RaceCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where RACCode  = '" & Trim(RaceCode) & "'" & _
              " AND Sex = '" & Trim(SexCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where RACCode IS NULL " & _
              " AND Sex = '" & Trim(SexCode) & "'"
    End If
    
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Race By Sex Results *******************************

'********************** List Race By Plan Results *******************************
Private Sub RacePlanListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    
    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT INSCode, PLNCode, RACCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By INSCode, PLNCode, RACCode ORDER BY INSCode "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!INSCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!INSCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!PLNCode) Then
                    Record.SubItems(1) = rst2!PLNCode
                End If
                    
                If Len(rst2!RACCode) Then
                    Record.SubItems(2) = rst2!RACCode
                End If
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(3) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(4) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(4) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(5) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(6) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountRacePlanRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(INSCode) Or INSCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where INSCode  = '" & Trim(INSCode) & "'" & _
              " AND PLNCode = '" & Trim(PlanCode) & "' AND RACCode  = '" & Trim(RaceCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where INSCode IS NULL " & _
              " AND PLNCode = '" & Trim(PlanCode) & "' AND RACCode  = '" & Trim(RaceCode) & "'"
    End If
    
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Race By Plan Results *******************************

'********************** List Race By Illness Results *******************************
Private Sub RaceIllnessListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT ClaimNatureCode, RACCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By ClaimNatureCode, RACCode ORDER BY ClaimNatureCode "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!ClaimNatureCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!ClaimNatureCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                    
                If Len(rst2!RACCode) Then
                    Record.SubItems(1) = rst2!RACCode
                End If
                
                If Len(rst2!TotalCases) Then
                    Record.SubItems(2) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountRaceIllnessRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(ClaimNatureCode) Or ClaimNatureCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where ClaimNatureCode  = '" & Trim(ClaimNatureCode) & "'" & _
              " AND RACCode  = '" & Trim(RaceCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where ClaimNatureCode IS NULL " & _
              " AND RACCode  = '" & Trim(RaceCode) & "'"
    End If
    
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Race By Illness Results *******************************

'********************** List Illness Results *******************************
Private Sub IllnessListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    lblCurrent = 0
    ListView2.listitems.Clear
          
    Sql = " SELECT ClaimNatureCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By ClaimNatureCode ORDER BY ClaimNatureCode "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!ClaimNatureCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!ClaimNatureCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                               
                If Len(rst2!TotalCases) Then
                    Record.SubItems(1) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(2) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                End If
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountIllnessRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    ListView3.listitems.Clear
    TotalDuration = 0
    
    If Len(ClaimNatureCode) And ClaimNatureCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where ClaimNatureCode  = '" & Trim(ClaimNatureCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where ClaimNatureCode IS NULL "
    End If
    
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
        'TotalDuration = 0
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Illness Results *******************************

'********************** List Agent Detail Results *******************************
Private Sub AgentListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim sql3 As String
Dim FirstICDCode As String
Dim FinalICDCode As String
Dim FirstDescription As String
Dim FinalDescription As String

        
    ListView2.listitems.Clear
    Current = 0
    Total = 0
            
    Sql = " SELECT PAYCode, MBMAgentCode, CASNumber, ClaimVersion, MBMNumber, " & _
          " MBMPatientCovID, PatientName, AttendDoctor, CLMTotalActual, " & _
          " CASDischargeDate, CASDateAdmitted, CLMTotalApproved, CLMRecoveryPaidbyM, " & _
          " FirstDiag, FinalDiag, InsCode, PLNCode, CASStatus " & _
          " From VIEW_CLMReport Where 0 = 0 "
    
    'sql2 = " Order By MBMAgentCode "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend '+ sql2
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    'If rst2.recordCount = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
    
    Do
              
       Do Until rst2.EOF
       
       Total = rst2.recordCount
       lblTotalRecord = Total
                      
        FirstICDCode = ""
        FinalICDCode = ""
        FirstDescription = ""
        FinalDescription = ""
                
            'With rst2
                    If Len(rst2!PAYCode) Then
                        Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                    Else
                        Set Record = ListView2.listitems.Add(, , "Empty")
                    End If
                    
                    If Len(rst2!MBMAgentCode) Then
                        Record.SubItems(1) = Trim(rst2!MBMAgentCode)
                    Else
                        Record.SubItems(1) = "Empty"
                    End If
                    
                    If Len(rst2!CASNumber & "-" & rst2!claimversion) Then
                       Record.SubItems(2) = rst2!CASNumber & "-" & rst2!claimversion
                    End If
                    
                    If Len(rst2!MBMNumber) Then
                       Record.SubItems(3) = rst2!MBMNumber
                    End If
                                    
                    If Len(rst2!MBMPatientCovID) Then
                       Record.SubItems(4) = rst2!MBMPatientCovID
                    End If
                    
                    If Len(rst2!PatientName) Then
                       Record.SubItems(5) = Trim(rst2!PatientName)
                    End If
                    
                    If Len(rst2!AttendDoctor) Then
                       Record.SubItems(6) = Trim(rst2!AttendDoctor)
                    End If
                    
                    If Len(rst2!CLMTotalActual) Then
                       Record.SubItems(7) = Format(rst2!CLMTotalActual, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(rst2!CLMTotalApproved) Then
                       Record.SubItems(8) = Format(rst2!CLMTotalApproved, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(rst2!CLMRecoveryPaidByM) Then
                       Record.SubItems(9) = Format(rst2!CLMRecoveryPaidByM, "#,##0.00;(#,##0.00)")
                    End If
                    
                    If Len(rst2!FirstDiag) Then
                       Record.SubItems(10) = rst2!FirstDiag
                       FirstICDCode = Record.SubItems(10)
                       sql3 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(FirstICDCode) & "'"
                       rst3.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                       If rst3.recordCount = 0 Then
                            rst3.Close
                            Set rst3 = Nothing
                       Else
                            FirstDescription = rst3!ICDDescription
                            rst3.Close
                            Set rst3 = Nothing
                       End If
                    End If
                    
                    If Len(FirstDescription) Then
                       Record.SubItems(11) = Trim(FirstDescription)
                    End If
                                                            
                    If Len(rst2!FinalDiag) Then
                       Record.SubItems(12) = rst2!FinalDiag
                       FinalICDCode = Record.SubItems(12)
                       sql3 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(FinalICDCode) & "'"
                       rst3.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                       If rst3.recordCount = 0 Then
                            rst3.Close
                            Set rst3 = Nothing
                       Else
                            FinalDescription = rst3!ICDDescription
                            rst3.Close
                            Set rst3 = Nothing
                       End If
                    End If
                    
                    If Len(FinalDescription) Then
                       Record.SubItems(13) = Trim(FinalDescription)
                    End If
                                        
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(14) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(15) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!INSCode) Then
                       Record.SubItems(16) = rst2!INSCode
                    End If
                    
                    If Len(rst2!PLNCode) Then
                       Record.SubItems(17) = rst2!PLNCode
                    End If
                                        
                    If Len(rst2!CASStatus) Then
                       Record.SubItems(18) = rst2!CASStatus
                    End If
                    
                    Current = Current + 1
                    lblCurrent = Current
                    
            'End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    intExit = 0
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub
'********************** List Agent Detail Results *******************************

'********************** List Agent Summary Results *******************************
Private Sub AgentSumListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    lblCurrent = 0
    ListView2.listitems.Clear
            
    If Option34.value = True Then
    
    Sql = " SELECT PAYCode, MBMAgentCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " Group By PAYCode, MBMAgentCode "
    
    ElseIf Option33.value = True Then
    
    Sql = " SELECT Top 50 PAYCode, MBMAgentCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " Group By PAYCode, MBMAgentCode ORDER BY TotalCases Desc "
    
    End If
    
       
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            'With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!MBMAgentCode) Then
                    Record.SubItems(1) = Trim(rst2!MBMAgentCode)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                                                         
                If Len(rst2!TotalCases) Then
                    Record.SubItems(2) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            'End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
       
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountAgentSumRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(AgentCode) Or AgentCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where MBMAgentCode  = '" & Trim(AgentCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where MBMAgentCode IS NULL "
    End If
    
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Agent Summary Results *******************************

'********************** List Agent with Same Doctor Results *******************************
Private Sub AgentDoctorListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

        
    lblCurrent = 0
    
    ListView2.listitems.Clear
            
    Sql = " SELECT PAYCode, MBMAgentCode, AttendDoctor, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " Group By PAYCode, MBMAgentCode, AttendDoctor "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            'With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!MBMAgentCode) Then
                    Record.SubItems(1) = Trim(rst2!MBMAgentCode)
                Else
                    Record.SubItems(1) = "Empty"
                End If
                    
                If Len(rst2!AttendDoctor) Then
                    Record.SubItems(2) = rst2!AttendDoctor
                    DoctorName = Trim(rst2!AttendDoctor)
                End If
                               
                If Len(rst2!TotalCases) Then
                    Record.SubItems(3) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(4) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(4) = rst2!NoofDay
                End If
                    
                If Len(rst2!TotalActual) Then
                    Record.SubItems(5) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(6) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            'End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountAgentDoctorRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(AgentCode) Or AgentCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where MBMAgentCode  = '" & Trim(AgentCode) & "'" & _
              " AND AttendDoctor = '" & Trim(DoctorName) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where MBMAgentCode IS NULL " & _
              " AND AttendDoctor = '" & Trim(DoctorName) & "'"
    End If
    
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Agent with Same Doctor Results *******************************

'********************** List Plan Summary Results *******************************
Private Sub PlanListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    lblCurrent = 0
    ListView2.listitems.Clear
        
    Sql = " SELECT PAYCode, PLNCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By PAYCode, PLNCode "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!PLNCode) Then
                    Record.SubItems(1) = rst2!PLNCode
                End If
                               
                If Len(rst2!TotalCases) Then
                    Record.SubItems(2) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub PlanPremiumListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    lblCurrent = 0
    ListView2.listitems.Clear
        
    Sql = " SELECT PAYCode, PLNCode, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay, SUM(MBMBasicPrem) AS BasicPremium, " & _
          " SUM(MBMPremium) AS Premium From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By PAYCode, PLNCode "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!PLNCode) Then
                    Record.SubItems(1) = rst2!PLNCode
                End If
                               
                If Len(rst2!TotalCases) Then
                    Record.SubItems(2) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(4) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(5) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
                If Len(rst2!BasicPremium) Then
                    Record.SubItems(6) = Format(rst2!BasicPremium, "#,##0.00")
                End If
                
                If Len(rst2!Premium) Then
                    Record.SubItems(7) = Format(rst2!Premium, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountPlanRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
        
    
    TotalDuration = 0
        
    ListView3.listitems.Clear
                
    If Len(PlanCode) Or PlanCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where PLNCode  = '" & Trim(PlanCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where PLNCode IS NULL "
    End If
    
        
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
                
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Plan Summary Results *******************************

'********************** List Plan Top 10 Hospital Results *******************************
Private Sub PlanTopHosListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    lblCurrent = 0
    ListView2.listitems.Clear
        
    Sql = " SELECT Top 10 PAYCode, PLNCode, HOSName, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By PAYCode, PLNCode, HOSName ORDER BY TotalCases Desc"
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!PLNCode) Then
                    Record.SubItems(1) = Trim(rst2!PLNCode)
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(2) = Trim(rst2!HOSName)
                End If
                           
                If Len(rst2!TotalCases) Then
                    Record.SubItems(3) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(4) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(4) = rst2!NoofDay
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(5) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(6) = Format(rst2!TotalApproved, "#,##0.00")
                End If
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountPlanTopHOSRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
        
    
    TotalDuration = 0
        
    ListView3.listitems.Clear
                
    If Len(PlanCode) Or PlanCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where PLNCode  = '" & Trim(PlanCode) & "'" & _
              " AND HOSName = '" & Trim(Hospital) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where PLNCode IS NULL " & _
              " AND HOSName = '" & Trim(Hospital) & "'"
    End If
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
                
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Plan Top 10 Hospital Results *******************************

'********************** List Plan Top 10 Final Diag Results *******************************
Private Sub PlanTopDiagListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim sql3 As String

    
    lblCurrent = 0
    ListView2.listitems.Clear
        
    Sql = " SELECT Top 10 PAYCode, PLNCode, FinalDiag, COUNT(CASNumber) as TotalCases, SUM(clmtotalactual) AS totalActual, " & _
          " SUM(clmtotalapproved) As TotalApproved, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
    
    sql2 = " Group By PAYCode, PLNCode, FinalDiag ORDER BY TotalCases Desc"
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!PLNCode) Then
                    Record.SubItems(1) = rst2!PLNCode
                End If
                
                If Len(rst2!FinalDiag) Then
                    Record.SubItems(2) = rst2!FinalDiag
                    DiagCode = rst2!FinalDiag
                    
                    sql3 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(DiagCode) & "'"
                    rst3.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                   
                   If rst3.recordCount = 0 Then
                        Record.SubItems(3) = ""
                        rst3.Close
                        Set rst3 = Nothing
                   Else
                        Record.SubItems(3) = Trim(rst3!ICDDescription)
                        rst3.Close
                        Set rst3 = Nothing
                   End If
                End If
                    
                If Len(rst2!TotalCases) Then
                    Record.SubItems(4) = rst2!TotalCases
                End If
                    
                If rst2!NoofDay = "0" Then
                    Record.SubItems(5) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(5) = rst2!NoofDay
                End If
                    
                If Len(rst2!TotalActual) Then
                    Record.SubItems(6) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(7) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountPlanTopDiagRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
        
    
    TotalDuration = 0
        
    ListView3.listitems.Clear
                
    If Len(PlanCode) Or PlanCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where PLNCode  = '" & Trim(PlanCode) & "'" & _
              " AND FinalDiag = '" & Trim(DiagCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where PLNCode IS NULL " & _
              " AND FianlDiag = '" & Trim(DiagCode) & "'"
    End If
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
                
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Plan Top 10 Final Diag Results *******************************

'********************** List Plan Type by Ageband Results *******************************
Private Sub PlanAgebandListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount

    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT PAYCode, PLNCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY PAYCode, PLNCode, Ageband "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
        
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!PLNCode) Then
                   Record.SubItems(1) = rst2!PLNCode
                End If
                    
                If Len(rst2!AgeBand) Then
                   Record.SubItems(2) = rst2!AgeBand
                End If
                                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(3) = rst2!TotalCases
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(4) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(4) = rst2!NoofDay
                End If
                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(5) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(6) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountPlanAgebandRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
        
    
    TotalDuration = 0
        
    ListView3.listitems.Clear
                
    If Len(PlanCode) Or PlanCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where PLNCode  = '" & Trim(PlanCode) & "'" & _
              " AND Ageband = '" & Trim(AgeCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where PLNCode IS NULL " & _
              " AND Ageband = '" & Trim(AgeCode) & "'"
    End If
    
        
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Plan Type By Ageband Results *******************************

'********************** List Plan Type by Insured Results *******************************
Private Sub PlanInsuredListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount

    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT PAYCode, PLNCode, INSCode, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY PAYCode, PLNCode, INSCode "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
        
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!INSCode) Then
                   Record.SubItems(1) = rst2!INSCode
                End If
                    
                If Len(rst2!PLNCode) Then
                   Record.SubItems(2) = rst2!PLNCode
                End If
                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(3) = rst2!TotalCases
                End If
                    
                If rst2!NoofDay = "0" Then
                    Record.SubItems(4) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(4) = rst2!NoofDay
                End If
                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(5) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(6) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub PlanInsuredListing2(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount

    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT PAYCode, PLNCode, INSCode, Ageband, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY PAYCode, PLNCode, INSCode, Ageband "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!INSCode) Then
                   Record.SubItems(1) = rst2!INSCode
                End If
                    
                If Len(rst2!PLNCode) Then
                   Record.SubItems(2) = rst2!PLNCode
                End If
                
                If Len(rst2!AgeBand) Then
                   Record.SubItems(3) = rst2!AgeBand
                End If
                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(4) = rst2!TotalCases
                End If
                    
                If rst2!NoofDay = "0" Then
                    Record.SubItems(5) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(5) = rst2!NoofDay
                End If
                                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(6) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(7) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountPlanInsuredRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
        
    
    TotalDuration = 0
        
    ListView3.listitems.Clear
                
    If Len(INSCode) Or INSCode = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where INSCode  = '" & Trim(INSCode) & "'" & _
              " AND PLNCode = '" & Trim(PlanCode) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where INSCode IS NULL " & _
              " AND PLNCode = '" & Trim(PlanCode) & "'"
    End If
    
        
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Plan Type By Insured Results *******************************

'********************** List Hospital Group Summary Results *******************************
Private Sub HOSGrpSumListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount

    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT HOSGRPName, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt, " & _
          " SUM(NoofDay) AS NoofDay From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY HOSGRPName ORDER BY HOSGRPName "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
        
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!HOSGRPName) Then
                    Set Record = ListView2.listitems.Add(, , rst2!HOSGRPName)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(1) = rst2!TotalCases
                End If
                    
                If rst2!NoofDay = "0" Then
                    Record.SubItems(2) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(2) = rst2!NoofDay
                End If
                                    
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(3) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(4) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountHOSGRPRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
        
    
    TotalDuration = 0
        
    ListView3.listitems.Clear
                
    If Len(HosGrp) Or HosGrp = "" Then
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where HOSGRPName  = '" & Trim(HosGrp) & "'"
    Else
        Sql = " Select CASDischargeDate, CASDateAdmitted From VIEW_CLMReport Where HOSGRPName IS NULL "
    End If
    
        
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                'If Len(TxtDay) Then
                    'Set Record = ListView3.listitems.Add(, , TxtDay)
                'End If
                'If Len(rst2!CASDateAdmitted) Then
                   'Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                'End If
                'If Len(TxtDay) Then
                   'Record.SubItems(2) = TxtDay
                'End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Hospital Group Summary Results *******************************

'********************** List Claim Type Summary Results *******************************
Private Sub ClaimTypeListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

    
    ListView2.listitems.Clear
        
    Sql = " SELECT PayTypeCode, COUNT(CASNumber) AS TotalCases, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " SUM(NoofDay) AS TotalDuration From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " Group By PayTypeCode ORDER BY PayTypeCode "
        
    strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        strCount = strCount + StrAppend
    Else
        Sql = Sql + sql2
        strCount = strCount + StrAppend
    End If
    
    Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    If rstCount!totalrecord = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    Else
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PayTypeCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PayTypeCode)
                    ClaimType = Trim(rst2!PayTypeCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                    ClaimType = ""
                End If
                               
                If Len(rst2!TotalCases) Then
                    Record.SubItems(1) = rst2!TotalCases
                End If
                
                If Len(rst2!TotalDuration) Then
                    Record.SubItems(2) = rst2!TotalDuration
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
            End With
            rst2.MoveNext
       Loop
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountClaimTypeSumRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    TotalDuration = 0
    
    ListView3.listitems.Clear
                
    If Len(ClaimType) Or ClaimType = "" Then
        Sql = " Select * From VIEW_AnalysisReportNew Where PayTypeCode  = '" & Trim(ClaimType) & "'"
    Else
        Sql = " Select * From VIEW_AnalysisReportNew Where PayTypeCode IS NULL "
    End If
    
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        TxtDay = ""
            If Len(rst2!CASDischargeDate) And Len(rst2!CASDateAdmitted) Then
                TxtDay = DateValue(rst2!CASDischargeDate) - DateValue((rst2!CASDateAdmitted))
            End If
            
            If TxtDay = "0" Then
                TxtDay = TxtDay + 1
            End If
            
            'With rst2
                If Len(rst2!CASDischargeDate) Then
                    Set Record = ListView3.listitems.Add(, , rst2!CASDischargeDate)
                End If
                If Len(rst2!CASDateAdmitted) Then
                   Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                If Len(TxtDay) Then
                   Record.SubItems(2) = TxtDay
                End If
                
                If Len(TxtDay) Then
                    TotalDuration = TotalDuration + CDbl(TxtDay)
                End If
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Claim Type Summary Results *******************************

'********************** List Claim Type By FollowUp Results *******************************
Private Sub ClaimTypeFollowUpListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

     
    ClaimType = ""
    ListView2.listitems.Clear
            
        
    Sql = " SELECT PayTypeCode, COUNT(CASNumber) AS TotalCases " & _
          " From VIEW_CLMReport Where (PayTypeCode = 'DC' or PayTypeCode = 'DR') " & _
          " OR (PayTypeCode = 'IC' or PayTypeCode = 'IR')"
          
    sql2 = " GROUP BY PayTypeCode ORDER BY PayTypeCode "
        
       
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PayTypeCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PayTypeCode)
                    ClaimType = Trim(rst2!PayTypeCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                    ClaimType = ""
                End If
                    
                Call CountClaimTypeFollowUpRecords(StrAppend)
                           
                If Len(rst2!TotalCases) Then
                    Record.SubItems(1) = rst2!TotalCases
                End If
                
                If Len(TotalFollowUp) Then
                    Record.SubItems(2) = TotalFollowUp
                End If
                
                If Len(TotalPreHos) Then
                    Record.SubItems(3) = TotalPreHos
                End If
                
            End With
            rst2.MoveNext
       Loop
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub CountClaimTypeFollowUpRecords(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
    
    TotalFollowUp = 0
    TotalPreHos = 0
    
    ListView3.listitems.Clear
                
    If Len(ClaimType) Or ClaimType = "" Then
        Sql = " Select PayTypeCode, ClaimType From VIEW_CLMReport Where PayTypeCode  = '" & Trim(ClaimType) & "'"
    Else
        Sql = " Select PayTypeCode, ClaimType From VIEW_CLMReport Where PayTypeCode IS NULL "
    End If
    
            
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
    End If
            
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
                  
            'With rst2
                If Len(rst2!PayTypeCode) Then
                    Set Record = ListView3.listitems.Add(, , rst2!PayTypeCode)
                End If
                
                If rst2!ClaimType = "FC" Then
                        
                    Record.SubItems(1) = rst2!ClaimType
                    TotalFollowUp = totalFllowup + 1
                    
                ElseIf rst2!ClaimType = "PC" Then
                    Record.SubItems(1) = rst2!ClaimType
                    TotalPreHos = TotalPreHos + 1
                End If
                
            
            'End With
            rst2.MoveNext
        Loop
    rst2.Close
    Set rst2 = Nothing
End Sub
'********************** List Claim Type By FollowUp Results *******************************

'********************** List Duration Summary *******************************
Private Sub DurationSum(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount

    ListView2.listitems.Clear
            
    Sql = " SELECT NoofDay, COUNT(CASNumber) AS TotalCases, " & _
          " SUM(CLMTotalActual) AS TotalActualAmt, " & _
          " SUM(CLMTotalApproved) AS TotalApprovedAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY NoofDay ORDER BY NoofDay "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
        
        Do Until rst2.EOF
        
            With rst2
            
                If rst2!NoofDay = "0" Then
                    Set Record = ListView2.listitems.Add(, , rst2!TotalCases)
                ElseIf Len(rst2!NoofDay) Then
                    Set Record = ListView2.listitems.Add(, , rst2!NoofDay)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(1) = rst2!TotalCases
                End If
                
                If Len(rst2!TotalActualAmt) Then
                   Record.SubItems(2) = Format(rst2!TotalActualAmt, "#,##0.00")
                End If
                
                If Len(rst2!TotalApprovedAmt) Then
                   Record.SubItems(3) = Format(rst2!TotalApprovedAmt, "#,##0.00")
                End If
                
            End With
            rst2.MoveNext
       Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub
'********************** List Duration Summary *******************************

'********************** List Duration Details *******************************
Private Sub DurationDetail(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim rst3 As New ADODB.Recordset
Dim FirstDiagCode As String
Dim FinalDiagCode As String
    
    
    Current = 0
    Total = 0
    FirstDiagCode = ""
    FinalDiagCode = ""
    ListView2.listitems.Clear
            
    Sql = " SELECT CASNumber, ClaimVersion, CASDateAdmitted, " & _
          " CASDischargeDate, NoofDay, PatientName, AttendDoctor, " & _
          " FirstDiag, FinalDiag From VIEW_CLMReport Where 0 = 0 " & StrAppend & ""
          
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
    'If rst2.recordCount = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
       
    Do
        
        Do Until rst2.EOF
        
            Total = rst2.recordCount
            lblTotalRecord = Total
    
            With rst2
            
                If Len(rst2!CASNumber) And Len(rst2!claimversion) Then
                    Set Record = ListView2.listitems.Add(, , rst2!CASNumber & "-" & rst2!claimversion)
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                   Record.SubItems(1) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASDischargeDate) Then
                   Record.SubItems(2) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                
                If rst2!NoofDay = "0" Then
                    Record.SubItems(3) = "1"
                ElseIf Len(rst2!NoofDay) Then
                    Record.SubItems(3) = rst2!NoofDay
                End If
                
                If Len(rst2!PatientName) Then
                   Record.SubItems(4) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!AttendDoctor) Then
                   Record.SubItems(5) = Trim(rst2!AttendDoctor)
                End If
                
                If Len(rst2!FirstDiag) Then
                   Record.SubItems(6) = Trim(rst2!FirstDiag)
                   FirstDiagCode = rst2!FirstDiag
                Else
                   Record.SubItems(6) = ""
                   FirstDiagCode = ""
                End If
                
                sql2 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(FirstDiagCode) & "'"
                rst3.Open sql2, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                If rst3.recordCount = 0 Then
                    rst3.Close
                    Set rst3 = Nothing
                    Record.SubItems(7) = ""
                Else
                    Description1 = Trim(rst3!ICDDescription)
                    Record.SubItems(7) = Trim(Description1)
                    rst3.Close
                    Set rst3 = Nothing
                End If
                
                If Len(rst2!FinalDiag) Then
                   Record.SubItems(8) = Trim(rst2!FinalDiag)
                   FinalDiagCode = rst2!FinalDiag
                Else
                   Record.SubItems(8) = ""
                   FinalDiagCode = ""
                End If
                
                sql2 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(FinalDiagCode) & "'"
                rst3.Open sql2, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                If rst3.recordCount = 0 Then
                    rst3.Close
                    Set rst3 = Nothing
                    Record.SubItems(9) = ""
                Else
                    Description2 = Trim(rst3!ICDDescription)
                    Record.SubItems(9) = Trim(Description2)
                    rst3.Close
                    Set rst3 = Nothing
                End If
                
                Current = Current + 1
                lblCurrent = Current
                
            End With
            rst2.MoveNext
            DoEvents
        
            If intExit = 1 Then
                Check = False
                Exit Do
            End If
            Loop
        Loop Until Check = False
    'End If
    intExit = 0
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True

End Sub '********************** List Duration Details *******************************

'********************** List Repudiation Codes Summary **********************
Private Sub RepcodeSummary(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount
    
    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT RepCode, RepDescription,COUNT(CASNumber) AS TotalCases " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY RepCode, Repdescription ORDER BY Repcode "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
               
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!REPCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!REPCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                                        
                If Len(rst2!REPDescription) Then
                   Record.SubItems(1) = rst2!REPDescription
                End If
                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(2) = rst2!TotalCases
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub '********************** List Repudiation Codes Summary **********************

'********************** List Repudiation Category Codes Summary **********************
Private Sub RepCatcodeSummary(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount
    
    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT RepCatCode, RepCatDescription,COUNT(CASNumber) AS TotalCases " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY RepCatCode, RepCatdescription ORDER BY RepCatcode "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
               
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!REPCatCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!REPCatCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                                        
                If Len(rst2!REPCatDescription) Then
                   Record.SubItems(1) = rst2!REPCatDescription
                End If
                    
                If Len(rst2!TotalCases) Then
                   Record.SubItems(2) = rst2!TotalCases
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub '********************** List Repudiation Category Codes Summary **********************
'****************************** Listing Summary  *********************************************
Private Sub ListingSummary(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount
    
    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT CASNumber, BordxRefNo, BordxDate, CASClaimability, PatientName, " & _
          " RepCatCode, RepCatDescription, RepCode, RepDescription, FirstDiag, " & _
          " FinalDiag, HosName, CASDateAdmitted, CASDischargeDate " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
        
    strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend
        strCount = strCount + StrAppend
    Else
        Sql = Sql
        strCount = strCount + StrAppend
    End If
    
    Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    If rstCount!totalrecord = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    Else
               
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!CASNumber) Then
                    Set Record = ListView2.listitems.Add(, , rst2!CASNumber)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!BordxRefNo) Then
                    Record.SubItems(1) = rst2!BordxRefNo
                End If
                
                If Len(rst2!BordxDate) Then
                    Record.SubItems(2) = rst2!BordxDate
                End If
                                
                If Len(rst2!CASClaimability) Then
                    Record.SubItems(3) = rst2!CASClaimability
                End If
                
                If Len(rst2!PatientName) Then
                    Record.SubItems(4) = rst2!PatientName
                End If
                
                If Len(rst2!REPCatCode) Then
                    Record.SubItems(5) = rst2!REPCatCode
                Else
                    'Set Record = ListView2.listitems.Add(, , "Empty")
                    Record.SubItems(5) = "Empty"
                End If
                
                If Len(rst2!REPCatDescription) Then
                    Record.SubItems(6) = rst2!REPCatDescription
                End If
                
                If Len(rst2!REPCode) Then
                    Record.SubItems(7) = rst2!REPCode
                Else
                    Record.SubItems(7) = "Empty"
                End If
                
                If Len(rst2!REPDescription) Then
                    Record.SubItems(8) = rst2!REPDescription
                End If
                
                If Len(rst2!FirstDiag) Then
                    Record.SubItems(9) = rst2!FirstDiag
                End If
                
                If Len(rst2!FinalDiag) Then
                    Record.SubItems(10) = rst2!FinalDiag
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(11) = rst2!HOSName
                End If
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(12) = rst2!CASDateAdmitted
                End If
                
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(13) = rst2!CASDischargeDate
                End If
            
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub ' **********************  Listing Summary  ************************************
' ******************************* Hospital Summary ************************************
Private Sub HospitalSummary(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount
    
    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT HOSName, COUNT(CASNumber) AS TotalCases " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY HOSName ORDER BY HOSName "
     
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
               
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!HOSName) Then
                    Set Record = ListView2.listitems.Add(, , rst2!HOSName)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                                        
                If Len(rst2!TotalCases) Then
                   Record.SubItems(1) = rst2!TotalCases
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub '********************************* Hospital Summary ***************************************
' ******************************* Payor Summary ************************************
Private Sub PayorSummary(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount
    
    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT PAYCode, COUNT(CASNumber) AS TotalCases " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY PAYCode ORDER BY PAYCode "
     
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
               
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!PAYCode) Then
                    Set Record = ListView2.listitems.Add(, , rst2!PAYCode)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                                        
                If Len(rst2!TotalCases) Then
                   Record.SubItems(1) = rst2!TotalCases
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub '******************************** Payor Summary ***************************************
'**************************************** Gender Summary **************************************
Private Sub GenderSummary(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount
    
    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT Sex, COUNT(CASNumber) AS TotalCases " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY Sex ORDER BY Sex "
     
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
               
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!Sex) Then
                    Set Record = ListView2.listitems.Add(, , rst2!Sex)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                                        
                If Len(rst2!TotalCases) Then
                   Record.SubItems(1) = rst2!TotalCases
                End If
                    
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
       Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub ' *************************** Gender Summary ********************************************
'************************************ AgeBand Summary *******************************************
Private Sub AgeBandSummary(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount
    
    lblCurrent = 0
    
    ListView2.listitems.Clear
            
    Sql = " SELECT AgeBand, COUNT(CASNumber) AS TotalCases " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY AgeBand ORDER BY AgeBand "
     
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
               
        Do Until rst2.EOF
        
            'With rst2
                If Len(rst2!AgeBand) Then
                    Set Record = ListView2.listitems.Add(, , rst2!AgeBand)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                                        
                If Len(rst2!TotalCases) Then
                   Record.SubItems(1) = rst2!TotalCases
                End If
                    
            'End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub


Private Sub FirstDiagSummary(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim sql3 As String
Dim strCount
    
    DiagCode = ""
    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT FirstDiag, COUNT(CASNumber) AS Total From VIEW_CLMReport Where 0 = 0 "
    sql2 = " GROUP BY FirstDiag ORDER BY FirstDiag "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
       
        Do Until rst2.EOF
    
            With rst2
                If Len(rst2!FirstDiag) Then
                    Set Record = ListView2.listitems.Add(, , rst2!FirstDiag)
                    DiagCode = rst2!FirstDiag
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                    DiagCode = ""
                End If
                
                sql3 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(DiagCode) & "'"
                rst3.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                If rst3.recordCount = 0 Then
                    rst3.Close
                    Set rst3 = Nothing
                    Record.SubItems(1) = ""
                Else
                    Description = Trim(rst3!ICDDescription)
                    Record.SubItems(1) = Trim(Description)
                    rst3.Close
                    Set rst3 = Nothing
                End If
                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
            
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub FinalDiagSummary(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim sql3 As String
Dim strCount
    
    DiagCode = ""
    lblCurrent = 0
    ListView2.listitems.Clear
            
    Sql = " SELECT FinalDiag, COUNT(CASNumber) AS Total From VIEW_CLMReport Where 0 = 0 "
    sql2 = " GROUP BY FinalDiag ORDER BY FinalDiag "
    
    'strCount = " select count(*) as totalRecord From VIEW_CLMReport where 0 = 0 "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
        'strCount = strCount + strAppend
    Else
        Sql = Sql + sql2
        'strCount = strCount + strAppend
    End If
    
    'Set rstCount = medixmasterconnWS.Execute(strCount)
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
        
    'If rstCount!totalrecord = 0 Then
        'MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        'CmdPrint.Enabled = True
        'CmdClear.Enabled = True
        'cmdExit.Enabled = True
    'Else
       
        Do Until rst2.EOF
    
            With rst2
                If Len(rst2!FinalDiag) Then
                    Set Record = ListView2.listitems.Add(, , rst2!FinalDiag)
                    DiagCode = rst2!FinalDiag
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                    DiagCode = ""
                End If
                
                sql3 = " SELECT ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(DiagCode) & "'"
                rst3.Open sql3, medixmasterconn, adOpenKeyset, adLockOptimistic
                       
                If rst3.recordCount = 0 Then
                    rst3.Close
                    Set rst3 = Nothing
                    Record.SubItems(1) = ""
                Else
                    Description = Trim(rst3!ICDDescription)
                    Record.SubItems(1) = Trim(Description)
                    rst3.Close
                    Set rst3 = Nothing
                End If
                
                If Len(rst2!Total) Then
                    Record.SubItems(2) = rst2!Total
                End If
            
            End With
            lblCurrent = lblCurrent + 1
            rst2.MoveNext
        Loop
    'End If
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub ICSummary(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim strCount
    
    lblCurrent = 0
    
    ListView2.listitems.Clear
            
    'Sql = " SELECT ICNo, CASNumber, ClaimVersion, PatientName, " & _
          " SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt, SUM(CONVERT(Money,CLMPayableByMedix2M)) AS OSAmt " & _
          " From VIEW_CLMReport Where CLMPayableByMedix2M < 0 "
    
    Sql = " SELECT ICNo, CASNumber, ClaimVersion, PatientName, " & _
          " SUM(CLMTotalActual) AS ActualAmt, SUM(CLMTotalApproved) AS ApprovedAmt, SUM(CONVERT(Money,CLMPayableByMedix2M)) AS OSAmt " & _
          " From VIEW_CLMReport Where 0 = 0 "
          
    sql2 = " GROUP BY ICNo, CASNumber, ClaimVersion, PatientName ORDER BY ICNo "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
               
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!ICNo) Then
                    Set Record = ListView2.listitems.Add(, , rst2!ICNo)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                
                If Len(rst2!CASNumber & "-" & rst2!claimversion) Then
                    Record.SubItems(1) = rst2!CASNumber & "-" & rst2!claimversion
                End If
                
                If Len(rst2!PatientName) Then
                   Record.SubItems(2) = Trim(rst2!PatientName)
                End If
                
                If Len(rst2!ActualAmt) Then
                   Record.SubItems(3) = Format(rst2!ActualAmt, "#,##0.00")
                End If
                                        
                If Len(rst2!ApprovedAmt) Then
                   Record.SubItems(4) = Format(rst2!ApprovedAmt, "#,##0.00")
                End If
                
                If Len(rst2!OSAmt) Then
                   Record.SubItems(5) = Format(rst2!OSAmt, "#,##0.00")
                End If
                    
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

'********************** List Bordx No *******************************
Private Sub BordxNoListing(Optional StrAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String

        
    lblCurrent = 0
    
    ListView2.listitems.Clear
            
    Sql = " SELECT Distinct BordxRefNo, BordxDate, BordxNoClaims, SUM(CLMTotalActual) " & _
          " AS totalActual, SUM(CLMTotalApproved) AS TotalApproved, " & _
          " BordxClaimAmt From VIEW_CLMReport2_1 Where BordxRefNo IS NOT NULL "
          
    sql2 = " Group By BordxRefNo, BordxDate, BordxNoClaims, BordxClaimAmt ORDER BY BordxRefNo "
    
    If Len(Trim(StrAppend)) Then
        Sql = Sql + StrAppend + sql2
    Else
        Sql = Sql + sql2
    End If
    
    rst2.Open Sql, medixmasterconnWS, adOpenKeyset, adLockOptimistic
    
        Do Until rst2.EOF
        
            'With rst2
                If Len(rst2!BordxRefNo) Then
                    Set Record = ListView2.listitems.Add(, , rst2!BordxRefNo)
                Else
                    Set Record = ListView2.listitems.Add(, , "Empty")
                End If
                    
                If Len(rst2!BordxDate) Then
                    Record.SubItems(1) = Format(rst2!BordxDate, "dd/mm/yyyy")
                End If
                               
                If Len(rst2!BordxNoClaims) Then
                    Record.SubItems(2) = Trim(rst2!BordxNoClaims)
                End If
                    
                If Len(rst2!TotalActual) Then
                    Record.SubItems(3) = Format(rst2!TotalActual, "#,##0.00")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(4) = Format(rst2!TotalApproved, "#,##0.00")
                End If
                
                If Len(rst2!BordxClaimAmt) Then
                    Record.SubItems(5) = Format(rst2!BordxClaimAmt, "#,##0.00")
                End If
                
            'End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    If ListView2.listitems.count = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
    End If
    rst2.Close
    Set rst2 = Nothing
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Public Sub ExportListViewtoExcel(lvwList As Control)
    Dim vntHeader As Variant
    Dim vntData As Variant
    Dim X As Long
    Dim Y As Long
    Dim intCol As Integer
    Dim lngRow As Long
    
    'Get Counts
    intCol = CInt(lvwList.ColumnHeaders.count - 1)
    lngRow = CLng(lvwList.listitems.count - 1)
    
    ReDim vntHeader(0)
    ReDim vntData(intCol, lngRow)
    
    'Create Header Array


    For X = 0 To intCol
        ReDim Preserve vntHeader(X)
        vntHeader(X) = lvwList.ColumnHeaders(X + 1).Text
    Next
    
    'Create Data Array


    For X = 0 To lngRow
        vntData(0, X) = lvwList.listitems.Item(X + 1).Text
        


        For Y = 1 To intCol
            vntData(Y, X) = lvwList.listitems.Item(X + 1).SubItems(Y)
        Next
    Next
    
    'Create Excel Object
    OpenExcel vntData, vntHeader
    
End Sub


Private Sub ExportRecords(vntData As Variant, vntHeader As Variant, ws As Worksheet)
    Dim lngRow As Long
    Dim intCol As Integer
    Dim varData As Variant
    Dim intStart As Integer
    'Select all Cells and and set the number


    '     format to string
        ws.Cells.Select
        ws.Cells.NumberFormat = "@"
        ws.Cells(1, 1).Select
        lngRow = UBound(vntData, 2) + 2
        intCol = UBound(vntData, 1) + 1
        intStart = 2 'Start from line 2
        'Freeze Row 2
        ws.Rows(2).Select
        ws.Activate
        'ActiveWindow..FreezePanes = True
        'Add Headers


        For X = 1 To intCol
            varData = vntHeader(X - 1)
            ws.Cells(1, X) = CStr(varData)
            ws.Cells(1, X).Font.Bold = True
        Next
        
        'Add Data


        For Y = 1 To intCol


            For X = intStart To lngRow
                varData = vntData(Y - 1, X - 2)


                If IsNull(varData) Then 'Make sure no null values, Excel will choke
                    'Add 1 to Move down a column
                    ws.Cells(X + 1, Y) = ""
                Else
                    ws.Cells(X + 1, Y) = CStr(varData) 'Convert to String to preserve formatting
                End If
            Next
        Next
        'Resize Columns to Fit
        ws.Columns.AutoFit
    End Sub


Private Sub OpenExcel(vntData As Variant, vntHeader As Variant)
    On Error GoTo Err_OpenExcel
    Dim objExcel As Excel.Application
    Dim objWrkSht As Worksheet
    Dim X As Integer
    'Create Excel Object
    Set objExcel = CreateObject("Excel.Application")
    'Add the Workbook
    objExcel.Workbooks.Add
    
    '========== Set Template ===============
        'Set objWorkbook = objExcel.Workbooks.Add(crdPath & "Template\DRCRListing\DRCRListing.XLT")
    '========== Set Template ===============
    
    Set objWrkSht = objExcel.ActiveWorkbook.Sheets(1)
    objExcel.Visible = True
    'Fill the Workbook with data
    ExportRecords vntData, vntHeader, objWrkSht
    objExcel.Interactive = True
    ' Clean up:
    Set objExlSht = Nothing
    Set objExcel = Nothing
Err_OpenExcel:


    Select Case ERR
        Case 0
        Case 439
        MsgBox "You must have Microsoft Excel installed on your PC.", vbCritical, "Application Not Found"
        Case Else
        MsgBox ERR & ": " & Error, vbCritical, "OpenExcel Error"
    End Select
End Sub

Private Sub txtActualAmt1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtActualAmt2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAppAmt1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub txtAppAmt2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub


