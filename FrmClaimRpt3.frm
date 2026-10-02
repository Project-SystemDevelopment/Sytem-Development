VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmClaimRpt3 
   Caption         =   "voon"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   DrawWidth       =   2
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Report Type"
      Height          =   735
      Left            =   0
      TabIndex        =   212
      Top             =   240
      Width           =   11895
      Begin VB.ComboBox cboReportSelection 
         Height          =   315
         Left            =   6960
         TabIndex        =   1
         Top             =   280
         Width           =   4000
      End
      Begin VB.ComboBox cboReportType 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   280
         Width           =   4000
      End
      Begin VB.Label Label82 
         Caption         =   "Report Selection"
         Height          =   255
         Left            =   6480
         TabIndex        =   213
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Frame Frame22 
      Caption         =   "Selection Criteria"
      Height          =   5235
      Left            =   0
      TabIndex        =   121
      Top             =   960
      Width           =   11865
      Begin VB.CheckBox ChkGrpCom2 
         Caption         =   "Check11"
         Height          =   255
         Left            =   7780
         TabIndex        =   214
         Top             =   1365
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
         TabIndex        =   30
         Top             =   1065
         Width           =   3120
      End
      Begin VB.CheckBox Check1 
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
         Left            =   11280
         TabIndex        =   191
         Top             =   1245
         Width           =   255
      End
      Begin VB.CheckBox Check2 
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
         Left            =   11280
         TabIndex        =   190
         Top             =   1485
         Width           =   255
      End
      Begin VB.CheckBox Check3 
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
         Left            =   11280
         TabIndex        =   189
         Top             =   1725
         Width           =   255
      End
      Begin VB.CheckBox Check4 
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
         Left            =   11280
         TabIndex        =   188
         Top             =   1965
         Width           =   255
      End
      Begin VB.CheckBox Check5 
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
         Left            =   11280
         TabIndex        =   187
         Top             =   2205
         Width           =   255
      End
      Begin VB.CheckBox Check6 
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
         Left            =   11280
         TabIndex        =   186
         Top             =   2445
         Width           =   255
      End
      Begin VB.CheckBox Check7 
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
         Left            =   11280
         TabIndex        =   185
         Top             =   2685
         Width           =   255
      End
      Begin VB.CheckBox Check8 
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
         Left            =   11280
         TabIndex        =   184
         Top             =   2925
         Width           =   255
      End
      Begin VB.CheckBox Check9 
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
         Left            =   11280
         TabIndex        =   183
         Top             =   3165
         Width           =   255
      End
      Begin VB.CheckBox Check10 
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
         Left            =   11280
         TabIndex        =   182
         Top             =   3435
         Width           =   255
      End
      Begin VB.CheckBox ChkPayTo 
         Height          =   255
         Left            =   11280
         TabIndex        =   117
         Top             =   300
         Width           =   255
      End
      Begin VB.CheckBox ChkCaseStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2880
         TabIndex        =   91
         Top             =   2180
         Width           =   255
      End
      Begin VB.CheckBox ChkPolStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2880
         TabIndex        =   89
         Top             =   1570
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
         Left            =   2880
         TabIndex        =   90
         Top             =   1880
         Width           =   255
      End
      Begin VB.CheckBox ChkMemType 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2880
         TabIndex        =   87
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox ChkINS 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2880
         TabIndex        =   86
         Top             =   500
         Width           =   255
      End
      Begin VB.CheckBox ChkHLT 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2880
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
         Left            =   2880
         TabIndex        =   93
         Top             =   2700
         Width           =   255
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
         Left            =   2880
         TabIndex        =   88
         Top             =   1305
         Width           =   255
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
         Left            =   2880
         TabIndex        =   92
         Top             =   2440
         Width           =   255
      End
      Begin VB.CheckBox ChkRepCat 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7560
         TabIndex        =   100
         Top             =   510
         Width           =   255
      End
      Begin VB.CheckBox ChkRep 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7560
         TabIndex        =   99
         Top             =   270
         Width           =   255
      End
      Begin VB.CheckBox ChkFinal 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2880
         TabIndex        =   98
         Top             =   4800
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
         Left            =   2880
         TabIndex        =   96
         Top             =   4280
         Width           =   375
      End
      Begin VB.CheckBox ChkBordx 
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
         Left            =   7560
         TabIndex        =   115
         Top             =   4605
         Width           =   255
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
         Left            =   7560
         TabIndex        =   109
         Top             =   3030
         Width           =   255
      End
      Begin VB.CheckBox ChkFirst 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2880
         TabIndex        =   97
         Top             =   4560
         Width           =   255
      End
      Begin VB.CheckBox ChkDoctor 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2880
         TabIndex        =   95
         Top             =   3240
         Width           =   255
      End
      Begin VB.CheckBox ChkHOSGrp 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7560
         TabIndex        =   105
         Top             =   1905
         Width           =   255
      End
      Begin VB.CheckBox ChkHOS 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7560
         TabIndex        =   104
         Top             =   1620
         Width           =   255
      End
      Begin VB.CheckBox ChkPolNo 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2880
         TabIndex        =   94
         Top             =   2980
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpCom 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7560
         TabIndex        =   103
         Top             =   1365
         Width           =   255
      End
      Begin VB.CheckBox ChkPayor 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7560
         TabIndex        =   102
         Top             =   1110
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
         Left            =   7560
         TabIndex        =   112
         Top             =   3750
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
         Left            =   7560
         TabIndex        =   111
         Top             =   3510
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
         Left            =   7560
         TabIndex        =   108
         Top             =   2730
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
         Left            =   7560
         TabIndex        =   110
         Top             =   3270
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
         Left            =   7560
         TabIndex        =   107
         Top             =   2445
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
         Left            =   7560
         TabIndex        =   106
         Top             =   2205
         Width           =   255
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "S - Sihat"
         Top             =   200
         Width           =   1680
      End
      Begin VB.CheckBox ChkClaimNature 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7560
         TabIndex        =   101
         Top             =   780
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
         Left            =   7560
         TabIndex        =   113
         Top             =   4095
         Width           =   255
      End
      Begin VB.CheckBox ChkCaseNo 
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
         Height          =   255
         Left            =   7560
         TabIndex        =   114
         Top             =   4335
         Width           =   255
      End
      Begin VB.CheckBox ChkBordxDate 
         Caption         =   "Check1"
         Height          =   255
         Left            =   7560
         TabIndex        =   116
         Top             =   4875
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
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   1680
      End
      Begin VB.ComboBox CboMemType 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1680
      End
      Begin VB.ComboBox cboSuppType 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1000
         Width           =   1680
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
         Left            =   1080
         TabIndex        =   7
         Top             =   1275
         Width           =   1680
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
         Left            =   1080
         TabIndex        =   8
         Top             =   1560
         Width           =   1680
      End
      Begin VB.ComboBox CboExpired 
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         Top             =   1850
         Width           =   1680
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
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   2130
         Width           =   1680
      End
      Begin VB.ComboBox CboClaimability 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   2400
         Width           =   1680
      End
      Begin VB.TextBox txtAgentCode 
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   12
         Top             =   2690
         Width           =   1680
      End
      Begin VB.TextBox txtPolNo 
         Height          =   285
         Left            =   1080
         MaxLength       =   25
         TabIndex        =   13
         Top             =   2940
         Width           =   1680
      End
      Begin VB.TextBox txtDOCName 
         Height          =   285
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   14
         Top             =   3200
         Width           =   1680
      End
      Begin VB.TextBox TxtPatientName 
         Height          =   285
         Left            =   1080
         MaxLength       =   60
         TabIndex        =   15
         Top             =   3450
         Width           =   1680
      End
      Begin VB.ComboBox CboStayType 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   3700
         Width           =   1680
      End
      Begin VB.ComboBox CboCashType 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   3980
         Width           =   1680
      End
      Begin VB.ComboBox CboGrpCom 
         Height          =   315
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   1350
         Width           =   3120
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
         TabIndex        =   32
         Top             =   1590
         Width           =   3120
      End
      Begin VB.ComboBox cboHOSGrp 
         Height          =   315
         Left            =   4320
         TabIndex        =   33
         Top             =   1875
         Width           =   3120
      End
      Begin VB.TextBox TxtPlan1 
         Height          =   315
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   18
         Top             =   4260
         Width           =   720
      End
      Begin VB.TextBox TxtPlan2 
         Height          =   315
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   19
         Top             =   4260
         Width           =   720
      End
      Begin VB.TextBox txtPreDiag1 
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   20
         Top             =   4550
         Width           =   720
      End
      Begin VB.TextBox txtPreDiag2 
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   21
         Top             =   4550
         Width           =   720
      End
      Begin VB.TextBox txtFinalDiag1 
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   22
         Top             =   4800
         Width           =   720
      End
      Begin VB.TextBox txtFinalDiag2 
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   23
         Top             =   4800
         Width           =   720
      End
      Begin VB.TextBox txtRepudCat1 
         Height          =   285
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   24
         Top             =   270
         Width           =   1200
      End
      Begin VB.TextBox txtRepudCat2 
         Height          =   285
         Left            =   6240
         MaxLength       =   4
         TabIndex        =   25
         Top             =   270
         Width           =   1200
      End
      Begin VB.TextBox txtRepudCode1 
         Height          =   285
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   26
         Top             =   525
         Width           =   1200
      End
      Begin VB.TextBox txtRepudCode2 
         Height          =   285
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   27
         Top             =   525
         Width           =   1200
      End
      Begin VB.TextBox TxtClaimNature1 
         Height          =   285
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   28
         Top             =   780
         Width           =   1200
      End
      Begin VB.TextBox TxtClaimNature2 
         Height          =   285
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   29
         Top             =   780
         Width           =   1200
      End
      Begin MSMask.MaskEdBox TxtDOA1 
         Height          =   315
         Left            =   4320
         TabIndex        =   34
         Top             =   2160
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOA2 
         Height          =   315
         Left            =   6240
         TabIndex        =   35
         Top             =   2160
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
         TabIndex        =   36
         Top             =   2445
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOD2 
         Height          =   315
         Left            =   6240
         TabIndex        =   37
         Top             =   2445
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
         TabIndex        =   38
         Top             =   2730
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
         TabIndex        =   39
         Top             =   2730
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPayEffDate1 
         Height          =   315
         Left            =   4320
         TabIndex        =   40
         Top             =   3015
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPayEffDate2 
         Height          =   315
         Left            =   6240
         TabIndex        =   41
         Top             =   3015
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
         TabIndex        =   42
         Top             =   3270
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
         TabIndex        =   43
         Top             =   3270
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
         TabIndex        =   44
         Top             =   3510
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
         TabIndex        =   45
         Top             =   3510
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
         TabIndex        =   46
         Top             =   3750
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
         TabIndex        =   47
         Top             =   3750
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
         TabIndex        =   48
         Top             =   4035
         Width           =   1200
      End
      Begin VB.TextBox TxtMBMNo2 
         Height          =   285
         Left            =   6240
         MaxLength       =   15
         TabIndex        =   49
         Top             =   4035
         Width           =   1200
      End
      Begin VB.TextBox TxtCaseNo1 
         Height          =   285
         Left            =   4320
         MaxLength       =   15
         TabIndex        =   50
         Top             =   4290
         Width           =   1200
      End
      Begin VB.TextBox TxtCaseNo2 
         Height          =   285
         Left            =   6240
         MaxLength       =   15
         TabIndex        =   51
         Top             =   4290
         Width           =   1200
      End
      Begin VB.TextBox TxtBordxNo1 
         Height          =   285
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   52
         Top             =   4545
         Width           =   1200
      End
      Begin VB.TextBox TxtBordxNo2 
         Height          =   285
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   53
         Top             =   4545
         Width           =   1200
      End
      Begin MSMask.MaskEdBox TxtBordxDate1 
         Height          =   315
         Left            =   4320
         TabIndex        =   54
         Top             =   4800
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtBordxDate2 
         Height          =   315
         Left            =   6240
         TabIndex        =   55
         Top             =   4800
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboPayTo 
         Height          =   315
         Left            =   9240
         TabIndex        =   56
         Top             =   240
         Width           =   1920
      End
      Begin VB.ComboBox CboCasePending 
         Height          =   315
         Left            =   9240
         TabIndex        =   57
         Top             =   540
         Width           =   1920
      End
      Begin VB.ComboBox CboClmPending 
         Height          =   315
         Left            =   9240
         TabIndex        =   58
         Top             =   840
         Width           =   1920
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   9300
         MaxLength       =   4
         TabIndex        =   59
         Top             =   1200
         Width           =   720
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   9300
         MaxLength       =   4
         TabIndex        =   61
         Top             =   1470
         Width           =   720
      End
      Begin VB.TextBox Text5 
         Height          =   300
         Left            =   9300
         MaxLength       =   4
         TabIndex        =   63
         Top             =   1725
         Width           =   720
      End
      Begin VB.TextBox Text7 
         Height          =   300
         Left            =   9300
         MaxLength       =   4
         TabIndex        =   65
         Top             =   1965
         Width           =   720
      End
      Begin VB.TextBox Text9 
         Height          =   300
         Left            =   9300
         MaxLength       =   4
         TabIndex        =   67
         Top             =   2205
         Width           =   720
      End
      Begin VB.TextBox Text11 
         Height          =   300
         Left            =   9300
         MaxLength       =   4
         TabIndex        =   69
         Top             =   2445
         Width           =   720
      End
      Begin VB.TextBox Text13 
         Height          =   300
         Left            =   9300
         MaxLength       =   4
         TabIndex        =   71
         Top             =   2685
         Width           =   720
      End
      Begin VB.TextBox Text15 
         Height          =   300
         Left            =   9300
         MaxLength       =   4
         TabIndex        =   73
         Top             =   2925
         Width           =   720
      End
      Begin VB.TextBox Text17 
         Height          =   300
         Left            =   9300
         MaxLength       =   4
         TabIndex        =   75
         Top             =   3165
         Width           =   720
      End
      Begin VB.TextBox Text19 
         Height          =   300
         Left            =   9300
         MaxLength       =   4
         TabIndex        =   77
         Top             =   3435
         Width           =   720
      End
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   60
         Top             =   1200
         Width           =   720
      End
      Begin VB.TextBox Text4 
         Height          =   300
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   62
         Top             =   1470
         Width           =   720
      End
      Begin VB.TextBox Text6 
         Height          =   300
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   64
         Top             =   1725
         Width           =   720
      End
      Begin VB.TextBox Text8 
         Height          =   300
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   66
         Top             =   1965
         Width           =   720
      End
      Begin VB.TextBox Text10 
         Height          =   300
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   68
         Top             =   2205
         Width           =   720
      End
      Begin VB.TextBox Text12 
         Height          =   300
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   70
         Top             =   2445
         Width           =   720
      End
      Begin VB.TextBox Text14 
         Height          =   300
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   72
         Top             =   2685
         Width           =   720
      End
      Begin VB.TextBox Text16 
         Height          =   300
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   74
         Top             =   2925
         Width           =   720
      End
      Begin VB.TextBox Text18 
         Height          =   300
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   76
         Top             =   3165
         Width           =   720
      End
      Begin VB.TextBox Text20 
         Height          =   300
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   78
         Top             =   3435
         Width           =   720
      End
      Begin VB.Label Label83 
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
         ForeColor       =   &H00FF0000&
         Height          =   135
         Left            =   7845
         TabIndex        =   215
         Top             =   120
         Width           =   180
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   10000
         TabIndex        =   211
         Top             =   1245
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "FolUp - DOD"
         Height          =   255
         Left            =   8040
         TabIndex        =   210
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "PreHosp - DOA"
         Height          =   255
         Left            =   8040
         TabIndex        =   209
         Top             =   1530
         Width           =   1095
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   10000
         TabIndex        =   208
         Top             =   1515
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "DOD - DOA"
         Height          =   255
         Left            =   8040
         TabIndex        =   207
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   10000
         TabIndex        =   206
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label27 
         Caption         =   "DOA - EffDate"
         Height          =   255
         Left            =   8040
         TabIndex        =   205
         Top             =   2070
         Width           =   1095
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   10000
         TabIndex        =   204
         Top             =   2055
         Width           =   375
      End
      Begin VB.Label Label44 
         Caption         =   "DOR - DOA"
         Height          =   255
         Left            =   8040
         TabIndex        =   203
         Top             =   2310
         Width           =   975
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   10000
         TabIndex        =   202
         Top             =   2295
         Width           =   375
      End
      Begin VB.Label Label56 
         Caption         =   "Exp - DOA"
         Height          =   255
         Left            =   8040
         TabIndex        =   201
         Top             =   2550
         Width           =   975
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   10000
         TabIndex        =   200
         Top             =   2535
         Width           =   375
      End
      Begin VB.Label Label63 
         Caption         =   "DPO - DOD"
         Height          =   255
         Left            =   8040
         TabIndex        =   199
         Top             =   2805
         Width           =   975
      End
      Begin VB.Label Label68 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   10000
         TabIndex        =   198
         Top             =   2805
         Width           =   375
      End
      Begin VB.Label Label69 
         Caption         =   "Bordx - DOD"
         Height          =   255
         Left            =   8040
         TabIndex        =   197
         Top             =   3045
         Width           =   975
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   10000
         TabIndex        =   196
         Top             =   3045
         Width           =   375
      End
      Begin VB.Label Label75 
         Caption         =   "InvDOR-InvDate"
         Height          =   255
         Left            =   8040
         TabIndex        =   195
         Top             =   3285
         Width           =   1335
      End
      Begin VB.Label Label76 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   10000
         TabIndex        =   194
         Top             =   3285
         Width           =   375
      End
      Begin VB.Label Label77 
         Caption         =   "Inv Date - DOD"
         Height          =   255
         Left            =   8040
         TabIndex        =   193
         Top             =   3510
         Width           =   1215
      End
      Begin VB.Label Label78 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   240
         Left            =   10005
         TabIndex        =   192
         Top             =   3510
         Width           =   375
      End
      Begin VB.Label Label81 
         Caption         =   "Clm Pending"
         Height          =   255
         Left            =   8040
         TabIndex        =   180
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label80 
         Caption         =   "Cas Pending"
         Height          =   255
         Left            =   8040
         TabIndex        =   179
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label79 
         Caption         =   "Pay To"
         Height          =   255
         Left            =   8040
         TabIndex        =   178
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label62 
         Caption         =   "Patient Name"
         Height          =   255
         Left            =   120
         TabIndex        =   177
         Top             =   3500
         Width           =   1095
      End
      Begin VB.Label Label61 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   176
         Top             =   4620
         Width           =   375
      End
      Begin VB.Label Label60 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   3360
         TabIndex        =   175
         Top             =   4620
         Width           =   855
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
         Left            =   2880
         TabIndex        =   174
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
         Left            =   7560
         TabIndex        =   173
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label55 
         Caption         =   "Agent Code"
         Height          =   255
         Left            =   120
         TabIndex        =   172
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   171
         Top             =   3765
         Width           =   375
      End
      Begin VB.Label Label42 
         Caption         =   "DOB"
         Height          =   255
         Left            =   3360
         TabIndex        =   170
         ToolTipText     =   "Payor Expiry Date"
         Top             =   3750
         Width           =   615
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   169
         Top             =   3510
         Width           =   375
      End
      Begin VB.Label Label40 
         Caption         =   "Date Join"
         Height          =   255
         Left            =   3360
         TabIndex        =   168
         ToolTipText     =   "Payor Expiry Date"
         Top             =   3510
         Width           =   855
      End
      Begin VB.Label Label39 
         Caption         =   "Doctor"
         Height          =   255
         Left            =   120
         TabIndex        =   167
         Top             =   3260
         Width           =   1095
      End
      Begin VB.Label Label37 
         Caption         =   "Hosp. Grp"
         Height          =   255
         Left            =   3360
         TabIndex        =   166
         Top             =   1950
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   3360
         TabIndex        =   165
         Top             =   1710
         Width           =   975
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5670
         TabIndex        =   164
         Top             =   615
         Width           =   375
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5670
         TabIndex        =   163
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   1740
         TabIndex        =   162
         Top             =   4830
         Width           =   375
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   1740
         TabIndex        =   161
         Top             =   4560
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Repud Code"
         Height          =   255
         Left            =   3360
         TabIndex        =   160
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Repud Cat"
         Height          =   255
         Left            =   3360
         TabIndex        =   159
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Final Diag"
         Height          =   255
         Left            =   120
         TabIndex        =   158
         Top             =   4830
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Pre. Diag"
         Height          =   255
         Left            =   120
         TabIndex        =   157
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Grp Comp"
         Height          =   255
         Left            =   3360
         TabIndex        =   156
         Top             =   1410
         Width           =   735
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
         TabIndex        =   155
         Top             =   2200
         Width           =   945
      End
      Begin VB.Label Label53 
         Caption         =   "Claimability"
         Height          =   255
         Left            =   120
         TabIndex        =   154
         Top             =   2500
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
         TabIndex        =   153
         Top             =   1640
         Width           =   1020
      End
      Begin VB.Label Label51 
         Caption         =   "Plan Code"
         Height          =   255
         Left            =   120
         TabIndex        =   152
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   1740
         TabIndex        =   151
         Top             =   4320
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
         TabIndex        =   150
         Top             =   1350
         Width           =   1020
      End
      Begin VB.Label Label38 
         Caption         =   "Payor"
         Height          =   225
         Left            =   3360
         TabIndex        =   149
         Top             =   1110
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
         TabIndex        =   148
         Top             =   550
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
         TabIndex        =   147
         Top             =   250
         Width           =   705
      End
      Begin VB.Label Label7 
         Caption         =   "Eff Date"
         Height          =   255
         Left            =   3360
         TabIndex        =   146
         ToolTipText     =   "Payor Effective Date"
         Top             =   3030
         Width           =   735
      End
      Begin VB.Label Label23 
         Caption         =   "Exp Date"
         Height          =   255
         Left            =   3360
         TabIndex        =   145
         ToolTipText     =   "Payor Expiry Date"
         Top             =   3270
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "DORC"
         Height          =   255
         Left            =   3360
         TabIndex        =   144
         ToolTipText     =   "Member Value Date"
         Top             =   2775
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   143
         Top             =   3030
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   142
         Top             =   3270
         Width           =   375
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   141
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label47 
         Caption         =   "DOA"
         Height          =   255
         Left            =   3360
         TabIndex        =   140
         ToolTipText     =   "Date of Admission"
         Top             =   2265
         Width           =   975
      End
      Begin VB.Label Label48 
         Caption         =   "DOD"
         Height          =   255
         Left            =   3360
         TabIndex        =   139
         ToolTipText     =   "Date of Discharged"
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5670
         TabIndex        =   138
         Top             =   870
         Width           =   375
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   137
         Top             =   2265
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Mem Type"
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   830
         Width           =   855
      End
      Begin VB.Label Label19 
         Height          =   255
         Left            =   2520
         TabIndex        =   135
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label25 
         Caption         =   "Exp Status"
         Height          =   255
         Left            =   120
         TabIndex        =   134
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "Policy No."
         Height          =   255
         Left            =   120
         TabIndex        =   133
         Top             =   3010
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Supp. Type"
         Height          =   255
         Left            =   120
         TabIndex        =   132
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Claim Nature"
         Height          =   255
         Left            =   3360
         TabIndex        =   131
         Top             =   870
         Width           =   975
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   195
         Left            =   5760
         TabIndex        =   130
         Top             =   2775
         Width           =   375
      End
      Begin VB.Label Label24 
         Caption         =   "Member No"
         Height          =   255
         Left            =   3360
         TabIndex        =   129
         Top             =   4065
         Width           =   1095
      End
      Begin VB.Label Label65 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   128
         Top             =   4065
         Width           =   375
      End
      Begin VB.Label Label66 
         Caption         =   "Claim No"
         Height          =   255
         Left            =   3360
         TabIndex        =   127
         Top             =   4350
         Width           =   1095
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   126
         Top             =   4350
         Width           =   375
      End
      Begin VB.Label Label70 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   3360
         TabIndex        =   125
         Top             =   4875
         Width           =   975
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5760
         TabIndex        =   124
         Top             =   4875
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Cash Type"
         Height          =   240
         Left            =   120
         TabIndex        =   123
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Stay Type"
         Height          =   255
         Left            =   120
         TabIndex        =   122
         Top             =   3800
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   118
      Top             =   7320
      Width           =   11895
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   216
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   6240
         TabIndex        =   80
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   375
         Left            =   7920
         TabIndex        =   82
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   375
         Left            =   7080
         TabIndex        =   81
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print to File"
         Height          =   375
         Left            =   8760
         TabIndex        =   83
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   375
         Left            =   10560
         TabIndex        =   85
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9720
         TabIndex        =   84
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   5400
         TabIndex        =   79
         Top             =   160
         Width           =   855
      End
      Begin VB.Label Label64 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1560
         TabIndex        =   119
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView ListView2 
      CausesValidation=   0   'False
      Height          =   1095
      Left            =   0
      TabIndex        =   120
      Top             =   6240
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1931
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
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   1764
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
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Claims Reports 3"
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
      Height          =   220
      Left            =   0
      TabIndex        =   181
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmClaimRpt3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private AA As Integer
Private intExit As Integer
Private GrpCompany As String
Private HOspitals As String
Private m_blnDirection(1 To 15) As Boolean

Private Sub ChkGrpCom2_Click()
Dim strsql As String
Dim PAYGRP As New ADODB.Recordset
Dim con As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String

If Len(Mid(cboPayorCode, 1, 2)) Then
    If ChkGrpCom2.Value = 1 Then
        CboGrpCom.Clear
     '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        'strAppendCase = "3"
        'StrAppend1 = " AND PAYCode = '" & Mid(cboPayorCode, 1, 2) & "'"
        'Set con.ActiveConnection = medixmasterconn
        'con.CommandText = "ClaimReport"
        'con.CommandType = adCmdStoredProc
        'con.CommandTimeout = 300
        'Set prm1 = con.CreateParameter("StrAppend1", adVarChar, adParamInput, 255, StrAppend1)
        'con.Parameters.Append prm1
        'Set prm2 = con.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
        'con.Parameters.Append prm2
        'PAYGRP.CursorLocation = adUseClient
        'PAYGRP.CursorType = adOpenForwardOnly
        'PAYGRP.CacheSize = 100
        'Set PAYGRP = con.Execute
        strsql = "select CoGRPName from CompanyGrp where CoGRPPaycode = '" & Mid(cboPayorCode, 1, 2) & "'"
        PAYGRP.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        
        If PAYGRP.recordCount Then
            Do Until PAYGRP.EOF
               CboGrpCom.AddItem Trim(PAYGRP!CoGRPName)
               PAYGRP.MoveNext
            Loop
        End If
        PAYGRP.Close
        Set PAYGRP = Nothing
     '   medixmasterconn.Close
     '   Set medixmasterconn = Nothing
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

Private Sub LoadPayor(Optional HLTCode As String)
    Dim PAY As New ADODB.Recordset
    Dim strAppend As String
    Dim cmd As New ADODB.Command
    Dim prm1 As ADODB.Parameter
    
    strAppend = strAppend & "And HLTCOde = '" & Mid(Trim(cboHLTCode), 1, 1) & "'"
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayorHLTType"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    Set prm1 = cmd.CreateParameter("strAppend", adVarChar, adParamInput, 2000, strAppend)
    cmd.Parameters.Append prm1
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
    
    cboPayorCode.Clear
    Do Until PAY.EOF
        With PAY
            cboPayorCode.AddItem !PAYCode & " - " & !PAYCompanyName
            PAY.MoveNext
        End With
    Loop
    
    PAY.Close
    Set PAY = Nothing
End Sub

Private Sub CmdBack_Click()
    ListView2.Height = 975
    ListView2.Left = 0
    ListView2.Top = 6240
    ListView2.Width = 11775
    Frame22.Visible = True
    Frame1.Visible = True
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
    Frame1.Visible = False
End Sub

Private Sub ListViewHeader1()

If cboReportSelection = cboReportSelection.List(0) Then
    ListView2.ColumnHeaders(1).Text = "Invoice No"
    ListView2.ColumnHeaders(2).Text = "Claim No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "DOA"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "DOD"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "Inv Date"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Inv DOR"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "Pre Hosp"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Text = "Post Hosp"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Text = "InvDOR - InvDate"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "InvDate - DOD"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "PreHosp - DOA"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "FolUp - DOA"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    
    For AA = 13 To 15
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboReportSelection = cboReportSelection.List(1) Then
    ListView2.ColumnHeaders(1).Text = "Case No"
    ListView2.ColumnHeaders(2).Text = "DOA"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "DOD"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "DPO"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "Eff Date"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Exp Date"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "DOD - DOA"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "DOA - Eff"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "DOR - DOA"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "Exp - DOA"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "DPO - DOD"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    
    For AA = 12 To 15
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboReportSelection = cboReportSelection.List(2) Then
    ListView2.ColumnHeaders(1).Text = "Claim No"
    ListView2.ColumnHeaders(2).Text = "DOA"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "DOD"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Text = "DOR"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "DPO"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Bordx Date"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "Eff Date"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Text = "Exp Date"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Text = "DOD - DOA"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "DOA - Eff"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "DOR - DOA"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "Exp - DOA"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "DPO - DOD"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "Bordx - DOD"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    
    For AA = 15 To 15
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboReportType = cboReportType.List(1) Then
    ListView2.ColumnHeaders(1).Text = "Claim No"
    ListView2.ColumnHeaders(2).Text = "Date"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "HospInvAmt"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "HospExcess"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "HospColExcess"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Pay2Hosp"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "MemInvActual"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "FinalTotActual"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "HospAppr"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "MemAppr"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "TotAppr"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "HospAmtChk"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "TotActualChk"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "MemAmtChk"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "TotActualChk"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
End If
End Sub

Private Sub ListViewHeader2()

If cboReportType = cboReportType.List(1) Then
    ListView2.ColumnHeaders(1).Text = "Claim No"
    ListView2.ColumnHeaders(2).Text = "Date"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Text = "HospInvAmt"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(4).Text = "HospExcess"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(5).Text = "HospColExcess"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(6).Text = "Pay2Hosp"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(7).Text = "MemInvActual"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(8).Text = "FinalTotActual"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(9).Text = "HospAppr"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(10).Text = "MemAppr"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(11).Text = "TotAppr"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(12).Text = "HospAmtChk"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(13).Text = "TotActualChk"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(14).Text = "MemAmtChk"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(15).Text = "TotActualChk"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
End If
End Sub

Private Sub ListViewHeader3()

If cboReportType = cboReportType.List(2) Then
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
    
    For AA = 7 To 15
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
End If
End Sub

Private Sub Form_Load()
    intExit = 0
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Call ComboListing
 '   medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    cboReportType.ListIndex = 0
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

Private Sub cmdPrint_Click()
    Screen.MousePointer = vbHourglass
        'Call ExportToExcelAnalysis1(ListView2)
    If ListView2.ListItems.Count <> 0 Then
        Call ExportListViewtoExcel(ListView2)
    Else
        MsgBox "Report cannot be printed without any record.", vbInformation
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSearch_Click()
On Error GoTo errRun
Dim strAppend As String
    strAppend = ""
    CmdPrint.Enabled = False
    CmdClear.Enabled = False
    cmdExit.Enabled = False
    cmdSearch.Enabled = False
    Screen.MousePointer = vbHourglass
 '   medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
 '   medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
    Call SearchRecords
 '   medixmasterconn.Close
'    Set medixmasterconn = Nothing
 '   medixmasterconnWS.Close
 '   Set medixmasterconnWS = Nothing
    Screen.MousePointer = vbDefault
    cmdSearch.Enabled = True
    Exit Sub

errRun:
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    ListView2.ListItems.Clear
    lblCurrent = ""
    CmdPrint.Enabled = True
    CmdClear.Enabled = True
    cmdExit.Enabled = True
    cmdSearch.Enabled = True
    medixmasterconn.Close
    Set medixmasterconn = Nothing
    medixmasterconnWS.Close
    Set medixmasterconnWS = Nothing
End Sub

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lblCurrent = ""
    Current = 0
    HospInvAmt = 0
    HospAppr = 0
    MemAppr = 0
    TotalAmt = 0
    HospAmtChk = 0
    MemAmtChk = 0
    TotActualChk = 0
    HospInvAmt = 0
    HospColExcess = 0
    PaytoHosp = 0
    MemInvActual = 0
    FinalTotActual = 0
    Check1 = 0
    Check2 = 0
    Check3 = 0
    Check4 = 0
    lblTotalRecord = ""
    ListView2.ListItems.Clear
End Sub
' =================== Command Button Click / Change ==========================


'============== ComboBox Listing =================================
Private Sub ComboListing()
Dim PAY As New ADODB.Recordset
Dim INS As New ADODB.Recordset
Dim PolStatus As New ADODB.Recordset
Dim HOS As New ADODB.Recordset
Dim HosGrp As New ADODB.Recordset
Dim cmd As New ADODB.Command
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchINS"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    INS.CursorLocation = adUseClient
    INS.CursorType = adOpenForwardOnly
    INS.CacheSize = 100
    Set INS = cmd.Execute
    
    Do Until INS.EOF
        With INS
            CboInsType.AddItem !INSCode & " - " & !INSDescription
            INS.MoveNext
        End With
    Loop
    INS.Close
    Set INS = Nothing
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPolStatus"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    PolStatus.CursorLocation = adUseClient
    PolStatus.CursorType = adOpenForwardOnly
    PolStatus.CacheSize = 100
    Set PolStatus = cmd.Execute
    
    Do Until PolStatus.EOF
        With PolStatus
            CboStatus.AddItem !pstatuscode & " - " & !pstatusDesc
            PolStatus.MoveNext
        End With
    Loop
    PolStatus.Close
    Set PolStatus = Nothing
    
        
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHOS"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    HOS.CursorLocation = adUseClient
    HOS.CursorType = adOpenForwardOnly
    HOS.CacheSize = 100
    Set HOS = cmd.Execute
    
    Do Until HOS.EOF
        With HOS
            cboHOSName.AddItem !HosName
            HOS.MoveNext
        End With
    Loop
    HOS.Close
    Set HOS = Nothing
    
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchHospGrp"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    HosGrp.CursorLocation = adUseClient
    HosGrp.CursorType = adOpenForwardOnly
    HosGrp.CacheSize = 100
    Set HosGrp = cmd.Execute
    
    Do Until HosGrp.EOF
        With HosGrp
            cboHOSGrp.AddItem !HOSGRPName
            HosGrp.MoveNext
        End With
    Loop
    HosGrp.Close
    Set HosGrp = Nothing
        
    Set cmd.ActiveConnection = medixmasterconn
    cmd.CommandText = "SearchPayor"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 300
    PAY.CursorLocation = adUseClient
    PAY.CursorType = adOpenForwardOnly
    PAY.CacheSize = 100
    Set PAY = cmd.Execute
        
    Do Until PAY.EOF
        With PAY
            cboPayorCode.AddItem !PAYCode & " - " & !PAYCompanyName
            PAY.MoveNext
        End With
    Loop
    PAY.Close
    Set PAY = Nothing
    
    cboHLTCode.AddItem "S" & " - " & "Sihat"
    cboHLTCode.AddItem "N" & " - " & "Non-Sihat"
    
    CboMemType.AddItem "Both"
    CboMemType.AddItem "P" & " - " & "Principal"
    CboMemType.AddItem "S" & " - " & "Supplementary"
    CboMemType.Text = "Both"
    
    
    cboDataStatus.AddItem "A - Active"
    cboDataStatus.AddItem "D - Delete"
    
    CboPayTo.AddItem "H - Hospital"
    CboPayTo.AddItem "M - Member"
    
    With CboClaimability
        .AddItem "A - Accepted"
        .AddItem "R - Rejected"
    End With
    
    With CboCaseStatus
        .AddItem "0 - Open"
        .AddItem "1 - Open & Close"
        .AddItem "2 - Delete"
        .AddItem "3 - All"
        .Text = "1 - Open & Close"
    End With
    
    CboExpired.AddItem "N - No"
    CboExpired.AddItem "Y - Yes"
    
    With CboStayType
        .AddItem "I - IN-PATIENT"
        .AddItem "O - OUT-PATIENT"
    End With
    
    With CboCashType
        .AddItem "C - CASHLESS"
        .AddItem "R - REIMBURSEMENT"
    End With
    
    CboCasePending.AddItem "N - No"
    CboCasePending.AddItem "Y - Yes"
    
    CboClmPending.AddItem "N - No"
    CboClmPending.AddItem "Y - Yes"
           
    cboSuppType.AddItem "Both"
    cboSuppType.AddItem "C" & " - " & "Combined Limit"
    cboSuppType.AddItem "S" & " - " & "Separated Limit"
    cboSuppType.Text = "Both"
    
    cboReportType.AddItem "Date"
    cboReportType.AddItem "CLM Value Control"
    cboReportType.AddItem "Bordx"
End Sub
'============== ComboBox Listing =================================

'=========== Option Button, Combo Box Change/KeyPress ===================

Private Sub cboHLTCode_Change()
    If InStr(cboHLTCode.Text, "null") Then
       cboHLTCode.Enabled = False
    End If
    
    If Len(Trim(cboHLTCode)) Then
    'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        LoadPayor (cboHLTCode)
   ' medixmasterconn.Close
   ' Set medixmasterconn = Nothing
    End If
End Sub

Private Sub cboHLTCode_Click()
    If Len(Trim(cboHLTCode)) Then
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        LoadPayor (cboHLTCode)
        'medixmasterconn.Close
        'Set medixmasterconn = Nothing
    End If
End Sub

Private Sub cboPayorCode_Click()
Dim strsql As String
Dim PAYGRP As New ADODB.Recordset
Dim con As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String

If Len(Mid(cboPayorCode, 1, 2)) Then
    If ChkGrpCom2.Value = 1 Then
        CboGrpCom.Clear
        'medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
        'strAppendCase = "3"
        'StrAppend1 = " AND PAYCode = '" & Mid(cboPayorCode, 1, 2) & "'"
        'Set con.ActiveConnection = medixmasterconn
        'con.CommandText = "ClaimReport"
        'con.CommandType = adCmdStoredProc
        'con.CommandTimeout = 300
        'Set prm1 = con.CreateParameter("StrAppend1", adVarChar, adParamInput, 255, StrAppend1)
        'con.Parameters.Append prm1
        'Set prm2 = con.CreateParameter("StrAppendCase", adVarChar, adParamInput, 255, strAppendCase)
        'con.Parameters.Append prm2
        'PAYGRP.CursorLocation = adUseClient
        'PAYGRP.CursorType = adOpenForwardOnly
        'PAYGRP.CacheSize = 100
        'Set PAYGRP = con.Execute
        strsql = "select CoGRPName from CompanyGrp where CoGRPPaycode = '" & Mid(cboPayorCode, 1, 2) & "'"
        PAYGRP.Open strsql, medixmasterconn, adOpenForwardOnly, adLockReadOnly
        
        If PAYGRP.recordCount Then
            Do Until PAYGRP.EOF
               CboGrpCom.AddItem Trim(PAYGRP!CoGRPName)
               PAYGRP.MoveNext
            Loop
        End If
        PAYGRP.Close
        Set PAYGRP = Nothing
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
    End If
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

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
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

Private Sub CboPayTo_Change()

    If CboPayTo.Text = "" Then
        If InStr(CboPayTo.Text, "null") Then
           CboPayTo.Enabled = False
        End If
    End If
End Sub

Private Sub cboPayTo_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(CboPayTo.Text, CboPayTo.SelStart) + Chr(KeyAscii) + Mid(CboPayTo.Text, CboPayTo.SelStart + CboPayTo.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboPayTo.ListCount - 1
 
        If NewText = LCase(Left(CboPayTo.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboPayTo.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboPayTo.Text = ValidValue
                CboPayTo.SelStart = 0
                CboPayTo.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboReportType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboReportType.Text, cboReportType.SelStart) + Chr(KeyAscii) + Mid(cboReportType.Text, cboReportType.SelStart + cboReportType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboReportType.ListCount - 1
 
        If NewText = LCase(Left(cboReportType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboReportType.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboReportType.Text = ValidValue
                cboReportType.SelStart = 0
                cboReportType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboReportSelection_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer

    If KeyAscii >= 32 And KeyAscii <> 127 Then
        NewText = LCase(Left(cboReportSelection.Text, cboReportSelection.SelStart) + Chr(KeyAscii) + Mid(cboReportSelection.Text, cboReportSelection.SelStart + cboReportSelection.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboReportSelection.ListCount - 1
 
        If NewText = LCase(Left(cboReportSelection.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboReportSelection.List(i)
        End If
    Next
        
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboReportSelection.Text = ValidValue
                cboReportSelection.SelStart = 0
                cboReportSelection.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboReportType_Click()

    cboReportSelection.Clear
    cboReportSelection.Visible = True
    ListView2.ListItems.Clear
    If cboReportType.ListIndex = 0 Then
        cboReportSelection.AddItem "By Invoice No"
        cboReportSelection.AddItem "By Case No"
        cboReportSelection.AddItem "By Claim No"
        cboReportSelection = cboReportSelection.List(0)
        cboReportSelection.ListIndex = 0
        lblCurrent = ""
        ListView2.ListItems.Clear
        Call ListViewHeader1
    ElseIf cboReportType.ListIndex = 1 Then
        cboReportSelection.Visible = False
        lblCurrent = ""
        ListView2.ListItems.Clear
        Call ListViewHeader2
    ElseIf cboReportType.ListIndex = 2 Then
        cboReportSelection.Visible = False
        lblCurrent = ""
        ListView2.ListItems.Clear
        Call ListViewHeader3
    End If
    
End Sub

Private Sub cboReportSelection_Click()
    
    ListView2.ListItems.Clear
    If cboReportSelection = cboReportSelection.List(0) Then
        lblCurrent = ""
        ListView2.ListItems.Clear
        CboClmPending.Enabled = True
        Call ListViewHeader1
    ElseIf cboReportSelection = cboReportSelection.List(1) Then
        lblCurrent = ""
        ListView2.ListItems.Clear
        CboClmPending.Text = ""
        CboClmPending.Enabled = False
        Call ListViewHeader1
    ElseIf cboReportSelection = cboReportSelection.List(2) Then
        lblCurrent = ""
        ListView2.ListItems.Clear
        CboClmPending.Enabled = True
        Call ListViewHeader1
    End If
End Sub
Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    
    If cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(0) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        End Select
    
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(1) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
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
        End Select
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(2) Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(4)
        Case 5
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(8)
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
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        End Select
    
    ElseIf cboReportType = cboReportType.List(1) Then
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
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
        End Select
    
    ElseIf cboReportType = cboReportType.List(2) Then
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
    End If
End Sub

'=========== Combo Box Change/KeyPress ===================

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumber(KeyAscii)
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

Private Sub txtBordxDate1_LostFocus()
    If IsDate(TxtBordxDate1) Then
    Else
        If TxtBordxDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtBordxDate1.SetFocus
        End If
    End If
End Sub

Private Sub txtBordxDate2_LostFocus()
    If IsDate(TxtBordxDate2) Then
    Else
        If TxtBordxDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtBordxDate2.SetFocus
        End If
    End If
End Sub


'=========== Check Date Format ===========================

Private Function SearchRecords()
Dim Sql As String
Dim strAppend As String

    strAppend = ""
    GrpCompany = ""
    HOspitals = ""

    If Len(cboReportType) = False Then
        MsgBox "Please select a Report Type", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        cmdSearch.Enabled = True
        Exit Function
    End If
    
    If cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(2) Then
        If Len(Trim(TxtBordxNo1)) Then
            strAppend = strAppend & " AND BordxRefNo >= '" & Trim(TxtBordxNo1) & "'"
        End If
    
        If Len(Trim(TxtBordxNo2)) Then
            strAppend = strAppend & " AND BordxRefNo <= '" & Trim(TxtBordxNo2) & "'"
        End If
        
        If (TxtBordxDate1) <> "__/__/____" And IsDate(TxtBordxDate1) = True _
            And (TxtBordxDate2) <> "__/__/____" And IsDate(TxtBordxDate2) = True Then
            Sql = ""
            Sql = " AND BordxDate >= '" & Format(Trim(TxtBordxDate1), "mm/dd/yyyy") & _
                  "' And BordxDate <= '" & Format(Trim(TxtBordxDate2), "mm/dd/yyyy") & "'"
            strAppend = strAppend + Sql
        End If
    
        If Len(Trim(TxtBordxDate1)) And IsDate(TxtBordxDate1) = True Then
            strAppend = strAppend & " AND BordxDate >= '" & Format(TxtBordxDate1, "mm/dd/yyyy") & "'"
        End If
        
        If Len(Trim(TxtBordxDate2)) And IsDate(TxtBordxDate2) = True Then
            strAppend = strAppend & " AND BordxDate <= '" & Format(TxtBordxDate2, "mm/dd/yyyy") & "'"
        End If
        
        If ChkBordx.Value = 1 Then
            strAppend = strAppend & " AND (BordxRefNo Is null OR BordxRefNo = '')"
        End If
        
        If ChkBordxDate.Value = 1 Then
            strAppend = strAppend & " AND (BordxDate Is null OR BordxDate = '')"
        End If
        
        If Len(Trim(Text15)) Then
            strAppend = strAppend & " AND BORDOD >= '" & Trim(Text15) & "'"
        End If
    
        If Len(Trim(Text16)) Then
            strAppend = strAppend & " AND BORDOD <= '" & Trim(Text16) & "'"
        End If
        
        If Check8.Value = 1 Then
            strAppend = strAppend & " AND (BORDOD Is null OR BORDOD = '')"
        End If
        
        If Mid(CboClmPending, 1, 1) = "N" Then
            strAppend = strAppend + " AND (CLMPending = 'N' OR CLMPending = '' OR CLMPending IS NULL)"
        ElseIf Mid(CboClmPending, 1, 1) = "Y" Then
            strAppend = strAppend + " AND CLMPending = 'Y'"
        End If
    End If
    
    If cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(0) Then
    
        If Len(Trim(CboPayTo)) Then
            strAppend = strAppend + " AND INVPayTo = '" & Trim(Mid(CboPayTo, 1, IIf(InStr(1, CboPayTo, "-") = 0, 4, InStr(1, CboPayTo, "-") - 1))) & "'"
        End If
    
        If ChkPayTo.Value = 1 Then
            Sql = ""
            Sql = "AND (INVPayTo Is null OR INVPayTo = '')" & ""
            strAppend = strAppend + Sql
        End If
        
        If Len(Trim(Text1)) Then
            strAppend = strAppend & " AND FOLDOD >= '" & Trim(Text1) & "'"
        End If
    
        If Len(Trim(Text2)) Then
            strAppend = strAppend & " AND FOLDOD <= '" & Trim(Text2) & "'"
        End If
        
        If Check1.Value = 1 Then
            strAppend = strAppend & " AND (FOLDOD Is null OR FOLDOD = '')"
        End If
        
        If Len(Trim(Text3)) Then
            strAppend = strAppend & " AND PREDOA >= '" & Trim(Text3) & "'"
        End If
    
        If Len(Trim(Text4)) Then
            strAppend = strAppend & " AND PREDOA <= '" & Trim(Text4) & "'"
        End If
        
        If Check2.Value = 1 Then
            strAppend = strAppend & " AND (PREDOA Is null OR PREDOA = '')"
        End If
        
        If Len(Trim(Text17)) Then
            strAppend = strAppend & " AND INVDORINV >= '" & Trim(Text17) & "'"
        End If
    
        If Len(Trim(Text18)) Then
            strAppend = strAppend & " AND INVDORINV <= '" & Trim(Text18) & "'"
        End If
        
        If Check9.Value = 1 Then
            strAppend = strAppend & " AND (INVDORINV Is null OR INVDORINV = '')"
        End If
        
        If Len(Trim(Text19)) Then
            strAppend = strAppend & " AND INVDOD >= '" & Trim(Text19) & "'"
        End If
    
        If Len(Trim(Text20)) Then
            strAppend = strAppend & " AND INVDOD <= '" & Trim(Text20) & "'"
        End If
        
        If Check10.Value = 1 Then
            strAppend = strAppend & " AND (INVDOD Is null OR INVDOD = '')"
        End If
        
        If Mid(CboClmPending, 1, 1) = "N" Then
            strAppend = strAppend + " AND (CLMPending = 'N' OR CLMPending = '' OR CLMPending IS NULL)"
        ElseIf Mid(CboClmPending, 1, 1) = "Y" Then
            strAppend = strAppend + " AND CLMPending = 'Y'"
        End If
        
    End If
    
    If Len(Mid(cboHLTCode, 1, 1)) Then
        strAppend = strAppend + " AND HLTCode ='" & Trim(Mid(cboHLTCode, 1, 1)) & "'"
    End If

    If Len(Mid(CboInsType, 1, 1)) Then
        strAppend = strAppend & " AND INSCode = '" & Trim(Mid(CboInsType, 1, 1)) & "'"
    End If
       
    If Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(cboSuppType) = "Both" Then
        strAppend = strAppend & " AND (MBMMemberType = 'P' Or MBMMemberType = 'Q') And MBMPatientCovID = '00' "
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
        strAppend = strAppend & " AND MBMMemberType = 'P'  And MBMPatientCovID = '00'"
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "P" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
        strAppend = strAppend & " AND MBMMemberType = 'Q' And MBMPatientCovID = '00'"
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(cboSuppType) = "Both" Then
        strAppend = strAppend & " AND (MBMCMemberType = 'S' OR MBMCMemberType = 'C')"
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
        strAppend = strAppend & " AND MBMCMemberType = 'C'"
    ElseIf Trim(Mid(CboMemType, 1, 1)) = "S" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
        strAppend = strAppend & " AND MBMCMemberType = 'S'"
    ElseIf Trim(CboMemType) = "Both" And Trim(Mid(cboSuppType, 1, 1)) = "C" Then
        strAppend = strAppend & " AND (MBMMemberType = 'P' OR MBMCMemberType = 'C')"
    ElseIf Trim(CboMemType) = "Both" And Trim(Mid(cboSuppType, 1, 1)) = "S" Then
        strAppend = strAppend & " AND (MBMMemberType = 'Q' OR MBMCMemberType = 'S')"
    End If
    
    If Len(Mid(cboDataStatus, 1, 1)) Then
        strAppend = strAppend & " AND MBMDataStatus = '" & Trim(Mid(cboDataStatus, 1, 1)) & "'"
    End If
    
    If Len(Mid(CboClaimability, 1, 1)) Then
        strAppend = strAppend & " AND CASClaimability = '" & Trim(Mid(CboClaimability, 1, 1)) & "'"
    End If

    If Len(Mid(CboStatus, 1, 1)) Then
        strAppend = strAppend & " AND MBMStatus = '" & Trim(Mid(CboStatus, 1, 1)) & "'"
    End If

    If Len(Mid(CboCaseStatus, 1, 1)) = 0 Then
        strAppend = strAppend & " AND CASStatus = 'O'"
    ElseIf Len(Mid(CboCaseStatus, 1, 1)) = 1 Then
        strAppend = strAppend & " AND (CASStatus = 'O' OR CASStatus = 'C')"
    ElseIf Len(Mid(CboCaseStatus, 1, 1)) = 2 Then
        strAppend = strAppend & " AND CASStatus = 'D'"
    End If
    
    If Len(Trim(Mid(CboExpired, 1, 1))) Then
          strAppend = strAppend & "AND MBMExpiredStatus = '" & Trim(Mid(CboExpired, 1, 1)) & "' "
    End If
    
    If Len(Trim(txtAgentCode)) Then
       strAppend = strAppend + " And  MBMAgentCode like '%" & ChkStr(txtAgentCode) & "%' "
    End If
    
    If Len(Trim(txtPolNo)) Then
       strAppend = strAppend + " And  MBMPolicyNo like '%" & ChkStr(txtPolNo) & "%' "
    End If
    
    If Len(Mid(cboPayorCode, 1, 2)) Then
        strAppend = strAppend & " AND PAYCode = '" & Trim(Mid(cboPayorCode, 1, 2)) & "'"
    Else
        strAppend = strAppend & " AND Substring(PAYCode,1,1) <> 'Z'"
    End If
           
    If Len(Trim(CboGrpCom)) Then
         GrpCompany = Replace(CboGrpCom, "'", "''")
         strAppend = strAppend + " AND MBMGroupCompany = '" & Trim(GrpCompany) & "'"
      End If
    
    If Len(Trim(cboHOSName)) Then
        HOspitals = Replace(cboHOSName, "'", "''")
        strAppend = strAppend + " AND HOSName = '" & Trim(HOspitals) & "'"
    End If
    
    If Len(Trim(cboHOSGrp)) Then
        strAppend = strAppend + " And HOSGRPName = '" & Trim(cboHOSGrp) & "'"
    End If
    
    If Len(Trim(txtDOCName)) Then
       strAppend = strAppend + " And  AttendDoctor like '%" & ChkStr(txtDOCName) & "%' "
    End If
    
    If Len(Trim(TxtPatientName)) Then
       strAppend = strAppend + " And  PatientName like '%" & ChkStr(TxtPatientName) & "%' "
    End If
    
    If Len(Trim(Mid(CboStayType, 1, 1))) Then
        strAppend = strAppend + " And CASStayType = '" & Trim(Mid(CboStayType, 1, 1)) & "'"
    End If
    
    If Len(Trim(Mid(CboCashType, 1, 1))) Then
        strAppend = strAppend + " And CASPayType = '" & Trim(Mid(CboCashType, 1, 1)) & "'"
    End If
    
    If Len(Trim(TxtDOA1)) > 0 And IsDate(TxtDOA1) = True Then
        strAppend = strAppend + " And CasDateAdmitted >= '" & Format(TxtDOA1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOA2)) > 0 And IsDate(TxtDOA2) = True Then
        strAppend = strAppend + " And CasDateAdmitted <= '" & Format(TxtDOA2, "mm/dd/yyyy") & "'"
    End If
       
    If (txtDOD1) <> "__/__/____" And IsDate(txtDOD1) = True _
        And (txtDOD2) <> "__/__/____" And IsDate(txtDOD2) = True Then
        Sql = ""
        Sql = " AND CasDischargeDate >= '" & Format(Trim(txtDOD1), "mm/dd/yyyy") & _
              "' And CasDischargeDate <= '" & Format(Trim(txtDOD2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + Sql
    End If
    
    If Len(Trim(txtDOD1)) > 0 And IsDate(txtDOD1) = True Then
        strAppend = strAppend + " And CasDischargeDate >= '" & Format(txtDOD1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOD2)) > 0 And IsDate(txtDOD2) = True Then
        strAppend = strAppend + " And CasDischargeDate <= '" & Format(txtDOD2, "mm/dd/yyyy") & "'"
    End If
    
    If (TxtDOR1) <> "__/__/____" And IsDate(TxtDOR1) = True _
        And (TxtDOR2) <> "__/__/____" And IsDate(TxtDOR2) = True Then
        Sql = ""
        Sql = " AND CaseRegDate >= '" & Format(Trim(TxtDOR1), "mm/dd/yyyy") & _
              "' And CaseRegDate <= '" & Format(Trim(TxtDOR2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + Sql
    End If
    
    If Len(Trim(TxtDOR1)) > 0 And IsDate(TxtDOR1) = True Then
        strAppend = strAppend + " And CaseRegDate >= '" & Format(TxtDOR1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtDOR2)) > 0 And IsDate(TxtDOR2) = True Then
        strAppend = strAppend + " And CaseRegDate <= '" & Format(TxtDOR2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtPreDiag1)) Then
        strAppend = strAppend + " And FirstDiag >= '" & Trim(txtPreDiag1) & "'"
    End If
    
    If Len(Trim(txtPreDiag2)) Then
        strAppend = strAppend + " And FirstDiag <= '" & Trim(txtPreDiag2) & "'"
    End If
    
    If Len(Trim(txtFinalDiag1)) Then
        strAppend = strAppend + " And FinalDiag >= '" & Trim(txtFinalDiag1) & "'"
    End If
    
    If Len(Trim(txtFinalDiag2)) Then
        strAppend = strAppend + " And FinalDiag <= '" & Trim(txtFinalDiag2) & "'"
    End If
    
    If Len(Trim(txtRepudCat1)) Then
        strAppend = strAppend + " And CASREPCode1 >= '" & Trim(txtRepudCat1) & "'"
    End If
    
    If Len(Trim(txtRepudCat2)) Then
        strAppend = strAppend + " And CASREPCode1 <= '" & Trim(txtRepudCat2) & "'"
    End If
    
    If Len(Trim(txtRepudCode1)) Then
        strAppend = strAppend + " And RepCatCode >= '" & Trim(txtRepudCode1) & "'"
    End If
    
    If Len(Trim(txtRepudCode2)) Then
        strAppend = strAppend + " And RepCatCode <= '" & Trim(txtRepudCode2) & "'"
    End If
    
    If Len(Trim(TxtClaimNature1)) Then
        strAppend = strAppend & " AND ClaimNatureCode >= '" & Trim(TxtClaimNature1) & "'"
    End If
    
    If Len(Trim(TxtClaimNature2)) Then
        strAppend = strAppend & " AND ClaimNatureCode <= '" & Trim(TxtClaimNature2) & "'"
    End If
    
    If Len(Trim(TxtPlan1)) Then
        strAppend = strAppend & " AND PLNCode >= '" & Trim(TxtPlan1) & "'"
    End If
    
    If Len(Trim(TxtPlan2)) Then
        strAppend = strAppend & " AND PLNCode <= '" & Trim(TxtPlan2) & "'"
    End If
    
    If Len(Trim(txtPolicyNo1)) Then
        strAppend = strAppend & " AND MBMPolicyNo >= '" & Trim(txtPolicyNo1) & "'"
    End If
    
    If Len(Trim(txtPolicyNo2)) Then
        strAppend = strAppend & " AND MBMPolicyNo <= '" & Trim(txtPolicyNo2) & "'"
    End If
    
    If Len(Trim(TxtPayEffDate1)) > 0 And IsDate(TxtPayEffDate1) = True Then
        strAppend = strAppend + " And MBMPayorEffDate >= '" & Format(TxtPayEffDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtPayEffDate2)) > 0 And IsDate(TxtPayEffDate2) = True Then
        strAppend = strAppend + " And MBMPayorEffDate <= '" & Format(TxtPayEffDate2, "mm/dd/yyyy") & "'"
    End If
       
    If (TxtPayEffDate1) <> "__/__/____" And IsDate(TxtPayEffDate1) = True _
        And (TxtPayEffDate2) <> "__/__/____" And IsDate(TxtPayEffDate2) = True Then
        Sql = ""
        Sql = " AND MBMPayorEffDate >= '" & Format(Trim(TxtPayEffDate1), "mm/dd/yyyy") & _
              "' And MBMPayorEffDate <= '" & Format(Trim(TxtPayEffDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + Sql
    End If
    
    If (TxtPayExpDate1) <> "__/__/____" And IsDate(TxtPayExpDate1) = True _
        And (TxtPayExpDate2) <> "__/__/____" And IsDate(TxtPayExpDate2) = True Then
        Sql = ""
        Sql = " AND MBMPayorExpDate >= '" & Format(Trim(TxtPayExpDate1), "mm/dd/yyyy") & _
              "' And MBMPayorExpDate <= '" & Format(Trim(TxtPayExpDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + Sql
    End If
    
    If Len(Trim(TxtPayExpDate1)) And IsDate(TxtPayExpDate1) = True Then
        strAppend = strAppend & " AND MBMPayorExpDate >= '" & Format(TxtPayExpDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtPayExpDate2)) And IsDate(TxtPayExpDate2) = True Then
        strAppend = strAppend & " AND MBMPayorExpDate <= '" & Format(TxtPayExpDate2, "mm/dd/yyyy") & "'"
    End If
    
    If (txtDateJoin1) <> "__/__/____" And IsDate(txtDateJoin1) = True _
        And (txtDateJoin2) <> "__/__/____" And IsDate(txtDateJoin2) = True Then
        Sql = ""
        Sql = " AND MBMFirstJoinedDate >= '" & Format(Trim(txtExpDate1), "mm/dd/yyyy") & _
              "' And MBMFirstJoinedDate <= '" & Format(Trim(txtExpDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + Sql
    End If
    
    If Len(Trim(txtDateJoin1)) And IsDate(txtDateJoin1) = True Then
        strAppend = strAppend & " AND MBMFirstJoinedDate >= '" & Format(txtDateJoin1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDateJoin2)) And IsDate(txtDateJoin2) = True Then
        strAppend = strAppend & " AND MBMFirstJoinedDate <= '" & Format(txtDateJoin2, "mm/dd/yyyy") & "'"
    End If
    
    If (txtDOB1) <> "__/__/____" And IsDate(txtDOB1) = True _
        And (txtDOB2) <> "__/__/____" And IsDate(txtDOB2) = True Then
        Sql = ""
        Sql = " AND DOB >= '" & Format(Trim(txtDOB1), "mm/dd/yyyy") & _
              "' And DOB <= '" & Format(Trim(txtDOB2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + Sql
    End If
    
    If Len(Trim(txtDOB1)) And IsDate(txtDOB1) = True Then
        strAppend = strAppend & " AND DOB >= '" & Format(txtDOB1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOB2)) And IsDate(txtDOB2) = True Then
        strAppend = strAppend & " AND DOB <= '" & Format(txtDOB2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(TxtCaseNo1)) Then
       strAppend = strAppend + " AND CASNumber >= '" & RemStr(Trim(TxtCaseNo1)) & "'"
    End If
    
    If Len(Trim(TxtCaseNo2)) Then
       strAppend = strAppend + " AND CASNumber <= '" & RemStr(Trim(TxtCaseNo2)) & "'"
    End If
    
    If Len(Trim(TxtMBMNo1)) Then
       strAppend = strAppend + " AND MBMNumber >= '" & RemStr(Trim(TxtMBMNo1)) & "'"
    End If
    
    If Len(Trim(TxtMBMNo2)) Then
       strAppend = strAppend + " AND MBMNumber <= '" & RemStr(Trim(TxtMBMNo2)) & "'"
    End If
    
    If Len(Trim(Text5)) Then
        strAppend = strAppend & " AND DayInsured >= '" & Trim(Text5) & "'"
    End If

    If Len(Trim(Text6)) Then
        strAppend = strAppend & " AND DayInsured <= '" & Trim(Text6) & "'"
    End If
    
    If Check3.Value = 1 Then
        strAppend = strAppend & " AND (DayInsured Is null OR DayInsured = '')"
    End If
    
    If Len(Trim(Text7)) Then
        strAppend = strAppend & " AND DOAEFF >= '" & Trim(Text7) & "'"
    End If

    If Len(Trim(Text8)) Then
        strAppend = strAppend & " AND DOAEFF <= '" & Trim(Text8) & "'"
    End If
    
    If Check4.Value = 1 Then
        strAppend = strAppend & " AND (DOAEFF Is null OR DOAEFF = '')"
    End If
    
    If Len(Trim(Text9)) Then
        strAppend = strAppend & " AND DORDOA >= '" & Trim(Text9) & "'"
    End If

    If Len(Trim(Text10)) Then
        strAppend = strAppend & " AND DORDOA <= '" & Trim(Text10) & "'"
    End If
    
    If Check5.Value = 1 Then
        strAppend = strAppend & " AND (DORDOA Is null OR DORDOA = '')"
    End If
    
    If Len(Trim(Text11)) Then
        strAppend = strAppend & " AND EXPDOA >= '" & Trim(Text11) & "'"
    End If

    If Len(Trim(Text12)) Then
        strAppend = strAppend & " AND EXPDOA <= '" & Trim(Text12) & "'"
    End If
    
    If Check6.Value = 1 Then
        strAppend = strAppend & " AND (EXPDOA Is null OR EXPDOA = '')"
    End If
    
    If Len(Trim(Text13)) Then
        strAppend = strAppend & " AND DPODOD >= '" & Trim(Text13) & "'"
    End If

    If Len(Trim(Text14)) Then
        strAppend = strAppend & " AND DPODOD <= '" & Trim(Text14) & "'"
    End If
    
    If Check7.Value = 1 Then
        strAppend = strAppend & " AND (DPODOD Is null OR DPODOD = '')"
    End If
    
    If ChkHLT.Value = 1 Then
        strAppend = strAppend & " AND (HLTCode Is null OR HLTCode = '')"
    End If
    
    If ChkINS.Value = 1 Then
        strAppend = strAppend & " AND (INSCode Is null OR INSCode = '')"
    End If
    
    If ChkMemType.Value = 1 Then
        strAppend = strAppend & " AND (MBMMemberType Is null OR MBMMemberType = '')"
    End If
    
    If ChkDataStatus.Value = 1 Then
        strAppend = strAppend & " AND (MBMDataStatus Is null OR MBMDataStatus = '')"
    End If
    
    If ChkClaimability.Value = 1 Then
        strAppend = strAppend & " AND (CASClaimability Is null OR CASClaimability = '')"
    End If
    
    If ChkPolStatus.Value = 1 Then
        strAppend = strAppend & " AND (MBMStatus Is null OR MBMStatus = '')"
    End If
    
    If ChkCaseStatus.Value = 1 Then
        strAppend = strAppend & " AND (CASStatus Is null OR CASStatus = '')"
    End If
    
    If ChkExpired.Value = 1 Then
        strAppend = strAppend & " AND (MBMExpiredStatus Is null OR MBMExpiredStatus = '')"
    End If
    
    If ChkAgentCode.Value = 1 Then
        strAppend = strAppend & " AND (MBMAgentCode Is null OR MBMAgentCode = '')"
    End If
    
    If ChkPolNo.Value = 1 Then
        strAppend = strAppend & " AND (MBMPolicyNo Is null OR MBMPolicyNo = '')"
    End If
    
    If ChkPayor.Value = 1 Then
        strAppend = strAppend & " AND (PAYCode Is null OR PAYCode = '')"
    End If
    
    If ChkGrpCom.Value = 1 Then
        strAppend = strAppend & " AND (MBMGroupCompany Is null OR MBMGroupCompany = '')"
    End If
    
    If ChkHOS.Value = 1 Then
        strAppend = strAppend & " AND (HOSName Is null OR HOSName = '')"
    End If
    
    If ChkHOSGrp.Value = 1 Then
        strAppend = strAppend & " AND (HOSGRPName Is null OR HOSGRPName = '')"
    End If
    
    If ChkDoctor.Value = 1 Then
        strAppend = strAppend & " AND (AttendDoctor Is null OR AttendDoctor = '')"
    End If
            
    If ChkDOA.Value = 1 Then
        strAppend = strAppend & " AND (CASDateAdmitted Is null OR CASDateAdmitted = '')"
    End If
    
    If ChkDOD.Value = 1 Then
        strAppend = strAppend & " AND (CASDischargeDate Is null OR CASDischargeDate = '')"
    End If
    
    If chkDOR.Value = 1 Then
        strAppend = strAppend & " AND (CaseRegDate Is null OR CaseRegDate = '')"
    End If
    
    If ChkFirst.Value = 1 Then
        strAppend = strAppend & " AND (FirstDiag Is null OR FirstDiag = '')"
    End If
    
    If ChkFinal.Value = 1 Then
        strAppend = strAppend & " AND (FinalDiag Is null OR FinalDiag = '')"
    End If
    
    If ChkRep.Value = 1 Then
        strAppend = strAppend & " AND (RepCode Is null OR RepCode = '')"
    End If
    
    If ChkRepCat.Value = 1 Then
        strAppend = strAppend & " AND (RepCatCode Is null OR RepCatCode = '')"
    End If
    
    If ChkPlan.Value = 1 Then
        strAppend = strAppend & " AND (PLNCode Is null OR PLNCode = '')"
    End If
    
    If ChkEffDate.Value = 1 Then
        strAppend = strAppend & " AND (MBMPayorEffDate Is null OR MBMPayorEffDate = '')"
    End If
    
    If chkExpDate.Value = 1 Then
        strAppend = strAppend & " AND (MBMPayorExpDate Is null OR MBMPayorExpDate = '')"
    End If
    
    If ChkJoinDate.Value = 1 Then
        strAppend = strAppend & " AND (MBMFirstJoinedDate Is null OR MBMFirstJoinedDate = '')"
    End If
    
    If ChkDOB.Value = 1 Then
        strAppend = strAppend & " AND (DOB Is null OR DOB = '')"
    End If
    
    If ChkClaimNature.Value = 1 Then
        strAppend = strAppend & " AND (ClaimNatureCode Is null OR ClaimNatureCode = '')"
    End If
    
    If Mid(CboCasePending, 1, 1) = "N" Then
        strAppend = strAppend + " AND (CASPending = 'N' OR CASPending = '' OR CASPending IS NULL)"
    ElseIf Mid(CboCasePending, 1, 1) = "Y" Then
        strAppend = strAppend + " AND CASPending = 'Y'"
    End If
    
    If cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(0) Then
        Call InvoiceNoListing(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(1) Then
        Call CaseNoListing(strAppend)
    ElseIf cboReportType = cboReportType.List(0) And cboReportSelection = cboReportSelection.List(2) Then
        Call ClaimNoListing(strAppend)
    ElseIf cboReportType = cboReportType.List(1) Then
        Call ClaimValueListing(strAppend)
    ElseIf cboReportType = cboReportType.List(2) Then
        Call BordxNoListing(strAppend)
    End If
End Function

Private Sub ClaimNoListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Sql As String
        
    ListView2.ListItems.Clear
    Current = 0
    strAppendCase = "3"
    
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "ClaimReport3"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    Do
              
       Do Until rst2.EOF
        
            With rst2
                    
                    Set Record = ListView2.ListItems.Add(, , rst2!CasNumber & "-" & rst2!claimversion)
                                                  
                    If Len(!CASDateAdmitted) Then
                       Record.SubItems(1) = Format(!CASDateAdmitted, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CASDischargeDate) Then
                       Record.SubItems(2) = Format(!CASDischargeDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CaseRegDate) Then
                       Record.SubItems(3) = Format(!CaseRegDate, "dd/mm/yyyy")
                    End If
                                    
                    If Len(!CASTokenDate) Then
                       Record.SubItems(4) = Format(!CASTokenDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!BordxDate) Then
                       Record.SubItems(5) = Format(!BordxDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEffDate) Then
                       Record.SubItems(6) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEXPDate) Then
                       Record.SubItems(7) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!DayInsured) Then
                       Record.SubItems(8) = !DayInsured
                    End If
                    
                    If Len(!DOAEFF) Then
                       Record.SubItems(9) = !DOAEFF
                    End If
                    
                    If Len(!DORDOA) Then
                       Record.SubItems(10) = !DORDOA
                    End If
                    
                    If Len(!EXPDOA) Then
                       Record.SubItems(11) = !EXPDOA
                    End If
                    
                    If Len(!DPODOD) Then
                       Record.SubItems(12) = !DPODOD
                    End If
                    
                    If Len(!BORDOD) Then
                       Record.SubItems(13) = !BORDOD
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
    
    If ListView2.ListItems.Count = 0 Then
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

Private Sub ClaimValueListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Sql As String
Dim HospAppr As String
Dim MemAppr As String
Dim TotalAmt As String
Dim HospAmtChk As String
Dim TotActualChk As String
Dim MemAmtChk As String
Dim HospInvAmt As String
Dim HospColExcess As String
Dim PaytoHosp As String
Dim MemInvActual As String
Dim FinalTotActual As String
Dim Check1 As String
Dim Check2 As String
Dim Check3 As String
Dim Check4 As String
        
    ListView2.ListItems.Clear
    Current = 0
    HospInvAmt = 0
    HospAppr = 0
    MemAppr = 0
    TotalAmt = 0
    HospAmtChk = 0
    MemAmtChk = 0
    TotActualChk = 0
    HospInvAmt = 0
    HospColExcess = 0
    PaytoHosp = 0
    MemInvActual = 0
    FinalTotActual = 0
    Check1 = 0
    Check2 = 0
    Check3 = 0
    Check4 = 0
              
    strAppendCase = "4"
    
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "ClaimReport3"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    Do
       Do Until rst2.EOF
        
            With rst2
                    
                    Set Record = ListView2.ListItems.Add(, , rst2!CasNumber & "-" & rst2!claimversion)
                                                  
                    If Len(!ClaimCloseDate) Then
                       Record.SubItems(1) = Format(!ClaimCloseDate, "dd/mm/yyyy")
                    End If
                    
                    If !InvPayTo = "H" Then
                        If Len(!INVAmount) Then
                            Record.SubItems(2) = Format(!INVAmount, "#,##0.00")
                            HospInvAmt = Format(!INVAmount, "#,##0.00")
                        Else
                            Record.SubItems(2) = "0.00"
                        End If
                    End If
                    
                    If Len(!CLMExcess) Then
                       Record.SubItems(3) = Format(!CLMExcess, "#,##0.00")
                    Else
                        Record.SubItems(3) = "0.00"
                    End If
                                    
                    If Len(!INVPaidByMem) Then
                       Record.SubItems(4) = Format(!INVPaidByMem, "#,##0.00")
                       HospColExcess = Format(!INVPaidByMem, "#,##0.00")
                    Else
                        Record.SubItems(4) = "0.00"
                    End If
                    
                    If Len(!CLMPayableByMedix2H) Then
                       Record.SubItems(5) = Format(!CLMPayableByMedix2H, "#,##0.00")
                       PaytoHosp = Format(!CLMPayableByMedix2H, "#,##0.00")
                    Else
                        Record.SubItems(5) = "0.00"
                    End If
                    
                    If Len(!CLMReimburse) Then
                       Record.SubItems(6) = Format(!CLMReimburse, "#,##0.00")
                       MemInvActual = Format(!CLMReimburse, "#,##0.00")
                    Else
                        Record.SubItems(6) = "0.00"
                    End If
                    
                    If Len(!CLMTotalActual) Then
                       Record.SubItems(7) = Format(!CLMTotalActual, "#,##0.00")
                       FinalTotActual = Format(!CLMTotalActual, "#,##0.00")
                    Else
                        Record.SubItems(7) = "0.00"
                    End If
                    
                    If Len(!CLMPayableByMedix2H) Then
                       Record.SubItems(8) = Format(!CLMPayableByMedix2H, "#,##0.00")
                       HospAppr = Format(!CLMPayableByMedix2H, "#,##0.00")
                    Else
                        Record.SubItems(8) = "0.00"
                    End If
                    
                    If Len(!CLMReimburse) Then
                       Record.SubItems(9) = Format(!CLMReimburse, "#,##0.00")
                       MemAppr = Format(!CLMReimburse, "#,##0.00")
                    Else
                        Record.SubItems(9) = "0.00"
                    End If
                    
                    TotalAmt = CDbl(HospAppr) + CDbl(MemAppr)
                    
                    Record.SubItems(10) = Format(TotalAmt, "#,##0.00")
                    
                    Check1 = CDbl(HospInvAmt) - CDbl(HospColExcess) - CDbl(PaytoHosp)
                    Record.SubItems(11) = Format(Check1, "#,##0.00")
                    
                    Check2 = CDbl(HospInvAmt) + CDbl(MemInvActual) - CDbl(FinalTotActual)
                    Record.SubItems(12) = Format(Check2, "#,##0.00")
                    
                    Check3 = CDbl(MemInvActual) - CDbl(MemAppr)
                    Record.SubItems(13) = Format(Check3, "#,##0.00")
                    
                    Check4 = CDbl(HospAppr) + CDbl(MemAppr) - CDbl(TotAmt)
                    Record.SubItems(14) = Format(Check4, "#,##0.00")
                    
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
    
    If ListView2.ListItems.Count = 0 Then
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

Private Sub CaseNoListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Sql As String
        
    ListView2.ListItems.Clear
    Current = 0
    strAppendCase = "2"
    
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "ClaimReport3"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    Do
              
       Do Until rst2.EOF
        
            With rst2
                    
                    Set Record = ListView2.ListItems.Add(, , rst2!CasNumber)
                                                  
                    If Len(!CASDateAdmitted) Then
                       Record.SubItems(1) = Format(!CASDateAdmitted, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CASDischargeDate) Then
                       Record.SubItems(2) = Format(!CASDischargeDate, "dd/mm/yyyy")
                    End If
                                    
                    If Len(!CASTokenDate) Then
                       Record.SubItems(3) = Format(!CASTokenDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEffDate) Then
                       Record.SubItems(4) = Format(!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!MBMPayorEXPDate) Then
                       Record.SubItems(5) = Format(!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!DayInsured) Then
                       Record.SubItems(6) = !DayInsured
                    End If
                    
                    If Len(!DOAEFF) Then
                       Record.SubItems(7) = !DOAEFF
                    End If
                    
                    If Len(!DORDOA) Then
                       Record.SubItems(8) = !DORDOA
                    End If
                    
                    If Len(!EXPDOA) Then
                       Record.SubItems(9) = !EXPDOA
                    End If
                    
                    If Len(!DPODOD) Then
                       Record.SubItems(10) = !DPODOD
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
    
    If ListView2.ListItems.Count = 0 Then
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

Private Sub InvoiceNoListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim Sql As String
        
    ListView2.ListItems.Clear
    Current = 0
    
    strAppendCase = "1"
    
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "ClaimReport3"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
    
    Do
              
       Do Until rst2.EOF
        
            With rst2
                    
                    If Len(!INVNo) Then
                        Set Record = ListView2.ListItems.Add(, , !INVNo)
                    Else
                        Set Record = ListView2.ListItems.Add(, , "Empty")
                    End If
                                                  
                    If Len(!CasNumber & "-" & !claimversion) Then
                        Record.SubItems(1) = (!CasNumber & "-" & !claimversion)
                    End If
                    
                    If Len(!CASDateAdmitted) Then
                       Record.SubItems(2) = Format(!CASDateAdmitted, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CASDischargeDate) Then
                       Record.SubItems(3) = Format(!CASDischargeDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!InvDate) Then
                       Record.SubItems(4) = Format(!InvDate, "dd/mm/yyyy")
                    End If
                                    
                    If Len(!INVEnteredDate) Then
                       Record.SubItems(5) = Format(!INVEnteredDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!InvDate) Then
                       Record.SubItems(6) = Format(!InvDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!INVFolUpDate) Then
                       Record.SubItems(7) = Format(!INVFolUpDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!INVDORINV) Then
                       Record.SubItems(8) = !INVDORINV
                    End If
                    
                    If Len(!INVDOD) Then
                       Record.SubItems(9) = !INVDOD
                    End If
                    
                    If Len(!PREDOA) Then
                       Record.SubItems(10) = !PREDOA
                    End If
                    
                    If Len(!FOLDOD) Then
                       Record.SubItems(11) = !FOLDOD
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
    
    If ListView2.ListItems.Count = 0 Then
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
Private Sub BordxNoListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim Sql As String
Dim sql2 As String
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
        
    lblCurrent = 0
    
    ListView2.ListItems.Clear
    strAppendCase = "5"
    
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "ClaimReport3"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 300
        Set prm1 = cmd.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, strAppend)
        cmd.Parameters.Append prm1
        Set prm2 = cmd.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
        cmd.Parameters.Append prm2
        rst2.CursorLocation = adUseClient
        rst2.CursorType = adOpenForwardOnly
        rst2.CacheSize = 100
        Set rst2 = cmd.Execute
        
        Do Until rst2.EOF
        
            With rst2
                If Len(rst2!BordxRefNo) Then
                    Set Record = ListView2.ListItems.Add(, , rst2!BordxRefNo)
                Else
                    Set Record = ListView2.ListItems.Add(, , "Empty")
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
                
            End With
            
            lblCurrent = lblCurrent + 1
            
            rst2.MoveNext
       Loop
    If ListView2.ListItems.Count = 0 Then
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

Private Sub CboCasePending_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboCasePending.Text, CboCasePending.SelStart) + Chr(KeyAscii) + Mid(CboCasePending.Text, CboCasePending.SelStart + CboCasePending.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboCasePending.ListCount - 1
 
        If NewText = LCase(Left(CboCasePending.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboCasePending.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboCasePending.Text = ValidValue
                CboCasePending.SelStart = 0
                CboCasePending.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboClmPending_Change()
    If InStr(CboClmPending.Text, "null") Then
       CboClmPending.Enabled = False
    End If
End Sub

Private Sub CboClmPending_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClmPending.Text, CboClmPending.SelStart) + Chr(KeyAscii) + Mid(CboClmPending.Text, CboClmPending.SelStart + CboClmPending.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboClmPending.ListCount - 1
 
        If NewText = LCase(Left(CboClmPending.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClmPending.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboClmPending.Text = ValidValue
                CboClmPending.SelStart = 0
                CboClmPending.SelLength = Len(ValidValue)
            End If
        End If
End Sub

