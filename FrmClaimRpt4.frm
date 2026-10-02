VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmClaimRpt4 
   Caption         =   "Claim Report 4"
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
   Begin VB.Frame Frame1 
      Caption         =   "Report Type"
      Height          =   735
      Left            =   0
      TabIndex        =   158
      Top             =   240
      Width           =   11895
      Begin VB.ComboBox cboReportSelection 
         Height          =   315
         Left            =   5880
         TabIndex        =   1
         Top             =   280
         Width           =   5000
      End
      Begin VB.ComboBox cboReportType 
         Height          =   315
         Left            =   600
         TabIndex        =   0
         Top             =   280
         Width           =   3000
      End
      Begin VB.Label Label5 
         Caption         =   "Report Selection"
         Height          =   255
         Left            =   5280
         TabIndex        =   159
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   0
      TabIndex        =   151
      Top             =   7320
      Width           =   11895
      Begin VB.TextBox lblCurrent 
         Height          =   285
         Left            =   240
         TabIndex        =   162
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   375
         Left            =   5400
         TabIndex        =   57
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   9720
         TabIndex        =   62
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   375
         Left            =   10560
         TabIndex        =   63
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print to File"
         Height          =   375
         Left            =   8760
         TabIndex        =   61
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         Height          =   375
         Left            =   7080
         TabIndex        =   59
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         Height          =   375
         Left            =   7920
         TabIndex        =   60
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "S&top"
         Height          =   375
         Left            =   6240
         TabIndex        =   58
         Top             =   160
         Width           =   855
      End
      Begin VB.Label Label64 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   1560
         TabIndex        =   152
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame22 
      Caption         =   "Selection Criteria"
      Height          =   5235
      Left            =   0
      TabIndex        =   95
      Top             =   960
      Width           =   11865
      Begin VB.CheckBox ChkGrpCom2 
         Caption         =   "Check1"
         Height          =   255
         Left            =   7440
         TabIndex        =   160
         Top             =   2180
         Width           =   255
      End
      Begin VB.TextBox TxtMBMNo1 
         Height          =   285
         Left            =   8640
         MaxLength       =   15
         TabIndex        =   48
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox TxtCaseNo1 
         Height          =   285
         Left            =   8640
         MaxLength       =   15
         TabIndex        =   50
         Top             =   495
         Width           =   1000
      End
      Begin VB.TextBox TxtMBMNo2 
         Height          =   285
         Left            =   10335
         MaxLength       =   15
         TabIndex        =   49
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox TxtCaseNo2 
         Height          =   285
         Left            =   10335
         MaxLength       =   15
         TabIndex        =   51
         Top             =   495
         Width           =   1000
      End
      Begin VB.ComboBox CboPending 
         Height          =   315
         Left            =   8640
         Sorted          =   -1  'True
         TabIndex        =   52
         Top             =   750
         Width           =   2685
      End
      Begin VB.CheckBox ChkPayor 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7200
         TabIndex        =   156
         Top             =   1920
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
         Top             =   1875
         Width           =   2800
      End
      Begin VB.CheckBox ChkPending 
         Caption         =   "Check1"
         Height          =   255
         Left            =   11400
         TabIndex        =   94
         Top             =   795
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox ChkBordxDate 
         Caption         =   "Check1"
         Height          =   255
         Left            =   11400
         TabIndex        =   93
         Top             =   1320
         Visible         =   0   'False
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
         Left            =   11400
         TabIndex        =   91
         Top             =   555
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
         Left            =   11400
         TabIndex        =   90
         Top             =   300
         Width           =   255
      End
      Begin VB.CheckBox ChkClaimNature 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7200
         TabIndex        =   79
         Top             =   1590
         Width           =   255
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   3
         Text            =   "N - Non Sihat"
         Top             =   200
         Width           =   1680
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
         Left            =   7200
         TabIndex        =   83
         Top             =   3015
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
         Left            =   7200
         TabIndex        =   84
         Top             =   3255
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
         Left            =   7200
         TabIndex        =   87
         Top             =   4080
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
         Left            =   7200
         TabIndex        =   85
         Top             =   3540
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
         Left            =   7200
         TabIndex        =   88
         Top             =   4320
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
         Left            =   7200
         TabIndex        =   89
         Top             =   4560
         Width           =   255
      End
      Begin VB.CheckBox ChkGrpCom 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7200
         TabIndex        =   80
         Top             =   2175
         Width           =   255
      End
      Begin VB.CheckBox ChkPolNo 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2880
         TabIndex        =   72
         Top             =   2980
         Width           =   255
      End
      Begin VB.CheckBox ChkHOS 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7200
         TabIndex        =   81
         Top             =   2430
         Width           =   255
      End
      Begin VB.CheckBox ChkHOSGrp 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7200
         TabIndex        =   82
         Top             =   2715
         Width           =   255
      End
      Begin VB.CheckBox ChkDoctor 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2880
         TabIndex        =   73
         Top             =   3240
         Width           =   255
      End
      Begin VB.CheckBox ChkFirst 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7200
         TabIndex        =   75
         Top             =   600
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
         Left            =   7200
         TabIndex        =   86
         Top             =   3840
         Width           =   255
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
         Left            =   11400
         TabIndex        =   92
         Top             =   1050
         Visible         =   0   'False
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
         Height          =   405
         Left            =   7200
         TabIndex        =   74
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox ChkFinal 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7200
         TabIndex        =   76
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox ChkRep 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7200
         TabIndex        =   77
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox ChkRepCat 
         Caption         =   "Check3"
         Height          =   255
         Left            =   7200
         TabIndex        =   78
         Top             =   1320
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
         TabIndex        =   70
         Top             =   2440
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
         TabIndex        =   66
         Top             =   1305
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
         TabIndex        =   71
         Top             =   2700
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
      Begin VB.CheckBox ChkINS 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2880
         TabIndex        =   64
         Top             =   500
         Width           =   255
      End
      Begin VB.CheckBox ChkMemType 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2880
         TabIndex        =   65
         Top             =   720
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
         TabIndex        =   68
         Top             =   1880
         Width           =   255
      End
      Begin VB.CheckBox ChkPolStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2880
         TabIndex        =   67
         Top             =   1570
         Width           =   255
      End
      Begin VB.CheckBox ChkCaseStatus 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2880
         TabIndex        =   69
         Top             =   2180
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
      Begin VB.TextBox TxtPlan1 
         Height          =   315
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   18
         Top             =   300
         Width           =   1050
      End
      Begin VB.TextBox TxtPlan2 
         Height          =   315
         Left            =   6135
         MaxLength       =   4
         TabIndex        =   19
         Top             =   300
         Width           =   1050
      End
      Begin VB.TextBox txtPreDiag1 
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   20
         Top             =   585
         Width           =   1050
      End
      Begin VB.TextBox txtPreDiag2 
         Height          =   285
         Left            =   6135
         MaxLength       =   3
         TabIndex        =   21
         Top             =   585
         Width           =   1050
      End
      Begin VB.TextBox txtFinalDiag1 
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   22
         Top             =   840
         Width           =   1050
      End
      Begin VB.TextBox txtFinalDiag2 
         Height          =   285
         Left            =   6135
         MaxLength       =   3
         TabIndex        =   23
         Top             =   840
         Width           =   1050
      End
      Begin VB.TextBox txtRepudCat1 
         Height          =   285
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   24
         Top             =   1080
         Width           =   1050
      End
      Begin VB.TextBox txtRepudCat2 
         Height          =   285
         Left            =   6135
         MaxLength       =   4
         TabIndex        =   25
         Top             =   1080
         Width           =   1050
      End
      Begin VB.TextBox txtRepudCode1 
         Height          =   285
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   26
         Top             =   1320
         Width           =   1050
      End
      Begin VB.TextBox txtRepudCode2 
         Height          =   285
         Left            =   6135
         MaxLength       =   10
         TabIndex        =   27
         Top             =   1320
         Width           =   1050
      End
      Begin VB.TextBox TxtClaimNature1 
         Height          =   285
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   28
         Top             =   1575
         Width           =   1050
      End
      Begin VB.TextBox TxtClaimNature2 
         Height          =   285
         Left            =   6135
         MaxLength       =   10
         TabIndex        =   29
         Top             =   1575
         Width           =   1050
      End
      Begin VB.ComboBox CboGrpCom 
         Height          =   315
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   2160
         Width           =   2800
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
         Top             =   2400
         Width           =   2800
      End
      Begin VB.ComboBox cboHOSGrp 
         Height          =   315
         Left            =   4320
         TabIndex        =   33
         Top             =   2685
         Width           =   2800
      End
      Begin MSMask.MaskEdBox TxtDOA1 
         Height          =   315
         Left            =   4320
         TabIndex        =   34
         Top             =   2970
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
         Left            =   4320
         TabIndex        =   36
         Top             =   3255
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
         Left            =   4320
         TabIndex        =   38
         Top             =   3540
         Width           =   1005
         _ExtentX        =   1773
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
         Top             =   3825
         Width           =   1005
         _ExtentX        =   1773
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
         Top             =   4080
         Width           =   1005
         _ExtentX        =   1773
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
         Top             =   4320
         Width           =   1005
         _ExtentX        =   1773
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
         Top             =   4560
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtBordxNo1 
         Height          =   285
         Left            =   8640
         MaxLength       =   10
         TabIndex        =   53
         Top             =   1035
         Visible         =   0   'False
         Width           =   1005
      End
      Begin MSMask.MaskEdBox TxtBordxDate1 
         Height          =   315
         Left            =   8640
         TabIndex        =   55
         Top             =   1275
         Visible         =   0   'False
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
         Left            =   6135
         TabIndex        =   35
         Top             =   2970
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
         Left            =   6135
         TabIndex        =   37
         Top             =   3255
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
         Left            =   6135
         TabIndex        =   39
         Top             =   3540
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
         Left            =   6135
         TabIndex        =   41
         Top             =   3840
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtPayExpDate2 
         Height          =   315
         Left            =   6135
         TabIndex        =   43
         Top             =   4080
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDateJoin2 
         Height          =   315
         Left            =   6135
         TabIndex        =   45
         Top             =   4320
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOB2 
         Height          =   315
         Left            =   6135
         TabIndex        =   47
         Top             =   4560
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtBordxNo2 
         Height          =   285
         Left            =   10335
         MaxLength       =   10
         TabIndex        =   54
         Top             =   1035
         Visible         =   0   'False
         Width           =   1000
      End
      Begin MSMask.MaskEdBox TxtBordxDate2 
         Height          =   315
         Left            =   10335
         TabIndex        =   56
         Top             =   1245
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label38 
         Caption         =   "Payor"
         Height          =   225
         Left            =   3360
         TabIndex        =   157
         Top             =   1935
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Case Pending"
         Height          =   240
         Left            =   7560
         TabIndex        =   154
         Top             =   795
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Stay Type"
         Height          =   255
         Left            =   120
         TabIndex        =   150
         Top             =   3800
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Cash Type"
         Height          =   240
         Left            =   120
         TabIndex        =   149
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9705
         TabIndex        =   148
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label70 
         Caption         =   "Bordx Date"
         Height          =   255
         Left            =   7560
         TabIndex        =   147
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9705
         TabIndex        =   146
         Top             =   520
         Width           =   375
      End
      Begin VB.Label Label66 
         Caption         =   "Claim No"
         Height          =   255
         Left            =   7560
         TabIndex        =   145
         Top             =   555
         Width           =   1095
      End
      Begin VB.Label Label65 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9705
         TabIndex        =   144
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label24 
         Caption         =   "Member No"
         Height          =   255
         Left            =   7560
         TabIndex        =   143
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   195
         Left            =   5505
         TabIndex        =   142
         Top             =   3585
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Claim Nature"
         Height          =   255
         Left            =   3360
         TabIndex        =   141
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Supp. Type"
         Height          =   255
         Left            =   120
         TabIndex        =   140
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "Policy No."
         Height          =   255
         Left            =   120
         TabIndex        =   139
         Top             =   3010
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Exp Status"
         Height          =   255
         Left            =   120
         TabIndex        =   138
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label19 
         Height          =   255
         Left            =   2520
         TabIndex        =   137
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Mem Type"
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   830
         Width           =   855
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5505
         TabIndex        =   135
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5505
         TabIndex        =   134
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label48 
         Caption         =   "DOD"
         Height          =   255
         Left            =   3360
         TabIndex        =   133
         ToolTipText     =   "Date of Discharged"
         Top             =   3330
         Width           =   975
      End
      Begin VB.Label Label47 
         Caption         =   "DOA"
         Height          =   255
         Left            =   3360
         TabIndex        =   132
         ToolTipText     =   "Date of Admission"
         Top             =   3075
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5505
         TabIndex        =   131
         Top             =   3330
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5505
         TabIndex        =   130
         Top             =   4080
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5505
         TabIndex        =   129
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "DORC"
         Height          =   255
         Left            =   3360
         TabIndex        =   128
         ToolTipText     =   "Member Value Date"
         Top             =   3585
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Exp Date"
         Height          =   255
         Left            =   3360
         TabIndex        =   127
         ToolTipText     =   "Payor Expiry Date"
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Eff Date"
         Height          =   255
         Left            =   3360
         TabIndex        =   126
         ToolTipText     =   "Payor Effective Date"
         Top             =   3840
         Width           =   735
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
         TabIndex        =   125
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
         TabIndex        =   124
         Top             =   550
         Width           =   690
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
         TabIndex        =   123
         Top             =   1350
         Width           =   1020
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5505
         TabIndex        =   122
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label51 
         Caption         =   "Plan Code"
         Height          =   255
         Left            =   3360
         TabIndex        =   121
         Top             =   360
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
         Left            =   120
         TabIndex        =   120
         Top             =   1640
         Width           =   1020
      End
      Begin VB.Label Label53 
         Caption         =   "Claimability"
         Height          =   255
         Left            =   120
         TabIndex        =   119
         Top             =   2500
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
         Left            =   120
         TabIndex        =   118
         Top             =   2200
         Width           =   945
      End
      Begin VB.Label Label6 
         Caption         =   "Grp Comp"
         Height          =   255
         Left            =   3360
         TabIndex        =   117
         Top             =   2220
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Pre. Diag"
         Height          =   255
         Left            =   3360
         TabIndex        =   116
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Final Diag"
         Height          =   255
         Left            =   3360
         TabIndex        =   115
         Top             =   870
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Repud Cat"
         Height          =   255
         Left            =   3360
         TabIndex        =   114
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Repud Code"
         Height          =   255
         Left            =   3360
         TabIndex        =   113
         Top             =   1425
         Width           =   975
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5505
         TabIndex        =   112
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5505
         TabIndex        =   111
         Top             =   870
         Width           =   375
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5505
         TabIndex        =   110
         Top             =   1170
         Width           =   375
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5505
         TabIndex        =   109
         Top             =   1425
         Width           =   375
      End
      Begin VB.Label Label36 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   3360
         TabIndex        =   108
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label37 
         Caption         =   "Hosp. Grp"
         Height          =   255
         Left            =   3360
         TabIndex        =   107
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label39 
         Caption         =   "Doctor"
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   3260
         Width           =   1095
      End
      Begin VB.Label Label40 
         Caption         =   "Date Join"
         Height          =   255
         Left            =   3360
         TabIndex        =   105
         ToolTipText     =   "Payor Expiry Date"
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5505
         TabIndex        =   104
         Top             =   4320
         Width           =   375
      End
      Begin VB.Label Label42 
         Caption         =   "DOB"
         Height          =   255
         Left            =   3360
         TabIndex        =   103
         ToolTipText     =   "Payor Expiry Date"
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   5505
         TabIndex        =   102
         Top             =   4575
         Width           =   375
      End
      Begin VB.Label Label55 
         Caption         =   "Agent Code"
         Height          =   255
         Left            =   120
         TabIndex        =   101
         Top             =   2760
         Width           =   1095
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
         Left            =   7200
         TabIndex        =   100
         Top             =   120
         Width           =   255
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
         TabIndex        =   99
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label60 
         Caption         =   "Bordx No"
         Height          =   255
         Left            =   7560
         TabIndex        =   98
         Top             =   1035
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label61 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   9705
         TabIndex        =   97
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label62 
         Caption         =   "Patient Name"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   3500
         Width           =   1095
      End
      Begin VB.Label Label15 
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
         Height          =   255
         Left            =   7450
         TabIndex        =   161
         Top             =   120
         Width           =   255
      End
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1095
      Left            =   0
      TabIndex        =   153
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
      NumItems        =   32
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
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Claims Reports 4"
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
      TabIndex        =   155
      Top             =   0
      Width           =   11805
   End
End
Attribute VB_Name = "FrmClaimRpt4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private AA As Integer
Private intExit As Integer
Private GrpCompany As String
Private strAppend As String
Private HOspitals As String
Private m_blnDirection(1 To 32) As Boolean

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
        Else
            CboGrpCom.Clear
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
    ListView2.Top = 300
    ListView2.Width = 11775
    Frame22.Visible = False
    Frame1.Visible = False
End Sub

Private Sub ListViewHeader1()

If cboReportSelection.ListIndex = 1 Then
    ListView2.ColumnHeaders(1).Text = "Case No"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Policy No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Covered ID"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "Prin. Name"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(6).Text = "Patient Name"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(7).Text = "Prin. Sex"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Text = "Prin. Ageband"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Text = "Patient Sex"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(10).Text = "Patient Ageband"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Text = "Insured Type"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Plan Code"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(13).Text = "Hospital"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(14).Text = "First Diag."
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Text = "Description"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(16).Text = "Final Diag."
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Text = "Description"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(18).Text = "Case Status"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Text = "Claimability"
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(20).Text = "Policy Status"
    ListView2.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(21).Text = "DOA"
    ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(22).Text = "DOD"
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Text = "DOR"
    ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(24).Text = "Delay Duration"
    ListView2.ColumnHeaders(24).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(25).Text = "Payor Eff. Date"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(26).Text = "Payor Exp. Date"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnCenter
    
    For AA = 27 To 32
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
    
ElseIf cboReportSelection.ListIndex = 0 Then
    ListView2.ColumnHeaders(1).Text = "Case & Wks No"
    ListView2.ColumnHeaders(2).Text = "Clm Version"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(3).Text = "Member No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Policy No"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(5).Text = "Relation"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Case Status"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "Claimability"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Text = "Member Name"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(9).Text = "Patient Name"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(10).Text = "Race"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Text = "Patient Sex"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "Patient Ageband"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Text = "Patient IC No"
    ListView2.ColumnHeaders(13).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(14).Text = "Patient DOB"
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Text = "Payor Eff Date"
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Text = "Payor Exp Date"
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Text = "DOR"
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(18).Text = "DOA"
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Text = "DOD"
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
    ListView2.ColumnHeaders(25).Text = "Reserve Amt"
    ListView2.ColumnHeaders(25).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(26).Text = "Bill Amt"
    ListView2.ColumnHeaders(26).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(27).Text = "Recovery Amt"
    ListView2.ColumnHeaders(27).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(28).Text = "Annual Limit"
    ListView2.ColumnHeaders(28).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(29).Text = "Payable by Medix2H"
    ListView2.ColumnHeaders(29).Alignment = lvwColumnRight
    ListView2.ColumnHeaders(30).Text = "Reason for admission"
    ListView2.ColumnHeaders(30).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(31).Text = "Underlying Dieses"
    ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
    
    For AA = 32 To 32
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA

ElseIf cboReportSelection.ListIndex = 2 Then
    ListView2.ColumnHeaders(1).Text = "Case No"
    ListView2.ColumnHeaders(2).Text = "Member No"
    ListView2.ColumnHeaders(2).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(3).Text = "Policy No"
    ListView2.ColumnHeaders(3).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(4).Text = "Data Status"
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Text = "Pending"
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Text = "Case Status"
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Text = "Claimability"
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Text = "Pay Type"
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Text = "Policy Status"
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(10).Text = "DOA"
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Text = "DOD"
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Text = "DOR"
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
        
    For AA = 13 To 31
        ListView2.ColumnHeaders(AA).Text = ""
    Next AA
End If

End Sub

Private Sub Form_Load()
    intExit = 0
  '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
    Call ComboListing
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
    Call ListViewHeader1
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
        If ListView2.ListItems.Count = 0 Then
            MsgBox "There aren't any items In the listview selected." _
            , vbOKOnly + vbInformation, "Export Failed"
        Else
            Call ExportListViewtoExcel(ListView2)
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
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
     '   medixmasterconnWS.Open "File Name=" & BOSPath & "Conn\mediconnectWS.UDL"
        Call SearchRecords
       ' medixmasterconn.Close
      '  Set medixmasterconn = Nothing
      '  medixmasterconnWS.Close
      '  Set medixmasterconnWS = Nothing
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
  '  medixmasterconn.Close
  '  Set medixmasterconn = Nothing
  '  medixmasterconnWS.Close
  '  Set medixmasterconnWS = Nothing
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
    cboReportSelection.Clear
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
    
    With CboClaimability
        .AddItem "1A" & " - " & "All Accepted"
        .AddItem "1N" & " - " & "All Not Decided"
        .AddItem "1P" & " - " & "All Pay&Claim"
        .AddItem "1R" & " - " & "All Repudiated"
        .AddItem "AA" & " - " & "Accepted"
        .AddItem "RR" & " - " & "Repudiated"
        .AddItem "CC" & " - " & "Cancelled"
        .AddItem "SA" & " - " & "Special Approval"
        .AddItem "PP" & " - " & "Pay & Claim"
        .AddItem "NN" & " - " & "Not Decided"
        .AddItem "NA" & " - " & "Not Decided/Accepted"
        .AddItem "NR" & " - " & "Not Decided/Repudiated"
        .AddItem "PA" & " - " & "Pay&Claim/Accepted"
        .AddItem "PR" & " - " & "Pay&Claim/Rejected"
        .AddItem "AR" & " - " & "Accepted/Rejected"
        .AddItem "RA" & " - " & "Rejected/Accepted"
        .AddItem "BR" & " - " & "BTBG/Rejected"
        .AddItem "AI" & " - " & "Accepted/Insurance"
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
    
    CboPending.AddItem "N - No"
    CboPending.AddItem "Y - Yes"
    
    With CboStayType
        .AddItem "I - IN-PATIENT"
        .AddItem "O - OUT-PATIENT"
    End With
    
    With CboCashType
        .AddItem "C - CASHLESS"
        .AddItem "R - REIMBURSEMENT"
    End With
           
    cboSuppType.AddItem "Both"
    cboSuppType.AddItem "C" & " - " & "Combined Limit"
    cboSuppType.AddItem "S" & " - " & "Separated Limit"
    cboSuppType.Text = "Both"
    
    cboReportType.AddItem "General Listing"
    
End Sub
'============== ComboBox Listing =================================

'=========== Option Button, Combo Box Change/KeyPress ===================

Private Sub cboHLTCode_Change()
    If InStr(cboHLTCode.Text, "null") Then
       cboHLTCode.Enabled = False
    End If
    
    If Len(Trim(cboHLTCode)) Then
       ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        LoadPayor (cboHLTCode)
      '  medixmasterconn.Close
      '  Set medixmasterconn = Nothing
    End If
End Sub

Private Sub cboHLTCode_Click()
    If Len(Trim(cboHLTCode)) Then
      '  medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        LoadPayor (cboHLTCode)
       ' medixmasterconn.Close
       ' Set medixmasterconn = Nothing
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
           ' medixmasterconn.Open "File Name=" & BOSPath & "Conn\mediconnect.UDL"
        
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
            'medixmasterconn.Close
            'Set medixmasterconn = Nothing
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

Private Sub CboReportType_Click()

    ListView2.ListItems.Clear
    If cboReportType.ListIndex = 0 Then
        lblCurrent = ""
        ListView2.ListItems.Clear
        cboReportSelection.Clear
        cboReportSelection.AddItem "Cases & Worksheets"
        'cboReportSelection.AddItem "Cases & Worksheets (with Remark)"
        cboReportSelection.AddItem "Cases without Worksheets"
        cboReportSelection.AddItem "Cases without Worksheets (Shorter ver.)"
        cboReportSelection = cboReportSelection.List(0)
        cboReportSelection.ListIndex = 0
        Call ListViewHeader1
    End If
End Sub

Private Sub cboReportSelection_Click()

    ListView2.ListItems.Clear
    If cboReportSelection.ListIndex = 0 Then
        lblCurrent = ""
        ListView2.ListItems.Clear
        Call ListViewHeader1
    ElseIf cboReportSelection.ListIndex = 1 Then
        lblCurrent = ""
        ListView2.ListItems.Clear
        Call ListViewHeader1
    ElseIf cboReportSelection.ListIndex = 2 Then
        lblCurrent = ""
        ListView2.ListItems.Clear
        Call ListViewHeader1
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

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    
    If cboReportType.ListIndex = 0 And cboReportSelection.ListIndex = 0 Then
        Select Case ColumnHeader.Index
        Case 1
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(1)
        Case 2
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(2)
        Case 3
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(3)
        Case 4
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(4)
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
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(12)
        Case 13
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(13)
        Case 14
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(14)
        Case 15
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(19)
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
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(29)
        Case 30
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(30)
        End Select
    
    ElseIf cboReportType.ListIndex = 0 And cboReportSelection.ListIndex = 1 Then
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
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(5)
        Case 6
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(9)
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
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(15)
        Case 16
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(16)
        Case 17
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(17)
        Case 18
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(18)
        Case 19
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(19)
        Case 20
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(20)
        Case 21
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(21)
        Case 22
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(22)
        Case 23
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(23)
        Case 24
            SortListView ListView2, ColumnHeader.Index, ldtNumber, m_blnDirection(24)
        Case 25
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(25)
        Case 26
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(26)
        End Select
        
    ElseIf cboReportType.ListIndex = 0 And cboReportSelection.ListIndex = 2 Then
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
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(6)
        Case 7
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(7)
        Case 8
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(8)
        Case 9
            SortListView ListView2, ColumnHeader.Index, ldtString, m_blnDirection(9)
        Case 10
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
        Case 11
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
        Case 12
            SortListView ListView2, ColumnHeader.Index, ldtDateTime, m_blnDirection(12)
        End Select
        
    End If
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

    strAppend = ""
    GrpCompany = ""
    HOspitals = ""
    
    If Len(cboReportType) = False Then
        MsgBox "Please select a Report Type.", vbOKOnly + vbCritical, DialogBoxTitle
        CmdPrint.Enabled = True
        CmdClear.Enabled = True
        cmdExit.Enabled = True
        cmdSearch.Enabled = True
        Exit Function
    End If
    
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
    
    If Mid(CboPending, 1, 1) = "N" Then
        strAppend = strAppend + " AND (CASPending = 'N' OR CASPending = '' OR CASPending IS NULL)"
    ElseIf Mid(CboPending, 1, 1) = "Y" Then
        strAppend = strAppend + " AND CASPending = 'Y'"
    End If
    
    'If Len(Mid(CboClaimability, 1, 1)) Then
    '    strAppend = strAppend & " AND CASClaimability = '" & Trim(Mid(CboClaimability, 1, 1)) & "'"
    'End If

    If Len(Mid(CboClaimability, 1, 2)) Then
        If Trim(Mid(CboClaimability, 1, 2)) = "1A" Then
            strAppend = strAppend & " AND (CASClaimability = 'AA' OR CASClaimability = 'SA' OR CASClaimability = 'NA' OR CASClaimability = 'PA' OR CASClaimability = 'RA')"
        ElseIf Trim(Mid(CboClaimability, 1, 2)) = "1R" Then
            strAppend = strAppend & " AND (CASClaimability = 'RR' OR CASClaimability = 'NR' OR CASClaimability = 'PR' OR CASClaimability = 'AR' OR CASClaimability = 'BR')"
        ElseIf Trim(Mid(CboClaimability, 1, 2)) = "1N" Then
            strAppend = strAppend & " AND (CASClaimability = 'NN' OR CASClaimability = 'NA' OR CASClaimability = 'NR')"
        ElseIf Trim(Mid(CboClaimability, 1, 2)) = "1P" Then
            strAppend = strAppend & " AND (CASClaimability = 'PP' OR CASClaimability = 'PA' OR CASClaimability = 'PR')"
        Else
            strAppend = strAppend & " AND CASClaimability = '" & Trim(Mid(CboClaimability, 1, 2)) & "'"
        End If
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
    
    If ChkPending.Value = 1 Then
        strAppend = strAppend & " AND (CASPending Is null OR CASPending = '')"
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
    
    If cboReportType.ListIndex = 0 And cboReportSelection.ListIndex = 0 Then
        Call CaseWksListing(strAppend)
    ElseIf cboReportType.ListIndex = 0 And cboReportSelection.ListIndex = 1 Then
        Call CaseNoListing(strAppend)
    ElseIf cboReportType.ListIndex = 0 And cboReportSelection.ListIndex = 2 Then
        Call CaseNoListing2(strAppend)
    End If
    
End Function

Private Sub CASFirstICD(StrAppend1 As String, strAppendCase As String, Record)
Dim rst3 As New ADODB.Recordset
Dim con As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter

Set con.ActiveConnection = medixmasterconn
    con.CommandText = "ClaimReport"
    con.CommandType = adCmdStoredProc
    con.CommandTimeout = 300
    Set prm1 = con.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
    con.Parameters.Append prm1
    Set prm2 = con.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    con.Parameters.Append prm2
    rst3.CursorLocation = adUseClient
    rst3.CursorType = adOpenForwardOnly
    rst3.CacheSize = 100
    Set rst3 = con.Execute

    If cboReportSelection.ListIndex = 0 Then
        If rst3.EOF = True Then
            Record.SubItems(21) = ""
            rst3.Close
            Set rst3 = Nothing
        Else
            Record.SubItems(21) = Trim(rst3!ICDDescription)
            rst3.Close
            Set rst3 = Nothing
        End If
    ElseIf cboReportSelection.ListIndex = 1 Then
        If rst3.EOF = True Then
            Record.SubItems(14) = ""
            rst3.Close
            Set rst3 = Nothing
        Else
            Record.SubItems(14) = Trim(rst3!ICDDescription)
            rst3.Close
            Set rst3 = Nothing
        End If
    ElseIf cboReportSelection.ListIndex = 2 Then
    End If

End Sub

Private Sub CASFinalICD(StrAppend1 As String, strAppendCase As String, Record)
Dim rst3 As New ADODB.Recordset
Dim con As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter

Set con.ActiveConnection = medixmasterconn
    con.CommandText = "ClaimReport"
    con.CommandType = adCmdStoredProc
    con.CommandTimeout = 300
    Set prm1 = con.CreateParameter("StrAppend", adVarChar, adParamInput, 2000, StrAppend1)
    con.Parameters.Append prm1
    Set prm2 = con.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
    con.Parameters.Append prm2
    rst3.CursorLocation = adUseClient
    rst3.CursorType = adOpenForwardOnly
    rst3.CacheSize = 100
    Set rst3 = con.Execute

    If cboReportSelection.ListIndex = 0 Then
        If rst3.EOF = True Then
            Record.SubItems(23) = ""
            rst3.Close
            Set rst3 = Nothing
        Else
            Record.SubItems(23) = Trim(rst3!ICDDescription)
            rst3.Close
            Set rst3 = Nothing
        End If
    ElseIf cboReportSelection.ListIndex = 1 Then
        If rst3.EOF = True Then
            Record.SubItems(16) = ""
            rst3.Close
            Set rst3 = Nothing
        Else
            Record.SubItems(16) = Trim(rst3!ICDDescription)
            rst3.Close
            Set rst3 = Nothing
        End If
    ElseIf cboReportSelection.ListIndex = 2 Then
    End If

End Sub

Private Sub CaseNoListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim FirstICDCode As String
Dim FinalICDCode As String
Dim FirstDescription As String
Dim FinalDescription As String
        
    ListView2.ListItems.Clear
    Current = 0
    Total = 0
              
    strAppendCase = "2"
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "ClaimReport4"
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
       
       Total = rst2.recordCount
       lblTotalRecord = Total
        
        FirstICDCode = ""
        FinalICDCode = ""
        FirstDescription = ""
        FinalDescription = ""
        
            With rst2
                    
                    If Len(rst2!CasNumber) Then
                        Set Record = ListView2.ListItems.Add(, , rst2!CasNumber)
                    Else
                        Set Record = ListView2.ListItems.Add(, , "Empty")
                    End If
                    
                    If Len(rst2!MBMNumber) Then
                       Record.SubItems(1) = rst2!MBMNumber
                    End If
                    
                    If Len(rst2!MBMPolicyNo) Then
                       Record.SubItems(2) = Trim(rst2!MBMPolicyNo)
                    End If
                                    
                    If Len(rst2!MBMPatientCovID) Then
                       Record.SubItems(3) = rst2!MBMPatientCovID
                    End If
                    
                    If Len(rst2!MBMName) Then
                       Record.SubItems(4) = Trim(rst2!MBMName)
                    End If
                    
                    If Len(rst2!PatientName) Then
                       Record.SubItems(5) = Trim(rst2!PatientName)
                    End If
                    
                    If Len(rst2!MBMSex) Then
                       Record.SubItems(6) = rst2!MBMSex
                    End If
                    
                    If Len(rst2!AGECode) Then
                       Record.SubItems(7) = rst2!AGECode
                    End If
                    
                    If Len(rst2!Sex) Then
                       Record.SubItems(8) = rst2!Sex
                    End If
                    
                    If Len(rst2!AgeBand) Then
                       Record.SubItems(9) = rst2!AgeBand
                    End If
                    
                    If Len(rst2!INSCode) Then
                       Record.SubItems(10) = rst2!INSCode
                    End If
                    
                    If Len(rst2!PLNCode) Then
                       Record.SubItems(11) = rst2!PLNCode
                    End If
                    
                    If Len(rst2!HosName) Then
                       Record.SubItems(12) = Trim(rst2!HosName)
                    End If
                    
                    If Len(rst2!FirstDiag) Then
                       Record.SubItems(13) = rst2!FirstDiag
                       FirstICDCode = Record.SubItems(13)
                       strAppendCase = "1"
                       StrAppend1 = " AND ICDCode = '" & Trim(FirstICDCode) & "'"
                       Call CASFirstICD(StrAppend1, strAppendCase, Record)
                    End If
                                                            
                    If Len(rst2!FinalDiag) Then
                       Record.SubItems(15) = rst2!FinalDiag
                       FinalICDCode = Record.SubItems(15)
                       strAppendCase = "1"
                       StrAppend1 = " AND ICDCode = '" & Trim(FinalICDCode) & "'"
                       Call CASFinalICD(StrAppend1, strAppendCase, Record)
                    End If
                    
                    If Len(rst2!CASStatus) Then
                       Record.SubItems(17) = rst2!CASStatus
                    End If
                    
                    If Len(rst2!CASClaimability) Then
                       Record.SubItems(18) = rst2!CASClaimability
                    End If
                    
                    If Len(rst2!MBMStatus) Then
                       Record.SubItems(19) = rst2!MBMStatus
                    End If
                    
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(20) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(21) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!CaseRegDate) Then
                        Record.SubItems(22) = Format(rst2!CaseRegDate, "dd/mm/yyyy")
                    End If
                    
                    If rst2!NoofDay = "0" Then
                        Record.SubItems(23) = "1"
                    ElseIf Len(rst2!NoofDay) Then
                        Record.SubItems(23) = rst2!NoofDay
                    End If
                    
                    If Len(rst2!MBMPayorEffDate) Then
                        Record.SubItems(24) = Format(rst2!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!MBMPayorEXPDate) Then
                        Record.SubItems(25) = Format(rst2!MBMPayorEXPDate, "dd/mm/yyyy")
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

Private Sub CaseWksListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
Dim FirstICDCode As String
Dim FinalICDCode As String
Dim FirstDescription As String
Dim FinalDescription As String
Dim StrAppend1 As String
        
    ListView2.ListItems.Clear
    Current = 0
    Total = 0
              
    strAppendCase = "1"
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "ClaimReport4"
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
       
       Total = rst2.recordCount
       lblTotalRecord = Total
        
        FirstICDCode = ""
        FinalICDCode = ""
        FirstDescription = ""
        FinalDescription = ""
        
            With rst2
                    
                    Set Record = ListView2.ListItems.Add(, , rst2!CasNumber)
                    
                    If Len(rst2!claimversion) Then
                       Record.SubItems(1) = rst2!claimversion
                    End If
                    
                    If Len(rst2!MBMNumber) Then
                       Record.SubItems(2) = rst2!MBMNumber
                    End If
                    
                    If Len(rst2!MBMPolicyNo) Then
                       Record.SubItems(3) = Trim(rst2!MBMPolicyNo)
                    End If
                                    
                    If Len(rst2!Relation) Then
                       Record.SubItems(4) = rst2!Relation
                    End If
                    
                    If Len(rst2!CASStatus) Then
                       Record.SubItems(5) = rst2!CASStatus
                    End If
                    
                    If Len(rst2!CASClaimability) Then
                       Record.SubItems(6) = rst2!CASClaimability
                    End If
                    
                    If Len(rst2!MBMName) Then
                       Record.SubItems(7) = Trim(rst2!MBMName)
                    End If
                    
                    If Len(rst2!PatientName) Then
                       Record.SubItems(8) = Trim(rst2!PatientName)
                    End If
                    
                    If Len(rst2!RACCode) Then
                       Record.SubItems(9) = rst2!RACCode
                    End If
                    
                    If Len(rst2!Sex) Then
                       Record.SubItems(10) = rst2!Sex
                    End If
                    
                    If Len(rst2!AgeBand) Then
                       Record.SubItems(11) = rst2!AgeBand
                    End If
                    
                    If Len(rst2!ICNo) Then
                       Record.SubItems(12) = Trim(rst2!ICNo)
                    End If
                    
                    If Len(rst2!DOB) Then
                       Record.SubItems(13) = Format(rst2!DOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!MBMPayorEffDate) Then
                        Record.SubItems(14) = Format(rst2!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!MBMPayorEXPDate) Then
                        Record.SubItems(15) = Format(rst2!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!CaseRegDate) Then
                        Record.SubItems(16) = Format(rst2!CaseRegDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(17) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(18) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!HosName) Then
                       Record.SubItems(19) = Trim(rst2!HosName)
                    End If
                    
                    If Len(rst2!FirstDiag) Then
                       Record.SubItems(20) = rst2!FirstDiag
                       FirstICDCode = Record.SubItems(20)
                       strAppendCase = "1"
                       StrAppend1 = " AND ICDCode = '" & Trim(FirstICDCode) & "'"
                       Call CASFirstICD(StrAppend1, strAppendCase, Record)
                    End If
                                                            
                    If Len(rst2!FinalDiag) Then
                       Record.SubItems(22) = rst2!FinalDiag
                       FinalICDCode = Record.SubItems(22)
                       StrAppend1 = " AND ICDCode = '" & Trim(FinalICDCode) & "'"
                       Call CASFinalICD(StrAppend1, strAppendCase, Record)
                    End If
                    
                    If Len(!CASReserveAmount) Then
                       Record.SubItems(24) = Format(!CASReserveAmount, "#,##0.00")
                    End If
                    
                    If Len(!CLMTotalActual) Then
                       Record.SubItems(25) = Format(!CLMTotalActual, "#,##0.00")
                    End If
                    
                    If Len(!CLMRecoveryPaidByM) Then
                       Record.SubItems(26) = Format(!CLMRecoveryPaidByM, "#,##0.00")
                    End If
                    
                    If Len(!AnnLimit) Then
                       Record.SubItems(27) = Format(!AnnLimit, "#,##0.00")
                    End If
                    
                    If Len(!CLMPayableByMedix2H) Then
                       Record.SubItems(28) = Format(!CLMPayableByMedix2H, "#,##0.00")
                    End If
                    
                    If Len(!CASAdmissionReason) Then
                       Record.SubItems(29) = Trim(!CASAdmissionReason)
                    End If
                    
                    If Len(!CASUDCode1) Then
                       Record.SubItems(30) = Trim(!CASUDCode1)
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

Private Sub CaseWksListing2(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim StrAppend1 As String
Dim Record As ListItem
Dim FirstICDCode As String
Dim FinalICDCode As String
Dim FirstDescription As String
Dim FinalDescription As String
Dim strCheck As Boolean
       
    ListView2.ListItems.Clear
    Current = 0
    Total = 0
    strCheck = False
              
    strAppendCase = "3"
    
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "ClaimReport4"
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
        
        Set con.ActiveConnection = medixmasterconn
        con.CommandText = "ClaimReport"
        con.CommandType = adCmdStoredProc
        con.CommandTimeout = 300
    
    Do
              
       Do Until rst2.EOF
       
       Total = rst2.recordCount
       lblTotalRecord = Total
        
        FirstICDCode = ""
        FinalICDCode = ""
        FirstDescription = ""
        FinalDescription = ""
        
            With rst2
                    
                    Set Record = ListView2.ListItems.Add(, , rst2!CasNumber)
                    
                    If Len(rst2!claimversion) Then
                       Record.SubItems(1) = rst2!claimversion
                    End If
                    
                    If Len(rst2!MBMNumber) Then
                       Record.SubItems(2) = rst2!MBMNumber
                    End If
                    
                    If Len(rst2!MBMPolicyNo) Then
                       Record.SubItems(3) = Trim(rst2!MBMPolicyNo)
                    End If
                                    
                    If Len(rst2!Relation) Then
                       Record.SubItems(4) = rst2!Relation
                    End If
                    
                    If Len(rst2!CASStatus) Then
                       Record.SubItems(5) = rst2!CASStatus
                    End If
                    
                    If Len(rst2!CASClaimability) Then
                       Record.SubItems(6) = rst2!CASClaimability
                    End If
                    
                    If Len(rst2!MBMName) Then
                       Record.SubItems(7) = Trim(rst2!MBMName)
                    End If
                    
                    If Len(rst2!PatientName) Then
                       Record.SubItems(8) = Trim(rst2!PatientName)
                    End If
                    
                    If Len(rst2!RACCode) Then
                       Record.SubItems(9) = rst2!RACCode
                    End If
                    
                    If Len(rst2!Sex) Then
                       Record.SubItems(10) = rst2!Sex
                    End If
                    
                    If Len(rst2!AgeBand) Then
                       Record.SubItems(11) = rst2!AgeBand
                    End If
                    
                    If Len(rst2!ICNo) Then
                       Record.SubItems(12) = Trim(rst2!ICNo)
                    End If
                    
                    If Len(rst2!DOB) Then
                       Record.SubItems(13) = Format(rst2!DOB, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!MBMPayorEffDate) Then
                        Record.SubItems(14) = Format(rst2!MBMPayorEffDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!MBMPayorEXPDate) Then
                        Record.SubItems(15) = Format(rst2!MBMPayorEXPDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!CaseRegDate) Then
                        Record.SubItems(16) = Format(rst2!CaseRegDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!CASDateAdmitted) Then
                        Record.SubItems(17) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!CASDischargeDate) Then
                        Record.SubItems(18) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(rst2!HosName) Then
                       Record.SubItems(19) = Trim(rst2!HosName)
                    End If
                    
                    If Len(rst2!FirstDiag) Then
                       Record.SubItems(20) = rst2!FirstDiag
                       FirstICDCode = Record.SubItems(20)
                       strAppendCase = "1"
                       StrAppend1 = Trim(FirstICDCode)
                       Set prm1 = con.CreateParameter("StrAppend1", adVarChar, adParamInput, 2000, StrAppend1)
                       con.Parameters.Append prm1
                       Set prm2 = con.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                       con.Parameters.Append prm2
                       If strCheck = False Then
                            rst3.CursorLocation = adUseClient
                            rst3.CursorType = adOpenForwardOnly
                            rst3.CacheSize = 100
                            Set rst3 = con.Execute
                        End If
                        
                       If rst3.EOF = True Then
                            Record.SubItems(21) = ""
                            strCheck = True
                       Else
                            Record.SubItems(21) = Trim(rst3!ICDDescription)
                            strCheck = True
                       End If
                    End If
                                                            
                    If Len(rst2!FinalDiag) Then
                       Record.SubItems(22) = rst2!FinalDiag
                       FinalICDCode = Record.SubItems(22)
                       StrAppend1 = Trim(FinalICDCode)
                       Set prm1 = con.CreateParameter("StrAppend1", adVarChar, adParamInput, 2000, StrAppend1)
                       con.Parameters.Append prm1
                       Set prm2 = con.CreateParameter("StrAppendCase", adVarChar, adParamInput, 2000, strAppendCase)
                       con.Parameters.Append prm2
                       
                       If rst3.EOF = True Then
                            Record.SubItems(23) = ""
                            strCheck = True
                       Else
                            Record.SubItems(23) = Trim(rst3!ICDDescription)
                            strCheck = True
                       End If
                    End If
                    
                    If Len(!CASReserveAmount) Then
                       Record.SubItems(24) = Format(!CASReserveAmount, "#,##0.00")
                    End If
                    
                    If Len(!CLMTotalActual) Then
                       Record.SubItems(25) = Format(!CLMTotalActual, "#,##0.00")
                    End If
                    
                    If Len(!CLMRecoveryPaidByM) Then
                       Record.SubItems(26) = Format(!CLMRecoveryPaidByM, "#,##0.00")
                    End If
                    
                    If Len(!AnnLimit) Then
                       Record.SubItems(27) = Format(!AnnLimit, "#,##0.00")
                    End If
                    
                    If Len(!CLMPayableByMedix2H) Then
                       Record.SubItems(28) = Format(!CLMPayableByMedix2H, "#,##0.00")
                    End If
                    
                    If Len(!CASAdmissionReason) Then
                       Record.SubItems(29) = Trim(!CASAdmissionReason)
                    End If
                    
                    If Len(!CASUDCode1) Then
                       Record.SubItems(30) = Trim(!CASUDCode1)
                    End If
                    
                    'If Len(!CASRemark) Then
                    '   Record.SubItems(31) = Trim(!CASRemark)
                    'End If
                    
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

Private Sub CaseNoListing2(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim prm1 As ADODB.Parameter
Dim prm2 As ADODB.Parameter
Dim strAppendCase As String
Dim Record As ListItem
        
    ListView2.ListItems.Clear
    Current = 0
    Total = 0
              
    strAppendCase = "4"
    
    Set cmd.ActiveConnection = medixmasterconnWS
        cmd.CommandText = "ClaimReport4"
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
       
       Total = rst2.recordCount
       lblTotalRecord = Total
        
            With rst2
                    
                    If Len(!CasNumber) Then
                        Set Record = ListView2.ListItems.Add(, , !CasNumber)
                    Else
                        Set Record = ListView2.ListItems.Add(, , "Empty")
                    End If
                    
                    If Len(!MBMNumber) Then
                       Record.SubItems(1) = !MBMNumber
                    End If
                    
                    If Len(!MBMPolicyNo) Then
                       Record.SubItems(2) = Trim(!MBMPolicyNo)
                    End If
                                    
                    If Len(!MBMDataStatus) Then
                       Record.SubItems(3) = !MBMDataStatus
                    End If
                    
                    If Len(!CASPending) Then
                       Record.SubItems(4) = !CASPending
                    End If
                    
                    If Len(!CASStatus) Then
                       Record.SubItems(5) = !CASStatus
                    End If
                    
                    If Len(!CASClaimability) Then
                       Record.SubItems(6) = !CASClaimability
                    End If
                    
                    If Len(!CASPAYType) Then
                       Record.SubItems(7) = !CASPAYType
                    End If
                    
                    If Len(!MBMStatus) Then
                       Record.SubItems(8) = !MBMStatus
                    End If
                    
                    If Len(!CASDateAdmitted) Then
                        Record.SubItems(9) = Format(!CASDateAdmitted, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CASDischargeDate) Then
                        Record.SubItems(10) = Format(!CASDischargeDate, "dd/mm/yyyy")
                    End If
                    
                    If Len(!CaseRegDate) Then
                        Record.SubItems(11) = Format(!CaseRegDate, "dd/mm/yyyy")
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

