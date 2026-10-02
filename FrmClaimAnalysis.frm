VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmClaimAnalysis 
   Caption         =   "Claim Analysis"
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
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   0
      TabIndex        =   89
      Top             =   0
      Width           =   11805
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Claim Analysis"
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
         TabIndex        =   90
         Top             =   90
         Width           =   3285
      End
   End
   Begin VB.Frame Frame10 
      Height          =   6810
      Left            =   120
      TabIndex        =   91
      Top             =   360
      Width           =   11655
      Begin VB.Frame Frame5 
         Caption         =   "Search By"
         Height          =   735
         Left            =   5280
         TabIndex        =   149
         Top             =   120
         Width           =   2415
         Begin VB.OptionButton Option2 
            Caption         =   "Case No"
            Height          =   195
            Left            =   1320
            TabIndex        =   151
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Claim No"
            Height          =   195
            Left            =   120
            TabIndex        =   150
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   1365
      End
      Begin VB.ComboBox CboRenewal 
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
         Left            =   4200
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   1365
      End
      Begin VB.CheckBox ChkDOA 
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
         Height          =   280
         Left            =   5640
         TabIndex        =   40
         Top             =   4200
         Width           =   495
      End
      Begin VB.CheckBox ChkDOD 
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
         Height          =   280
         Left            =   5640
         TabIndex        =   43
         Top             =   4560
         Width           =   495
      End
      Begin VB.CheckBox ChkFinal 
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
         Left            =   5640
         TabIndex        =   37
         Top             =   3900
         Width           =   495
      End
      Begin VB.CheckBox ChkFirst 
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
         Left            =   5640
         TabIndex        =   34
         Top             =   3600
         Width           =   495
      End
      Begin VB.CheckBox ChkRenew 
         Caption         =   "*"
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
         Left            =   5640
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox TxtAgeband2 
         Height          =   315
         Left            =   9120
         MaxLength       =   2
         TabIndex        =   49
         Top             =   960
         Width           =   1245
      End
      Begin VB.TextBox TxtAgeband1 
         Height          =   315
         Left            =   7440
         MaxLength       =   2
         TabIndex        =   48
         Top             =   960
         Width           =   1245
      End
      Begin VB.TextBox txtSearch 
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
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   1
         Top             =   240
         Width           =   3825
      End
      Begin VB.OptionButton txtMBMNo1 
         Caption         =   "Mem Num"
         Height          =   240
         Left            =   4080
         TabIndex        =   5
         Top             =   600
         Width           =   1020
      End
      Begin VB.OptionButton txtMBMPolicyNo1 
         Caption         =   "Policy Num"
         Height          =   240
         Left            =   2880
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton txtMBMName6 
         Caption         =   "Mem Name"
         Height          =   240
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton txtIcBcPc1 
         Caption         =   "IC Num"
         Height          =   240
         Left            =   1920
         TabIndex        =   3
         Top             =   600
         Width           =   825
      End
      Begin VB.CheckBox chkExpDate 
         Caption         =   "*"
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
         Left            =   10440
         TabIndex        =   57
         Top             =   1845
         Width           =   615
      End
      Begin VB.CheckBox chkDOR 
         Caption         =   "*"
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
         Left            =   10440
         TabIndex        =   60
         Top             =   2130
         Width           =   615
      End
      Begin VB.CheckBox chkBordxDate 
         Caption         =   "*"
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
         Left            =   10440
         TabIndex        =   63
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox chkBatch 
         Caption         =   "*"
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
         Left            =   10440
         TabIndex        =   92
         Top             =   2700
         Width           =   495
      End
      Begin VB.CheckBox chkEndtBordx 
         Caption         =   "*"
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
         Left            =   10440
         TabIndex        =   68
         Top             =   3000
         Width           =   495
      End
      Begin VB.CheckBox chkEndtBatch 
         Caption         =   "*"
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
         Left            =   10440
         TabIndex        =   71
         Top             =   3300
         Width           =   495
      End
      Begin VB.CheckBox chkPCode 
         Caption         =   "*"
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
         Left            =   10440
         TabIndex        =   74
         Top             =   3600
         Width           =   495
      End
      Begin VB.CheckBox chkPolYear 
         Caption         =   "*"
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
         Left            =   10440
         TabIndex        =   77
         Top             =   3900
         Width           =   615
      End
      Begin VB.CheckBox ChkEffDate 
         Caption         =   "*"
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
         Left            =   10440
         TabIndex        =   54
         Top             =   1560
         Width           =   495
      End
      Begin VB.CheckBox ChkOCP 
         Caption         =   "*"
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
         Left            =   5640
         TabIndex        =   23
         Top             =   2410
         Width           =   495
      End
      Begin VB.CheckBox ChkRace 
         Caption         =   "*"
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
         Left            =   5640
         TabIndex        =   17
         Top             =   1840
         Width           =   495
      End
      Begin VB.CheckBox ChkSex 
         Caption         =   "*"
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
         Left            =   5640
         TabIndex        =   11
         Top             =   1260
         Width           =   495
      End
      Begin VB.CheckBox ChkEndt 
         Caption         =   "*"
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
         Left            =   5640
         TabIndex        =   14
         Top             =   1560
         Width           =   495
      End
      Begin VB.CheckBox ChkPre 
         Caption         =   "*"
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
         Left            =   5640
         TabIndex        =   20
         Top             =   2130
         Width           =   495
      End
      Begin VB.CheckBox ChkMarital 
         Caption         =   "*"
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
         Left            =   5640
         TabIndex        =   29
         Top             =   3000
         Width           =   495
      End
      Begin VB.CheckBox ChkState 
         Caption         =   "*"
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
         Left            =   5640
         TabIndex        =   26
         Top             =   2700
         Width           =   495
      End
      Begin VB.ComboBox CboSex 
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
         Left            =   4200
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   1260
         Width           =   1365
      End
      Begin VB.ComboBox cboEndorseStatus 
         Height          =   315
         Left            =   4200
         TabIndex        =   13
         Top             =   1560
         Width           =   1365
      End
      Begin VB.ComboBox CboRace 
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
         Left            =   4200
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   1840
         Width           =   1365
      End
      Begin VB.ComboBox CboPreMode 
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
         Left            =   4200
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   2130
         Width           =   1365
      End
      Begin VB.ComboBox CboOCP 
         Height          =   315
         Left            =   4200
         TabIndex        =   22
         Top             =   2400
         Width           =   1365
      End
      Begin VB.ComboBox CboState 
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
         Left            =   4200
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   2700
         Width           =   1365
      End
      Begin VB.ComboBox CboMarital 
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
         Left            =   4200
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   3000
         Width           =   1365
      End
      Begin VB.ComboBox cboINSCode 
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1260
         Width           =   1365
      End
      Begin VB.ComboBox cboHLTCode 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   12
         Text            =   "S"
         Top             =   1560
         Width           =   1365
      End
      Begin VB.ComboBox CboAdultChild 
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Top             =   1840
         Width           =   1365
      End
      Begin VB.ComboBox cboPrincipalType 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2130
         Width           =   1365
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
         Left            =   1560
         TabIndex        =   21
         Text            =   "A- Active"
         Top             =   2400
         Width           =   1365
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
         Left            =   1560
         TabIndex        =   24
         Top             =   2700
         Width           =   1365
      End
      Begin VB.ComboBox CboMbmType 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3000
         Width           =   1365
      End
      Begin VB.TextBox TxtPlan1 
         Height          =   315
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   50
         Top             =   1260
         Width           =   1245
      End
      Begin MSMask.MaskEdBox TxtEffDate1 
         Height          =   315
         Left            =   7440
         TabIndex        =   52
         Top             =   1560
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtExpDate1 
         Height          =   315
         Left            =   7440
         TabIndex        =   55
         Top             =   1845
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOR1 
         Height          =   315
         Left            =   7440
         TabIndex        =   58
         Top             =   2130
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtBordx1 
         Height          =   315
         Left            =   7440
         TabIndex        =   61
         Top             =   2400
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtBatch1 
         Height          =   315
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   64
         Top             =   2700
         Width           =   1245
      End
      Begin MSMask.MaskEdBox TxtEndtBordx1 
         Height          =   315
         Left            =   7440
         TabIndex        =   66
         Top             =   3000
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtPlan2 
         Height          =   315
         Left            =   9120
         MaxLength       =   4
         TabIndex        =   51
         Top             =   1260
         Width           =   1245
      End
      Begin MSMask.MaskEdBox TxtEffDate2 
         Height          =   315
         Left            =   9120
         TabIndex        =   53
         Top             =   1560
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtExpDate2 
         Height          =   315
         Left            =   9120
         TabIndex        =   56
         Top             =   1845
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDOR2 
         Height          =   315
         Left            =   9120
         TabIndex        =   59
         Top             =   2130
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtBordx2 
         Height          =   315
         Left            =   9120
         TabIndex        =   62
         Top             =   2400
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtBatch2 
         Height          =   315
         Left            =   9120
         MaxLength       =   4
         TabIndex        =   65
         Top             =   2700
         Width           =   1245
      End
      Begin MSMask.MaskEdBox TxtEndtBordx2 
         Height          =   315
         Left            =   9120
         TabIndex        =   67
         Top             =   3000
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TxtEndtBatch1 
         Height          =   315
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   69
         Top             =   3300
         Width           =   1245
      End
      Begin VB.TextBox TxtEndtBatch2 
         Height          =   315
         Left            =   9120
         MaxLength       =   4
         TabIndex        =   70
         Top             =   3300
         Width           =   1245
      End
      Begin VB.TextBox TxtPCode1 
         Height          =   315
         Left            =   7440
         MaxLength       =   5
         TabIndex        =   72
         Top             =   3600
         Width           =   1245
      End
      Begin VB.TextBox TxtPCode2 
         Height          =   315
         Left            =   9120
         MaxLength       =   5
         TabIndex        =   73
         Top             =   3600
         Width           =   1245
      End
      Begin VB.TextBox TxtPolYear1 
         Height          =   315
         Left            =   7440
         MaxLength       =   2
         TabIndex        =   75
         Top             =   3900
         Width           =   1245
      End
      Begin VB.TextBox TxtPolYear2 
         Height          =   315
         Left            =   9120
         MaxLength       =   2
         TabIndex        =   76
         Top             =   3900
         Width           =   1245
      End
      Begin VB.TextBox TxtCaseNo 
         Height          =   315
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   30
         Top             =   3300
         Width           =   1365
      End
      Begin VB.TextBox TxtCaseNo2 
         Height          =   315
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   31
         Top             =   3300
         Width           =   1365
      End
      Begin VB.TextBox TxtFirst1 
         Height          =   315
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   32
         Top             =   3600
         Width           =   1365
      End
      Begin VB.TextBox TxtFirst2 
         Height          =   315
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   33
         Top             =   3600
         Width           =   1365
      End
      Begin VB.TextBox TxtFinal2 
         Height          =   315
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   36
         Top             =   3900
         Width           =   1365
      End
      Begin VB.TextBox TxtFinal1 
         Height          =   315
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   35
         Top             =   3900
         Width           =   1365
      End
      Begin MSMask.MaskEdBox txtDOADate1 
         Height          =   315
         Left            =   1560
         TabIndex        =   38
         Top             =   4200
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDOADate2 
         Height          =   315
         Left            =   4200
         TabIndex        =   39
         Top             =   4200
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODDate1 
         Height          =   315
         Left            =   1560
         TabIndex        =   41
         Top             =   4500
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDODDate2 
         Height          =   315
         Left            =   4200
         TabIndex        =   42
         Top             =   4500
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboHos 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   44
         Top             =   4800
         Width           =   1365
      End
      Begin VB.ComboBox CboClaimType 
         Height          =   315
         Left            =   4200
         Sorted          =   -1  'True
         TabIndex        =   45
         Top             =   4800
         Width           =   1365
      End
      Begin VB.ComboBox CboClaimability 
         Height          =   315
         Left            =   4200
         TabIndex        =   47
         Top             =   5100
         Width           =   1365
      End
      Begin VB.ComboBox CboCaseStatus 
         Height          =   315
         Left            =   1560
         TabIndex        =   46
         Top             =   5100
         Width           =   1365
      End
      Begin VB.Label Label54 
         Caption         =   "Claimability"
         Height          =   255
         Left            =   3240
         TabIndex        =   146
         Top             =   5200
         Width           =   1215
      End
      Begin VB.Label Label53 
         Caption         =   "Case Status"
         Height          =   255
         Left            =   600
         TabIndex        =   145
         Top             =   5200
         Width           =   975
      End
      Begin VB.Label Label52 
         Caption         =   "Claim Type"
         Height          =   255
         Left            =   3240
         TabIndex        =   144
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label Label51 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   600
         TabIndex        =   143
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   142
         Top             =   4605
         Width           =   375
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   141
         Top             =   4305
         Width           =   375
      End
      Begin VB.Label Label48 
         Caption         =   "DOD"
         Height          =   255
         Left            =   600
         TabIndex        =   140
         ToolTipText     =   "Date of Discharged"
         Top             =   4600
         Width           =   615
      End
      Begin VB.Label Label47 
         Caption         =   "DOA"
         Height          =   255
         Left            =   600
         TabIndex        =   139
         ToolTipText     =   "Date of Admission"
         Top             =   4300
         Width           =   1215
      End
      Begin VB.Label Label46 
         Caption         =   "Final Diag."
         Height          =   255
         Left            =   600
         TabIndex        =   138
         Top             =   4000
         Width           =   855
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   137
         Top             =   4000
         Width           =   375
      End
      Begin VB.Label Label44 
         Caption         =   "First Diag."
         Height          =   255
         Left            =   600
         TabIndex        =   136
         Top             =   3675
         Width           =   855
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   135
         Top             =   3675
         Width           =   375
      End
      Begin VB.Label Label42 
         Caption         =   "Case No"
         Height          =   255
         Left            =   600
         TabIndex        =   134
         Top             =   3400
         Width           =   735
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         Caption         =   "To"
         Height          =   255
         Left            =   3120
         TabIndex        =   133
         Top             =   3400
         Width           =   375
      End
      Begin VB.Label Label25 
         Caption         =   "Sex"
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
         Left            =   3240
         TabIndex        =   132
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label Label24 
         Caption         =   "Marital"
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
         Left            =   3240
         TabIndex        =   131
         Top             =   3120
         Width           =   930
      End
      Begin VB.Label Label22 
         Caption         =   "Renewal Ind"
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
         Left            =   3240
         TabIndex        =   130
         Top             =   990
         Width           =   930
      End
      Begin VB.Label Label21 
         Caption         =   "Policy Year"
         Height          =   255
         Left            =   6480
         TabIndex        =   129
         Top             =   4030
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   128
         Top             =   4030
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "Age Band"
         Height          =   255
         Left            =   6480
         TabIndex        =   127
         Top             =   990
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   126
         Top             =   990
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Post Code"
         Height          =   255
         Left            =   6480
         TabIndex        =   125
         Top             =   3700
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   124
         Top             =   3700
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "Batch"
         Height          =   255
         Left            =   6480
         TabIndex        =   123
         Top             =   2800
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   122
         Top             =   2800
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Bordx"
         Height          =   255
         Left            =   6480
         TabIndex        =   121
         Top             =   2530
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   120
         Top             =   2530
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   119
         Top             =   2250
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   118
         Top             =   1930
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   117
         Top             =   1620
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "DOR"
         Height          =   255
         Left            =   6480
         TabIndex        =   116
         ToolTipText     =   "Date of Discharged"
         Top             =   2250
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Exp Date"
         Height          =   255
         Left            =   6480
         TabIndex        =   115
         ToolTipText     =   "Date of Admission"
         Top             =   1930
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Eff Date"
         Height          =   255
         Left            =   6480
         TabIndex        =   114
         ToolTipText     =   "Date of Registered"
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label Label76 
         Caption         =   "Endt. Status"
         Height          =   255
         Left            =   3240
         TabIndex        =   113
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Supp. Type"
         Height          =   255
         Left            =   600
         TabIndex        =   112
         Top             =   2250
         Width           =   975
      End
      Begin VB.Label Label73 
         Caption         =   "Payor "
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
         Left            =   600
         TabIndex        =   111
         Top             =   990
         Width           =   705
      End
      Begin VB.Label Label74 
         Caption         =   "Ins Type"
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
         Left            =   600
         TabIndex        =   110
         Top             =   1305
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Mem Type"
         Height          =   255
         Left            =   600
         TabIndex        =   109
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Hlt Type"
         Height          =   255
         Left            =   600
         TabIndex        =   108
         Top             =   1650
         Width           =   975
      End
      Begin VB.Label Label67 
         Caption         =   "Search"
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
         Left            =   600
         TabIndex        =   107
         Top             =   240
         Width           =   810
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
         Left            =   600
         TabIndex        =   106
         Top             =   2865
         Width           =   1020
      End
      Begin VB.Label Label30 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   105
         Top             =   3100
         Width           =   375
      End
      Begin VB.Label Label31 
         Caption         =   "Endt. Bordx"
         Height          =   255
         Left            =   6480
         TabIndex        =   104
         Top             =   3100
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   103
         Top             =   3380
         Width           =   375
      End
      Begin VB.Label Label33 
         Caption         =   "Endt. Batch"
         Height          =   255
         Left            =   6480
         TabIndex        =   102
         Top             =   3380
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Adult/Child"
         Height          =   255
         Left            =   600
         TabIndex        =   101
         Top             =   1950
         Width           =   975
      End
      Begin VB.Label Label34 
         Caption         =   "State"
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
         Left            =   3240
         TabIndex        =   100
         Top             =   2865
         Width           =   810
      End
      Begin VB.Label Label35 
         Caption         =   "Race"
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
         Left            =   3240
         TabIndex        =   99
         Top             =   1950
         Width           =   690
      End
      Begin VB.Label Label36 
         Caption         =   "To"
         Height          =   255
         Left            =   8760
         TabIndex        =   98
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label37 
         Caption         =   "Plan Code"
         Height          =   255
         Left            =   6480
         TabIndex        =   97
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label38 
         Caption         =   "Occupation"
         Height          =   255
         Left            =   3240
         TabIndex        =   96
         Top             =   2550
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "Pre. Mode"
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
         Left            =   3240
         TabIndex        =   95
         Top             =   2250
         Width           =   810
      End
      Begin VB.Label Label39 
         Caption         =   "* - Search Empty Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   6480
         Width           =   2295
      End
      Begin VB.Label Label40 
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
         Left            =   600
         TabIndex        =   93
         Top             =   2550
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   7200
      Width           =   11655
      Begin VB.CommandButton CmdClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9720
         TabIndex        =   83
         Top             =   200
         Width           =   840
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5400
         TabIndex        =   78
         Top             =   200
         Width           =   840
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Height          =   435
         Left            =   10560
         TabIndex        =   84
         Top             =   200
         Width           =   840
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   435
         Left            =   6240
         TabIndex        =   79
         Top             =   200
         Width           =   840
      End
      Begin VB.CommandButton CmdBack 
         Caption         =   "&Back"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7080
         TabIndex        =   80
         Top             =   200
         Width           =   840
      End
      Begin VB.CommandButton CmdView 
         Caption         =   "&View"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7920
         TabIndex        =   81
         Top             =   200
         Width           =   840
      End
      Begin VB.CommandButton CmdPrintFile 
         Caption         =   "Print to &File"
         Height          =   435
         Left            =   8760
         TabIndex        =   82
         Top             =   200
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView11 
         Height          =   375
         Left            =   3600
         TabIndex        =   148
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ICD"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label27 
         Caption         =   "Record(s)"
         Height          =   255
         Left            =   2640
         TabIndex        =   88
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label28 
         Caption         =   "of"
         Height          =   255
         Left            =   1320
         TabIndex        =   87
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   240
         TabIndex        =   86
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1560
         TabIndex        =   85
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lstMBMInquiryList 
      Height          =   6720
      Left            =   0
      TabIndex        =   147
      Top             =   480
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   11853
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
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
         Text            =   "Claim Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Member Number"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Member Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Covered ID"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Hospital"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "First Diag."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Description"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Final Diag."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Description"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "DOA"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "DOD"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "Claim Type"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Case Status"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "Claimability"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "Payable to Hosp."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Text            =   "Payable to MBM"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Total Approved"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Text            =   "Total Actual"
         Object.Width           =   2117
      EndProperty
   End
End
Attribute VB_Name = "FrmClaimAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intExit As Integer
Public FirstICD As String
Public FinalICD As String
Public FirstICDDescription As String
Public FinalICDDescription As String
Private m_blnDirection(1 To 18) As Boolean

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

Private Sub Form_Load()
    Frame10.Visible = True
    lstMBMInquiryList.Visible = False
    intExit = 0
    Call ComboListing
    Call LoadInsuredType
    Call LoadPayor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    intExit = 1
End Sub

'===============Start Command Button ===========================

Private Sub CmdClear_Click()
    Call ClearForm(Me)
    lstMBMInquiryList.listitems.Clear
    txtMBMName6.value = True
    lblCurrent = ""
    lblTotal = ""
End Sub

Private Sub CmdBack_Click()
    Frame10.Visible = True
    lstMBMInquiryList.Visible = False
End Sub

Private Sub CmdView_Click()
    Frame10.Visible = False
    lstMBMInquiryList.Visible = True
End Sub

Private Sub CmdPrintFile_Click()
    Call ExportToExcelClaimAnalysis(lstMBMInquiryList)
End Sub


Private Sub cmdSearch_Click()
    cmdStop.Enabled = True
    CmdClear.Enabled = False
    lblCurrent = ""
    lblTotal = ""
                
    If Len(txtSearch) >= 3 Then
        CmdClear.Enabled = False
        cmdExit.Enabled = False
        Frame10.Visible = False
        lstMBMInquiryList.Visible = True
        Screen.MousePointer = vbHourglass
        Call SearchRecord
        Screen.MousePointer = Default
    Else
        MsgBox "Search String must be at least 3 characters!", vbInformation
        Frame10.Visible = True
        lstMBMInquiryList.Visible = False
        txtSearch.SetFocus
        cmdStop.Enabled = False
    End If
End Sub

Private Sub CmdExit_Click()
'Dim RecordExit
'    RecordExit = MsgBox("Quit this form?", vbYesNo + vbExclamation)
'    If RecordExit = vbYes Then
'        bExit = True
        Unload Me
'    End If
End Sub

Private Sub CmdStop_Click()
    intExit = 1
    cmdExit.Enabled = True
    CmdClear.Enabled = True
End Sub
'=============== End Command Button ===========================

'=============== Start ComboListing ===========================

Private Sub ComboListing()
Dim EndtStatus As New ADODB.Recordset
Dim PolStatus As New ADODB.Recordset
Dim State As New ADODB.Recordset
Dim OCP As New ADODB.Recordset
Dim PAYTYPE As New ADODB.Recordset
Dim HOS As New ADODB.Recordset
Dim sqlEndtStatus, sqlPolStatus, sqlState, sqlOCP As String

    
    sqlEndtStatus = " select statuscode ,statusDesc from Status"
    EndtStatus.Open sqlEndtStatus, medixmasterconn, adOpenKeyset, adLockOptimistic
        
    Do Until EndtStatus.EOF
        cboEndorseStatus.AddItem ChkStr(EndtStatus!StatusCode) & " - " & ChkStr(EndtStatus!statusDesc)
        EndtStatus.MoveNext
    Loop
    
    EndtStatus.Close
    Set EndtStatus = Nothing
    
    sqlPolStatus = " select pstatuscode ,pstatusDesc from PolicyStatus"
    PolStatus.Open sqlPolStatus, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until PolStatus.EOF
        CboStatus.AddItem (PolStatus!pStatusCode) & " - " & (PolStatus!pstatusDesc)
        PolStatus.MoveNext
    Loop
    
    PolStatus.Close
    Set PolStatus = Nothing
        
    sqlState = " Select STADescription from STA"
    State.Open sqlState, medixmasterconn, adOpenKeyset, adLockOptimistic
        
    Do Until State.EOF
        CboState.AddItem ChkStr(State!STADescription)
        State.MoveNext
    Loop
    
    State.Close
    Set State = Nothing
    
    sqlOCP = " Select OCPCode, OCPDescription from OCP"
    OCP.Open sqlOCP, medixmasterconn, adOpenKeyset, adLockOptimistic
        
    Do Until OCP.EOF
        CboOCP.AddItem ChkStr(OCP!OCPCode) & " - " & ChkStr(OCP!OCPDescription)
        OCP.MoveNext
    Loop
    
    OCP.Close
    Set OCP = Nothing
    
    PAYTYPE.Open "Select PayTypeCode from PAYTYPE", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until PAYTYPE.EOF
        If Len(PAYTYPE!PayTypeCode) Then
            CboClaimType.AddItem PAYTYPE!PayTypeCode
        End If
        PAYTYPE.MoveNext
    Loop
    PAYTYPE.Close
    Set PAYTYPE = Nothing
    
    HOS.Open "Select HOSName from HOS", medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until HOS.EOF
        CboHos.AddItem Trim(HOS!HOSName)
        HOS.MoveNext
    Loop
    HOS.Close
    Set HOS = Nothing
    
    CboMbmType.AddItem "Both"
    CboMbmType.AddItem "P" & " - " & "Principal"
    CboMbmType.AddItem "S" & " - " & "Supplementary"
    CboMbmType.Text = "Both"
    
    cboPrincipalType.AddItem "Both"
    cboPrincipalType.AddItem "C" & " - " & "Combined Limit"
    cboPrincipalType.AddItem "S" & " - " & "Separated Limit"
    cboPrincipalType.Text = "C" & " - " & "Combined Limit"
    
    cboHLTCode.AddItem "S"
    cboHLTCode.AddItem "N"
    
    CboSex.AddItem "M" & " - " & "Male"
    CboSex.AddItem "F" & " - " & "Female"
    
    CboPreMode.AddItem "A - Annually"
    CboPreMode.AddItem "Q - Quarterly"
    CboPreMode.AddItem "S - Semi Annual"
    CboPreMode.AddItem "M - Monthly"
    
    CboRenewal.AddItem "N - No"
    CboRenewal.AddItem "Y - Yes"
    
    cboDataStatus.AddItem "A - Active"
    cboDataStatus.AddItem "D - Delete"
    
    CboMarital.AddItem "S - Single"
    CboMarital.AddItem "M - Married"
    
    CboRace.AddItem "C - Chinese"
    CboRace.AddItem "I - Indian"
    CboRace.AddItem "M - Malay"
    CboRace.AddItem "O - Other"
        
    CboAdultChild.AddItem "A - Adult"
    CboAdultChild.AddItem "C - Child"
    
    With CboCaseStatus
        .AddItem "O - Open"
        .AddItem "C - Close"
        .AddItem "D - Delete"
    End With
    
    With CboClaimability
        .AddItem "A - Accepted"
        .AddItem "R - Rejected"
    End With
    
End Sub
'=============== End ComboListing ===========================

'=============== Start Load Payor  ===========================

Private Sub LoadPayor(Optional HLTCode As String)
    Dim PAY As New ADODB.Recordset
    If Mid(ChkStr(HLTCode), 1, 1) = "N" Then
        sqlPAY = "select distinct PAYCode , PAYCompanyName from View_PAY_PLN "
        sqlPAY = sqlPAY + " Where HLTCODE = 'N'"
    Else
        sqlPAY = "select PAYCode , PAYCompanyName from PAY "
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
'=============== End Load Payor ===================================


'=============== Start Load Insured Type ===========================

Private Sub LoadInsuredType()
Dim strsql As String
Dim HLTCode As String
Dim INS As New ADODB.Recordset
    
    cboINSCode.Clear
    strsql = "Select INSCode, INSDescription from INS "
    INS.Open strsql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    Do Until INS.EOF
        cboINSCode.AddItem ChkStr(INS!INSCode) & " - " & ChkStr(INS!INSDescription)
        INS.MoveNext
    Loop
    INS.Close
    Set INS = Nothing
End Sub
'=============== End Load Insured Type ===========================

'=============== Start Sorting ListView ===========================

Private Sub lstMBMInquiryList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
m_blnDirection(ColumnHeader.Index) = Not m_blnDirection(ColumnHeader.Index)
    Select Case ColumnHeader.Index
    Case 1
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(1)
    Case 2
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(2)
    Case 3
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(3)
    Case 4
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtNumber, m_blnDirection(4)
    Case 5
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(5)
    Case 6
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(6)
    Case 7
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(7)
    Case 8
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(8)
    Case 9
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(9)
    Case 10
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtDateTime, m_blnDirection(10)
    Case 11
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtDateTime, m_blnDirection(11)
    Case 12
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(12)
    Case 13
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(13)
    Case 14
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtString, m_blnDirection(14)
    Case 15
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtNumber, m_blnDirection(15)
    Case 16
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtNumber, m_blnDirection(16)
    Case 17
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtNumber, m_blnDirection(17)
    Case 18
        SortListView lstMBMInquiryList, ColumnHeader.Index, ldtNumber, m_blnDirection(18)
    End Select
End Sub
'=============== End Sorting ListView ===========================

'---------------- Start Check Date Format -------------------------

Private Sub TxtEffDate1_LostFocus()
    If IsDate(TxtEffDate1) Then
    Else
        If TxtEffDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtEffDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtEffDate2_LostFocus()
    If IsDate(TxtEffDate2) Then
    Else
        If TxtEffDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtEffDate2.SetFocus
        End If
    End If
End Sub

Private Sub TxtExpDate1_LostFocus()
    If IsDate(TxtExpDate1) Then
    Else
        If TxtExpDate1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtExpDate1.SetFocus
        End If
    End If
End Sub

Private Sub TxtExpDate2_LostFocus()
    If IsDate(TxtExpDate2) Then
    Else
        If TxtExpDate2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtExpDate2.SetFocus
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

Private Sub TxtBordx1_LostFocus()
    If IsDate(TxtBordx1) Then
    Else
        If TxtBordx1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtBordx1.SetFocus
        End If
    End If
End Sub

Private Sub TxtBordx2_LostFocus()
    If IsDate(TxtBordx2) Then
    Else
        If TxtBordx2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtBordx2.SetFocus
        End If
    End If
End Sub

Private Sub TxtEndtBordx1_LostFocus()
    If IsDate(TxtEndtBordx1) Then
    Else
        If TxtEndtBordx1 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtEndtBordx1.SetFocus
        End If
    End If
End Sub

Private Sub TxtEndtBordx2_LostFocus()
    If IsDate(TxtEndtBordx2) Then
    Else
        If TxtEndtBordx2 <> "__/__/____" Then
            MsgBox "Invalid Date Format! ", vbCritical
            TxtEndtBordx2.SetFocus
        End If
    End If
End Sub
'----------------End Check Date Format -------------------------

'---------------- Start ComboBox Change/ KeyPress -------------------------

Private Sub cboINSCode_Change()
    If InStr(cboINSCode.Text, "null") Then
       cboINSCode.Enabled = False
    End If
End Sub

Private Sub cboRace_Change()
    If InStr(CboRace.Text, "null") Then
       CboRace.Enabled = False
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

Private Sub cboPayorCode_Change()
    If InStr(cboPayorCode.Text, "null") Then
       cboPayorCode.Enabled = False
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

Private Sub cboINSCode_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboINSCode.Text, cboINSCode.SelStart) + Chr(KeyAscii) + Mid(cboINSCode.Text, cboINSCode.SelStart + cboINSCode.SelLength + 1))
 
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

Private Sub cboRace_KeyPress(KeyAscii As Integer)
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
Private Sub cboPrincipalType_Change()
    If InStr(cboPrincipalType.Text, "null") Then
       cboPrincipalType.Enabled = False
    End If
End Sub

Private Sub cboPrincipalType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    'NewText = LCase(Left(cboPrincipalType.Text, cboPrincipalType.SelStart) + Chr(KeyAscii) + Mid(cboPrincipalType.Text, cboPrincipalType.SelStart + cboPrincipalType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboPrincipalType.ListCount - 1
 
        If NewText = LCase(Left(cboPrincipalType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboPrincipalType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboPrincipalType.Text = ValidValue
                cboPrincipalType.SelStart = 0
                cboPrincipalType.SelLength = Len(ValidValue)
            End If
        End If
End Sub


Private Sub cboHLTCode_KeyPress(KeyAscii As Integer)
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

Private Sub cboStatus_Change()
    If InStr(CboStatus.Text, "null") Then
       CboStatus.Enabled = False
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

Private Sub cboDataStatus_Change()
    If InStr(cboDataStatus.Text, "null") Then
       cboDataStatus.Enabled = False
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

Private Sub CboMbmType_Change()
    If InStr(CboMbmType.Text, "null") Then
       CboMbmType.Enabled = False
    End If
End Sub

Private Sub CboMbmType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    'NewText = LCase(Left(CboMbmType.Text, CboMbmType.SelStart) + Chr(KeyAscii) + Mid(CboMbmType.Text, CboMbmType.SelStart + CboMbmType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboMbmType.ListCount - 1
 
        If NewText = LCase(Left(CboMbmType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboMbmType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboMbmType.Text = ValidValue
                CboMbmType.SelStart = 0
                CboMbmType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboRenewal_Change()
    If InStr(CboRenewal.Text, "null") Then
       CboRenewal.Enabled = False
    End If
End Sub

Private Sub CboRenewal_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboRenewal.Text, CboRenewal.SelStart) + Chr(KeyAscii) + Mid(CboRenewal.Text, CboRenewal.SelStart + CboRenewal.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboRenewal.ListCount - 1
 
        If NewText = LCase(Left(CboRenewal.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboRenewal.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboRenewal.Text = ValidValue
                CboRenewal.SelStart = 0
                CboRenewal.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboPreMode_Change()
    If InStr(CboPreMode.Text, "null") Then
       CboPreMode.Enabled = False
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

Private Sub CboSex_Change()
    If InStr(CboSex.Text, "null") Then
       CboSex.Enabled = False
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

Private Sub CboMarital_Change()
    If InStr(CboMarital.Text, "null") Then
       CboMarital.Enabled = False
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

Private Sub CboOCP_Change()
    If InStr(CboOCP.Text, "null") Then
       CboOCP.Enabled = False
    End If
End Sub

Private Sub CboOCP_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboOCP.Text, CboOCP.SelStart) + Chr(KeyAscii) + Mid(CboOCP.Text, CboOCP.SelStart + CboOCP.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboOCP.ListCount - 1
 
        If NewText = LCase(Left(CboOCP.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboOCP.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboOCP.Text = ValidValue
                CboOCP.SelStart = 0
                CboOCP.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboEndorseStatus_Change()
    If InStr(cboEndorseStatus.Text, "null") Then
       cboEndorseStatus.Enabled = False
    End If
End Sub

Private Sub cboEndorseStatus_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(cboEndorseStatus.Text, cboEndorseStatus.SelStart) + Chr(KeyAscii) + Mid(cboEndorseStatus.Text, cboEndorseStatus.SelStart + cboEndorseStatus.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To cboEndorseStatus.ListCount - 1
 
        If NewText = LCase(Left(cboEndorseStatus.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = cboEndorseStatus.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                cboEndorseStatus.Text = ValidValue
                cboEndorseStatus.SelStart = 0
                cboEndorseStatus.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub CboAdultChild_Change()
    If InStr(CboAdultChild.Text, "null") Then
       CboAdultChild.Enabled = False
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

Private Sub CboCaseStatus_Change()
    If InStr(CboCaseStatus.Text, "null") Then
       CboCaseStatus.Enabled = False
    End If
End Sub

Private Sub cboClaimability_Change()
    If InStr(CboClaimability.Text, "null") Then
       CboClaimability.Enabled = False
    End If
End Sub

Private Sub cboClaimability_KeyPress(KeyAscii As Integer)
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

Private Sub CboState_Change()
    If InStr(CboState.Text, "null") Then
       CboState.Enabled = False
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

Private Sub CboClaimType_Change()
    If InStr(CboClaimType.Text, "null") Then
       CboClaimType.Enabled = False
    End If
End Sub

Private Sub CboClaimType_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboClaimType.Text, CboClaimType.SelStart) + Chr(KeyAscii) + Mid(CboClaimType.Text, CboClaimType.SelStart + CboClaimType.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboClaimType.ListCount - 1
 
        If NewText = LCase(Left(CboClaimType.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboClaimType.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboClaimType.Text = ValidValue
                CboClaimType.SelStart = 0
                CboClaimType.SelLength = Len(ValidValue)
            End If
        End If
End Sub

Private Sub cboHOS_Change()
    If InStr(CboHos.Text, "null") Then
       CboHos.Enabled = False
    End If
End Sub

Private Sub cboHOS_KeyPress(KeyAscii As Integer)
Dim NewText As String
Dim ValidCount As Integer
Dim ValidValue As String
Dim i As Integer
    If KeyAscii >= 32 And KeyAscii <> 127 Then
    NewText = LCase(Left(CboHos.Text, CboHos.SelStart) + Chr(KeyAscii) + Mid(CboHos.Text, CboHos.SelStart + CboHos.SelLength + 1))
 
    ValidCount = 0
    ValidValue = ""
 
    For i = 0 To CboHos.ListCount - 1
 
        If NewText = LCase(Left(CboHos.List(i), Len(NewText))) Then
                ValidCount = ValidCount + 1
                ValidValue = CboHos.List(i)
        End If
        Next
        If ValidCount <= 1 Then KeyAscii = 0
            If ValidCount = 1 Then
                CboHos.Text = ValidValue
                CboHos.SelStart = 0
                CboHos.SelLength = Len(ValidValue)
            End If
        End If
End Sub
'---------------- End ComboBox Change/ KeyPress -------------------------

Private Function SearchRecord()
Dim strsql As String
Dim strAppend As String
Dim sql As String
Dim CAS As New ADODB.Recordset
    
    strAppend = ""
    
        
    If Len(Trim(TxtCaseNo)) Then
       strAppend = strAppend + " AND CASNumber >= '" & RemStr(Trim(TxtCaseNo)) & "'"
    End If
    
    If Len(Trim(TxtCaseNo2)) Then
       strAppend = strAppend + " AND CASNumber <= '" & RemStr(Trim(TxtCaseNo2)) & "'"
    End If
                
    If Mid(CboCaseStatus, 1, 1) = "C" Then
        strAppend = strAppend + " And CASStatus = 'C'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "D" Then
        strAppend = strAppend + " And CASStatus = 'D'"
    ElseIf Mid(CboCaseStatus, 1, 1) = "O" Then
        strAppend = strAppend + " And CASStatus = 'O'"
    'Else
        'strAppend = strAppend + " And CASStatus = 'C' OR CASStatus = 'O'"
        
    End If
    If CboCaseStatus = "" Then
        If Mid(CboClaimability, 1, 1) = "R" Then
            strAppend = strAppend + " Or CASClaimability = 'R'"
        End If
        
        If Mid(CboClaimability, 1, 1) = "A" Then
            strAppend = strAppend + " Or CASClaimability = 'A'"
        End If
    Else
        If Mid(CboClaimability, 1, 1) = "R" Then
            strAppend = strAppend + " And CASClaimability = 'R'"
        End If
        
        If Mid(CboClaimability, 1, 1) = "A" Then
            strAppend = strAppend + " And CASClaimability = 'A'"
        End If
    End If
    
    If ChkDOA.value = 1 Then
        sql = ""
        sql = "AND CasDateAdmitted Is null " & ""
        strAppend = strAppend + sql
    End If
    
    
    If ChkDOD.value = 1 Then
        sql = ""
        sql = "AND CasDischargeDate Is null " & ""
        strAppend = strAppend + sql
    End If
    
    If (txtDOADate1) <> "__/__/____" And IsDate(txtDOADate1) = True _
        And (txtDOADate2) <> "__/__/____" And IsDate(txtDOADate2) = True Then
        sql = ""
        sql = " AND CasDateAdmitted >= '" & Format(Trim(txtDOADate1), "mm/dd/yyyy") & _
              "' And CasDateAdmitted <= '" & Format(Trim(txtDOADate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDOADate1)) > 0 And IsDate(txtDOADate1) = True Then
        strAppend = strAppend + " And CasDateAdmitted >= '" & Format(txtDOADate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDOADate2)) > 0 And IsDate(txtDOADate2) = True Then
        strAppend = strAppend + " And CasDateAdmitted <= '" & Format(txtDOADate2, "mm/dd/yyyy") & "'"
    End If
    
    If (txtDODDate1) <> "__/__/____" And IsDate(txtDODDate1) = True _
        And (txtDODDate2) <> "__/__/____" And IsDate(txtDODDate2) = True Then
        sql = ""
        sql = " AND CasDischargeDate >= '" & Format(Trim(txtDODDate1), "mm/dd/yyyy") & _
              "' And CasDischargeDate <= '" & Format(Trim(txtDODDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(txtDODDate1)) > 0 And IsDate(txtDODDate1) = True Then
        strAppend = strAppend + " And CasDischargeDate >= '" & Format(txtDODDate1, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(txtDODDate2)) > 0 And IsDate(txtDODDate2) = True Then
        strAppend = strAppend + " And CasDischargeDate <= '" & Format(txtDODDate2, "mm/dd/yyyy") & "'"
    End If
    
    If Len(Trim(CboClaimType)) Then
        strAppend = strAppend + " And PayTypeCode = '" & Trim(Mid(CboClaimType, 1, IIf(InStr(1, CboClaimType, "-") = 0, 4, InStr(1, CboClaimType, "-") - 1))) & "'"
    End If
    
    If Len(Trim(TxtFirst2)) Then
        strAppend = strAppend + " AND FirstDiag <= '" & Trim(TxtFirst2) & "'"
    End If
    
    If ChkFirst.value = 1 Then
        sql = ""
        sql = "AND FirstDiag Is null OR FirstDiag = '' " & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(TxtFinal1)) Then
        strAppend = strAppend + " AND FinalDiag >= '" & Trim(TxtFinal1) & "'"
    End If
    
    If Len(Trim(TxtFinal2)) Then
        strAppend = strAppend + " AND FinalDiag <= '" & Trim(TxtFinal2) & "'"
    End If
    
    If ChkFinal.value = 1 Then
        sql = ""
        sql = "AND FinalDiag Is null OR FinalDiag = '' " & ""
        strAppend = strAppend + sql
    End If
    
    If Len(Trim(CboHos)) Then
        strAppend = strAppend + " And HOSName = '" & Trim(CboHos) & "'"
    End If
     
    If Len(Trim(CboClaimType)) Then
        strAppend = strAppend + " And PayTypeCode = '" & Trim(Mid(CboClaimType, 1, IIf(InStr(1, CboClaimType, "-") = 0, 4, InStr(1, CboClaimType, "-") - 1))) & "'"
    End If
     
     
     '--------------------------------------------------------------------------------------------------------------------------------------------
      If Len(Trim(cboHLTCode)) Then
          strAppend = strAppend & "AND HLTCode= '" & Trim(cboHLTCode) & "' "
      End If
      
      If Len(Trim(CboState)) Then
          strAppend = strAppend & "AND STACode= '" & Trim(CboState) & "' "
      End If
      
      If ChkState.value = 1 Then
        sql = ""
        sql = "AND STACode Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If Len(Trim(cboDataStatus)) Then
          strAppend = strAppend & "AND MBMDataStatus = '" & Trim(Mid(cboDataStatus, 1, InStr(1, cboDataStatus, "-") - 1)) & "' "
      End If
      
      If Len(Trim(Mid(CboSex, 1, 1))) Then
          strAppend = strAppend & "AND Sex = '" & Trim(Mid(CboSex, 1, 1)) & "' "
      End If
      
      If ChkSex.value = 1 Then
        sql = ""
        sql = "AND Sex Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If Len(Trim(Mid(CboOCP, 1, 2))) Then
          strAppend = strAppend & "AND Occupation= '" & Trim(Mid(CboOCP, 1, 2)) & "' "
      End If
      
      If ChkOCP.value = 1 Then
        sql = ""
        sql = "AND Occupation Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If Len(Trim(Mid(CboAdultChild, 1, 1))) Then
          strAppend = strAppend & "AND AdultChild= '" & Trim(Mid(CboAdultChild, 1, 1)) & "' "
      End If
      
      If Len(Trim(Mid(CboRace, 1, 1))) Then
          strAppend = strAppend & "AND RACCode = '" & Trim(Mid(CboRace, 1, 1)) & "' "
      End If
      
      If ChkRace.value = 1 Then
        sql = ""
        sql = "AND RACCode Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If Len(Trim(TxtPCode1)) Then
        strAppend = strAppend + " AND MBMPostCode >= '" & Trim(TxtPCode1) & "'"
      End If
    
      If Len(Trim(TxtPCode2)) Then
        strAppend = strAppend + " AND MBMPostCode <= '" & Trim(TxtPCode2) & "'"
      End If
      
      If chkPCode.value = 1 Then
        sql = ""
        sql = "AND MBMPostCode Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If Len(Trim(TxtAgeband1)) Then
        strAppend = strAppend + " AND AGECode >= '" & Trim(TxtAgeband1) & "'"
      End If
    
      If Len(Trim(TxtAgeband2)) Then
        strAppend = strAppend + " AND AGECode <= '" & Trim(TxtAgeband2) & "'"
      End If
      
            
      If Len(Trim(TxtPolYear1)) Then
        strAppend = strAppend + " AND MBMPolicyYear >= '" & Trim(TxtPolYear1) & "'"
      End If
    
      If Len(Trim(TxtPolYear2)) Then
        strAppend = strAppend + " AND MBMPolicyYear <= '" & Trim(TxtPolYear2) & "'"
      End If
      
      If chkPolYear.value = 1 Then
        sql = ""
        sql = "AND MBMPolicyYear Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If Len(Trim(TxtPlan1)) Then
        strAppend = strAppend + " AND PLNCode >= '" & Trim(TxtPlan1) & "'"
      End If
    
      If Len(Trim(TxtPlan2)) Then
        strAppend = strAppend + " AND PLNCode <= '" & Trim(TxtPlan2) & "'"
      End If
      
                     
     '--------------------------------------------------------------------------------------------------------------------------------------------
      If Len(Mid(cboPayorCode, 1, 2)) Then
         'If InStr(1, cboPayorCode, "-") <> 0 Then
           'strAppend = strAppend & " AND PAYCode = '" & Trim(Mid(cboPayorCode, 1, InStr(1, cboPayorCode, "-") - 1)) & "' "
         'Else
             strAppend = strAppend & "AND PAYCode = '" & Mid(cboPayorCode, 1, 2) & "' "
         'End If
      End If
      
            
      If (TxtEffDate1) <> "__/__/____" And IsDate(TxtEffDate1) = True _
        And (TxtEffDate2) <> "__/__/____" And IsDate(TxtEffDate2) = True Then
        sql = ""
        sql = " AND MBMPayorEffDate >= '" & Format(Trim(TxtEffDate1), "mm/dd/yyyy") & _
              "' And MBMPayorEffDate <= '" & Format(Trim(TxtEffDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtEffDate1)) > 0 And IsDate(TxtEffDate1) = True Then
        strAppend = strAppend + " And MBMPayorEffDate >= '" & Format(TxtEffDate1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtEffDate2)) > 0 And IsDate(TxtEffDate2) = True Then
        strAppend = strAppend + " And MBMPayorEffDate <= '" & Format(TxtEffDate2, "mm/dd/yyyy") & "'"
      End If
      
      If ChkEffDate.value = 1 Then
        sql = ""
        sql = "AND MBMPayorEffDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If (TxtExpDate1) <> "__/__/____" And IsDate(TxtExpDate1) = True _
        And (TxtExpDate2) <> "__/__/____" And IsDate(TxtExpDate2) = True Then
        sql = ""
        sql = " AND MBMPayorExpDate >= '" & Format(Trim(TxtExpDate1), "mm/dd/yyyy") & _
              "' And MBMPayorExpDate <= '" & Format(Trim(TxtExpDate2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtExpDate1)) > 0 And IsDate(TxtExpDate1) = True Then
        strAppend = strAppend + " And MBMPayorExpDate >= '" & Format(TxtExpDate1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtExpDate2)) > 0 And IsDate(TxtExpDate2) = True Then
        strAppend = strAppend + " And MBMPayorExpDate <= '" & Format(TxtExpDate2, "mm/dd/yyyy") & "'"
      End If
      
      If chkExpDate.value = 1 Then
        sql = ""
        sql = "AND MBMPayorExpDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If (TxtDOR1) <> "__/__/____" And IsDate(TxtDOR1) = True _
        And (TxtDOR2) <> "__/__/____" And IsDate(TxtDOR2) = True Then
        sql = ""
        sql = " AND MBMValueDate >= '" & Format(Trim(TxtDOR1), "mm/dd/yyyy") & _
              "' And MBMValueDate <= '" & Format(Trim(TxtDOR2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtDOR1)) > 0 And IsDate(TxtDOR1) = True Then
        strAppend = strAppend + " And MBMValueDate >= '" & Format(TxtDOR1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtDOR2)) > 0 And IsDate(TxtDOR2) = True Then
        strAppend = strAppend + " And MBMValueDate <= '" & Format(TxtDOR2, "mm/dd/yyyy") & "'"
      End If
      
      If chkDOR.value = 1 Then
        sql = ""
        sql = "AND MBMValueDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
    '--------------------------------------------------------------------------------------------------------------------------------------------
      If Len(Trim(cboINSCode)) Then
           strAppend = strAppend & " AND INSCode = '" & Trim(Mid(cboINSCode, 1, IIf(InStr(1, cboINSCode, "-") = 0, 1, InStr(1, cboINSCode, "-") - 1))) & "' "
      End If
        
      If Len(Trim(CboPreMode)) Then
           strAppend = strAppend & " AND MBMInstallment = '" & Trim(Mid(CboPreMode, 1, IIf(InStr(1, CboPreMode, "-") = 0, 1, InStr(1, CboPreMode, "-") - 1))) & "' "
      End If
      
      If ChkPre.value = 1 Then
        sql = ""
        sql = "AND MBMInstallment Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If Len(Trim(Mid(CboMarital, 1, 1))) Then
           strAppend = strAppend & " AND MBMMaritalStatus = '" & Trim(Mid(CboMarital, 1, 1)) & "' "
      End If
      
      If ChkMarital.value = 1 Then
        sql = ""
        sql = "AND MBMMaritalStatus Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If Len(Trim(CboRenewal)) Then
           strAppend = strAppend & " AND MBMRenewFlag = '" & Trim(Mid(CboRenewal, 1, IIf(InStr(1, CboRenewal, "-") = 0, 1, InStr(1, CboRenewal, "-") - 1))) & "' "
      End If
      
      If ChkRenew.value = 1 Then
        sql = ""
        sql = "AND MBMRenewFlag Is null " & ""
        strAppend = strAppend + sql
      End If
     
     '--------------------------------------------------------------------------------------------------------------------------------------------
      If Len(Trim(CboStatus)) Then
          strAppend = strAppend & "AND MBMStatus = '" & Trim(Mid(CboStatus, 1, InStr(1, CboStatus, "-") - 1)) & "' "
      End If
     
      If Len(Trim(txtSearch)) Then
         If txtMBMNo1.value = True Then
             strAppend = strAppend & " AND  MBMNumber LIKE  '%" & RemStr(txtSearch) & "%' "
         ElseIf txtMBMPolicyNo1.value = True Then
             strAppend = strAppend & " AND MBMPolicyNo like '%" & RemStr(txtSearch) & "%' "
         ElseIf txtMBMName6.value = True Then
             strAppend = strAppend & " AND Name LIKE '%" & RemStr(txtSearch) & "%' "
         ElseIf txtIcBcPc1.value = True Then
             strAppend = strAppend & " AND  ICNo like  '%" & RemStr(txtSearch) & "%' "
         End If
      End If
     
     '--------------------------------------------------------------------------------------------------------------------------------------------
      'If Len(Trim(cboEndorseStatus)) Then
             'strAppend = strAppend & " AND MBMAmendEdtType = '" & Trim(Mid(cboEndorseStatus, 1, IIf(InStr(1, cboEndorseStatus, "-") = 0, 1, InStr(1, cboEndorseStatus, "-") - 1))) & "' "
      'End If
      
      'If ChkEndt.value = 1 Then
        'Sql = ""
        'Sql = "AND MBMAmendEdtType Is null " & ""
        'strAppend = strAppend + Sql
      'End If
     
     '--------------------------------------------------------------------------------------------------------------------------------------------
      'If Len(Trim(cboPrincipalType)) And InStr(1, cboPrincipalType, "-") <> 0 Then
           'If Trim(Mid(cboPrincipalType, 1, InStr(1, cboPrincipalType, "-") - 1)) = "C" Then
               'strAppend = strAppend & " AND (MBMMemberType = 'P' OR MBMMemberType = 'C')"
           'ElseIf Trim(Mid(cboPrincipalType, 1, InStr(1, cboPrincipalType, "-") - 1)) = "S" Then
               'strAppend = strAppend & " AND (MBMMemberType = 'Q')"
           'End If
      'End If
      
      If Len(Trim(cboPrincipalType)) And InStr(1, cboPrincipalType, "-") <> 0 Then
           If Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(Mid(CboMbmType, 1, 1)) = "P" Then
               strAppend = strAppend & " AND (MemberType = 'P')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(Mid(CboMbmType, 1, 1)) = "S" Then
               strAppend = strAppend & " AND (MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "C" And Trim(CboMbmType) = "Both" Then
               strAppend = strAppend & " AND (MemberType = 'P' or MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(Mid(CboMbmType, 1, 1)) = "P" Then
               strAppend = strAppend & " AND (MemberType = 'Q')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(Mid(CboMbmType, 1, 1)) = "S" Then
               strAppend = strAppend & " AND (MemberType = 'C')"
           ElseIf Trim(Mid(cboPrincipalType, 1, 1)) = "S" And Trim(CboMbmType) = "Both" Then
               strAppend = strAppend & " AND (MemberType = 'Q' or MemberType = 'C')"
           End If
      End If

'==========Search Bordx Date & Batch No====================================================

     'tempsqlstrBrodx = ""
     'tempsqlstrEndt = ""

      If (TxtBordx1) <> "__/__/____" And IsDate(TxtBordx1) = True _
        And (TxtBordx2) <> "__/__/____" And IsDate(TxtBordx2) = True Then
        sql = ""
        sql = " AND MBMBordxDate >= '" & Format(Trim(TxtBordx1), "mm/dd/yyyy") & _
              "' And MBMBordxDate <= '" & Format(Trim(TxtBordx2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtBordx1)) > 0 And IsDate(TxtBordx1) = True Then
        strAppend = strAppend + " And MBMBordxDate >= '" & Format(TxtBordx1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtBordx2)) > 0 And IsDate(TxtBordx2) = True Then
        strAppend = strAppend + " And MBMBordxDate <= '" & Format(TxtBordx2, "mm/dd/yyyy") & "'"
      End If
      
      If chkBordxDate.value = 1 Then
        sql = ""
        sql = "AND MBMBordxDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If Len(Trim(TxtBatch1)) Then
        strAppend = strAppend + " AND MBMBatchNo >= '" & Trim(TxtBatch1) & "'"
      End If
    
      If Len(Trim(TxtBatch2)) Then
        strAppend = strAppend + " AND MBMBatchNo <= '" & Trim(TxtBatch2) & "'"
      End If
      
      If chkBatch.value = 1 Then
        sql = ""
        sql = "AND MBMBatchNo Is null " & ""
        strAppend = strAppend + sql
      End If

'========== End Search Bordx Date & Batch No====================================================
      
      
'=====Search Endt. Bordx Date & Endt Batch No ==================================================================
        
     If (TxtEndtBordx1) <> "__/__/____" And IsDate(TxtEndtBordx1) = True _
        And (TxtEndtBordx2) <> "__/__/____" And IsDate(TxtEndtBordx2) = True Then
        sql = ""
        sql = " AND MBMAEndtBordxDate >= '" & Format(Trim(TxtEndtBordx1), "mm/dd/yyyy") & _
              "' And MBMAEndtBordxDate <= '" & Format(Trim(TxtEndtBordx2), "mm/dd/yyyy") & "'"
        strAppend = strAppend + sql
      End If
    
      If Len(Trim(TxtEndtBordx1)) > 0 And IsDate(TxtEndtBordx1) = True Then
        strAppend = strAppend + " And MBMAEndtBordxDate >= '" & Format(TxtEndtBordx1, "mm/dd/yyyy") & "'"
      End If
    
      If Len(Trim(TxtEndtBordx2)) > 0 And IsDate(TxtEndtBordx2) = True Then
        strAppend = strAppend + " And MBMAEndtBordxDate <= '" & Format(TxtEndtBordx2, "mm/dd/yyyy") & "'"
      End If
      
      If chkEndtBordx.value = 1 Then
        sql = ""
        sql = "AND MBMAEndtBordxDate Is null " & ""
        strAppend = strAppend + sql
      End If
      
      If Len(Trim(TxtEndtBatch1)) Then
        strAppend = strAppend + " AND MBMAEndtBatchNo >= '" & Trim(TxtBatch1) & "'"
      End If
    
      If Len(Trim(TxtEndtBatch2)) Then
        strAppend = strAppend + " AND MBMAEndtBatchNo <= '" & Trim(TxtBatch2) & "'"
      End If
      
      If chkEndtBatch.value = 1 Then
        sql = ""
        sql = "AND MBMAEndtBatchNo Is null " & ""
        strAppend = strAppend + sql
      End If
    
    CaseCriteria = strAppend
    Call ClaimAnalysisListing(strAppend)
   
End Function

Private Sub ClaimAnalysisListing(Optional strAppend As String)
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
Dim sql2 As String
Dim ClaimAdmin As Boolean
Dim Total As String
Dim Current As String

    Total = 0
    Current = 0
    
    lstMBMInquiryList.listitems.Clear

    If Option1.value = True Then
    
    sql = "Select CASNumber, ClaimVersion, MBMNumber, Name, CASStatus, MBMPremium, " & _
          " CASDateAdmitted, CasDateAdmitted, PayTypeCode, PLNCode, " & _
          " ICNo, Sex, MBMMCOFee, HLTCode, CASDischargeDate, HOSName, " & _
          " Relation, BordxDate, BatchNo, MBMDataStatus, FirstDiag, " & _
          " FinalDiag, MBMPolicyNo, MBMPostCode, MBMENDtEffDate, " & _
          " PAYCode, CASClaimability, MBMPayorEffDate, MBMPAyorExpDate, " & _
          " MBMInstallment, AGEBand, AdultChild, MBMRenewFlag, STACode, " & _
          " MBMStatus, MBMValueDate, MBMENDtBordxDate, MBMLastEndorseStatus, " & _
          " MBMPolicyYear, RACCode, Sex, MBMMaritalStatus, Occupation, " & _
          " MemberType, MBMPatientCovID, CLMPayableByMedix2H, CLMPayableByMedix2M, " & _
          " CLMTotalApproved, CLMTOtalActual " & _
          " From VIEW_ClaimAnalysis where (CASNumber is not null and ClaimVersion is not null)"
              
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend
    End If
              
    ElseIf Option2.value = True Then
    
    sql = " SELECT * From VIEW_ClaimAnalysis2 WHERE 0 = 0 "
    
    sql2 = " AND CASNumber Is Not Null "
    
        
    If Len(Trim(strAppend)) Then
        sql = sql + strAppend + sql2
    End If
    
    End If
    
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
   
    If rst2.recordCount = 0 Then
        MsgBox "No record found!", vbOKOnly + vbCritical, DialogBoxTitle
    Else
         
    Do
        Do Until rst2.EOF
        Total = rst2.recordCount
        lblTotal = Total
        
            With rst2
                                    
            If Option1.value = True Then
                Set Record = lstMBMInquiryList.listitems.Add(, , rst2!CASNumber & "-" & rst2!claimversion)
                                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(1) = rst2!MBMNumber
                End If
                               
                If Len(rst2!Name) Then
                    Record.SubItems(2) = rst2!Name
                End If
                
                If Len(rst2!MBMPatientCovID) Then
                    Record.SubItems(3) = rst2!MBMPatientCovID
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(4) = rst2!HOSName
                End If
                
                If Len(rst2!FirstDiag) Then
                    Record.SubItems(5) = rst2!FirstDiag
                    FirstICD = Record.SubItems(5)
                    Call SearchFirstICD
                End If
                
                Record.SubItems(6) = FirstICDDescription
                
                If Len(rst2!FinalDiag) Then
                    Record.SubItems(7) = rst2!FinalDiag
                    FinalICD = Record.SubItems(7)
                    Call SearchFinalICD
                End If
                
                Record.SubItems(8) = FinalICDDescription
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(9) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(10) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!PayTypeCode) Then
                    Record.SubItems(11) = rst2!PayTypeCode
                End If
                
                If Len(rst2!CASStatus) Then
                    Record.SubItems(12) = rst2!CASStatus
                End If
                
                If Len(rst2!CASClaimability) Then
                    Record.SubItems(13) = rst2!CASClaimability
                End If
                
                If Len(rst2!CLMPayableByMedix2H) Then
                    Record.SubItems(14) = Format(rst2!CLMPayableByMedix2H, "#,##0.00;(#,##0.00)")
                End If
                
                If Len(rst2!CLMPayableByMedix2M) Then
                    Record.SubItems(15) = Format(rst2!CLMPayableByMedix2M, "#,##0.00;(#,##0.00)")
                End If
                
                If Len(rst2!CLMTotalApproved) Then
                    Record.SubItems(16) = Format(rst2!CLMTotalApproved, "##,#0.00")
                End If
                
                If Len(rst2!CLMTotalActual) Then
                    Record.SubItems(17) = Format(rst2!CLMTotalActual, "##,#0.00")
                End If
                
                Current = Current + 1
                lblCurrent = Current
                
                FirstICD = ""
                FirstICDDescription = ""
                FinalICD = ""
                FinalICDDescription = ""
                
            ElseIf Option2.value = True Then
                Set Record = lstMBMInquiryList.listitems.Add(, , rst2!CASNumber)
                                
                If Len(rst2!MBMNumber) Then
                    Record.SubItems(1) = rst2!MBMNumber
                End If
                               
                If Len(rst2!Name) Then
                    Record.SubItems(2) = rst2!Name
                End If
                
                If Len(rst2!MBMPatientCovID) Then
                    Record.SubItems(3) = rst2!MBMPatientCovID
                End If
                
                If Len(rst2!HOSName) Then
                    Record.SubItems(4) = rst2!HOSName
                End If
                
                If Len(rst2!FirstDiag) Then
                    Record.SubItems(5) = rst2!FirstDiag
                    FirstICD = Record.SubItems(5)
                    Call SearchFirstICD
                End If
                
                Record.SubItems(6) = FirstICDDescription
                
                If Len(rst2!FinalDiag) Then
                    Record.SubItems(7) = rst2!FinalDiag
                    FinalICD = Record.SubItems(7)
                    Call SearchFinalICD
                End If
                
                Record.SubItems(8) = FinalICDDescription
                
                If Len(rst2!CASDateAdmitted) Then
                    Record.SubItems(9) = Format(rst2!CASDateAdmitted, "dd/mm/yyyy")
                End If
                
                If Len(rst2!CASDischargeDate) Then
                    Record.SubItems(10) = Format(rst2!CASDischargeDate, "dd/mm/yyyy")
                End If
                
                If Len(rst2!PayTypeCode) Then
                    Record.SubItems(11) = rst2!PayTypeCode
                End If
                
                If Len(rst2!CASStatus) Then
                    Record.SubItems(12) = rst2!CASStatus
                End If
                
                If Len(rst2!CASClaimability) Then
                    Record.SubItems(13) = rst2!CASClaimability
                End If
                
                If Len(rst2!PayableByMedix2H) Then
                    Record.SubItems(14) = Format(rst2!PayableByMedix2H, "#,##0.00;(#,##0.00)")
                End If
                
                If Len(rst2!PayableByMedix2M) Then
                    Record.SubItems(15) = Format(rst2!PayableByMedix2M, "#,##0.00;(#,##0.00)")
                End If
                
                If Len(rst2!TotalApproved) Then
                    Record.SubItems(16) = Format(rst2!TotalApproved, "##,#0.00")
                End If
                
                If Len(rst2!TotalActual) Then
                    Record.SubItems(17) = Format(rst2!TotalActual, "##,#0.00")
                End If
                
                Current = Current + 1
                lblCurrent = Current
                
                FirstICD = ""
                FirstICDDescription = ""
                FinalICD = ""
                FinalICDDescription = ""
            End If
            
            End With
            rst2.MoveNext
            DoEvents
                If intExit = 1 Then
                   Check = False
                   Exit Do
                End If
        Loop
    Loop Until Check = False
    End If
    intExit = 0
    rst2.Close
    Set rst2 = Nothing
    CmdClear.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub SearchFirstICD()
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
        
        
    ListView11.listitems.Clear
                
    sql = " Select ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(FirstICD) & "'"
    
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
    If rst2.recordCount = 0 Then
        'MsgBox Not "No record found for this ICD Code : " & Trim(FirstICD), vbOKOnly + vbCritical, DialogBoxTitle
    Else
            With rst2

                Set Record = ListView11.listitems.Add(, , rst2!ICDCode)
                
                If Len(rst2!ICDDescription) Then
                    Record.SubItems(1) = Trim(rst2!ICDDescription)
                    FirstICDDescription = Record.SubItems(1)
                End If
        
            End With
            rst2.MoveNext
    End If
    rst2.Close
    Set rst2 = Nothing
End Sub

Private Sub SearchFinalICD()
Dim rst2 As New ADODB.Recordset
Dim Record As ListItem
Dim sql As String
        
        
    ListView11.listitems.Clear
                
    sql = " Select ICDCode, ICDDescription From ICD Where ICDCode = '" & Trim(FinalICD) & "'"
    
            
    rst2.Open sql, medixmasterconn, adOpenKeyset, adLockOptimistic
    
     If rst2.recordCount = 0 Then
        'MsgBox Not "No record found for this ICD Code : " & Trim(FinalICD), vbOKOnly + vbCritical, DialogBoxTitle
     Else
            With rst2
                Set Record = ListView11.listitems.Add(, , rst2!ICDCode)
                
                If Len(rst2!ICDDescription) Then
                    Record.SubItems(1) = Trim(rst2!ICDDescription)
                    FinalICDDescription = Record.SubItems(1)
                End If
                
                
            End With
            rst2.MoveNext
    End If
    rst2.Close
    Set rst2 = Nothing
End Sub
